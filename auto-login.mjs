import { mkdir, readFile } from 'node:fs/promises';
import { dirname } from 'node:path';
import { createInterface } from 'node:readline/promises';
import { openTeams, waitForChatList, PROFILE_DIR, AUTH_PATH } from './teams.mjs';
import { acquireCommandLock, connectableDaemon } from './daemon.mjs';

// Usage:
//   nix develop .#playwright --command node auto-login.mjs
//
// Same as manual-login.mjs, but fills the account, password and MFA method
// automatically. The email and password are read from a git-ignored ".env"
// file next to this script:
//
//   TEAMS_EMAIL=you@example.com
//   TEAMS_PASSWORD=your-password
//
// The MFA one-time code cannot be known ahead of time, so once the code is
// sent you are asked to type it into the console.
//
// Runs headless using the persistent profile. Because a persistent profile
// drops session cookies when reopened (and not all tenants offer "Stay signed
// in?"), the auth session is also captured to a storageState file.
//
// The daemon has to be stopped first: two browsers on one profile directory
// write over each other's stored session, so a login captured next to a running
// daemon is not reliably the session that survives.

// Reads the git-ignored ".env" file into a plain object. Kept deliberately
// small — only the two keys this script needs — rather than pulling in a
// dependency. Lines that are blank or start with "#" are ignored; a leading
// "export " and surrounding quotes are stripped so the file also works when
// sourced by a shell.
async function readEnvFile(path) {
  let raw;
  try {
    raw = await readFile(path, 'utf8');
  } catch (err) {
    if (err.code === 'ENOENT') {
      throw new Error(
        `Credentials file "${path}" not found. Create it with:\n` +
          '  TEAMS_EMAIL=you@example.com\n' +
          '  TEAMS_PASSWORD=your-password',
      );
    }
    throw err;
  }

  const env = {};
  for (const line of raw.split('\n')) {
    const trimmed = line.trim();
    if (!trimmed || trimmed.startsWith('#')) continue;
    const eq = trimmed.indexOf('=');
    if (eq === -1) continue;
    const key = trimmed.slice(0, eq).replace(/^export\s+/, '').trim();
    let value = trimmed.slice(eq + 1).trim();
    if (
      (value.startsWith('"') && value.endsWith('"')) ||
      (value.startsWith("'") && value.endsWith("'"))
    ) {
      value = value.slice(1, -1);
    }
    env[key] = value;
  }
  return env;
}

const env = await readEnvFile('.env');
const email = env.TEAMS_EMAIL;
const password = env.TEAMS_PASSWORD;
if (!email || !password) {
  console.error('".env" must define both TEAMS_EMAIL and TEAMS_PASSWORD.');
  process.exit(1);
}

// Held for the whole run. The record alone would miss the case this most needs
// to catch: a daemon is written down only once it is usable, so for the minute
// or more it takes to boot, its browser already owns the profile while nothing
// says so. A command that is starting one holds this lock across the spawn,
// which is what makes the lock — not the record — the thing to wait on.
const releaseLock = await acquireCommandLock();

// The record still has to be consulted for a daemon that finished starting and
// is now sitting idle: it holds the profile without holding the lock. Asked via
// connectableDaemon() so that a record left over from before a reboot, whose pid
// has since been reused, does not refuse a login and point at a --stop that
// would signal an unrelated process.
const daemon = await connectableDaemon();
if (daemon) {
  console.log(`The Teams daemon (pid ${daemon.pid}) is holding the browser profile "${PROFILE_DIR}".`);
  console.log('Stop it first: node teams-daemon.mjs --stop');
  await releaseLock();
  process.exit(1);
}

// Neither check covers a teams-daemon.mjs someone started by hand in a third
// terminal; Chromium's own profile lock is the only backstop there.

await mkdir(dirname(AUTH_PATH), { recursive: true });

// No session to restore — this is the script that creates one. It must be this
// process's own browser, too: the daemon's is already signed in, which is the
// opposite of what a login is for.
const { context, page } = await openTeams({ headless: true, restoreAuth: false, daemon: false });

// Asks for the MFA code on the console. Kept until Verify is clicked so a
// mistyped code can be retried without restarting the whole login.
async function promptForCode() {
  const rl = createInterface({ input: process.stdin, output: process.stdout });
  try {
    let code = '';
    while (!code) {
      code = (await rl.question('Enter the MFA code sent to your phone: ')).trim();
    }
    return code;
  } finally {
    rl.close();
  }
}

try {
  console.log(`Signing in as ${email}...`);

  // The account picker lists remembered accounts by email; a fresh profile
  // shows an email textbox instead. Handle both so the script does not depend
  // on the profile having been used before.
  const accountTile = page.locator(`[data-test-id="${email}"]`);
  const emailInput = page.getByRole('textbox', { name: /email|someone@example/i });
  await Promise.race([
    accountTile.waitFor({ state: 'visible', timeout: 60000 }),
    emailInput.waitFor({ state: 'visible', timeout: 60000 }),
  ]);
  if (await accountTile.isVisible()) {
    await accountTile.click();
  } else {
    await emailInput.fill(email);
    await page.getByRole('button', { name: 'Next' }).click();
  }

  await page.getByRole('textbox', { name: 'Enter the password for' }).fill(password);
  await page.getByRole('button', { name: 'Sign in' }).click();

  // Choose SMS as the verification method. The phone button's name carries a
  // masked number that differs per account, so match on the "Text" prefix.
  await page.getByRole('link', { name: 'Sign in another way' }).click();
  await page.getByRole('button', { name: /^Text/ }).click();

  console.log('An MFA code has been sent to your phone.');
  const code = await promptForCode();
  await page.getByRole('textbox', { name: 'Enter code' }).fill(code);
  await page.getByRole('button', { name: 'Verify' }).click();

  // Entra ID may interpose a "Stay signed in?" prompt between MFA and the SPA.
  // It blocks the redirect to Teams, and a headless run has nobody to click it,
  // so answer it here. Tenants that do not show it simply have nothing to click.
  await page.getByRole('button', { name: 'Yes' }).click({ timeout: 15000 }).catch(() => {});

  // Everything from here on is machine work — the human input is already done —
  // so the budget is far shorter than manual-login.mjs's, which has to cover
  // someone typing at the keyboard. Failing fast beats hanging with no output.
  console.log('Waiting for the chat list to appear (up to 2 minutes)...');
  await waitForChatList(page, { timeout: 120000 });

  // Capture the session once (this briefly opens a page per origin to read
  // localStorage — expected, and it only happens this one time).
  await context.storageState({ path: AUTH_PATH });
  console.log(`Login captured — cookies saved to "${AUTH_PATH}", profile at "${PROFILE_DIR}".`);

  // Close the browser ourselves now that the session is saved. Closing the
  // context flushes the persistent profile to disk, so there is nothing left
  // for the user to do by hand.
  await context.close();
  console.log('Browser closed.');
} catch (err) {
  await context.close();
  throw err;
} finally {
  await releaseLock();
}
