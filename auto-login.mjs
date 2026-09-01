import { readFile } from 'node:fs/promises';
import { dirname, join } from 'node:path';
import { createInterface } from 'node:readline/promises';
import { fileURLToPath } from 'node:url';
import { beginLogin, waitForChatList, waitForEmojiCatalog, PROFILE_DIR, AUTH_PATH } from './teams.mjs';

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
// sent you are asked to type it into the console. Text-message MFA is the only
// method driven here; an account with no phone method registered has to use
// manual-login.mjs.
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
// sourced by a shell. Values are trimmed, so a value with leading or trailing
// spaces has to be quoted — which .env.example says.
//
// A missing file is reported as null rather than thrown, so that the caller can
// present it the same way as the other configuration mistakes: a setup step the
// user has not done yet is not a bug, and does not want a stack trace.
async function readEnvFile(path) {
  let raw;
  try {
    raw = await readFile(path, 'utf8');
  } catch (err) {
    if (err.code === 'ENOENT') return null;
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

// Resolved against this script rather than the working directory, so the file
// really is the one "next to this script" that the comment above and the README
// both promise — running the script by absolute path from elsewhere finds it
// just the same.
const ENV_PATH = join(dirname(fileURLToPath(import.meta.url)), '.env');

const env = await readEnvFile(ENV_PATH);
if (!env) {
  console.error(
    `Credentials file "${ENV_PATH}" not found. Create it with:\n` +
      '  TEAMS_EMAIL=you@example.com\n' +
      '  TEAMS_PASSWORD=your-password',
  );
  process.exit(1);
}

const email = env.TEAMS_EMAIL;
const password = env.TEAMS_PASSWORD;
if (!email || !password) {
  console.error(`"${ENV_PATH}" must define both TEAMS_EMAIL and TEAMS_PASSWORD.`);
  process.exit(1);
}

// Headless, since every step of this login is filled in by the script.
const { context, page, releaseLock } = await beginLogin({ headless: true });

// Waits for whichever of the given locators becomes visible first and returns
// its key. Every wait is subscribed to, so the ones that lose the race do not
// reject unobserved once their own timeout expires.
async function firstVisible(locators, timeout) {
  const names = Object.keys(locators);
  try {
    return await Promise.any(
      names.map((name) => locators[name].waitFor({ state: 'visible', timeout }).then(() => name)),
    );
  } catch {
    throw new Error(`Timed out after ${timeout} ms — none of these appeared: ${names.join(', ')}.`);
  }
}

// How many times a rejected MFA code may be re-entered before giving up. The
// alternative to re-prompting is a whole new login and a fresh text message, so
// a mistyped digit should not cost that — but the loop must not run forever
// either, since "still on the code page" is also what a stuck page looks like.
const MAX_CODE_ATTEMPTS = 3;
// How long a submitted code gets to be judged before "no verdict" is given up
// on. Generous, because the page has to make a round trip to the identity
// provider before it can say anything.
const CODE_VERDICT_TIMEOUT_MS = 60_000;
// How long each look at the page waits for the code box to go away before
// checking the alerts instead. Doubles as the interval between those looks.
const CODE_POLL_MS = 500;
// How long the code box gets to go away after an alert has been raised. An
// alert is not proof of rejection — the MFA pages announce things that are not
// errors — so a code the page went on to accept still gets to win.
const CODE_ACCEPT_GRACE_MS = 5_000;

// Whether the locator went away within the budget. A detached element counts as
// hidden, which is what a page navigating away from the code entry produces.
async function becameHidden(locator, timeout) {
  return locator
    .waitFor({ state: 'hidden', timeout })
    .then(() => true)
    .catch(() => false);
}

// What the page's alert regions currently say, flattened into one string that
// can be compared verbatim. Whitespace is normalised so that a re-render which
// only reflows the same message does not read as a new one, and empty regions
// are dropped, since a live region that is merely present is not a message.
async function alertSummary(alerts) {
  try {
    const texts = await alerts.allInnerTexts();
    return texts.map((text) => text.replace(/\s+/g, ' ').trim()).filter(Boolean).join(' | ');
  } catch {
    return '';
  }
}

// Waits for the page to say whether the code that was just submitted is good.
//
// An accepted code navigates away from the code-entry page; a rejected one
// leaves the box in place and raises an error beside it. The catch is that
// nothing is guaranteed to clear that error when the next code is typed, so
// from the second attempt onwards "an alert is visible" says nothing about the
// code just submitted — only that an earlier one was wrong. What counts is
// therefore not the presence of a message but a message that is *new*: one that
// is non-empty and reads differently from the last one seen, starting with
// `staleAlerts`, the snapshot taken before Verify was clicked. Comparing the
// text, rather than waiting for the old banner to disappear first, keeps the
// retry working on pages that never take the banner down at all.
//
// A message going away only moves that baseline along, it is never a verdict —
// a page that does clear the banner while the next code is typed would
// otherwise look like it had just rejected it. Moving the baseline is also what
// lets the same message be recognised a second time, which is exactly what
// mistyping twice in a row produces.
//
// The code box gets the last word either way: it is checked first on every
// look, and once more after a new message appears, so an announcement that was
// not an error cannot fail a code the page went on to accept. No verdict inside
// the budget is treated as a rejection, which costs a re-prompt at worst.
async function codeAccepted(codeInput, alerts, staleAlerts) {
  const deadline = Date.now() + CODE_VERDICT_TIMEOUT_MS;
  let lastAlerts = staleAlerts;
  while (Date.now() < deadline) {
    if (await becameHidden(codeInput, CODE_POLL_MS)) return true;
    const currentAlerts = await alertSummary(alerts);
    if (currentAlerts && currentAlerts !== lastAlerts) {
      return await becameHidden(codeInput, CODE_ACCEPT_GRACE_MS);
    }
    lastAlerts = currentAlerts;
  }
  return false;
}

// Asks for the MFA code on the console, submits it, and re-prompts while the
// code-entry page is still showing. The readline interface is deliberately held
// open across the Verify click, which is what makes the retry possible.
async function submitCode(page, codeInput) {
  const rl = createInterface({ input: process.stdin, output: process.stdout });
  const alerts = page.getByRole('alert');
  try {
    for (let attempt = 1; attempt <= MAX_CODE_ATTEMPTS; attempt++) {
      let code = '';
      while (!code) {
        code = (await rl.question('Enter the MFA code sent to your phone: ')).trim();
      }

      // A verification that finished after its budget ran out lands while the
      // next code is being typed, and filling a box that is on its way out
      // would fail a login that actually succeeded.
      if (attempt > 1 && (await becameHidden(codeInput, CODE_POLL_MS))) return;

      // Snapshotted before the click, so that a banner left over from the
      // previous attempt cannot be read as this attempt's verdict.
      const staleAlerts = await alertSummary(alerts);
      await codeInput.fill(code);
      await page.getByRole('button', { name: 'Verify' }).click();

      if (await codeAccepted(codeInput, alerts, staleAlerts)) return;

      console.log('That code was not accepted. Check the message and try again.');
    }

    // The guard above only rescues a wrong verdict at the top of a *later*
    // attempt, and the last one has none after it. A verdict errs towards
    // rejection by design, so give the code box the last word here too rather
    // than failing a login the page went on to accept.
    if (await becameHidden(codeInput, CODE_POLL_MS)) return;

    throw new Error(`The MFA code was not accepted after ${MAX_CODE_ATTEMPTS} attempts.`);
  } finally {
    rl.close();
  }
}

try {
  console.log(`Signing in as ${email}...`);

  // Three ways the profile can land here. The account picker lists remembered
  // accounts by email and a fresh profile shows an email textbox instead — but
  // a profile whose tokens are still good is waved straight through to the chat
  // list, since only the session cookies are dropped when it is reopened. That
  // third landing is not a failure, it is the state this run is trying to reach,
  // so it goes to the capture rather than waiting out a sign-in form that is
  // never going to appear.
  const accountTile = page.locator(`[data-test-id="${email}"]`);
  const emailInput = page.getByRole('textbox', { name: /email|someone@example/i });
  const chatList = page.getByRole('treeitem').first();
  const landing = await firstVisible(
    { 'account tile': accountTile, 'email box': emailInput, 'chat list': chatList },
    60000,
  );

  if (landing === 'chat list') {
    console.log('The profile is still signed in — no credentials or MFA needed.');
  } else {
    if (landing === 'account tile') {
      await accountTile.click();
    } else {
      await emailInput.fill(email);
      await page.getByRole('button', { name: 'Next' }).click();
    }

    await page.getByRole('textbox', { name: 'Enter the password for' }).fill(password);
    await page.getByRole('button', { name: 'Sign in' }).click();

    // A rejected password never reaches the MFA step: the page stays where it is
    // with an error banner. Race that banner against the two shapes the MFA step
    // can take, so a wrong value in ".env" is reported as a wrong value instead
    // of surfacing later as a timeout on a locator unrelated to the cause.
    const passwordError = page.locator('#passwordError');
    const anotherWay = page.getByRole('link', { name: 'Sign in another way' });
    const codeInput = page.getByRole('textbox', { name: 'Enter code' });
    const step = await firstVisible(
      { 'password error': passwordError, 'method choice': anotherWay, 'code entry': codeInput },
      60000,
    );
    if (step === 'password error') {
      throw new Error(`Sign-in was rejected: ${(await passwordError.innerText()).trim()}`);
    }

    // Choose SMS as the verification method. The link is only offered when the
    // account has more than one method registered — with text as the only or
    // default method, Entra ID goes straight to the code-entry page. The phone
    // button's name carries a masked number that differs per account, so match
    // on the "Text" prefix.
    if (step === 'method choice') {
      await anotherWay.click();
      await page.getByRole('button', { name: /^Text/ }).click();
    }

    console.log('An MFA code has been sent to your phone.');
    await submitCode(page, codeInput);

    // Entra ID may interpose a "Stay signed in?" prompt between MFA and the SPA.
    // It blocks the redirect to Teams, and a headless run has nobody to click
    // it, so answer it here. Tenants that do not show it have nothing to click.
    await page.getByRole('button', { name: 'Yes' }).click({ timeout: 15000 }).catch(() => {});
  }

  // Everything from here on is machine work — the human input is already done —
  // so the budget is far shorter than manual-login.mjs's, which has to cover
  // someone typing at the keyboard. Failing fast beats hanging with no output.
  console.log('Waiting for the chat list to appear (up to 2 minutes)...');
  await waitForChatList(page, { timeout: 120000 });

  // Capture the session once (this briefly opens a page per origin to read
  // localStorage — expected, and it only happens this one time).
  await context.storageState({ path: AUTH_PATH });
  console.log(`Login captured — cookies saved to "${AUTH_PATH}", profile at "${PROFILE_DIR}".`);

  // The chat list is not the client having finished loading, and this is the
  // one script that closes the browser itself the moment it appears — the
  // daemon stays up for its idle timeout and manual-login.mjs waits to be
  // closed by hand, so neither of those cuts a background sync short. Closing
  // here in the seconds the emoji catalog is being written is what leaves the
  // half-filled catalog the reaction commands then have to repair, so the sync
  // is given its chance before the browser goes.
  //
  // After the capture rather than before it: the session is what this run is
  // for and is already on disk by now, so nothing about the waiting below can
  // cost it.
  await waitForEmojiCatalog(page);

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
