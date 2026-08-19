import { mkdir } from 'node:fs/promises';
import { dirname } from 'node:path';
import { openTeams, waitForChatList, PROFILE_DIR, AUTH_PATH } from './teams.mjs';
import { acquireCommandLock, connectableDaemon } from './daemon.mjs';

// Usage:
//   nix develop .#playwright --command node manual-login.mjs
//
// Opens Teams in a visible browser using the persistent profile. Because a
// persistent profile drops session cookies when reopened (and not all tenants
// offer "Stay signed in?"), the auth session is also captured to a storageState
// file.
//
// Log in manually (including MFA). Once the chat list appears the session is
// saved automatically; you can then close the browser window.
//
// The daemon has to be stopped first: two browsers on one profile directory
// write over each other's stored session, so a login captured next to a running
// daemon is not reliably the session that survives.

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
const { context, page } = await openTeams({ headless: false, restoreAuth: false, daemon: false });

try {
  console.log('Log in manually in the browser window that just opened.');
  console.log('Waiting for you to finish logging in (up to 10 minutes)...');
  await waitForChatList(page, { timeout: 600000 });

  // Capture the session once (this briefly opens a page per origin to read
  // localStorage — expected, and it only happens this one time).
  await context.storageState({ path: AUTH_PATH });
  console.log(`Login captured — cookies saved to "${AUTH_PATH}", profile at "${PROFILE_DIR}".`);
  console.log('You can close the browser window now.');

  // Wait for the window to close so the persistent profile is flushed too.
  await new Promise((resolve) => context.on('close', resolve));
} catch (err) {
  await context.close();
  throw err;
} finally {
  await releaseLock();
}
