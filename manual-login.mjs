import { mkdir } from 'node:fs/promises';
import { dirname } from 'node:path';
import { openTeams, waitForChatList, PROFILE_DIR, AUTH_PATH } from './teams.mjs';
import { readInfo, isAlive } from './daemon.mjs';

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

// Refused here rather than left to the browser, which does not reliably refuse
// it — and, when it does, says nothing about the daemon.
const daemon = await readInfo();
if (daemon?.pid && isAlive(daemon.pid)) {
  console.log(`The Teams daemon (pid ${daemon.pid}) is holding the browser profile "${PROFILE_DIR}".`);
  console.log('Stop it first: node teams-daemon.mjs --stop');
  process.exit(1);
}

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
}
