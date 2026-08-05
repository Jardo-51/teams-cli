import { mkdir } from 'node:fs/promises';
import { dirname } from 'node:path';
import { openTeams, waitForChatList, PROFILE_DIR, AUTH_PATH } from './teams.mjs';

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

await mkdir(dirname(AUTH_PATH), { recursive: true });

// No session to restore — this is the script that creates one.
const { context, page } = await openTeams({ headless: false, restoreAuth: false });

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
