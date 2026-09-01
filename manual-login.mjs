import { beginLogin, waitForChatList, waitForEmojiCatalog, PROFILE_DIR, AUTH_PATH } from './teams.mjs';

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

// Visible, because the whole point is that a person drives it.
const { context, page, releaseLock } = await beginLogin({ headless: false });

try {
  console.log('Log in manually in the browser window that just opened.');
  console.log('Waiting for you to finish logging in (up to 10 minutes)...');
  await waitForChatList(page, { timeout: 600000 });

  // Capture the session once (this briefly opens a page per origin to read
  // localStorage — expected, and it only happens this one time).
  await context.storageState({ path: AUTH_PATH });
  console.log(`Login captured — cookies saved to "${AUTH_PATH}", profile at "${PROFILE_DIR}".`);

  // Waited for before the window is offered up, not after: Teams is still
  // writing its emoji catalog for a few seconds after the chat list appears,
  // and a window closed in that gap leaves the catalog half-filled. Nobody
  // reading "you can close it now" should have to know that.
  await waitForEmojiCatalog(page);

  // Both of the steps below are for a window that is still open, and after a
  // wait of up to a minute that is no longer a given: waitForEmojiCatalog()
  // returns quietly when the user closes the window while it is running, which
  // on the one script whose ordinary ending is the user closing the window is
  // an ordinary way for it to end. Saying "you can close it now" at a window
  // that is already gone reads as the script not having noticed, and the wait
  // below would never end — the context's close event fired while the window
  // went, and a listener attached afterwards waits for an event that has
  // already been and gone, so the finally that releases the command lock is
  // never reached.
  if (!page.isClosed()) {
    console.log('You can close the browser window now.');

    // Wait for the window to close so the persistent profile is flushed too.
    await new Promise((resolve) => context.on('close', resolve));
  }
} catch (err) {
  await context.close();
  throw err;
} finally {
  await releaseLock();
}
