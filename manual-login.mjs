import { chromium } from '@playwright/test';
import { mkdir } from 'node:fs/promises';
import { dirname } from 'node:path';

// Usage:
//   nix develop .#playwright --command node manual-login.mjs
//
// Opens Teams in a visible browser using a persistent profile at
// $TEAMS_PROFILE (default ".profile") for localStorage/cache. Because a
// persistent profile drops session cookies when reopened (and this tenant does
// not offer "Stay signed in?"), the auth session is also captured to a
// storageState file at $TEAMS_AUTH (default ".auth/user.json").
//
// Log in manually (including MFA). Once the chat list appears the session is
// saved automatically; you can then close the browser window.

const PROFILE_DIR = process.env.TEAMS_PROFILE || '.profile';
const AUTH_PATH = process.env.TEAMS_AUTH || '.auth/user.json';

await mkdir(dirname(AUTH_PATH), { recursive: true });

const context = await chromium.launchPersistentContext(PROFILE_DIR, {
  headless: false,
  viewport: { width: 1400, height: 900 },
});
const page = context.pages()[0] ?? await context.newPage();

try {
  console.log('Opening Teams — log in manually in the browser window that just opened.');
  await page.goto('https://teams.microsoft.com/v2/?ctx=chat', { waitUntil: 'domcontentloaded', timeout: 120000 });

  // Login is complete once the chat list renders. Give plenty of time for
  // manual sign-in + MFA.
  console.log('Waiting for you to finish logging in (up to 10 minutes)...');
  await page.getByRole('treeitem').first().waitFor({ state: 'visible', timeout: 600000 });

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
