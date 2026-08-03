import { chromium } from '@playwright/test';
import { mkdir } from 'node:fs/promises';
import { dirname } from 'node:path';

// Usage:
//   nix develop .#playwright --command node manual-login.mjs
//
// Opens Teams in a visible browser using a persistent profile at
// $TEAMS_PROFILE (default ".profile") for localStorage/cache. Because a
// persistent profile drops session cookies when reopened (and this tenant does
// not offer "Stay signed in?"), the auth cookies are also snapshotted to a
// storageState file at $TEAMS_AUTH (default ".auth/user.json") while the browser
// is open — the last snapshot before you close wins.
//
// Log in manually (including MFA), then simply close the browser window.

const PROFILE_DIR = process.env.TEAMS_PROFILE || '.profile';
const AUTH_PATH = process.env.TEAMS_AUTH || '.auth/user.json';

await mkdir(dirname(AUTH_PATH), { recursive: true });

const context = await chromium.launchPersistentContext(PROFILE_DIR, {
  headless: false,
  viewport: { width: 1400, height: 900 },
});
const page = context.pages()[0] ?? await context.newPage();

let closed = false;
context.on('close', () => { closed = true; });

console.log('Opening Teams — log in manually, then close the browser window when you are done.');
await page.goto('https://teams.microsoft.com/v2/', { waitUntil: 'domcontentloaded', timeout: 120000 });

// Snapshot cookies to the storageState file until the browser is closed.
while (!closed) {
  await new Promise((r) => setTimeout(r, 3000));
  try {
    await context.storageState({ path: AUTH_PATH });
  } catch {
    break; // context is closing/closed
  }
}

console.log(`Session saved to profile "${PROFILE_DIR}" and cookies to "${AUTH_PATH}".`);
