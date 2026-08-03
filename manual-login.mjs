import { chromium } from '@playwright/test';
import { mkdir } from 'node:fs/promises';
import { existsSync } from 'node:fs';
import { dirname } from 'node:path';
import { createInterface } from 'node:readline/promises';

// Usage:
//   nix develop .#playwright --command node manual-login.mjs
//
// Opens Teams in a visible browser so you can log in manually (including MFA).
// An existing session at $TEAMS_AUTH (default ".auth/user.json") is loaded if
// present, so you may only need to refresh it. When you're done, return to this
// terminal and press Enter — the session is saved back to the same path used by
// post-message.mjs.

const AUTH_PATH = process.env.TEAMS_AUTH || '.auth/user.json';

const browser = await chromium.launch({ headless: false });
const context = await browser.newContext({
  storageState: existsSync(AUTH_PATH) ? AUTH_PATH : undefined,
  viewport: { width: 1400, height: 900 },
});
const page = await context.newPage();

try {
  console.log('Opening Teams — log in manually in the browser window that just opened.');
  await page.goto('https://teams.microsoft.com/v2/', { waitUntil: 'domcontentloaded', timeout: 120000 });

  const rl = createInterface({ input: process.stdin, output: process.stdout });
  await rl.question('\nWhen you have finished logging in, press Enter here to save the session... ');
  rl.close();

  await mkdir(dirname(AUTH_PATH), { recursive: true });
  await context.storageState({ path: AUTH_PATH });
  console.log(`Session saved to ${AUTH_PATH}`);
} finally {
  await browser.close();
}
