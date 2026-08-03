import { chromium } from '@playwright/test';
import { readFile } from 'node:fs/promises';

// Usage:
//   nix develop .#playwright --command node post-message.mjs "<chat name>" "<message>" [--dry-run]
//
// Posts <message> into the Teams chat whose name matches <chat name>.
// With --dry-run the message is typed into the compose box but NOT sent,
// so you can confirm the correct chat is targeted before anything goes out.
//
// The browser profile (localStorage, cache) is a persistent directory at
// $TEAMS_PROFILE (default ".profile"). Persistent profiles drop session cookies
// on reopen, so auth cookies are additionally loaded from the storageState file
// at $TEAMS_AUTH (default ".auth/user.json"). Both are created/refreshed by
// manual-login.mjs.

const PROFILE_DIR = process.env.TEAMS_PROFILE || '.profile';
const AUTH_PATH = process.env.TEAMS_AUTH || '.auth/user.json';

const args = process.argv.slice(2);
const dryRun = args.includes('--dry-run');
const positional = args.filter(a => a !== '--dry-run');
const [chatName, message] = positional;

if (!chatName || !message) {
  console.log('Usage: node post-message.mjs "<chat name>" "<message>" [--dry-run]');
  process.exit(1);
}

const context = await chromium.launchPersistentContext(PROFILE_DIR, {
  headless: true,
  viewport: { width: 1400, height: 900 },
});

// Persistent profiles drop session cookies on reopen (and this tenant has no
// "Stay signed in?"), so restore the full auth state — cookies plus per-origin
// localStorage, where MSAL keeps its tokens — from the storageState file. The
// profile itself still provides cache/warm-start.
try {
  const state = JSON.parse(await readFile(AUTH_PATH, 'utf8'));
  if (state.cookies?.length) await context.addCookies(state.cookies);
  for (const { origin, localStorage } of state.origins ?? []) {
    if (!localStorage?.length) continue;
    await context.addInitScript((data) => {
      if (location.origin === data.origin) {
        for (const { name, value } of data.items) {
          try { window.localStorage.setItem(name, value); } catch {}
        }
      }
    }, { origin, items: localStorage });
  }
} catch {
  console.log(`No saved session at "${AUTH_PATH}" — run manual-login.mjs first.`);
}

const page = context.pages()[0] ?? await context.newPage();

try {
  console.log('Opening Teams...');
  await page.goto('https://teams.microsoft.com/v2/?ctx=chat', { waitUntil: 'domcontentloaded', timeout: 120000 });
  await page.waitForTimeout(18000); // let the SPA hydrate

  // Find the chat in the left rail by (partial, case-insensitive) name.
  // Group headers (e.g. "Favorites", "Chats") are also treeitems that CONTAIN
  // the chat rows, so we must pick the leaf: a matching treeitem that has no
  // nested treeitem inside it.
  console.error(`Looking for chat: "${chatName}"`);
  const matching = page.getByRole('treeitem').filter({ hasText: new RegExp(escapeRegExp(chatName), 'i') });
  await matching.first().waitFor({ state: 'visible', timeout: 30000 });

  let chatItem = null;
  const count = await matching.count();
  for (let i = 0; i < count; i++) {
    const candidate = matching.nth(i);
    const nested = await candidate.getByRole('treeitem').count();
    if (nested === 0) { chatItem = candidate; break; }
  }
  if (!chatItem) {
    throw new Error(`No leaf chat row matched "${chatName}" (only group headers matched).`);
  }

  const resolvedName = (await chatItem.innerText()).split('\n')[0].trim();
  console.log(`Matched chat: "${resolvedName}"`);
  await chatItem.click();
  await page.waitForTimeout(6000); // let the conversation load

  // Locate the compose box (CKEditor contenteditable).
  const composer = page.locator(
    '[data-tid="ckeditor"] [contenteditable="true"], div[role="textbox"][contenteditable="true"], [contenteditable="true"][data-tid="ckeditor"]'
  ).first();
  await composer.waitFor({ state: 'visible', timeout: 30000 });
  await composer.click();
  await composer.type(message, { delay: 15 });

  if (dryRun) {
    console.log('DRY RUN: message typed but not sent.');
  } else {
    await page.keyboard.press('Enter');
    await page.waitForTimeout(4000);
    console.log(`Sent to "${resolvedName}".`);
  }
} finally {
  await context.close();
}

function escapeRegExp(s) {
  return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}
