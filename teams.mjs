import { chromium } from '@playwright/test';
import { readFile } from 'node:fs/promises';

// Shared plumbing for the Teams scripts: launching a browser with a restored
// session, and opening a chat by name.

// The browser profile (localStorage, cache) is a persistent directory at
// $TEAMS_PROFILE. Persistent profiles drop session cookies on reopen, so auth
// cookies are additionally stored in the storageState file at $TEAMS_AUTH.
// Both are created/refreshed by manual-login.mjs.
export const PROFILE_DIR = process.env.TEAMS_PROFILE || '.profile';
export const AUTH_PATH = process.env.TEAMS_AUTH || '.auth/user.json';

const TEAMS_URL = 'https://teams.microsoft.com/v2/?ctx=chat';

// Launches a browser on the persistent profile, restores the saved session and
// navigates to Teams. Pass restoreAuth: false when capturing a new session.
export async function openTeams({ headless = true, restoreAuth = true } = {}) {
  const context = await chromium.launchPersistentContext(PROFILE_DIR, {
    headless,
    viewport: { width: 1400, height: 900 },
  });

  if (restoreAuth) await restoreSession(context);

  const page = context.pages()[0] ?? await context.newPage();

  console.log('Opening Teams...');
  await page.goto(TEAMS_URL, { waitUntil: 'domcontentloaded', timeout: 120000 });

  return { context, page };
}

// Restores the full auth state — cookies plus per-origin localStorage, where
// MSAL keeps its tokens — from the storageState file. The persistent profile
// itself still provides cache/warm-start.
async function restoreSession(context) {
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
}

// The SPA is ready once the left rail renders. A generous timeout covers manual
// sign-in (including MFA) when the session still has to be established.
export async function waitForChatList(page, { timeout = 120000 } = {}) {
  await page.getByRole('treeitem').first().waitFor({ state: 'visible', timeout });
}

// Opens the chat whose name matches chatName (partial, case-insensitive) and
// returns the name it resolved to.
export async function openChat(page, chatName) {
  console.log(`Looking for chat: "${chatName}"`);

  // Group headers (e.g. "Favorites", "Chats") are also treeitems that CONTAIN
  // the chat rows, so we must pick the leaf: a matching treeitem that has no
  // nested treeitem inside it.
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

  // Wait for the conversation itself rather than its messages, so that opening
  // an empty chat does not stall.
  await page.locator('[data-tid="message-pane-list-viewport"]').first()
    .waitFor({ state: 'visible', timeout: 60000 });
  await page.waitForTimeout(4000);

  return resolvedName;
}

function escapeRegExp(s) {
  return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}
