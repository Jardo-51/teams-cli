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

// How much of the viewport height each scroll step moves. Kept below 1 so
// consecutive rendered windows overlap and nothing falls between them.
const SCROLL_STEP_FRACTION = 0.8;
// How far back the pane is scrolled looking for older messages. Each step only
// covers part of a viewport, so covering a period of days takes a lot of them.
export const MAX_SCROLL_STEPS = 300;
// How long the pane gets to deliver older history once it sits at the top of
// the loaded range, before "nothing arrived" is accepted as the start of the
// conversation. Deliberately generous: a fetch that is merely slow looks
// exactly like a chat with no more history, and mistaking one for the other
// cuts the search short silently.
const OLDER_HISTORY_GRACE_MS = 20_000;

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
  } catch (err) {
    // A missing file is the ordinary first-run case. Anything else — corrupt
    // JSON, no read permission, a cookie rejected by addCookies — needs its own
    // message, since running manual-login.mjs would not necessarily fix it and
    // the caller otherwise sits in waitForChatList() until it times out.
    if (err.code === 'ENOENT') {
      console.log(`No saved session at "${AUTH_PATH}" — run manual-login.mjs first.`);
    } else {
      console.log(`Could not restore the session from "${AUTH_PATH}": ${err.message}`);
    }
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
  // an empty chat does not stall. Callers then wait on whatever they actually
  // need — the messages, or the compose box — so there is nothing to sleep for
  // here.
  await page.locator('[data-tid="message-pane-list-viewport"]').first()
    .waitFor({ state: 'visible', timeout: 60000 });

  return resolvedName;
}

// Scrolls the message pane up by roughly a viewport. Returns the scrollTop
// before and after the move plus the pane's scrolling geometry, or null when
// the viewport element could not be found.
export function scrollUp(page) {
  return page.evaluate((fraction) => {
    const viewport = document.querySelector('[data-tid="message-pane-list-viewport"]');
    if (!viewport) return null;
    // Step up by roughly a viewport rather than jumping to the top. The pane is
    // virtualised and only the rendered window is readable, so a jump would
    // skip everything between the old window and the new one.
    const before = viewport.scrollTop;
    viewport.scrollTop = Math.max(0, before - viewport.clientHeight * fraction);
    const overflowY = getComputedStyle(viewport).overflowY;
    return {
      before,
      after: viewport.scrollTop,
      clientHeight: viewport.clientHeight,
      overflowY,
      scrollable: ['auto', 'scroll', 'overlay'].includes(overflowY),
    };
  }, SCROLL_STEP_FRACTION);
}

// The pane's scroll extent and the id of the topmost rendered message — the two
// things that move when older history is added to the list. Null when the
// viewport element could not be found.
function readPaneState(page) {
  return page.evaluate(() => {
    const viewport = document.querySelector('[data-tid="message-pane-list-viewport"]');
    if (!viewport) return null;
    const oldest = document.querySelector('[data-tid="chat-pane-message"]');
    return {
      scrollHeight: viewport.scrollHeight,
      oldestMid: oldest?.getAttribute('data-mid') ?? null,
    };
  });
}

// Waits for an older-history fetch to land while the pane sits at the top of
// the loaded range. Returns true if more history arrived within the grace
// period — scrollTop cannot show this, since it is pinned at 0 either way.
export async function waitForOlderHistory(page) {
  const before = await readPaneState(page);
  if (!before) throw viewportGoneError();

  console.log('At the top of the loaded history — waiting for older messages...');
  const deadline = Date.now() + OLDER_HISTORY_GRACE_MS;
  while (Date.now() < deadline) {
    await page.waitForTimeout(1000);
    const now = await readPaneState(page);
    if (!now) throw viewportGoneError();
    if (now.scrollHeight > before.scrollHeight || now.oldestMid !== before.oldestMid) return true;
  }
  return false;
}

export function viewportGoneError() {
  return new Error(
    'The message pane viewport ([data-tid="message-pane-list-viewport"]) was not '
    + 'found, so the history cannot be scrolled. The Teams DOM has probably changed.'
  );
}

// Brings a message into the rendered window, scrolling back through the history
// — and waiting for older history to load — until it appears. Callers start at
// the newest end and work backwards, so this only ever scrolls up. Returns false
// if the message could not be reached.
export async function scrollMessageIntoView(page, mid) {
  const message = page.locator(`[data-tid="chat-pane-message"][data-mid="${mid}"]`).first();
  for (let step = 0; step < MAX_SCROLL_STEPS; step++) {
    if (await message.count() > 0) {
      await message.scrollIntoViewIfNeeded().catch(() => {});
      return true;
    }
    const scrolled = await scrollUp(page);
    // A vanished viewport is a broken selector, not a missing message: returning
    // false here would have the caller report a DOM change as a bad message id.
    if (!scrolled) throw viewportGoneError();
    // Parked at the top of what is loaded: the pane cannot move any further, but
    // older history may still be on its way, so give the fetch a chance before
    // concluding the message is not there.
    if (scrolled.after === scrolled.before) {
      if (!await waitForOlderHistory(page)) return false;
      continue;
    }
    await page.waitForTimeout(1500);
  }
  return false;
}

function escapeRegExp(s) {
  return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}
