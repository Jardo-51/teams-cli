import { chromium } from '@playwright/test';
import { readFile, writeFile, mkdir } from 'node:fs/promises';
import { dirname } from 'node:path';

// Usage:
//   nix develop .#playwright --command node read-chat-messages.mjs "<chat name>" "<period>" "<output file>"
//
// Reads the recent messages of the Teams chat whose name matches <chat name>
// and writes them to <output file> as JSON.
//
// <period> is a relative time span ending "now": "<number><unit>" where unit is
// m (minutes), h (hours) or d (days) — e.g. "10m", "6h", "2d".
//
// Reading is done through the Teams UI, so it only ever hovers and scrolls;
// it never clicks a reaction pill (clicking one would toggle your own
// reaction). Note that opening a chat marks its messages as read, which is
// inherent to reading them through the web client.
//
// The browser profile (localStorage, cache) is a persistent directory at
// $TEAMS_PROFILE (default ".profile"). Persistent profiles drop session cookies
// on reopen, so auth cookies are additionally loaded from the storageState file
// at $TEAMS_AUTH (default ".auth/user.json"). Both are created/refreshed by
// manual-login.mjs.

const PROFILE_DIR = process.env.TEAMS_PROFILE || '.profile';
const AUTH_PATH = process.env.TEAMS_AUTH || '.auth/user.json';

// How far back the pane is scrolled looking for the start of the period.
const MAX_SCROLL_STEPS = 60;
// Consecutive scrolls that load nothing new before we assume the top of the
// conversation has been reached.
const MAX_STAGNANT_SCROLLS = 3;

const [chatName, period, outputFile] = process.argv.slice(2);

if (!chatName || !period || !outputFile) {
  console.log('Usage: node read-chat-messages.mjs "<chat name>" "<period>" "<output file>"');
  console.log('  <period> e.g. "10m" (minutes), "6h" (hours), "2d" (days)');
  process.exit(1);
}

const periodMs = parsePeriod(period);
if (periodMs === null) {
  console.log(`Invalid period "${period}" — expected a number followed by m, h or d (e.g. "10m", "6h", "2d").`);
  process.exit(1);
}

const cutoff = Date.now() - periodMs;
console.log(`Reading messages since ${new Date(cutoff).toISOString()} (last ${period}).`);

const context = await chromium.launchPersistentContext(PROFILE_DIR, {
  headless: true,
  viewport: { width: 1400, height: 900 },
});

// Persistent profiles drop session cookies on reopen, so restore the full auth
// state — cookies plus per-origin localStorage, where MSAL keeps its tokens —
// from the storageState file. The profile itself still provides cache/warm-start.
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
  console.log(`Looking for chat: "${chatName}"`);
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

  // Wait for the conversation to render.
  await page.locator('[data-tid="chat-pane-message"]').first().waitFor({ state: 'visible', timeout: 60000 });
  await page.waitForTimeout(4000);

  // Scroll back through the history until the oldest loaded message predates
  // the cutoff (or the conversation starts). The pane is virtualised, so
  // messages are accumulated as they appear rather than read once at the end.
  const collected = new Map();
  let stagnant = 0;

  for (let step = 0; step < MAX_SCROLL_STEPS; step++) {
    let added = 0;
    for (const m of await extractMessages(page)) {
      if (!collected.has(m.id)) { collected.set(m.id, m); added++; }
    }

    const oldest = Math.min(...[...collected.values()].map(m => Date.parse(m.time)).filter(Number.isFinite));
    if (Number.isFinite(oldest) && oldest < cutoff) break;

    stagnant = added > 0 ? 0 : stagnant + 1;
    if (stagnant >= MAX_STAGNANT_SCROLLS) {
      console.log('Reached the beginning of the conversation.');
      break;
    }

    console.log(`Loading older messages (${collected.size} so far)...`);
    await page.evaluate(() => {
      const viewport = document.querySelector('[data-tid="message-pane-list-viewport"]');
      if (viewport) viewport.scrollTop = 0;
    });
    await page.waitForTimeout(2500);
  }

  // Consecutive messages from the same person are grouped and only the first
  // carries the author name, so carry the last known author forward.
  const all = [...collected.values()]
    .filter(m => Number.isFinite(Date.parse(m.time)))
    .sort((a, b) => Date.parse(a.time) - Date.parse(b.time));

  let lastAuthor = '';
  for (const m of all) {
    if (m.author) lastAuthor = m.author;
    else m.author = lastAuthor;
  }

  const inRange = all.filter(m => Date.parse(m.time) >= cutoff);
  console.log(`${inRange.length} message(s) in the last ${period}.`);

  // Reactions are only in the DOM while the pill's flyout is open. Hovering
  // opens it; clicking would toggle our own reaction, so we never click.
  const withReactions = inRange.filter(m => m.hasReactions).length;
  if (withReactions) console.log(`Reading reactions for ${withReactions} message(s)...`);

  const messages = [];
  for (const m of inRange) {
    const { hasReactions, ...rest } = m;
    messages.push({ ...rest, reactions: hasReactions ? await readReactions(page, m.id) : [] });
  }

  await mkdir(dirname(outputFile), { recursive: true }).catch(() => {});
  await writeFile(outputFile, JSON.stringify(messages, null, 2) + '\n', 'utf8');
  console.log(`Wrote ${messages.length} message(s) to "${outputFile}".`);
} finally {
  await context.close();
}

// Reads every message currently rendered in the pane.
function extractMessages(page) {
  return page.evaluate(() => {
    const messages = [];
    for (const msg of document.querySelectorAll('[data-tid="chat-pane-message"]')) {
      const mid = msg.getAttribute('data-mid');
      if (!mid) continue;

      const item = msg.closest('[data-tid="chat-pane-item"]');
      // Teams message ids are the send time in epoch milliseconds, which is the
      // fallback if the rendered <time> element is missing.
      const timeEl = document.getElementById(`timestamp-${mid}`) ?? item?.querySelector('time[datetime]');
      const iso = timeEl?.getAttribute('datetime')
        || (/^\d+$/.test(mid) ? new Date(Number(mid)).toISOString() : '');

      const authorEl = document.getElementById(`author-${mid}`) ?? item?.querySelector('[data-tid="message-author-name"]');
      const contentEl = document.getElementById(`content-${mid}`) ?? msg.querySelector('[data-message-content]');

      messages.push({
        id: mid,
        time: iso,
        author: authorEl?.textContent?.trim() ?? '',
        body: (contentEl?.innerText ?? contentEl?.textContent ?? '').trim(),
        hasReactions: !!msg.querySelector('[data-tid="diverse-reaction-pill-button"]'),
      });
    }
    return messages;
  });
}

// Hovering a reaction pill opens a flyout listing who reacted with it.
async function readReactions(page, mid) {
  const message = page.locator(`[data-tid="chat-pane-message"][data-mid="${mid}"]`).first();
  if (await message.count() === 0) {
    console.log(`  message ${mid} is no longer rendered — skipping its reactions.`);
    return [];
  }

  const userList = page.locator('[data-tid="diverse-reaction-user-list"]');
  const pills = message.locator('[data-tid="diverse-reaction-pill-button"]');
  const reactions = [];

  for (let i = 0; i < await pills.count(); i++) {
    const pill = pills.nth(i);

    // The <img> alt is the emoji character itself; itemid (e.g. "1f4dd_memo")
    // is the fallback when the alt text is empty.
    const emoji = await pill.locator('img[itemid]').first()
      .evaluate(img => img.getAttribute('alt') || img.getAttribute('itemid') || '')
      .catch(() => '');

    // The pill is labelled e.g. "2 Memo reactions.", which tells us how many
    // entries the flyout should end up listing.
    const expected = await pill.evaluate((el) => {
      const label = document.getElementById(el.getAttribute('aria-labelledby'))?.textContent ?? '';
      return Number(/^\s*(\d+)/.exec(label)?.[1] ?? 0);
    }).catch(() => 0);

    let authors = [];
    for (let attempt = 0; attempt < 3 && authors.length < Math.max(expected, 1); attempt++) {
      // Make sure a previous flyout is gone, so its entries can't be read here.
      await closeFlyout(page, userList);
      try {
        await pill.scrollIntoViewIfNeeded();
        await pill.hover();
        await userList.first().waitFor({ state: 'visible', timeout: 5000 });
        authors = await readFlyoutAuthors(page, expected);
      } catch {
        // Flyout did not open on this attempt; the next one retries it.
      }
    }

    if (authors.length < expected) {
      console.log(`  read ${authors.length}/${expected} author(s) of the "${emoji}" reaction on message ${mid}.`);
    }
    for (const author of authors) reactions.push({ author, emoji });

    // Move away so the flyout closes before the next pill is hovered,
    // otherwise its entries would be attributed to the wrong emoji.
    await closeFlyout(page, userList);
  }

  return reactions;
}

// Entries render progressively, so poll until the expected number shows up.
async function readFlyoutAuthors(page, expected) {
  let authors = [];
  for (let i = 0; i < 12; i++) {
    authors = await page.evaluate(() =>
      [...document.querySelectorAll('[data-tid="diverse-reaction-user-list-item"]')]
        .map(el => el.innerText.trim())
        .filter(Boolean)
    );
    if (authors.length >= expected) break;
    await page.waitForTimeout(500);
  }
  return authors;
}

async function closeFlyout(page, userList) {
  await page.mouse.move(5, 5);
  await userList.first().waitFor({ state: 'hidden', timeout: 5000 }).catch(() => {});
}

function parsePeriod(value) {
  const match = /^(\d+)\s*([mhd])$/i.exec(value.trim());
  if (!match) return null;
  const amount = Number(match[1]);
  if (!amount) return null;
  const unit = { m: 60_000, h: 3_600_000, d: 86_400_000 }[match[2].toLowerCase()];
  return amount * unit;
}

function escapeRegExp(s) {
  return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}
