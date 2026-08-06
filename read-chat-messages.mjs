import { writeFile, mkdir } from 'node:fs/promises';
import { dirname } from 'node:path';
import { openTeams, waitForChatList, openChat } from './teams.mjs';

// Usage:
//   nix develop .#playwright --command node read-chat-messages.mjs "<chat name>" "<period>" "<output file>" [--without-reactions-only]
//
// Reads the recent messages of the Teams chat whose name matches <chat name>
// and writes them to <output file> as JSON.
//
// <period> is a relative time span ending "now": "<number><unit>" where unit is
// m (minutes), h (hours) or d (days) — e.g. "10m", "6h", "2d".
//
// With --without-reactions-only the output is limited to the messages that
// nobody has reacted to.
//
// Reading is done through the Teams UI, so it only ever hovers and scrolls;
// it never clicks a reaction pill (clicking one would toggle your own
// reaction). Note that opening a chat marks its messages as read, which is
// inherent to reading them through the web client.

// How much of the viewport height each scroll step moves. Kept below 1 so
// consecutive rendered windows overlap and nothing falls between them.
const SCROLL_STEP_FRACTION = 0.8;
// How far back the pane is scrolled looking for the start of the period. Each
// step only covers part of a viewport, so covering a period of days takes a
// lot of them.
const MAX_SCROLL_STEPS = 300;
// Consecutive scrolls that load nothing new before we assume the top of the
// conversation has been reached.
const MAX_STAGNANT_SCROLLS = 3;
// How far apart two messages may be and still plausibly belong to the same
// author group. Teams only groups messages that are close in time, so a
// message with no author name that is further than this from its predecessor
// lost its group header rather than being part of that group.
const AUTHOR_GROUP_WINDOW_MS = 5 * 60_000;

const args = process.argv.slice(2);
const withoutReactionsOnly = args.includes('--without-reactions-only');
const [chatName, period, outputFile] = args.filter(a => !a.startsWith('-'));

// Catch a mistyped flag instead of silently reading it as the chat name.
const unknownFlag = args.find(a => a.startsWith('-') && a !== '--without-reactions-only');

if (!chatName || !period || !outputFile || unknownFlag) {
  if (unknownFlag) console.log(`Unknown option "${unknownFlag}".`);
  console.log('Usage: node read-chat-messages.mjs "<chat name>" "<period>" "<output file>" [--without-reactions-only]');
  console.log('  <period>                   e.g. "10m" (minutes), "6h" (hours), "2d" (days)');
  console.log('  --without-reactions-only   only messages that nobody reacted to');
  process.exit(1);
}

const periodMs = parsePeriod(period);
if (periodMs === null) {
  console.log(`Invalid period "${period}" — expected a number followed by m, h or d (e.g. "10m", "6h", "2d").`);
  process.exit(1);
}

const cutoff = Date.now() - periodMs;
console.log(`Reading messages since ${new Date(cutoff).toISOString()} (last ${period}).`);

const { context, page } = await openTeams();

try {
  await waitForChatList(page);
  await openChat(page, chatName);

  // Give the messages themselves a moment; a chat with none stays empty and
  // simply yields no results.
  await page.locator('[data-tid="chat-pane-message"]').first()
    .waitFor({ state: 'visible', timeout: 30000 })
    .catch(() => console.log('No messages are rendered in this chat.'));

  // Scroll back through the history until the oldest loaded message predates
  // the cutoff (or the conversation starts). The pane is virtualised, so
  // messages are accumulated as they appear rather than read once at the end.
  const collected = new Map();
  let stagnant = 0;
  let atTop = false;
  let oldest = Infinity;
  // Whether the scroll loop got far enough back to cover the whole period,
  // rather than running out of steps part-way.
  let periodCovered = false;

  for (let step = 0; step < MAX_SCROLL_STEPS; step++) {
    let added = 0;
    for (const m of await extractMessages(page)) {
      if (!collected.has(m.id)) { collected.set(m.id, m); added++; }
    }

    oldest = Math.min(...[...collected.values()].map(m => Date.parse(m.time)).filter(Number.isFinite));
    if (Number.isFinite(oldest) && oldest < cutoff) { periodCovered = true; break; }

    stagnant = added > 0 ? 0 : stagnant + 1;
    if (stagnant >= MAX_STAGNANT_SCROLLS) {
      // Nothing new loaded — but that only means the start of the conversation
      // if the pane is genuinely at the top of what it has. Otherwise older
      // messages just stopped arriving, and calling it the beginning would pass
      // off a truncated read as a complete one.
      if (!atTop) {
        throw new Error(
          `No new messages after ${MAX_STAGNANT_SCROLLS} consecutive scrolls, `
          + 'but the pane is not at the top of the loaded history. Older messages '
          + 'are loading too slowly, or the pane is not scrolling.'
        );
      }
      // There is nothing older to read, so the period is covered as fully as
      // this conversation allows.
      console.log('Reached the beginning of the conversation.');
      periodCovered = true;
      break;
    }

    console.log(`Loading older messages (${collected.size} so far)...`);
    const scrolled = await scrollUp(page);
    if (!scrolled) {
      throw new Error(
        'The message pane viewport ([data-tid="message-pane-list-viewport"]) was not '
        + 'found, so the history cannot be scrolled. The Teams DOM has probably changed.'
      );
    }
    // An unchanged scrollTop means the pane was already at the top of the
    // loaded range; anything else means the scroll really happened.
    atTop = scrolled.after === scrolled.before;
    await page.waitForTimeout(2500);
  }

  // Running out of scroll steps truncates the result, which otherwise looks
  // exactly like a quiet chat — say so rather than letting the message count
  // be the only clue.
  if (!periodCovered) {
    const reached = Number.isFinite(oldest) ? new Date(oldest).toISOString() : 'nothing at all';
    console.log(
      `WARNING: stopped after ${MAX_SCROLL_STEPS} scrolls, only reaching ${reached}, `
      + `so the last ${period} is NOT fully covered. The output is truncated.`
    );
  }

  // Consecutive messages from the same person are grouped and only the first
  // carries the author name, so carry the last known author forward.
  const all = [...collected.values()]
    .filter(m => Number.isFinite(Date.parse(m.time)))
    .sort((a, b) => Date.parse(a.time) - Date.parse(b.time));

  let lastAuthor = '';
  let lastTime = -Infinity;
  for (const m of all) {
    const time = Date.parse(m.time);
    if (m.author) {
      lastAuthor = m.author;
    } else if (time - lastTime <= AUTHOR_GROUP_WINDOW_MS) {
      m.author = lastAuthor;
    } else {
      // Too far from the previous message to belong to its group, so the header
      // this one belongs to was never collected. Naming the previous author
      // would confidently name the wrong person; leave it — and the rest of
      // this group — unattributed instead.
      lastAuthor = '';
      m.author = '';
    }
    lastTime = time;
  }

  const inRange = all.filter(m => Date.parse(m.time) >= cutoff);
  console.log(`${inRange.length} message(s) in the last ${period}.`);

  // Whether a message has any reaction is already visible on the message
  // itself, so filtering happens before the reaction authors are read — with
  // the flag, nothing selected has reactions and the pass below is skipped.
  const selected = withoutReactionsOnly ? inRange.filter(m => !m.hasReactions) : inRange;
  if (withoutReactionsOnly) console.log(`${selected.length} of them without reactions.`);

  // Reactions are only in the DOM while the pill's flyout is open. Hovering
  // opens it; clicking would toggle our own reaction, so we never click.
  const withReactions = selected.filter(m => m.hasReactions).length;
  if (withReactions) console.log(`Reading reactions for ${withReactions} message(s)...`);

  // The scroll pass above left the pane at the top of the loaded history, so
  // the messages we want — the newest ones — are no longer rendered. Walk back
  // down and read newest-first, bringing each message into the rendered window
  // just before its pills are hovered.
  const reactionsById = new Map();
  if (withReactions) {
    await scrollToNewest(page);
    for (const m of [...selected].reverse()) {
      if (m.hasReactions) reactionsById.set(m.id, await readReactions(page, m.id));
    }
  }

  const messages = [];
  for (const m of selected) {
    const { hasReactions, ...rest } = m;
    // null (rather than []) where the reactions could not be read, so that a
    // failure is visible in the output instead of looking like "nobody reacted".
    messages.push({ ...rest, reactions: hasReactions ? reactionsById.get(m.id) ?? null : [] });
  }

  await mkdir(dirname(outputFile), { recursive: true }).catch(() => {});
  await writeFile(outputFile, JSON.stringify(messages, null, 2) + '\n', 'utf8');
  console.log(`Wrote ${messages.length} message(s) to "${outputFile}".`);
} finally {
  await context.close();
}

// Scrolls the pane up by roughly a viewport. Returns the scrollTop before and
// after the move, or null when the viewport element could not be found.
function scrollUp(page) {
  return page.evaluate((fraction) => {
    const viewport = document.querySelector('[data-tid="message-pane-list-viewport"]');
    if (!viewport) return null;
    // Step up by roughly a viewport rather than jumping to the top. The pane is
    // virtualised and only the rendered window is readable, so a jump would
    // skip everything between the old window and the new one.
    const before = viewport.scrollTop;
    viewport.scrollTop = Math.max(0, before - viewport.clientHeight * fraction);
    return { before, after: viewport.scrollTop };
  }, SCROLL_STEP_FRACTION);
}

// Jumps back to the newest messages at the bottom of the pane.
async function scrollToNewest(page) {
  await page.evaluate(() => {
    const viewport = document.querySelector('[data-tid="message-pane-list-viewport"]');
    if (viewport) viewport.scrollTop = viewport.scrollHeight;
  });
  await page.waitForTimeout(2500);
}

// Brings a message back into the rendered window. Callers work newest-first, so
// this only ever has to walk up from wherever the previous message left the
// pane. Returns false if the message could not be reached.
async function scrollMessageIntoView(page, mid) {
  const message = page.locator(`[data-tid="chat-pane-message"][data-mid="${mid}"]`).first();
  for (let step = 0; step < MAX_SCROLL_STEPS; step++) {
    if (await message.count() > 0) {
      await message.scrollIntoViewIfNeeded().catch(() => {});
      return true;
    }
    const scrolled = await scrollUp(page);
    if (!scrolled || scrolled.after === scrolled.before) return false;
    await page.waitForTimeout(1500);
  }
  return false;
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

// Hovering a reaction pill opens a flyout listing who reacted with it. Returns
// null when the message could not be brought back into the rendered window, so
// that "could not read" stays distinguishable from "nobody reacted".
async function readReactions(page, mid) {
  if (!await scrollMessageIntoView(page, mid)) {
    console.log(`  message ${mid} could not be brought back into view — its reactions are unknown.`);
    return null;
  }

  const message = page.locator(`[data-tid="chat-pane-message"][data-mid="${mid}"]`).first();
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
    let stuckFlyout = false;
    for (let attempt = 0; attempt < 3 && authors.length < Math.max(expected, 1); attempt++) {
      // A previous flyout still on screen would have its entries read here and
      // attributed to this emoji, so skip the pill rather than read through it.
      if (!await closeFlyout(page, userList)) { stuckFlyout = true; break; }
      try {
        await pill.scrollIntoViewIfNeeded();
        await pill.hover();
        await userList.first().waitFor({ state: 'visible', timeout: 5000 });
        authors = await readFlyoutAuthors(page, userList, expected);
      } catch {
        // Flyout did not open on this attempt; the next one retries it.
      }
    }

    if (stuckFlyout) {
      console.log(`  previous flyout did not close — skipping the "${emoji}" reaction on message ${mid}.`);
      continue;
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

// Entries render progressively, so poll until the expected number shows up. The
// entries are read from within the flyout element rather than from the whole
// document, so that a lingering flyout cannot contribute names to this one.
async function readFlyoutAuthors(page, userList, expected) {
  let authors = [];
  for (let i = 0; i < 12; i++) {
    // Two flyouts open at once makes attribution guesswork; no names beat names
    // filed under the wrong emoji.
    if (await userList.count() > 1) return [];

    authors = await userList.first().evaluate(list =>
      [...list.querySelectorAll('[data-tid="diverse-reaction-user-list-item"]')]
        .map(el => el.innerText.trim())
        .filter(Boolean)
    ).catch(() => []);

    if (authors.length >= expected) break;
    await page.waitForTimeout(500);
  }
  return authors;
}

// Moves the pointer away from the pill and waits for the flyout to go. Returns
// false if it is still on screen, since its entries would then be read as part
// of whatever is hovered next.
function closeFlyout(page, userList) {
  return page.mouse.move(5, 5)
    .then(() => userList.first().waitFor({ state: 'hidden', timeout: 5000 }))
    .then(() => true, () => false);
}

function parsePeriod(value) {
  const match = /^(\d+)\s*([mhd])$/i.exec(value.trim());
  if (!match) return null;
  const amount = Number(match[1]);
  if (!amount) return null;
  const unit = { m: 60_000, h: 3_600_000, d: 86_400_000 }[match[2].toLowerCase()];
  return amount * unit;
}
