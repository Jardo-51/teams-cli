// The parts of the commands that are only data: the arguments they are started
// with, the period they are asked for, and the messages the page has rendered.
// None of it opens a browser, so all of it can be run — and covered by
// `node --test` — without one.
//
// collectRenderedMessages is the exception that proves the rule: it runs inside
// the page, sent there by page.evaluate(). A function serialised into the page
// arrives without the module around it, so everything it needs is either inside
// its own body or handed to it — which is also why it takes the document it
// reads rather than only ever reaching for the global one.

// How far apart two messages may be and still plausibly belong to the same
// author group. Teams only groups messages that are close in time, so a message
// with no author name that is further than this from its predecessor lost its
// group header rather than being part of that group.
export const AUTHOR_GROUP_WINDOW_MS = 5 * 60_000;

// The arguments of read-chat-messages.mjs. Positional arguments are taken as
// the ones that do not start with "-", so that a mistyped flag is reported as
// unknown instead of being silently read as the chat name — which is what
// unknownFlag is for; the caller prints it along with its own usage.
export function parseReadChatMessagesArgs(args) {
  const withoutReactionsOnly = args.includes('--without-reactions-only');
  const [chatName, period, outputFile] = args.filter(a => !a.startsWith('-'));
  const unknownFlag = args.find(a => a.startsWith('-') && a !== '--without-reactions-only');
  return { chatName, period, outputFile, withoutReactionsOnly, unknownFlag };
}

// The arguments of post-message.mjs.
export function parsePostMessageArgs(args) {
  const dryRun = args.includes('--dry-run');
  const [chatName, message] = args.filter(a => a !== '--dry-run');
  return { chatName, message, dryRun };
}

// The message ids react-to-message.mjs and unreact-to-message.mjs are given,
// which take the same arguments: one id, or several as a comma-separated list.
// Blank entries — a trailing or a doubled comma — are dropped rather than
// refused, since they say nothing about which messages are meant, and a
// repeated id is collapsed: its second turn would only find what the first one
// left and report it as needing nothing.
//
// Returns { error } rather than throwing, so the caller can print it the way it
// prints its own usage.
export function parseMessageIds(messageIdList) {
  const ids = [...new Set(messageIdList.split(',').map(id => id.trim()).filter(Boolean))];
  if (!ids.length) {
    return { error: `No message id in "${messageIdList}" — expected an id, or several as a comma-separated list.` };
  }
  // The ids end up inside CSS attribute selectors, so anything that could break
  // out of one is refused rather than escaped — no message id legitimately
  // contains such characters.
  for (const id of ids) {
    if (!/^[A-Za-z0-9_.:-]+$/.test(id)) {
      return { error: `Invalid message id "${id}" — expected the id read-chat-messages.mjs reports, e.g. "1785922526738".` };
    }
  }
  return { ids };
}

// Why the emoji argument cannot be used, or null when it can be. Same reasoning
// as for the ids: it too is put into a CSS attribute selector.
export function emojiArgumentError(emoji) {
  if (/["'\\]/.test(emoji)) {
    return `Invalid emoji "${emoji}" — expected a single emoji character, e.g. "👍".`;
  }
  // An emoji name ("thumbsup") or a word passes the check above and would only
  // be refused minutes later, after the browser has opened and the picker has
  // been walked. Every emoji lies outside ASCII, so that one cheap test rejects
  // plain text here; anything finer is left to the picker lookup.
  if (!/[^\x00-\x7F]/.test(emoji)) {
    return `Invalid emoji "${emoji}" — expected the emoji character itself, e.g. "👍", not its name.`;
  }
  return null;
}

// A relative time span ending "now", as "<number><unit>" where the unit is m
// (minutes), h (hours) or d (days) — e.g. "10m", "6h", "2d". Returns the span in
// milliseconds, or null for anything that is not one of those. A span of zero is
// refused too: it names no messages at all, so it is a typo rather than a
// request.
export function parsePeriod(value) {
  const match = /^(\d+)\s*([mhd])$/i.exec(value.trim());
  if (!match) return null;
  const amount = Number(match[1]);
  if (!amount) return null;
  const unit = { m: 60_000, h: 3_600_000, d: 86_400_000 }[match[2].toLowerCase()];
  return amount * unit;
}

// Consecutive messages from the same person are grouped and only the first
// carries the author name, so the last known author is carried forward through
// the ones that follow it. Takes the messages in time order and fills in their
// author in place.
export function carryForwardAuthors(messages) {
  let lastAuthor = '';
  let lastTime = -Infinity;
  for (const m of messages) {
    if (m.author) {
      lastAuthor = m.author;
    } else if (m.ts - lastTime <= AUTHOR_GROUP_WINDOW_MS) {
      m.author = lastAuthor;
    } else {
      // Too far from the previous message to belong to its group, so the header
      // this one belongs to was never collected. Naming the previous author
      // would confidently name the wrong person; leave it — and the rest of
      // this group — unattributed instead.
      lastAuthor = '';
      m.author = '';
    }
    lastTime = m.ts;
  }
}

// Reads every message rendered in the given document, which in a run is the
// page's own. Each of the ids and data-tid attributes below is an assumption
// about what Teams renders, and a silent one — a message whose markup does not
// match is simply not returned — which is why they are pinned by fixtures in
// the tests rather than only ever exercised against the live client.
export function collectRenderedMessages(doc = document) {
  const messages = [];
  for (const msg of doc.querySelectorAll('[data-tid="chat-pane-message"]')) {
    const mid = msg.getAttribute('data-mid');
    if (!mid) continue;

    const item = msg.closest('[data-tid="chat-pane-item"]');
    // Teams message ids are the send time in epoch milliseconds, which is the
    // fallback if the rendered <time> element is missing.
    const timeEl = doc.getElementById(`timestamp-${mid}`) ?? item?.querySelector('time[datetime]');
    // Date only covers ±8.64e15 ms, so a longer numeric id would make
    // toISOString() throw and take the whole run down with it.
    const fromMid = /^\d+$/.test(mid) ? new Date(Number(mid)) : null;
    const iso = timeEl?.getAttribute('datetime')
      || (fromMid && Number.isFinite(fromMid.getTime()) ? fromMid.toISOString() : '');

    const authorEl = doc.getElementById(`author-${mid}`) ?? item?.querySelector('[data-tid="message-author-name"]');
    const contentEl = doc.getElementById(`content-${mid}`) ?? msg.querySelector('[data-message-content]');

    // Parsed once here and carried alongside the ISO string, so the rest of
    // the script sorts, compares and filters without re-parsing.
    const ts = Date.parse(iso);

    // innerText only carries the anchor's display text, which Teams truncates
    // for long links (e.g. "https://.../…"), so the href is read separately to
    // keep full URLs. Both are kept because the text is the human-facing label.
    const links = contentEl
      ? [...contentEl.querySelectorAll('a[href]')].map(a => ({
          text: (a.textContent ?? '').trim(),
          href: a.getAttribute('href'),
        }))
      : [];

    messages.push({
      id: mid,
      time: iso,
      ts: Number.isFinite(ts) ? ts : null,
      author: authorEl?.textContent?.trim() ?? '',
      body: (contentEl?.innerText ?? contentEl?.textContent ?? '').trim(),
      links,
      hasReactions: !!msg.querySelector('[data-tid="diverse-reaction-pill-button"]'),
    });
  }
  return messages;
}
