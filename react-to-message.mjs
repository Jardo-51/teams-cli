import { openTeams, waitForChatList, openChat, scrollMessageIntoView } from './teams.mjs';

// Usage:
//   nix develop .#playwright --command node react-to-message.mjs "<chat name>" "<message id>" "<emoji>"
//
// Reacts with <emoji> to the message <message id> in the Teams chat whose name
// matches <chat name>. The id and the emoji are the ones read-chat-messages.mjs
// reports, so its output can be fed straight back in.
//
// Reacting is a toggle in Teams, so a reaction we already left is never clicked
// again — that would take it back. Such a run reports the existing reaction and
// changes nothing.

// How often the walk back through the history may pause for a fetch of older
// messages. The target can be arbitrarily far back, so the pauses are needed
// here; the cap only bounds how much of a long conversation one run pages in.
const MAX_HISTORY_WAITS = 100;
// How far the reaction picker is scrolled looking for the emoji. The emoji list
// is virtualised and only the rendered window is searchable, so it has to be
// walked a viewport at a time; the full list takes some sixty steps.
const MAX_PICKER_SCROLL_STEPS = 200;
// How much of the picker's height each of those steps moves. Kept below 1 so
// consecutive rendered windows overlap and no row falls between them.
const PICKER_SCROLL_FRACTION = 0.8;
// How long the picker gets to render its emoji. Its frame appears first, so this
// covers the list arriving, not the popup opening.
const PICKER_TIMEOUT_MS = 15000;
// How long each scroll step of the picker is given to render the emoji it moved
// into view, before that window is searched.
const PICKER_SETTLE_MS = 250;
// How long the hover toolbar of the message gets to open.
const MESSAGE_ACTIONS_TIMEOUT_MS = 15000;
// How long the reaction is given to show up on the message after the emoji is
// clicked — it is only really applied once the server has taken it.
const REACTION_TIMEOUT_MS = 15000;
// How long the message itself is given to render after the history walk reports
// it, before anything is read off it.
const MESSAGE_TIMEOUT_MS = 15000;
// How long an already-rendered message is given to bring its reaction row with
// it. A message nobody reacted to has no pills at all, so running out here is an
// ordinary outcome rather than a failure.
const REACTION_SETTLE_MS = 5000;

const [chatName, messageId, emoji] = process.argv.slice(2);

if (!chatName || !messageId || !emoji) {
  console.log('Usage: node react-to-message.mjs "<chat name>" "<message id>" "<emoji>"');
  console.log('  <message id>   the "id" of a message, as reported by read-chat-messages.mjs');
  console.log('  <emoji>        the emoji character to react with, e.g. "👍"');
  process.exit(1);
}

// Both values end up inside CSS attribute selectors, so anything that could
// break out of one is refused rather than escaped — no message id or emoji
// legitimately contains these characters.
if (!/^[A-Za-z0-9_.:-]+$/.test(messageId)) {
  console.log(`Invalid message id "${messageId}" — expected the id read-chat-messages.mjs reports, e.g. "1785922526738".`);
  process.exit(1);
}
if (/["'\\]/.test(emoji)) {
  console.log(`Invalid emoji "${emoji}" — expected a single emoji character, e.g. "👍".`);
  process.exit(1);
}

const { context, page } = await openTeams();

try {
  await waitForChatList(page);
  const resolvedName = await openChat(page, chatName);

  // The pane opens at the newest messages, so an older target is reached by
  // scrolling back — the same walk read-chat-messages.mjs makes.
  console.log(`Looking for message ${messageId}...`);
  if (!await scrollMessageIntoView(page, messageId, { maxHistoryWaits: MAX_HISTORY_WAITS })) {
    throw new Error(
      `Message ${messageId} was not found in "${resolvedName}" — the history was scrolled back `
      + 'as far as it goes without the message turning up. Check that the id belongs to this '
      + 'chat and that the message has not been deleted.'
    );
  }

  const message = page.locator(`[data-tid="chat-pane-message"][data-mid="${messageId}"]`).first();
  console.log(`Found: ${await describe(page, messageId)}`);

  // The pane is virtualised, so the message may have been mounted a moment ago
  // with its reaction row still to come. Reading the pills before they render
  // would conclude "not reacted" and click the emoji — and since reacting is a
  // toggle, that would take back the reaction that was already there.
  await settleReactions(message);

  // Reactions are stored per person, and clicking one we already left removes
  // it. Ours are the pills the client marks as pressed.
  const ownPill = message.locator('[data-tid="diverse-reaction-pill-button"][aria-pressed="true"]')
    .filter({ has: emojiImage(page) });
  const applied = await ownPill.count() === 0 && await react(page, message, messageId, ownPill);
  console.log(applied
    ? `Reacted with "${emoji}" to message ${messageId} in "${resolvedName}".`
    : `Already reacted with "${emoji}" to this message — leaving it as it is.`);
} finally {
  await context.close();
}

// Gives a freshly rendered message time to render its reaction row, so that the
// pills can be read off it. A message nobody has reacted to never grows one, so
// waiting for the row is a settle rather than a requirement.
async function settleReactions(message) {
  await message.waitFor({ state: 'visible', timeout: MESSAGE_TIMEOUT_MS });
  await message.locator('[data-tid="diverse-reaction-pill-button"]').first()
    .waitFor({ state: 'visible', timeout: REACTION_SETTLE_MS })
    .catch(() => {});
}

// Applies the reaction: opens the message's reaction picker, finds the emoji in
// it and waits for the reaction to land on the message. Returns false without
// clicking anything if our reaction turns out to be there after all.
async function react(page, message, mid, ownPill) {
  const actions = await openMessageActions(page, message, mid);

  await actions.locator('[data-tid="expanded-reactions-picker-entry"]').click();
  // The picker's frame appears before its emoji do, so wait for the list
  // itself — searching it while it is still empty would find nothing.
  const picker = page.locator('[data-tid="reaction-picker-root"]');
  await picker.locator('[data-tid^="emoticon-button-"]').first()
    .waitFor({ state: 'visible', timeout: PICKER_TIMEOUT_MS });

  const button = await findEmojiButton(page, picker);
  if (!button) {
    throw new Error(
      `The emoji "${emoji}" is not in the reaction picker. Pass the emoji character itself `
      + '(the "emoji" value read-chat-messages.mjs reports), not its name.'
    );
  }
  // Opening the picker and walking it gave the message plenty of time to finish
  // rendering, so the pills are asked once more right before the click — this is
  // the last moment at which a reaction we had missed can still be spared.
  if (await ownPill.count() > 0) return false;

  await button.click();

  // The pill only appears once the reaction has been accepted, so waiting for
  // it is what tells us the reaction was actually left rather than just clicked.
  await ownPill.first()
    .waitFor({ state: 'visible', timeout: REACTION_TIMEOUT_MS })
    // The cause is carried along, so a wait that failed for some other reason —
    // a crashed page, a closed target — is not read as a reaction gone missing.
    .catch((err) => {
      throw new Error(
        `The "${emoji}" reaction is not on message ${mid} ${REACTION_TIMEOUT_MS / 1000}s after `
        + 'clicking it, so either it was not saved, or it was already there unrendered and the '
        + 'click took it back. Check the chat before retrying.',
        { cause: err }
      );
    });
  return true;
}

// Hovers the message to raise its action toolbar. The toolbar is rendered
// outside the message, in a popover whose id carries the message id — which is
// the only thing tying the two together, so it is matched on rather than
// assumed: reacting to whatever else is hovered would be worse than failing.
async function openMessageActions(page, message, mid) {
  await message.scrollIntoViewIfNeeded();
  await message.hover();

  const actions = page.locator(`[data-tid="message-actions-container"][id="${mid}-popover-surface"]`);
  await actions.waitFor({ state: 'visible', timeout: MESSAGE_ACTIONS_TIMEOUT_MS }).catch((err) => {
    throw new Error(
      `The action toolbar of message ${mid} did not open on hover (no visible `
      + `[data-tid="message-actions-container"] with id "${mid}-popover-surface"). The Teams DOM `
      + 'has probably changed.',
      { cause: err }
    );
  });
  return actions;
}

// Finds the emoji's button in the open picker, scrolling the list until it
// shows up. Returns null if the whole list was walked without a match.
async function findEmojiButton(page, picker) {
  // The picker keeps hidden copies of its grids in the DOM, and those cannot be
  // clicked, so only what is actually rendered counts as a match.
  const buttons = picker.locator('[data-tid^="emoticon-button-"]:visible').filter({ has: emojiImage(page) });

  for (let step = 0; step < MAX_PICKER_SCROLL_STEPS; step++) {
    if (await buttons.count() > 0) return buttons.first();

    const scrolled = await scrollPicker(page, PICKER_SCROLL_FRACTION);
    if (scrolled.reason === 'no-content') {
      throw new Error(
        'The reaction picker has no emoji list ([data-tid="unified-picker-emojis-content"]), so '
        + 'no emoji could be searched. The Teams DOM has probably changed.'
      );
    }
    // A list with nothing to scroll — one that fits on screen, or a filtered
    // set — was searched in full above, so it is a plain "not in the picker",
    // the same as reaching the bottom.
    if (scrolled.reason === 'not-scrollable' || scrolled.after === scrolled.before) return null;
    await page.waitForTimeout(PICKER_SETTLE_MS);
  }
  return null;
}

// Scrolls the picker's emoji list down by a fraction of its height. Returns the
// scrollTop before and after the move, or a reason why nothing moved — kept
// apart because "the picker is not there" is a DOM change while "there is
// nothing to scroll" is an ordinary short list.
function scrollPicker(page, fraction) {
  return page.evaluate((fraction) => {
    const content = document.querySelector('[data-tid="reaction-picker-root"] [data-tid="unified-picker-emojis-content"]');
    if (!content) return { reason: 'no-content' };
    // The element that scrolls is an unnamed wrapper inside the emoji content,
    // so it is picked out by being the one with something to scroll.
    const scrollers = [...content.querySelectorAll('*')].filter(el =>
      el.scrollHeight > el.clientHeight && ['auto', 'scroll', 'overlay'].includes(getComputedStyle(el).overflowY));
    if (!scrollers.length) return { reason: 'not-scrollable' };
    // Nested wrappers can overflow by a few pixels each, and the first in
    // document order is not necessarily the emoji grid — the grid is the one
    // with a whole virtualised list's worth of scrolling left in it.
    const scroller = scrollers.reduce((widest, el) =>
      el.scrollHeight - el.clientHeight > widest.scrollHeight - widest.clientHeight ? el : widest);
    const before = scroller.scrollTop;
    scroller.scrollTop = before + scroller.clientHeight * fraction;
    return { before, after: scroller.scrollTop };
  }, fraction);
}

// Matches the <img> Teams renders an emoji as. Its alt is the emoji character,
// which may or may not carry the U+FE0F variation selector — the same emoji
// either way, so both spellings are accepted.
function emojiImage(page) {
  const bare = emoji.replace(/\uFE0F/g, '');
  return page.locator(`img[alt="${bare}"], img[alt="${bare}\uFE0F"]`);
}

// A short "author: body" line, so the run says which message it acted on rather
// than only echoing back the id.
function describe(page, mid) {
  return page.evaluate((mid) => {
    const msg = document.querySelector(`[data-tid="chat-pane-message"][data-mid="${mid}"]`);
    const item = msg?.closest('[data-tid="chat-pane-item"]');
    const author = document.getElementById(`author-${mid}`)?.textContent?.trim()
      || item?.querySelector('[data-tid="message-author-name"]')?.textContent?.trim()
      || '(unknown author)';
    const content = document.getElementById(`content-${mid}`) ?? msg?.querySelector('[data-message-content]');
    const body = (content?.innerText ?? content?.textContent ?? '').trim().replace(/\s+/g, ' ');
    return `${author}: ${body.length > 80 ? body.slice(0, 80) + '…' : body}`;
  }, mid);
}
