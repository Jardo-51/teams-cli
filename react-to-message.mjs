import {
  openTeams, waitForChatList, openChat, scrollToNewest, scrollMessageIntoView, messageSelector,
  messageLocator,
} from './teams.mjs';

// Usage:
//   nix develop .#playwright --command node react-to-message.mjs "<chat name>" "<message ids>" "<emoji>"
//
// Reacts with <emoji> to the messages <message ids> in the Teams chat whose name
// matches <chat name>. Several ids can be given as one comma-separated list, in
// which case every one of those messages gets the reaction. They are worked
// through newest first, whatever order they are given in, so one walk back
// through the history covers the whole list. The ids and the emoji are the ones
// read-chat-messages.mjs reports, so its output can be fed straight back in.
//
// Reacting is a toggle in Teams, so a reaction we already left is never clicked
// again — that would take it back. Such a message reports the existing reaction
// and is left as it is.
//
// A message that cannot be reached does not stop the ones after it: the run
// works through the whole list and fails at the end with what went wrong. A
// failure that is not about the message at all — the emoji is not in the
// picker, the pane cannot be walked — does stop it, since every id left would
// only meet the same failure again.

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

const [chatName, messageIdList, emoji] = process.argv.slice(2);

if (!chatName || !messageIdList || !emoji) {
  console.log('Usage: node react-to-message.mjs "<chat name>" "<message ids>" "<emoji>"');
  console.log('  <message ids>  the "id" of a message, as reported by read-chat-messages.mjs,');
  console.log('                 or several of them as a comma-separated list');
  console.log('  <emoji>        the emoji character to react with, e.g. "👍"');
  process.exit(1);
}

// The same message twice would have the second turn find the reaction the first
// one left and report it as already there, so the list is reduced to its
// distinct ids. Blank entries — a trailing or doubled comma — are dropped
// rather than refused, since they say nothing about which messages are meant.
const messageIds = [...new Set(messageIdList.split(',').map(id => id.trim()).filter(Boolean))];

if (!messageIds.length) {
  console.log(`No message id in "${messageIdList}" — expected an id, or several as a comma-separated list.`);
  process.exit(1);
}

// Both values end up inside CSS attribute selectors, so anything that could
// break out of one is refused rather than escaped — no message id or emoji
// legitimately contains these characters.
for (const messageId of messageIds) {
  if (!/^[A-Za-z0-9_.:-]+$/.test(messageId)) {
    console.log(`Invalid message id "${messageId}" — expected the id read-chat-messages.mjs reports, e.g. "1785922526738".`);
    process.exit(1);
  }
}
if (/["'\\]/.test(emoji)) {
  console.log(`Invalid emoji "${emoji}" — expected a single emoji character, e.g. "👍".`);
  process.exit(1);
}
// An emoji name ("thumbsup") or a word passes the check above and would only be
// refused minutes later, after the browser has opened and the picker has been
// walked. Every emoji lies outside ASCII, so that one cheap test rejects plain
// text here; anything finer is left to the picker lookup.
if (!/[^\x00-\x7F]/.test(emoji)) {
  console.log(`Invalid emoji "${emoji}" — expected the emoji character itself, e.g. "👍", not its name.`);
  process.exit(1);
}

// The history walk only ever scrolls back, so the ids are worked through newest
// first: every target after the first is then older than where the pane already
// stands, and one walk carries on through the whole list instead of each id
// sending it back to the newest messages. The ids are epoch milliseconds, so
// ordering them numerically orders them in time; an id that is not a plain
// number carries no such order, so a list holding one is left in the order it
// was given.
const orderedIds = messageIds.every(id => /^\d+$/.test(id))
  ? [...messageIds].sort((a, b) => Number(b) - Number(a))
  : messageIds;

const { page, close } = await openTeams();

try {
  await waitForChatList(page);
  const resolvedName = await openChat(page, chatName);

  let reacted = 0;
  let alreadyReacted = 0;
  const failures = [];

  for (const messageId of orderedIds) {
    try {
      if (await reactToMessage(page, messageId, resolvedName)) {
        reacted++;
        console.log(`Reacted with "${emoji}" to message ${messageId} in "${resolvedName}".`);
      } else {
        alreadyReacted++;
        console.log(`Already reacted with "${emoji}" to message ${messageId} — leaving it as it is.`);
      }
    } catch (err) {
      // One unreachable message must not cost the rest of the list its
      // reaction, so what went wrong is kept and the run moves on. The
      // collected failures are raised together once the list is done.
      console.log(`Could not react to message ${messageId}: ${err.message}`);
      failures.push(err);
      // Unless the failure was never about this message: the same verdict
      // awaits every id left, each after another full walk back through the
      // history, so the run stops here and reports what it has rather than
      // proving the same thing over and over.
      if (isSystemic(page, err)) {
        console.log('This says nothing about the remaining messages either — stopping here.');
        break;
      }
    }
  }

  if (messageIds.length > 1) {
    // A run stopped by a systemic failure leaves ids it never looked at, and
    // counting them as anything else would misreport what happened to them.
    const notAttempted = messageIds.length - reacted - alreadyReacted - failures.length;
    console.log(
      `${messageIds.length} message(s): ${reacted} reacted, ${alreadyReacted} already reacted, `
      + `${failures.length} failed` + (notAttempted ? `, ${notAttempted} not attempted.` : '.')
    );
  }
  if (failures.length) {
    throw new AggregateError(
      failures,
      `Could not react to ${failures.length} of ${messageIds.length} message(s) in "${resolvedName}".`
    );
  }
} finally {
  await close();
}

// Leaves our reaction on one message: finds it in the history, reads the pills
// it already carries and clicks the emoji unless we reacted with it before.
// Returns whether the reaction was applied by this run.
async function reactToMessage(page, mid, resolvedName) {
  console.log(`Looking for message ${mid}...`);
  if (!await findMessage(page, mid)) {
    throw new Error(
      `Message ${mid} was not found in "${resolvedName}" — the history was scrolled back `
      + 'as far as it goes without the message turning up. Check that the id belongs to this '
      + 'chat and that the message has not been deleted.'
    );
  }

  const message = messageLocator(page, mid);
  console.log(`Found: ${await describe(page, mid)}`);

  // The pane is virtualised, so the message may have been mounted a moment ago
  // with its reaction row still to come. Reading the pills before they render
  // would conclude "not reacted" and click the emoji — and since reacting is a
  // toggle, that would take back the reaction that was already there.
  await settleReactions(message);

  // Reactions are stored per person, and clicking one we already left removes
  // it. Ours are the pills the client marks as pressed.
  const ownPill = message.locator('[data-tid="diverse-reaction-pill-button"][aria-pressed="true"]')
    .filter({ has: emojiImage(page) });
  return await ownPill.count() === 0 && await react(page, message, mid, ownPill);
}

// Whether a failure was about the run rather than about the message it happened
// on — the emoji is not in the picker, the pane cannot be walked at all, the
// browser is gone. The scripts mark such errors as they raise them; a page that
// has closed under us is the same verdict arrived at from the other side.
function isSystemic(page, err) {
  return err?.systemic === true || page.isClosed();
}

// Whether the pane still stands where the chat opened it, at the newest
// messages. The walk below only ever goes back, so this is what says whether
// what the walk has already passed can still be reached without returning to
// the bottom first.
let paneAtNewest = true;

// Brings the message into the pane. The pane opens at the newest messages and
// the walk only ever goes back, so an older target is reached by scrolling —
// the same walk read-chat-messages.mjs makes. Since the ids are handled newest
// first, the walk carries on from where the previous message left it rather
// than starting over.
//
// A target that is neither rendered nor older than the pane lies ahead of it,
// and that direction is only reachable from the newest end, so a walk that came
// up empty is tried once more from there. A walk that started at the newest end
// has already seen the whole history, so it is not repeated — an id that
// belongs to another chat costs one walk, not two.
async function findMessage(page, mid) {
  const startedAtNewest = paneAtNewest;
  // Either walk may leave the pane part way back through the history.
  paneAtNewest = false;

  if (await scrollMessageIntoView(page, mid, { maxHistoryWaits: MAX_HISTORY_WAITS })) return true;
  if (startedAtNewest) return false;

  console.log(`Message ${mid} is not behind the pane — looking again from the newest messages...`);
  await scrollToNewest(page);
  return scrollMessageIntoView(page, mid, { maxHistoryWaits: MAX_HISTORY_WAITS });
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

  // Forced past the actionability check on purpose: for a message at the top of
  // the pane the toolbar renders under the pinned-message banner, which sits on
  // top of the "More reactions" button and intercepts the click. The button is
  // the right target — it is just visually overlapped — so the receives-events
  // check is the wrong guard here and would only time out.
  await actions.locator('[data-tid="expanded-reactions-picker-entry"]').click({ force: true });

  try {
    // The picker's frame appears before its emoji do, so wait for the list
    // itself — searching it while it is still empty would find nothing.
    const picker = page.locator('[data-tid="reaction-picker-root"]');
    await picker.locator('[data-tid^="emoticon-button-"]').first()
      .waitFor({ state: 'visible', timeout: PICKER_TIMEOUT_MS });

    const button = await findEmojiButton(page, picker);
    if (!button) {
      // The picker holds the same emoji for every message, so this verdict is
      // about the emoji that was asked for, not about this message.
      throw Object.assign(new Error(
        `The emoji "${emoji}" is not in the reaction picker. Pass the emoji character itself `
        + '(the "emoji" value read-chat-messages.mjs reports), not its name.'
      ), { systemic: true });
    }
    // Opening the picker and walking it gave the message plenty of time to
    // finish rendering, so the pills are asked once more right before the click
    // — this is the last moment at which a reaction we had missed can still be
    // spared.
    if (await ownPill.count() > 0) return false;

    await button.click();

    // The pill only appears once the reaction has been accepted, so waiting for
    // it is what tells us the reaction was actually left rather than just
    // clicked.
    await ownPill.first()
      .waitFor({ state: 'visible', timeout: REACTION_TIMEOUT_MS })
      // The cause is carried along, so a wait that failed for some other
      // reason — a crashed page, a closed target — is not read as a reaction
      // gone missing.
      .catch((err) => {
        throw new Error(
          `The "${emoji}" reaction is not on message ${mid} ${REACTION_TIMEOUT_MS / 1000}s after `
          + 'clicking it, so either it was not saved, or it was already there unrendered and the '
          + 'click took it back. Check the chat before retrying.',
          { cause: err }
        );
      });
    return true;
  } finally {
    // The picker is a modal popup: left open it covers the message pane, and
    // the next message cannot even be hovered through it — one bad emoji would
    // cost the rest of the list its reaction for a reason of its own making.
    // A dismissal that itself fails must not replace the failure that led here.
    await page.keyboard.press('Escape').catch(() => {});
  }
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
      throw Object.assign(new Error(
        'The reaction picker has no emoji list ([data-tid="unified-picker-emojis-content"]), so '
        + 'no emoji could be searched. The Teams DOM has probably changed.'
      ), { systemic: true });
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
  return page.evaluate(({ mid, selector }) => {
    const msg = document.querySelector(selector);
    const item = msg?.closest('[data-tid="chat-pane-item"]');
    const author = document.getElementById(`author-${mid}`)?.textContent?.trim()
      || item?.querySelector('[data-tid="message-author-name"]')?.textContent?.trim()
      || '(unknown author)';
    const content = document.getElementById(`content-${mid}`) ?? msg?.querySelector('[data-message-content]');
    const body = (content?.innerText ?? content?.textContent ?? '').trim().replace(/\s+/g, ' ');
    return `${author}: ${body.length > 80 ? body.slice(0, 80) + '…' : body}`;
  }, { mid, selector: messageSelector(mid) });
}
