import {
  REACTION_TIMEOUT_MS, actOnMessages, createMessageFinder, describeMessage, emojiArgumentError,
  findPickerButton, messageLocator, openChat, openReactionPicker, openTeams, ownReactionPills,
  parseMessageIds, pickerButtons, settleReactions, waitForChatList,
} from './teams.mjs';

// Usage:
//   nix develop .#playwright --command node unreact-to-message.mjs "<chat name>" "<message ids>" "<emoji>"
//
// Takes back the <emoji> reaction we left on the messages <message ids> in the
// Teams chat whose name matches <chat name> — the counterpart of
// react-to-message.mjs, taking the same arguments. Several ids can be given as
// one comma-separated list, in which case every one of those messages loses the
// reaction. They are worked through newest first, whatever order they are given
// in, so one walk back through the history covers the whole list.
//
// Only our own reaction can be taken back: Teams offers no way to remove
// someone else's, so a message we have not reacted to with that emoji is
// reported and left alone rather than failed.
//
// A message that cannot be reached does not stop the ones after it: the run
// works through the whole list and fails at the end with what went wrong. A
// failure that is not about the message at all — the pane cannot be walked, the
// browser is gone — does stop it, since every id left would only meet the same
// failure again.

const [chatName, messageIdList, emoji] = process.argv.slice(2);

if (!chatName || !messageIdList || !emoji) {
  console.log('Usage: node unreact-to-message.mjs "<chat name>" "<message ids>" "<emoji>"');
  console.log('  <message ids>  the "id" of a message, as reported by read-chat-messages.mjs,');
  console.log('                 or several of them as a comma-separated list');
  console.log('  <emoji>        the emoji character whose reaction to take back, e.g. "👍"');
  process.exit(1);
}

const { ids: messageIds, error: idError } = parseMessageIds(messageIdList);
if (idError) {
  console.log(idError);
  process.exit(1);
}
const emojiError = emojiArgumentError(emoji);
if (emojiError) {
  console.log(emojiError);
  process.exit(1);
}

const { page, close } = await openTeams();

try {
  await waitForChatList(page);
  const resolvedName = await openChat(page, chatName);
  const findMessage = createMessageFinder(page);

  await actOnMessages(page, messageIds, mid => removeReaction(page, mid, resolvedName, findMessage), {
    changed: 'removed',
    unchanged: 'not reacted',
    couldNot: 'remove the reaction from',
    chatName: resolvedName,
  });
} finally {
  await close();
}

// Takes our reaction off one message: finds it in the history, reads the pills
// it carries and clicks the emoji again — clicking one we already left is what
// removes it — unless there is no such reaction of ours to begin with. Returns
// whether this run removed anything.
async function removeReaction(page, mid, resolvedName, findMessage) {
  console.log(`Looking for message ${mid}...`);
  if (!await findMessage(mid)) {
    throw new Error(
      `Message ${mid} was not found in "${resolvedName}" — the history was scrolled back `
      + 'as far as it goes without the message turning up. Check that the id belongs to this '
      + 'chat and that the message has not been deleted.'
    );
  }

  const message = messageLocator(page, mid);
  console.log(`Found: ${await describeMessage(page, mid)}`);

  // Reading the pills before they render would conclude that there is nothing
  // of ours to remove and leave the reaction sitting there, reporting the
  // message as done.
  await settleReactions(message);

  const ownPills = ownReactionPills(page, message, emoji);
  const removed = await hasOwnPill(ownPills) && await unreact(page, message, mid, ownPills);
  if (removed) console.log(`Removed our "${emoji}" reaction from message ${mid} in "${resolvedName}".`);
  else console.log(`Message ${mid} carries no "${emoji}" reaction of ours — nothing to remove.`);
  return removed;
}

// How long our own pill is given to turn up before the message is called
// unreacted.
const OWN_PILL_SETTLE_MS = 5000;

// Whether the message carries a reaction of ours to take back.
//
// A pill that has not rendered yet looks exactly like one that was never left,
// and settleReactions does not tell the two apart: it waits for the message's
// *first* pill, so on a message already showing someone else's reaction it is
// satisfied before ours has arrived — or before aria-pressed has settled on it.
// So an absence is given a second look rather than believed on the first read.
// This is the one verdict of this command that is reported as a success and
// then never revisited: the reacting command re-reads the pills right before it
// clicks and waits for the outcome afterwards, whereas "nothing to remove"
// leaves the reaction sitting there and the run looking as if it had done its
// job.
async function hasOwnPill(ownPills) {
  if (await ownPills.count() > 0) return true;
  await ownPills.first().waitFor({ state: 'visible', timeout: OWN_PILL_SETTLE_MS }).catch(() => {});
  return await ownPills.count() > 0;
}

// How long the message's reaction row is given to finish re-rendering before
// the pill's absence is read as the removal. The row is rebuilt whenever a
// reaction changes, and a locator asked for the pill in the middle of that sees
// nothing — the same answer a removed reaction gives.
const REMOVAL_SETTLE_MS = 1000;

// Removes the reaction: works out which of the picker's buttons applied it,
// clicks that one again and waits for the pill to go. Returns false without
// clicking anything if the reaction turns out to be gone after all.
async function unreact(page, message, mid, ownPills) {
  // Which button left the pill, rather than which button shows the emoji that
  // was asked for. Several of the picker's buttons render the same character —
  // it holds three 👏 — and clicking one that merely looks right would add a
  // second reaction instead of taking this one back.
  const itemId = await reactionItemId(ownPills.first(), mid);
  if (!itemId) {
    throw new Error(
      `Could not tell which reaction the "${emoji}" pill on message ${mid} is (no usable itemid `
      + 'on it), so the picker button that would take it back cannot be identified. The Teams '
      + 'DOM has probably changed.'
    );
  }

  // The pill of that one reaction, which is what has to disappear. "Ours, with
  // this emoji" is not specific enough to wait on: two buttons rendering one
  // character leave two pills that look alike, and only one of them is going.
  const pill = pressedPill(message, itemId, mid);

  let pickerOpen = false;
  try {
    const picker = await openReactionPicker(page, message, mid);
    pickerOpen = true;

    const button = await findPickerButton(page, reactionButtons(page, picker, itemId));
    if (!button) {
      // Said of this reaction rather than of the run: another message in the
      // list may well carry one the picker does have a button for.
      throw new Error(
        `The reaction on message ${mid} ("${itemId}") is not among the picker's emoji, so there is `
        + 'no button to click to take it back. The Teams DOM has probably changed.'
      );
    }
    // Opening the picker and walking it gave the reaction plenty of time to be
    // taken back from somewhere else, so the pill is asked for once more right
    // before the click — clicking now would put the reaction back on.
    if (await pill.count() === 0) return false;

    await button.click();
    // Clicking an emoji closes the picker, so from here on there is nothing
    // left to dismiss.
    pickerOpen = false;

    // A pill goes away entirely when the last person un-reacts, but only loses
    // its pressed state when others reacted with it too — so what is waited on
    // is the pressed pill, which covers both.
    await pill.first()
      .waitFor({ state: 'detached', timeout: REACTION_TIMEOUT_MS })
      // The cause is carried along, so a wait that failed for some other
      // reason — a crashed page, a closed target — is not read as a reaction
      // that would not go.
      .catch((err) => {
        throw new Error(
          `The "${emoji}" reaction is still on message ${mid} ${REACTION_TIMEOUT_MS / 1000}s after `
          + 'clicking it, so either it was not taken back, or the click landed on a different '
          + 'reaction and added one. Check the chat before retrying.',
          { cause: err }
        );
      });

    // Nothing matching the locator is not the same thing as the reaction being
    // gone, and two states other than a successful removal produce it: the pane
    // is virtualised, so a message unmounted after the picker closed takes
    // every selector inside it with it, and a reaction row polled in the middle
    // of the re-render a reaction change triggers is momentarily empty too.
    // Both would pass the wait above and have the message reported as done
    // while it still carries the reaction. So the absence is only believed of a
    // message that is still there to carry it, and the pill is asked for once
    // more after a pause long enough for a row that was re-rendering to have
    // put it back.
    await page.waitForTimeout(REMOVAL_SETTLE_MS);
    if (await message.count() === 0) {
      throw new Error(
        `Message ${mid} left the message pane while the "${emoji}" reaction was being taken back, `
        + 'so whether it was really removed could not be established. Read the chat back before '
        + 'retrying.'
      );
    }
    if (await pill.count() > 0) {
      throw new Error(
        `The "${emoji}" reaction is on message ${mid} again once its reaction row had re-rendered, `
        + 'so the pill going missing right after the click was that re-render rather than the '
        + 'removal. Check the chat before retrying.'
      );
    }
    return true;
  } finally {
    // The picker is a modal popup: left open it covers the message pane, and
    // the next message cannot even be hovered through it — one bad reaction
    // would cost the rest of the list its turn for a reason of its own making.
    // Only the ways out that leave it standing are dismissed, so that a run
    // that went well sends no stray keystroke into the chat. A dismissal that
    // itself fails must not replace the failure that led here.
    if (pickerOpen) await page.keyboard.press('Escape').catch(() => {});
  }
}

// The itemid of the reaction a pill shows — Teams' name for that exact
// reaction, and the suffix of the picker button that applies it ("praying" for
// [data-tid="emoticon-button-praying"]). It is on the <img> the pill renders,
// and, failing that, in the id of the element labelling the pill, which spells
// "message-<itemid>-<mid>".
//
// Null when neither is there, and equally when what is there could break out of
// the CSS attribute selector it is about to be put into — no itemid legitimately
// contains such characters.
//
// A read that fails outright is left to propagate rather than folded into that
// null: the pill detaching between being counted and being asked, the page
// going away, a throw inside the callback all say nothing about what the pill
// carries, and answering them with "no usable itemid on it, the Teams DOM has
// probably changed" would send whoever hits it looking for a change that never
// happened.
async function reactionItemId(pill, mid) {
  const itemId = await pill.evaluate((el, mid) => {
    const img = el.querySelector('img[itemid]');
    if (img) return img.getAttribute('itemid');
    const labelId = el.getAttribute('aria-labelledby') ?? '';
    return new RegExp(`^message-(.+)-${mid}$`).exec(labelId)?.[1] ?? null;
  }, mid);

  return itemId && /^[A-Za-z0-9_.:-]+$/.test(itemId) ? itemId : null;
}

// Our pill for one exact reaction: the pressed one carrying that itemid, found
// either by the <img> inside it or by the id of the element labelling it.
function pressedPill(message, itemId, mid) {
  const own = '[data-tid="diverse-reaction-pill-button"][aria-pressed="true"]';
  return message.locator(
    `${own}:has(img[itemid="${itemId}"]), ${own}[aria-labelledby="message-${itemId}-${mid}"]`
  );
}

// The picker's button for one exact reaction. The itemid is the suffix of the
// button's data-tid, and it is on the <img> inside the button as well; both are
// taken, so a Teams rename of either spelling still leaves the button findable.
// The emoji last reacted with sits in the picker's "Recent" grid, so this is
// usually a match without any scrolling at all.
function reactionButtons(page, picker, itemId) {
  return pickerButtons(picker).filter({ has: page.locator(`img[itemid="${itemId}"]`) })
    .or(picker.locator(`[data-tid="emoticon-button-${itemId}"]:visible`));
}
