import {
  REACTION_TIMEOUT_MS, actOnMessages, closeReactionOverflow, createMessageFinder, describeMessage,
  emojiArgumentError, emojiImage, ensureEmojiCatalog, messageLocator, openChat,
  openReactionOverflow, openTeams, ownReactionPills, parseMessageIds, reactionOverflowButton,
  settleReactions, waitForChatList, withOpenPopup,
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

// How long our own pill is given to turn up before the message is called
// unreacted.
const OWN_PILL_SETTLE_MS = 5000;

// How long the message's reaction row is given to finish re-rendering before
// what it shows is read as the removal. The row is rebuilt whenever a reaction
// changes, and a locator asked in the middle of that sees nothing — the same
// answer a removed reaction gives.
const REMOVAL_SETTLE_MS = 1000;

// How often the message is asked whether the reaction taken back from behind
// its "+N" has really gone from it.
const OVERFLOW_POLL_MS = 250;

// How long the "+N" is given to turn up before the message is taken to be
// hiding nothing. Shorter than the wait our own pill gets above, because what
// it covers is smaller: the overflow button is part of the same reaction row as
// the pills, and settleReactions has already waited for that row's first pill,
// so this is the button rendering a tick behind the pills beside it rather than
// the row arriving at all.
const OVERFLOW_SETTLE_MS = 2000;

// How many reactions of ours one message may have taken off it in a single run,
// applied to each of the two places they can be. It is a bound rather than a
// limit anyone should reach: the picker holds three buttons for the most
// crowded character and none of them can be applied twice. What it is really
// for is keeping a reaction that the removal does not actually clear from being
// clicked round after round.
const MAX_OWN_REACTIONS = 10;

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
  // Before the chat is opened, so that a catalog that has to be synced again
  // costs one reload here rather than one after the history has been walked.
  await ensureEmojiCatalog(page);
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

// Takes our reaction off one message: finds it in the history and clicks the
// pill it left — clicking a reaction we already left is what takes it back —
// unless there is no such reaction of ours to begin with. A message showing no
// pill for it has not thereby been left alone: past six distinct reactions
// Teams hides the rest behind a "+N", and ours may be one of them. Returns
// whether this run removed anything.
async function removeReaction(page, mid, resolvedName, findMessage) {
  console.log(`Looking for message ${mid}...`);
  if (!await findMessage(mid)) {
    throw new Error(
      `Message ${mid} was not found in "${resolvedName}" — the history was walked back as far `
      + 'as one run goes without it turning up, and the walk said above what stopped it. Check '
      + 'that the id belongs to this chat and that the message has not been deleted.'
    );
  }

  const message = messageLocator(page, mid);
  console.log(`Found: ${await describeMessage(page, mid)}`);

  // Reading the pills before they render would conclude that there is nothing
  // of ours to remove and leave the reaction sitting there, reporting the
  // message as done.
  await settleReactions(message);

  const removed = await unreactAll(page, message, mid);
  if (removed === 1) console.log(`Removed our "${emoji}" reaction from message ${mid} in "${resolvedName}".`);
  else if (removed > 1) console.log(`Removed our ${removed} "${emoji}" reactions from message ${mid} in "${resolvedName}".`);
  else console.log(`Message ${mid} carries no "${emoji}" reaction of ours — nothing to remove.`);
  return removed > 0;
}

// Takes back every reaction of ours on the message that renders <emoji>, and
// returns how many of them this run removed.
//
// Usually that is the one, but the character does not identify a reaction:
// several of the picker's buttons render the same one, so a message can carry
// two pressed pills of ours that look alike. Neither command produces such a
// pair — react-to-message.mjs leaves a message that already carries any own
// pill with the emoji alone — but reacting by hand in the client does, and it
// is exactly the case this command exists to get right. Taking only the first
// one back while reporting the message as removed would leave the other sitting
// there with nothing in the output hinting at it.
//
// The two places a reaction of ours can be are worked in a fixed order: what
// the "+N" overflow is hiding goes first, and only then the pills. A removal
// from the overflow promotes a reaction that was hidden into the rendered row,
// and were that promoted pill another look-alike of ours, the pill route would
// count as many pills after its click as before it — a removal that worked,
// read as one that did not. Draining the overflow first leaves nothing of ours
// behind to be promoted into the row that is being counted, and nothing can put
// one back there: what a later removal promotes is by then someone else's.
async function unreactAll(page, message, mid) {
  const ownPills = ownReactionPills(page, message, emoji);
  let removed = await hasOverflow(message)
    ? await unreactOverflowed(page, message, mid)
    : 0;

  while (removed < MAX_OWN_REACTIONS) {
    // The second look is only worth its wait while nothing has been removed
    // yet: after a removal an empty reaction row is the expected end of this
    // loop rather than a pill that may still be on its way.
    if (!await hasOwnPill(ownPills, removed === 0)) break;
    await unreactPill(page, message, mid, ownPills);
    removed++;
  }
  return removed;
}

// Whether the message is hiding reactions behind a "+N".
//
// Absence gets the second look the pills get below, and for the same reason.
// settleReactions waits for the message's *first* pill, so a reaction row that
// renders its pills before the overflow button beside them satisfies it while
// the "+N" is still to come — and this read is what decides whether the
// overflow is opened at all. Believed on the first look, an overflow that was
// merely late is an overflow never opened, and a reaction of ours sitting
// behind it is reported as "nothing to remove": the false success issue #23 was
// filed to kill, reached from the other side.
async function hasOverflow(message) {
  const overflow = reactionOverflowButton(message);
  if (await overflow.count() > 0) return true;
  await overflow.first().waitFor({ state: 'visible', timeout: OVERFLOW_SETTLE_MS }).catch(() => {});
  return await overflow.count() > 0;
}

// Whether the message renders a pill for a reaction of ours to take back.
//
// A pill that has not rendered yet looks exactly like one that was never left,
// and settleReactions does not tell the two apart: it waits for the message's
// *first* pill, so on a message already showing someone else's reaction it is
// satisfied before ours has arrived — or before aria-pressed has settled on it.
// So with <settle> an absence is given a second look rather than believed on
// the first read. This is the one verdict of this command that is reported as a
// success and then never revisited: the reacting command re-reads the pills
// right before it clicks and waits for the outcome afterwards, whereas "nothing
// to remove" leaves the reaction sitting there and the run looking as if it had
// done its job.
async function hasOwnPill(ownPills, settle) {
  if (await ownPills.count() > 0) return true;
  if (!settle) return false;
  await ownPills.first().waitFor({ state: 'visible', timeout: OWN_PILL_SETTLE_MS }).catch(() => {});
  return await ownPills.count() > 0;
}

// Takes back one of the reactions the message renders a pill for, by clicking
// that pill: a reaction is stored per person and clicking one we left is what
// removes it, so the pill is both the thing that says the reaction is ours and
// the control that takes it back.
//
// Which of several look-alike pills went is not asked. The click takes back the
// reaction of the pill it landed on, whichever that was, and the loop above
// comes round for whatever is left — so what has to be established is that the
// message carries one fewer of them than it did, not which one it lost. That is
// what the nth() waits on: with one gone there is no longer an nth pill for it
// to match, and a locator matching nothing is a locator that is detached.
//
// What is counted is the pressed pills, and that is what makes the count a fair
// reading of the removal: a pill goes off the message entirely when the last
// person un-reacts, but only loses its pressed state when other people reacted
// with it too, and counting only the pressed ones covers both.
async function unreactPill(page, message, mid, ownPills) {
  const before = await ownPills.count();
  await ownPills.first().click();

  await ownPills.nth(before - 1)
    .waitFor({ state: 'detached', timeout: REACTION_TIMEOUT_MS })
    // The cause is carried along, so a wait that failed for some other reason —
    // a crashed page, a closed target — is not read as a reaction that would
    // not go.
    .catch((err) => {
      throw new Error(
        `The "${emoji}" reaction is still on message ${mid} ${REACTION_TIMEOUT_MS / 1000}s after `
        + 'its pill was clicked, so it was not taken back. Check the chat before retrying.',
        { cause: err }
      );
    });

  await confirmPillRemoved(page, message, mid, ownPills, before);
}

// Waits for the pill's absence to be the removal rather than a moment in a
// re-render, and says what it means when it is not.
//
// Nothing matching the locator is not the same thing as the reaction being
// gone, and two states other than a successful removal produce it: the pane is
// virtualised, so a message unmounted after the click takes every selector
// inside it with it, and a reaction row polled in the middle of the re-render a
// reaction change triggers is momentarily empty too. Both would pass the wait
// above and have the message reported as done while it still carries the
// reaction. So the absence is only believed of a message that is still there to
// carry it, and the pills are counted once more after a pause long enough for a
// row that was re-rendering to have put them back.
async function confirmPillRemoved(page, message, mid, ownPills, before) {
  await page.waitForTimeout(REMOVAL_SETTLE_MS);
  if (await message.count() === 0) {
    throw new Error(
      `Message ${mid} left the message pane while the "${emoji}" reaction was being taken back, `
      + 'so whether it was really removed could not be established. Read the chat back before '
      + 'retrying.'
    );
  }
  if (await ownPills.count() >= before) {
    throw new Error(
      `The "${emoji}" reaction is on message ${mid} again once its reaction row had re-rendered, `
      + 'so the pill going missing right after the click was that re-render rather than the '
      + 'removal. Check the chat before retrying.'
    );
  }
}

// Takes back every reaction of ours with <emoji> that the message's "+N" is
// hiding, and returns how many that was. Usually none: what the overflow hides
// is everyone's reactions, and only our own can be taken back.
//
// This is the only way to such a reaction. Teams renders no pill for it, so
// there is nothing on the message to click and nothing to read its emoji off —
// to the pills alone, a reaction sitting in the overflow looks exactly like a
// reaction that is not there.
async function unreactOverflowed(page, message, mid) {
  const menu = await openReactionOverflow(page, message, mid);

  return withOpenPopup(menu, () => closeReactionOverflow(page, message), async () => {
    const rows = ourOverflowRows(page, menu);
    let removed = 0;

    // The menu is not opened again between removals: it survives the click that
    // takes one back and re-renders with the rows that are left, so what is
    // still there is read off the open menu. It closes by itself along with the
    // overflow, when the reaction removed was the last one being hidden — and a
    // menu that is gone matches no rows, which ends the loop.
    while (removed < MAX_OWN_REACTIONS && await rows.count() > 0) {
      // Read before the click, because it is what the removal is confirmed
      // against afterwards.
      const hidden = await hiddenReactionCount(message);
      // No "+N" on the message means nothing is being hidden, so there is
      // nothing left here to take back — whatever rows the menu is still
      // showing, it is showing them over a message that has moved on. It has to
      // be stopped on rather than clicked through, because zero is a baseline
      // no confirmation can pass: confirmOverflowRemoved asks whether the count
      // is still at or above what it was, and every count is at or above zero,
      // so the confirmation would poll out the full timeout and then report a
      // reaction the click had in fact taken back as one that would not go. The
      // moment is reachable without anything being wrong: the overflow can
      // collapse a tick before the menu it opened unmounts, and the reaction
      // row is momentarily empty mid-re-render, which is the same thing
      // REMOVAL_SETTLE_MS is ridden out everywhere else in this file.
      if (hidden === 0) break;
      if (hidden === null) {
        throw new Error(
          `The "+N" reaction overflow of message ${mid} does not say how many reactions it is `
          + 'hiding, so a reaction taken back from it could not be confirmed as gone. The Teams '
          + 'DOM has probably changed.'
        );
      }

      await rows.first().locator('[data-tid="remove-reaction-button"]').click();
      await confirmOverflowRemoved(page, message, mid, hidden);
      removed++;
    }
    return removed;
  });
}

// The overflow's rows for a reaction of ours with <emoji>: the ones showing
// that emoji and carrying the button that takes a reaction back.
//
// The menu holds a row per person per reaction, and only our own can be taken
// back, so it is that button which says which of the rows are ours — matched on
// rather than assumed of every row, since in a group chat the overflow hides
// other people's reactions beside ours. The emoji is matched the way it is on
// the pills, on the img Teams renders it as, rather than on the row's label:
// the label spells the reaction's name ("Clapping"), which is whatever language
// the client is in.
function ourOverflowRows(page, menu) {
  return menu.locator('[data-tid="diverse-reaction-user-list-item"]')
    .filter({ has: emojiImage(page, emoji) })
    .filter({ has: page.locator('[data-tid="remove-reaction-button"]') });
}

// How many reactions the message's "+N" is hiding, read off the button itself
// ("+2" → 2) — digits rather than its label, which spells the count out in the
// client's language. Zero when the message has no overflow at all, and null
// when it has one that says something this cannot read, so that a confirmation
// is never built on a number that was not there.
async function hiddenReactionCount(message) {
  const overflow = reactionOverflowButton(message);
  if (await overflow.count() === 0) return 0;
  const digits = /\d+/.exec(await overflow.first().innerText().catch(() => ''));
  return digits ? Number(digits[0]) : null;
}

// Waits for the message to really be hiding one reaction fewer than it was.
//
// The menu is no evidence by itself: it stays open over a removal, and the row
// leaving it says only that the menu re-rendered. So what is waited on is the
// message: either its "+N" counts down, or it goes altogether, which is what
// happens when the reaction removed was the one that pushed the row over the
// six Teams renders pills for. And as after a pill click, the answer is only
// believed once the reaction row has finished re-rendering: a row polled in the
// middle of that can be without its overflow for a moment.
async function confirmOverflowRemoved(page, message, mid, before) {
  const stillHidden = async () => {
    const hidden = await hiddenReactionCount(message);
    return hidden === null || hidden >= before;
  };

  const deadline = Date.now() + REACTION_TIMEOUT_MS;
  while (await stillHidden()) {
    if (Date.now() > deadline) {
      throw new Error(
        `The "${emoji}" reaction is still behind the "+N" of message ${mid} `
        + `${REACTION_TIMEOUT_MS / 1000}s after it was removed from the overflow, so it was not `
        + 'taken back. Check the chat before retrying.'
      );
    }
    await page.waitForTimeout(OVERFLOW_POLL_MS);
  }

  await page.waitForTimeout(REMOVAL_SETTLE_MS);
  if (await message.count() === 0) {
    throw new Error(
      `Message ${mid} left the message pane while the "${emoji}" reaction was being taken back, `
      + 'so whether it was really removed could not be established. Read the chat back before '
      + 'retrying.'
    );
  }
  if (await stillHidden()) {
    throw new Error(
      `The "${emoji}" reaction is behind the "+N" of message ${mid} again once its reaction row `
      + 'had re-rendered, so the overflow counting down right after the click was that re-render '
      + 'rather than the removal. Check the chat before retrying.'
    );
  }
}
