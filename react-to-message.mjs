import {
  REACTION_TIMEOUT_MS, actOnMessages, clickPickerButton, createMessageFinder, describeMessage,
  emojiArgumentError, emojiImage, messageLocator, openChat, openTeams, ownReactionPills,
  parseMessageIds, pickerButtons, settleReactions, waitForChatList,
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
// and is left as it is. To take one back, use unreact-to-message.mjs.
//
// A message that cannot be reached does not stop the ones after it: the run
// works through the whole list and fails at the end with what went wrong. A
// failure that is not about the message at all — the emoji is not in the
// picker, the pane cannot be walked — does stop it, since every id left would
// only meet the same failure again.

const [chatName, messageIdList, emoji] = process.argv.slice(2);

if (!chatName || !messageIdList || !emoji) {
  console.log('Usage: node react-to-message.mjs "<chat name>" "<message ids>" "<emoji>"');
  console.log('  <message ids>  the "id" of a message, as reported by read-chat-messages.mjs,');
  console.log('                 or several of them as a comma-separated list');
  console.log('  <emoji>        the emoji character to react with, e.g. "👍"');
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

  await actOnMessages(page, messageIds, mid => reactToMessage(page, mid, resolvedName, findMessage), {
    changed: 'reacted',
    unchanged: 'already reacted',
    couldNot: 'react to',
    chatName: resolvedName,
  });
} finally {
  await close();
}

// Leaves our reaction on one message: finds it in the history, reads the pills
// it already carries and clicks the emoji unless we reacted with it before.
// Returns whether the reaction was applied by this run.
async function reactToMessage(page, mid, resolvedName, findMessage) {
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

  // Reading the pills before they render would conclude "not reacted" and click
  // the emoji — and since reacting is a toggle, that would take back the
  // reaction that was already there.
  await settleReactions(message);

  const ownPills = ownReactionPills(page, message, emoji);
  const applied = await ownPills.count() === 0 && await react(page, message, mid, ownPills);
  if (applied) console.log(`Reacted with "${emoji}" to message ${mid} in "${resolvedName}".`);
  else console.log(`Already reacted with "${emoji}" to message ${mid} — leaving it as it is.`);
  return applied;
}

// Applies the reaction: opens the message's reaction picker, finds the emoji in
// it and waits for the reaction to land on the message. Returns false without
// clicking anything if our reaction turns out to be there after all. The picker
// itself is driven from teams.mjs, which both reaction commands share; what is
// left here is this command's half of it — the button to click, and what the
// message has to do about it.
function react(page, message, mid, ownPills) {
  return clickPickerButton(page, message, mid, {
    buttons: picker => pickerButtons(picker).filter({ has: emojiImage(page, emoji) }),

    // The picker holds the same emoji for every message, so this verdict is
    // about the emoji that was asked for, not about this message.
    notInPicker: () => Object.assign(new Error(
      `The emoji "${emoji}" is not in the reaction picker. Pass the emoji character itself `
      + '(the "emoji" value read-chat-messages.mjs reports), not its name.'
    ), { systemic: true }),

    // The last moment at which a reaction we had missed can still be spared:
    // clicking the emoji now would take it back rather than leave it.
    stillNeeded: async () => await ownPills.count() === 0,

    // The pill only appears once the reaction has been accepted, so waiting for
    // it is what tells us the reaction was actually left rather than just
    // clicked.
    confirm: () => ownPills.first()
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
      }),
  });
}
