import { openTeams, waitForChatList, openChat, composerLocator, clearComposer, pasteIntoComposer } from './teams.mjs';

// Usage:
//   nix develop .#playwright --command node post-message.mjs "<chat name>" "<message>" [--dry-run]
//
// Posts <message> into the Teams chat whose name matches <chat name>. It may
// span several lines; the whole of it is posted as a single message, verbatim.
// With --dry-run the message is put into the compose box but NOT sent,
// so you can confirm the correct chat is targeted before anything goes out.

const args = process.argv.slice(2);
const dryRun = args.includes('--dry-run');
const positional = args.filter(a => a !== '--dry-run');
const [chatName, message] = positional;

// A message of nothing but whitespace is rejected here rather than sent: Teams
// treats such a compose box as empty and refuses the Enter, so it would go no
// further than a command that reported a message it never delivered.
if (!chatName || !message?.trim()) {
  console.log('Usage: node post-message.mjs "<chat name>" "<message>" [--dry-run]');
  process.exit(1);
}

const { page, close } = await openTeams();

try {
  await waitForChatList(page);
  // Posting never reads the message pane, so it does not wait for the pane to
  // settle at the newest end. What makes this safe is openChat confirming the
  // switch itself — the compose box of the chat being left stays visible for
  // most of a second, and typing into that one would post to the wrong chat.
  const resolvedName = await openChat(page, chatName, { atNewest: false });

  const composer = composerLocator(page);
  await composer.waitFor({ state: 'visible', timeout: 30000 });
  // The message goes in at the caret, so anything already in the box becomes
  // part of the message that goes out. The page reset clears the draft of the
  // chat the previous command left open; this covers the rest — a draft Teams
  // itself synced in from another client, or one in a chat this page has not
  // had open. No draft is worth risking a wrong message for.
  await clearComposer(composer);
  // Whatever <message> holds goes out as one message: pasteIntoComposer keeps a
  // newline in it from reaching the composer as the Enter that sends.
  await pasteIntoComposer(composer, message);

  if (dryRun) {
    console.log('DRY RUN: message put in the compose box but not sent.');
  } else {
    await page.keyboard.press('Enter');
    await page.waitForTimeout(4000);
    console.log(`Sent to "${resolvedName}".`);
  }
} finally {
  await close();
}
