import { openTeams, waitForChatList, openChat, composerLocator, clearComposer } from './teams.mjs';

// Usage:
//   nix develop .#playwright --command node post-message.mjs "<chat name>" "<message>" [--dry-run]
//
// Posts <message> into the Teams chat whose name matches <chat name>.
// With --dry-run the message is typed into the compose box but NOT sent,
// so you can confirm the correct chat is targeted before anything goes out.

const args = process.argv.slice(2);
const dryRun = args.includes('--dry-run');
const positional = args.filter(a => a !== '--dry-run');
const [chatName, message] = positional;

if (!chatName || !message) {
  console.log('Usage: node post-message.mjs "<chat name>" "<message>" [--dry-run]');
  process.exit(1);
}

const { page, close } = await openTeams();

try {
  await waitForChatList(page);
  const resolvedName = await openChat(page, chatName);

  const composer = composerLocator(page);
  await composer.waitFor({ state: 'visible', timeout: 30000 });
  // Typing puts the text at the caret, so anything already in the box becomes
  // part of the message that goes out. The page reset clears the draft of the
  // chat the previous command left open; this covers the rest — a draft Teams
  // itself synced in from another client, or one in a chat this page has not
  // had open. Nothing typed here is worth risking a wrong message for.
  await clearComposer(composer);
  await composer.type(message, { delay: 15 });

  if (dryRun) {
    console.log('DRY RUN: message typed but not sent.');
  } else {
    await page.keyboard.press('Enter');
    await page.waitForTimeout(4000);
    console.log(`Sent to "${resolvedName}".`);
  }
} finally {
  await close();
}
