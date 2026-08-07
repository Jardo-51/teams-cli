import { openTeams, waitForChatList, openChat } from './teams.mjs';

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

  // Locate the compose box (CKEditor contenteditable).
  const composer = page.locator(
    '[data-tid="ckeditor"] [contenteditable="true"], div[role="textbox"][contenteditable="true"], [contenteditable="true"][data-tid="ckeditor"]'
  ).first();
  await composer.waitFor({ state: 'visible', timeout: 30000 });
  await composer.click();
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
