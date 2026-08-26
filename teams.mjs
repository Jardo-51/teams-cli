import { chromium } from '@playwright/test';
import { mkdir, readFile } from 'node:fs/promises';
import { dirname } from 'node:path';
import {
  CDP_ENDPOINT, DAEMON_ENABLED, acquireCommandLock, connectableDaemon, ensureDaemon, stopDaemon,
  touchActivity,
} from './daemon.mjs';

// Shared plumbing for the Teams scripts: launching a browser with a restored
// session, and opening a chat by name.

// The browser profile (localStorage, cache) is a persistent directory at
// $TEAMS_PROFILE. Persistent profiles drop session cookies on reopen, so auth
// cookies are additionally stored in the storageState file at $TEAMS_AUTH.
// Both are created/refreshed by manual-login.mjs.
export const PROFILE_DIR = process.env.TEAMS_PROFILE || '.profile';
export const AUTH_PATH = process.env.TEAMS_AUTH || '.auth/user.json';

const TEAMS_URL = 'https://teams.microsoft.com/v2/?ctx=chat';
// What a Teams tab's URL begins with, for picking it out of the shared browser's
// pages. The path is rewritten by the SPA as it is navigated; the origin is not.
const TEAMS_ORIGIN = 'https://teams.microsoft.com/';

// How long the daemon's page gets to show its chat list before it is treated as
// stale and reloaded. A healthy tab has it rendered already, so this only has to
// cover a client that is busy, not one that is booting.
const HEALTH_CHECK_TIMEOUT_MS = 15_000;
// How long the message pane gets to settle after being scrolled to the newest
// messages, so the virtualised list has rendered that end before it is read.
const PANE_SETTLE_MS = 2500;
// How long a clicked chat gets to become the one on screen. Generous next to the
// sub-second switch it normally waits for, but not unbounded: this is also what
// a chat whose title is spelled differently from its row would wait out on every
// single command, so it trades a slow worst case against a slow common one.
const CHAT_SWITCH_TIMEOUT_MS = 15_000;

// How much of the viewport height each scroll step moves. Kept below 1 so
// consecutive rendered windows overlap and nothing falls between them.
const SCROLL_STEP_FRACTION = 0.8;
// How far back the pane is scrolled looking for older messages. Each step only
// covers part of a viewport, so covering a period of days takes a lot of them.
export const MAX_SCROLL_STEPS = 300;
// How long the pane gets to deliver older history once it sits at the top of
// the loaded range, before "nothing arrived" is accepted as the start of the
// conversation. Deliberately generous: a fetch that is merely slow looks
// exactly like a chat with no more history, and mistaking one for the other
// cuts the search short silently.
const OLDER_HISTORY_GRACE_MS = 20_000;

// Gives a command a Teams page to work on, plus the close() it must call when
// it is done.
//
// By default that page belongs to the shared browser the daemon keeps up, which
// is attached to over CDP and left running afterwards — booting the SPA is what
// makes a command slow, and doing it once per burst of commands rather than once
// per command is the whole point. Pass daemon: false to launch a browser of this
// command's own instead: that is what the daemon itself does, what manual-login
// needs (it must not attach to a browser that is already signed in), and what
// TEAMS_DAEMON=0 falls back to.
//
// Pass restoreAuth: false when capturing a new session, and args to add
// command-line switches to a browser being launched.
export async function openTeams({
  headless = true, restoreAuth = true, daemon = DAEMON_ENABLED, args = [],
} = {}) {
  if (daemon) {
    // None of these mean anything for a browser this process did not launch, and
    // ignoring them would make restoreAuth: false return the opposite of what it
    // asks for — the daemon's signed-in page. manual-login.mjs is only correct
    // today because it also passes daemon: false; that coupling should not be
    // something the next caller has to know.
    if (!restoreAuth || !headless || args.length) {
      throw new Error(
        'restoreAuth, headless and args only apply to a browser the command launches itself — '
        + 'pass daemon: false to get one.'
      );
    }
    return attachToDaemon();
  }

  const context = await chromium.launchPersistentContext(PROFILE_DIR, {
    headless,
    args,
    viewport: { width: 1400, height: 900 },
  });

  if (restoreAuth) await restoreSession(context);

  const page = context.pages()[0] ?? await context.newPage();

  console.log('Opening Teams...');
  await page.goto(TEAMS_URL, { waitUntil: 'domcontentloaded', timeout: 120000 });

  return { context, page, close: () => context.close() };
}

// The preamble both login scripts share, differing only in whether the browser
// is visible: take the command lock, refuse to run beside a daemon that holds
// the profile, and open a browser of this process's own with no session
// restored. Returns that browser plus the lock's release, which the caller has
// to call when it is done.
//
// Kept here rather than copied into each script because the reasoning below is
// the delicate part, and two copies of it drift.
export async function beginLogin({ headless }) {
  // Held for the whole run. The record alone would miss the case this most needs
  // to catch: a daemon is written down only once it is usable, so for the minute
  // or more it takes to boot, its browser already owns the profile while nothing
  // says so. A command that is starting one holds this lock across the spawn,
  // which is what makes the lock — not the record — the thing to wait on.
  const releaseLock = await acquireCommandLock();

  // The record still has to be consulted for a daemon that finished starting and
  // is now sitting idle: it holds the profile without holding the lock. Asked via
  // connectableDaemon() so that a record left over from before a reboot, whose pid
  // has since been reused, does not refuse a login and point at a --stop that
  // would signal an unrelated process.
  const daemon = await connectableDaemon();
  if (daemon) {
    console.log(`The Teams daemon (pid ${daemon.pid}) is holding the browser profile "${PROFILE_DIR}".`);
    console.log('Stop it first: node teams-daemon.mjs --stop');
    await releaseLock();
    process.exit(1);
  }

  // Neither check covers a teams-daemon.mjs someone started by hand in a third
  // terminal; Chromium's own profile lock is the only backstop there.

  await mkdir(dirname(AUTH_PATH), { recursive: true });

  // No session to restore — this is what the login scripts create. It must be
  // this process's own browser, too: the daemon's is already signed in, which is
  // the opposite of what a login is for.
  const { context, page } = await openTeams({ headless, restoreAuth: false, daemon: false });
  return { context, page, releaseLock };
}

// Attaches to the daemon's browser, starting one if none is running. The command
// lock is taken first and held until close(), so that only one command drives
// the shared page at a time — and, because it is held across the "is a daemon
// there?" check too, so that simultaneous commands do not both try to spawn one.
async function attachToDaemon() {
  const release = await acquireCommandLock();
  let browser = null;
  try {
    await touchActivity();
    browser = await connectToDaemon();
    const context = browser.contexts()[0];
    if (!context) {
      throw new Error('The shared browser has no browser context — it is not the daemon\'s browser.');
    }
    const page = await refreshPage(context);
    return {
      context,
      page,
      // Closing a CDP connection disconnects this client and leaves the browser
      // running — the behaviour this whole design rests on, and verified against
      // the pinned Playwright version rather than assumed. Closing the context
      // here would take the daemon's browser down with the first command.
      close: async () => {
        await browser.close().catch(() => {});
        await touchActivity();
        await release();
      },
    };
  } catch (err) {
    if (browser) await browser.close().catch(() => {});
    await release();
    throw err;
  }
}

// Attaches to the daemon's browser, retrying once with a fresh daemon. A daemon
// that answered the "is one there?" check can still be gone by the time we
// connect — it exited on its idle timeout just as we arrived, or its record
// outlived it. Retrying is what keeps that race from surfacing as a bare
// connection error the user can do nothing with.
async function connectToDaemon() {
  const endpoint = await ensureDaemon();
  try {
    return await chromium.connectOverCDP(endpoint);
  } catch (err) {
    // An externally managed browser is not ours to restart.
    if (CDP_ENDPOINT) throw err;
    console.log('The shared browser did not accept the connection — starting a new one...');
    // Stopped rather than merely forgotten: the recorded process may still be
    // running but unusable, and starting a second daemon on the same profile
    // while it is would leave two browsers writing over the same session.
    const stopped = await stopDaemon();
    // A daemon that ignored SIGTERM is still holding the profile, so the retry
    // below would start that second browser. Said here rather than left to be
    // discovered as whatever the new browser fails with, which names neither the
    // old daemon nor the reason.
    if (!stopped.stopped && stopped.reason === 'timeout') {
      throw new Error(
        `The shared browser is unusable and its daemon (pid ${stopped.pid}) did not stop when asked. `
        + 'Kill it before running another command, so that a second browser is not started on the '
        + 'same profile directory.',
        { cause: err }
      );
    }
    return chromium.connectOverCDP(await ensureDaemon());
  }
}

// Puts the shared page back into the state a freshly launched browser used to
// provide by construction. Everything a command relied on getting for free —
// nothing open, a live session — has to be re-established here, since the
// previous command left the page however it left it and it may have been sitting
// there for hours since.
async function refreshPage(context) {
  const open = context.pages().filter(p => !p.isClosed());
  // Picked by URL rather than by position: on a browser the command had just
  // launched itself, the first page was the Teams tab by construction, which a
  // browser that outlives the command no longer gives us. A consent popup, a
  // "continue in the desktop app" window or a link a previous run followed can
  // sit in front of it — and the reload below would then navigate that popup
  // and leave the signed-in tab open beside it.
  const teams = open.find(p => p.url().startsWith(TEAMS_ORIGIN));
  // Anything else open is reused rather than replaced with a fresh tab, so that
  // a signed-in tab caught mid-redirect through the login origin is navigated
  // back to Teams instead of being abandoned for a second one.
  const page = teams ?? open[0] ?? await context.newPage();
  // Nothing in the daemon ever closes a page, so a stray left here stays for the
  // daemon's whole life and is a candidate again on every command after this.
  for (const stray of open) if (stray !== page) await stray.close().catch(() => {});

  // Whatever the last command left open: a reaction picker, a hover flyout.
  await page.keyboard.press('Escape').catch(() => {});

  if (!teams) {
    console.log('The shared browser has no Teams tab — opening one...');
    await loadTeams(page);
  } else if (!await showsChatList(page)) {
    // A tab that has been alive across a suspend can be signed out or wedged.
    // Reload it once here, so that a stale daemon surfaces as one slow command
    // rather than as a confusing failure deep inside the calling script.
    console.log('The shared browser is not showing the chat list — reloading Teams...');
    await loadTeams(page);
  }

  // Teams keeps the open chat's draft in the compose box, and on a shared page
  // that draft outlives the command that typed it — a --dry-run, or a run
  // interrupted between typing and sending. Escape does not clear a CKEditor, so
  // the next post would be typed onto the end of it and sent as one message.
  const composer = composerLocator(page);
  if (await composer.isVisible().catch(() => false)) await clearComposer(composer);

  return page;
}

// Whether the page has the chat list up right now. A healthy tab has it
// rendered already, so this is a health check rather than a wait for a boot.
function showsChatList(page) {
  return waitForChatList(page, { timeout: HEALTH_CHECK_TIMEOUT_MS }).then(() => true, () => false);
}

async function loadTeams(page) {
  await page.goto(TEAMS_URL, { waitUntil: 'domcontentloaded', timeout: 120000 });
  await waitForChatList(page);
}

// The compose box (a CKEditor contenteditable). Named in one place because both
// the command that types into it and the page reset that empties it need it.
export function composerLocator(page) {
  return page.locator(
    '[data-tid="ckeditor"] [contenteditable="true"], div[role="textbox"][contenteditable="true"], [contenteditable="true"][data-tid="ckeditor"]'
  ).first();
}

// Empties the compose box. Select-all and delete rather than fill(): CKEditor
// keeps its own model, so writing the element's text does not reach it.
export async function clearComposer(composer) {
  await composer.click();
  await composer.press('ControlOrMeta+A');
  await composer.press('Backspace');
}

// Restores the full auth state — cookies plus per-origin localStorage, where
// MSAL keeps its tokens — from the storageState file. The persistent profile
// itself still provides cache/warm-start.
async function restoreSession(context) {
  try {
    const state = JSON.parse(await readFile(AUTH_PATH, 'utf8'));
    if (state.cookies?.length) await context.addCookies(state.cookies);
    for (const { origin, localStorage } of state.origins ?? []) {
      if (!localStorage?.length) continue;
      await context.addInitScript((data) => {
        if (location.origin === data.origin) {
          for (const { name, value } of data.items) {
            try { window.localStorage.setItem(name, value); } catch {}
          }
        }
      }, { origin, items: localStorage });
    }
  } catch (err) {
    // A missing file is the ordinary first-run case. Anything else — corrupt
    // JSON, no read permission, a cookie rejected by addCookies — needs its own
    // message, since running manual-login.mjs would not necessarily fix it and
    // the caller otherwise sits in waitForChatList() until it times out.
    if (err.code === 'ENOENT') {
      console.log(`No saved session at "${AUTH_PATH}" — run manual-login.mjs first.`);
    } else {
      console.log(`Could not restore the session from "${AUTH_PATH}": ${err.message}`);
    }
  }
}

// The SPA is ready once the left rail renders. A generous timeout covers manual
// sign-in (including MFA) when the session still has to be established.
export async function waitForChatList(page, { timeout = 120000 } = {}) {
  await page.getByRole('treeitem').first().waitFor({ state: 'visible', timeout });
}

// Opens the chat whose name matches chatName (partial, case-insensitive) and
// returns the name it resolved to.
//
// atNewest puts the message pane back at the newest messages before returning.
// Callers that walk the history need it; the one that only types into the
// compose box does not, and it is not what makes the chat switch safe — see
// waitForChatOpen below.
export async function openChat(page, chatName, { atNewest = true } = {}) {
  console.log(`Looking for chat: "${chatName}"`);

  // Group headers (e.g. "Favorites", "Chats") are also treeitems that CONTAIN
  // the chat rows, so we must pick the leaf: a matching treeitem that has no
  // nested treeitem inside it.
  const matching = page.getByRole('treeitem').filter({ hasText: new RegExp(escapeRegExp(chatName), 'i') });
  await matching.first().waitFor({ state: 'visible', timeout: 30000 });

  let chatItem = null;
  const count = await matching.count();
  for (let i = 0; i < count; i++) {
    const candidate = matching.nth(i);
    const nested = await candidate.getByRole('treeitem').count();
    if (nested === 0) { chatItem = candidate; break; }
  }
  if (!chatItem) {
    throw new Error(`No leaf chat row matched "${chatName}" (only group headers matched).`);
  }

  const resolvedName = (await chatItem.innerText()).split('\n')[0].trim();
  console.log(`Matched chat: "${resolvedName}"`);
  await chatItem.click();

  await waitForChatOpen(page, resolvedName);

  // Wait for the conversation itself rather than its messages, so that opening
  // an empty chat does not stall. Callers then wait on whatever they actually
  // need — the messages, or the compose box — so there is nothing to sleep for
  // here.
  await page.locator('[data-tid="message-pane-list-viewport"]').first()
    .waitFor({ state: 'visible', timeout: 60000 });

  // Callers that read the history start at the newest messages and work
  // backwards. On a shared page that is not a given: reopening the chat a
  // previous command scrolled far up can leave the pane where that command left
  // it, and reading or reacting from the middle of the history is exactly the
  // kind of intermittent failure that is painful to reproduce.
  if (atNewest) await scrollToNewest(page);

  return resolvedName;
}

// Waits until the chat that was clicked is the one actually on screen.
//
// Clicking a chat row does not swap the conversation synchronously: the previous
// chat's message pane AND its compose box stay mounted and visible while the new
// one loads — measured at roughly 700-900ms on a warm page. So "a message pane is
// visible" and "a compose box is visible" are both true of the chat we are
// leaving, and a caller that acts on them straight after the click reads the
// wrong history, or types a message into the wrong conversation.
//
// The document title names the open chat ("Chat | <name> | Microsoft Teams") and
// changes in the same frame as the pane's messages and the compose box, which
// makes it the positive signal the DOM otherwise does not offer: the rows carry
// no aria-selected, and the panes and boxes of two chats are indistinguishable by
// attribute.
async function waitForChatOpen(page, resolvedName) {
  try {
    await page.waitForFunction(
      (name) => document.title.includes(name),
      resolvedName,
      { timeout: CHAT_SWITCH_TIMEOUT_MS },
    );
  } catch {
    // The title is Teams' to change, and a chat whose row text is not spelled
    // the same way there would otherwise fail every command outright. Falling
    // back to a plain settle is what this did before the title was used at all,
    // so a signal that stops working costs the guarantee, not the feature.
    console.log(
      `The window title ("${await page.title()}") never named "${resolvedName}", so the chat switch `
      + 'could not be confirmed — falling back to a fixed wait.'
    );
    await page.waitForTimeout(PANE_SETTLE_MS);
  }
}

// Jumps to the newest messages at the bottom of the pane.
export async function scrollToNewest(page) {
  await page.evaluate(() => {
    const viewport = document.querySelector('[data-tid="message-pane-list-viewport"]');
    if (viewport) viewport.scrollTop = viewport.scrollHeight;
  });
  await page.waitForTimeout(PANE_SETTLE_MS);
}

// How a message is addressed by its id — the one selector all three scripts
// depend on, so that a Teams rename breaks them in a single place rather than
// one at a time. The string form is for code running inside the page.
export function messageSelector(mid) {
  return `[data-tid="chat-pane-message"][data-mid="${mid}"]`;
}

export function messageLocator(page, mid) {
  return page.locator(messageSelector(mid)).first();
}

// Scrolls the message pane up by roughly a viewport. Returns the scrollTop
// before and after the move plus the pane's scrolling geometry, or null when
// the viewport element could not be found.
export function scrollUp(page) {
  return page.evaluate((fraction) => {
    const viewport = document.querySelector('[data-tid="message-pane-list-viewport"]');
    if (!viewport) return null;
    // Step up by roughly a viewport rather than jumping to the top. The pane is
    // virtualised and only the rendered window is readable, so a jump would
    // skip everything between the old window and the new one.
    const before = viewport.scrollTop;
    viewport.scrollTop = Math.max(0, before - viewport.clientHeight * fraction);
    const overflowY = getComputedStyle(viewport).overflowY;
    return {
      before,
      after: viewport.scrollTop,
      clientHeight: viewport.clientHeight,
      overflowY,
      scrollable: ['auto', 'scroll', 'overlay'].includes(overflowY),
    };
  }, SCROLL_STEP_FRACTION);
}

// The pane's scroll extent and the id of the topmost rendered message — the two
// things that move when older history is added to the list. Null when the
// viewport element could not be found.
function readPaneState(page) {
  return page.evaluate(() => {
    const viewport = document.querySelector('[data-tid="message-pane-list-viewport"]');
    if (!viewport) return null;
    const oldest = document.querySelector('[data-tid="chat-pane-message"]');
    return {
      scrollHeight: viewport.scrollHeight,
      oldestMid: oldest?.getAttribute('data-mid') ?? null,
    };
  });
}

// Waits for an older-history fetch to land while the pane sits at the top of
// the loaded range. Returns true if more history arrived within the grace
// period — scrollTop cannot show this, since it is pinned at 0 either way.
export async function waitForOlderHistory(page) {
  const before = await readPaneState(page);
  if (!before) throw viewportGoneError();

  console.log('At the top of the loaded history — waiting for older messages...');
  const deadline = Date.now() + OLDER_HISTORY_GRACE_MS;
  while (Date.now() < deadline) {
    await page.waitForTimeout(1000);
    const now = await readPaneState(page);
    if (!now) throw viewportGoneError();
    if (now.scrollHeight > before.scrollHeight || now.oldestMid !== before.oldestMid) return true;
  }
  return false;
}

export function viewportGoneError() {
  return new Error(
    'The message pane viewport ([data-tid="message-pane-list-viewport"]) was not '
    + 'found, so the history cannot be scrolled. The Teams DOM has probably changed.'
  );
}

// A pane that cannot scroll at all reports an unchanged scrollTop for every
// step, which reads exactly like the top of the history. Nothing can be inferred
// from such a pane, so callers stop rather than draw that conclusion.
export function paneNotScrollableError({ clientHeight, overflowY }) {
  return new Error(
    'The message pane viewport ([data-tid="message-pane-list-viewport"]) is not a '
    + `scroll container (height ${clientHeight}px, overflow-y "${overflowY}"), so the `
    + 'history cannot be scrolled. The Teams DOM has probably changed.'
  );
}

// Brings a message into the rendered window, scrolling back through the history
// until it appears. Callers start at the newest end and work backwards, so this
// only ever scrolls up. Returns false if the message could not be reached.
//
// maxHistoryWaits caps how often the walk may sit at the top of the loaded range
// waiting for a fetch of older history. A caller hunting for a message it has
// already seen knows the message is above it, not older, so it wants none of
// those waits — hence the default of zero, which makes "not in the loaded
// history" an immediate false rather than a fetch of the whole conversation.
export async function scrollMessageIntoView(page, mid, { maxHistoryWaits = 0 } = {}) {
  const message = messageLocator(page, mid);
  let historyWaits = 0;
  for (let step = 0; step < MAX_SCROLL_STEPS; step++) {
    if (await message.count() > 0) {
      await message.scrollIntoViewIfNeeded().catch(() => {});
      return true;
    }
    const scrolled = await scrollUp(page);
    // A vanished viewport is a broken selector, not a missing message: returning
    // false here would have the caller report a DOM change as a bad message id.
    if (!scrolled) throw viewportGoneError();
    if (!scrolled.clientHeight || !scrolled.scrollable) throw paneNotScrollableError(scrolled);
    // Parked at the top of what is loaded: the pane cannot move any further, but
    // older history may still be on its way, so give the fetch a chance before
    // concluding the message is not there.
    if (scrolled.after === scrolled.before) {
      if (historyWaits++ >= maxHistoryWaits) return false;
      if (!await waitForOlderHistory(page)) return false;
      continue;
    }
    await page.waitForTimeout(1500);
  }
  return false;
}

function escapeRegExp(s) {
  return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}
