import { chromium } from '@playwright/test';
import { mkdir, readFile } from 'node:fs/promises';
import { dirname } from 'node:path';
import {
  CDP_ENDPOINT, DAEMON_ENABLED, acquireCommandLock, connectableDaemon, ensureDaemon, stopDaemon,
  touchActivity,
} from './daemon.mjs';

// Shared plumbing for the Teams scripts: launching a browser with a restored
// session, opening a chat by name, walking back through its history, and the
// message-reaction helpers the two reaction commands work from.

// The browser profile (localStorage, cache) is a persistent directory at
// $TEAMS_PROFILE. Persistent profiles drop session cookies on reopen, so auth
// cookies are additionally stored in the storageState file at $TEAMS_AUTH.
// Both are created/refreshed by manual-login.mjs.
export const PROFILE_DIR = process.env.TEAMS_PROFILE || '.profile';
export const AUTH_PATH = process.env.TEAMS_AUTH || '.auth/user.json';

const TEAMS_URL = 'https://teams.microsoft.com/v2/?ctx=chat';
// The origins a Teams tab's URL can have, for picking it out of the shared
// browser's pages. The path is rewritten by the SPA as it is navigated; the
// origin is not. TEAMS_URL redirects to teams.cloud.microsoft, though, and the
// client has moved between hosts before, so this lists the origins Teams lands
// on rather than the one it is asked for. A hint rather than the last word, for
// the same reason: a tab on a host missing from here is still recognised, by
// asking the page itself — see refreshPage().
const TEAMS_ORIGINS = new Set([
  'https://teams.cloud.microsoft',
  'https://teams.microsoft.com',
]);

// How long the daemon's page gets to show its chat list before it is treated as
// stale and reloaded. A healthy tab has it rendered already, so this only has to
// cover a client that is busy, not one that is booting.
const HEALTH_CHECK_TIMEOUT_MS = 15_000;
// How long a page on an origin TEAMS_ORIGINS does not list gets to prove it is a
// Teams tab after all. Shorter than the health check above, which covers a known
// Teams tab that is merely busy: this one is asked of pages that are usually not
// Teams at all, and every one that isn't waits it out in full before the tab
// that is can be found.
const TEAMS_PROBE_TIMEOUT_MS = 3000;
// How long the message pane gets to settle after being scrolled to the newest
// messages, so the virtualised list has rendered that end before it is read.
const PANE_SETTLE_MS = 2500;
// How long a clicked chat gets to become the one on screen. Generous next to the
// sub-second switch it normally waits for, but not unbounded: this is also what
// a chat whose title is spelled differently from its row would wait out on every
// single command, so it trades a slow worst case against a slow common one.
const CHAT_SWITCH_TIMEOUT_MS = 15_000;
// How long a message pasted into the compose box gets to render there. Running
// it out means Teams stopped honouring the synthetic paste, so the message was
// never composed and there is nothing to send.
const COMPOSER_PASTE_TIMEOUT_MS = 10_000;

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
  // Compared as parsed origins rather than as URL prefixes, so that neither a
  // lookalike host (https://teams.microsoft.com.example/) matches nor a URL with
  // no path fails to. URL.parse() returns null rather than throwing, so an
  // about:blank page is simply not a match.
  let teams = open.find(p => TEAMS_ORIGINS.has(URL.parse(p.url())?.origin));
  // The origin list is a hint, not the only thing that may recognise a Teams
  // tab. It has gone stale before, and a signed-in tab on a host missing from it
  // would not merely be reloaded: with a popup open in front, it is the popup
  // that gets picked below and navigated, while the tab holding the session is
  // closed as a stray. So when no origin matches, ask the pages themselves —
  // showsChatList() answers "is this a working Teams tab?" without reference to
  // the URL. An unknown host then costs a probe of what is open, rather than a
  // reload and a lost session tab.
  // Set when the probe finds the tab, since passing it is also the answer to the
  // health check further down, which is then not asked twice.
  let healthy = false;
  if (!teams) {
    for (const candidate of open) {
      if (await showsChatList(candidate, { timeout: TEAMS_PROBE_TIMEOUT_MS })) {
        teams = candidate;
        healthy = true;
        break;
      }
    }
  }
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
  } else if (!healthy && !await showsChatList(page)) {
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
function showsChatList(page, { timeout = HEALTH_CHECK_TIMEOUT_MS } = {}) {
  return waitForChatList(page, { timeout }).then(() => true, () => false);
}

async function loadTeams(page) {
  await page.goto(TEAMS_URL, { waitUntil: 'domcontentloaded', timeout: 120000 });
  await waitForChatList(page);
}

// The compose box (a CKEditor contenteditable). Named in one place because the
// command that writes a message into it and the page reset that empties it both
// reach for it.
const COMPOSER_SELECTOR =
  '[data-tid="ckeditor"] [contenteditable="true"], div[role="textbox"][contenteditable="true"], [contenteditable="true"][data-tid="ckeditor"]';

export function composerLocator(page) {
  return page.locator(COMPOSER_SELECTOR).first();
}

// Empties the compose box. Select-all and delete rather than fill(): CKEditor
// keeps its own model, so writing the element's text does not reach it.
export async function clearComposer(composer) {
  await composer.click();
  await composer.press('ControlOrMeta+A');
  await composer.press('Backspace');
}

// Puts <text> into the compose box at the caret, as a paste rather than as
// keystrokes. Typing it would leak two of the composer's own behaviours into
// the message: a literal newline arrives as Enter, which sends what has been
// typed so far and makes the rest a second message, and a line starting "- "
// (or "1. ") is turned into a list as it is written, a reflow that swallows the
// keystroke after it. A paste hands the whole string over in one go, so its
// newlines become line breaks within the single message being composed and
// nothing in it is auto-formatted.
// What a link contributes to the composer's text is not what was pasted in:
// Teams makes an anchor of it, and shortens a long label to "https://…/x?query…"
// while the anchor still carries the whole URL. So the wait below compares the
// prose around the links rather than the pasted string itself — otherwise a
// message carrying a long link reads as one that never arrived. A "www." link
// is matched too, since Teams linkifies (and so may shorten) those as well, and
// so is any token holding the ellipsis it shortens with.
const LINK_LIKE = /\S*(?::\/\/|www\.|…)\S*/g;

function proseOf(text) {
  return text.replace(LINK_LIKE, ' ').replace(/\s+/g, ' ').trim();
}

export async function pasteIntoComposer(composer, text) {
  await composer.click();
  await composer.evaluate((el, value) => {
    const data = new DataTransfer();
    data.setData('text/plain', value);
    el.dispatchEvent(new ClipboardEvent('paste', { clipboardData: data, bubbles: true, cancelable: true }));
  }, text);

  // A message of nothing but whitespace renders no text to wait for.
  if (!text.trim()) return;
  // CKEditor takes the paste synchronously but renders it on its own schedule,
  // so whatever the caller does next would be racing that render. What is
  // waited for is <text> itself, not the box merely holding something: a draft
  // clearComposer failed to empty satisfies "not empty" on the first poll, and
  // so does the first character of a multi-line paste that is still rendering.
  // Either would let the caller's Enter send a message other than the one it
  // was given, and report it as sent. Whitespace is normalised on both sides
  // because the composer lays the paste out in paragraphs of its own, and the
  // links are cut out of both because their labels are not what was pasted
  // either — see LINK_LIKE.
  //
  // The element the paste went to is handed to the wait rather than looked up
  // again from inside it: re-resolving the selector on every poll would happily
  // settle on a different box that also matches — the composer of the chat being
  // left, or an inline message-edit field — and assert about the wrong one.
  const expected = proseOf(text);
  const handle = await composer.elementHandle();
  try {
    await composer.page().waitForFunction(
      ({ el, wanted, linkLike }) => {
        const rendered = el.innerText ?? '';
        // An empty box is a paste that did not take. For a message that is
        // nothing but a link there is no prose left to compare, so this is
        // also all that can be checked about one.
        if (!rendered.trim()) return false;
        const prose = rendered.replace(new RegExp(linkLike, 'g'), ' ').replace(/\s+/g, ' ').trim();
        return prose.includes(wanted);
      },
      { el: handle, wanted: expected, linkLike: LINK_LIKE.source },
      { timeout: COMPOSER_PASTE_TIMEOUT_MS },
    );
  } catch (cause) {
    // The friendly wording covers the case worth naming, but it is not the only
    // way out of that wait — a closed page, a detached frame or a mistake in the
    // predicate itself all land here too, and relabelling those would turn a
    // stack trace into a confident wrong diagnosis. The cause carries the real
    // one, including whether the wait timed out or failed at once.
    throw new Error('The message never appeared in the compose box.', { cause });
  } finally {
    await handle.dispose();
  }
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

// Both pane errors are marked systemic: a pane that cannot be walked cannot be
// walked for any message, so a caller working through a list of them stops at
// the first rather than reaching the same verdict once per id.
export function viewportGoneError() {
  return Object.assign(new Error(
    'The message pane viewport ([data-tid="message-pane-list-viewport"]) was not '
    + 'found, so the history cannot be scrolled. The Teams DOM has probably changed.'
  ), { systemic: true });
}

// A pane that cannot scroll at all reports an unchanged scrollTop for every
// step, which reads exactly like the top of the history. Nothing can be inferred
// from such a pane, so callers stop rather than draw that conclusion.
export function paneNotScrollableError({ clientHeight, overflowY }) {
  return Object.assign(new Error(
    'The message pane viewport ([data-tid="message-pane-list-viewport"]) is not a '
    + `scroll container (height ${clientHeight}px, overflow-y "${overflowY}"), so the `
    + 'history cannot be scrolled. The Teams DOM has probably changed.'
  ), { systemic: true });
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

// --- Message reactions -----------------------------------------------------
//
// What react-to-message.mjs and unreact-to-message.mjs share, which is nearly
// everything: the same arguments, the same walk back through the history, the
// same hover toolbar and the same emoji picker. The two differ only in which of
// the picker's buttons they end up clicking and in what they then expect to
// happen to the pill, so the plumbing lives here and each command keeps just
// its half of that difference.

// How often the walk back through the history may pause for a fetch of older
// messages. The target can be arbitrarily far back, so the pauses are needed
// here; the cap only bounds how much of a long conversation one run pages in.
const MAX_HISTORY_WAITS = 100;
// How far the reaction picker is scrolled looking for a button. The emoji list
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
// How long the message itself is given to render after the history walk reports
// it, before anything is read off it.
const MESSAGE_TIMEOUT_MS = 15000;
// How long an already-rendered message is given to bring its reaction row with
// it. A message nobody reacted to has no pills at all, so running out here is an
// ordinary outcome rather than a failure.
const REACTION_SETTLE_MS = 5000;
// How long the message is given to show the change after an emoji is clicked —
// a reaction is only really applied, or really taken back, once the server has
// accepted it.
export const REACTION_TIMEOUT_MS = 15000;
// How long the client is given to sync its emoji catalog: after an incomplete
// one has been dropped, and after a login, before the browser is closed on it.
// The sync runs on its own once the tab is up, so this only bounds how long the
// waiting side gives it.
const EMOJI_CATALOG_SYNC_TIMEOUT_MS = 60000;
// How often the catalog is asked whether that sync has filled it in.
const EMOJI_CATALOG_POLL_MS = 2000;
// How long the catalog has to stay exactly as it is before the sync writing it
// is taken to have finished. Three polls: the categories land within a second
// of one another, so a catalog this still is one nothing is writing to, not one
// caught between two categories.
const EMOJI_CATALOG_SETTLE_MS = 6000;
// How long into a login's wait the sync is taken to have started by, before
// which a still catalog is not taken as a finished one. Stillness can only ever
// be asked of the catalogs that exist when it is asked, and the profile
// directory outlives a re-login: a profile that has already held one account
// carries that account's complete catalog, which sits there perfectly still
// from the very first poll while the catalog the login is actually waiting for
// has not been created yet. Waiting this out first is what keeps a wait from
// settling on the wrong one and closing the browser on the sync it exists to
// protect. Measured at 5.5-8s between the chat list and the first emoji
// records, on a warm profile and a cold one alike, so this is comfortably over
// twice the delay it covers; the cost is that a login whose catalog was already
// complete waits here rather than returning at the first stillness.
//
// The same moment answers the opposite question, being where a catalog that has
// never once been readable stops meaning "the sync has not got to it yet" and
// starts meaning "there is nothing here this code can find" — the storage
// assumptions below having gone stale. Past this point the sync has either
// written something or it is not going to.
const EMOJI_CATALOG_SYNC_START_BY_MS = 20000;
// The catalog's own category for the emoji we reached for most recently. It is
// empty until we have picked one, which is a state of ours rather than a gap in
// the catalog, so it is not counted as one.
const RECENT_EMOJI_CATEGORY_ID = 'recent';
// Where the catalog lives in the profile's IndexedDB: a database whose name
// carries this fragment (the rest of it names the signed-in user), holding the
// list of categories in one store and the emoji themselves in the other. These
// are assumptions about storage that is Teams' own and undocumented, so they
// are the first things to go stale when Teams ships a change to it — which is
// why they are named here together rather than written out where they are read.
// They were read off a signed-in web client in September 2026; nothing warns
// when they stop matching, so a catalog check that has started reporting
// "could not be read" is the sign to look here first.
const EMOJI_DB_NAME_FRAGMENT = 'emoji-manager';
const EMOJI_METADATA_STORE = 'teams-emoji-metadata';
const EMOJI_STORE = 'teams-emoji';

// The message ids a command was given: one id, or several as a comma-separated
// list. Blank entries — a trailing or a doubled comma — are dropped rather than
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

// The ids of a run in the order they should be worked through: newest first,
// whatever order they were given in. The history walk only ever scrolls back,
// so every target after the first is then older than where the pane already
// stands, and one walk carries on through the whole list instead of each id
// sending it back to the newest messages.
//
// The ids are epoch milliseconds, so ordering them numerically orders them in
// time; an id that is not a plain number carries no such order, so a list
// holding one is left exactly as it was given.
function orderNewestFirst(messageIds) {
  return messageIds.every(id => /^\d+$/.test(id))
    ? [...messageIds].sort((a, b) => Number(b) - Number(a))
    : messageIds;
}

// A find(mid) that brings messages into the rendered window, for a command
// working through a list of them. The pane opens at the newest messages and the
// walk only ever goes back, so the finder remembers where it left the pane:
// targets handed to it newest first are then reached by one walk carrying on
// rather than by a walk per id.
//
// A target that is neither rendered nor older than the pane lies ahead of it,
// and that direction is only reachable from the newest end, so a walk that came
// up empty is tried once more from there. A walk that started at the newest end
// has already seen the whole history, so it is not repeated — an id that
// belongs to another chat costs one walk, not two.
export function createMessageFinder(page, { maxHistoryWaits = MAX_HISTORY_WAITS } = {}) {
  // Where openChat leaves the pane, and the only thing the walks below move
  // away from.
  let paneAtNewest = true;

  return async function findMessage(mid) {
    const startedAtNewest = paneAtNewest;
    // Either walk may leave the pane part way back through the history.
    paneAtNewest = false;

    if (await scrollMessageIntoView(page, mid, { maxHistoryWaits })) return true;
    if (startedAtNewest) return false;

    console.log(`Message ${mid} is not behind the pane — looking again from the newest messages...`);
    await scrollToNewest(page);
    return scrollMessageIntoView(page, mid, { maxHistoryWaits });
  };
}

// Works through a list of message ids, one at a time. act(mid) does whatever
// the command does to one message and returns whether that changed anything; it
// reports each message itself, since only it knows what to call the outcome.
// What is kept here is the bookkeeping both commands need — one message that
// cannot be reached must not cost the rest of the list its turn — plus the
// tally and the failure they raise at the end.
//
// The words are passed in because the commands speak of different things: "2
// reacted, 1 already reacted" against "2 removed, 1 not reacted".
export async function actOnMessages(page, messageIds, act, { changed, unchanged, couldNot, chatName }) {
  const orderedIds = orderNewestFirst(messageIds);
  let changedCount = 0;
  let unchangedCount = 0;
  const failures = [];

  for (const [index, mid] of orderedIds.entries()) {
    try {
      if (await act(mid)) changedCount++;
      else unchangedCount++;
    } catch (err) {
      // One unreachable message must not cost the rest of the list its turn, so
      // what went wrong is kept and the run moves on. The collected failures
      // are raised together once the list is done.
      console.log(`Could not ${couldNot} message ${mid}: ${err.message}`);
      failures.push(err);
      // Unless the failure was never about this message: the same verdict
      // awaits every id left, each after another full walk back through the
      // history, so the run stops here and reports what it has rather than
      // proving the same thing over and over.
      if (isSystemic(page, err)) {
        if (index < orderedIds.length - 1) {
          console.log('This says nothing about the remaining messages either — stopping here.');
        }
        break;
      }
    }
  }

  if (messageIds.length > 1) {
    // A run stopped by a systemic failure leaves ids it never looked at, and
    // counting them as anything else would misreport what happened to them.
    const notAttempted = messageIds.length - changedCount - unchangedCount - failures.length;
    console.log(
      `${messageIds.length} message(s): ${changedCount} ${changed}, ${unchangedCount} ${unchanged}, `
      + `${failures.length} failed` + (notAttempted ? `, ${notAttempted} not attempted.` : '.')
    );
  }
  // A single-id run has nothing to aggregate: the wrapper would make its
  // headline a count and push the sentence saying what to do about it into the
  // [errors] array underneath. Raised unchanged, it reads as it always did.
  if (messageIds.length === 1 && failures.length === 1) throw failures[0];
  if (failures.length) {
    throw new AggregateError(
      failures,
      `Could not ${couldNot} ${failures.length} of ${messageIds.length} message(s) in "${chatName}".`
    );
  }
}

// Whether a failure was about the run rather than about the message it happened
// on — the emoji is not in the picker, the pane cannot be walked at all, the
// browser is gone. The scripts mark such errors as they raise them; a page that
// has closed under us is the same verdict arrived at from the other side.
function isSystemic(page, err) {
  return err?.systemic === true || page.isClosed();
}

// A short "author: body" line, so a command says which message it acted on
// rather than only echoing back the id.
export function describeMessage(page, mid) {
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

// Gives a freshly rendered message time to render its reaction row, so that the
// pills can be read off it. The pane is virtualised, so a message may have been
// mounted a moment ago with its reactions still to come — and since reacting is
// a toggle, a command that read the pills too early would click the very
// reaction it meant to leave alone. A message nobody has reacted to never grows
// a row, so this is a settle rather than a requirement.
export async function settleReactions(message) {
  await message.waitFor({ state: 'visible', timeout: MESSAGE_TIMEOUT_MS });
  await message.locator('[data-tid="diverse-reaction-pill-button"]').first()
    .waitFor({ state: 'visible', timeout: REACTION_SETTLE_MS })
    .catch(() => {});
}

// The message's <emoji> pills that we left ourselves. Reactions are stored per
// person and clicking one we already left takes it back, so whether one of them
// is ours is what both commands turn on. The client marks ours as pressed —
// an attribute, so it says the same in whatever language the client is in.
export function ownReactionPills(page, message, emoji) {
  return message.locator('[data-tid="diverse-reaction-pill-button"][aria-pressed="true"]')
    .filter({ has: emojiImage(page, emoji) });
}

// Matches the <img> Teams renders an emoji as. Its alt is the emoji character,
// which may or may not carry the U+FE0F variation selector — the same emoji
// either way, so both spellings are accepted.
export function emojiImage(page, emoji) {
  const bare = emoji.replace(/\uFE0F/g, '');
  return page.locator(`img[alt="${bare}"], img[alt="${bare}\uFE0F"]`);
}

// Waits for the client to finish writing its emoji catalog, so that a browser
// closed on it moments later does not cut the sync in half.
//
// The chat list rendering is not the client having finished loading. The
// catalog is fetched into IndexedDB seconds after the rail is up: on a run
// measured here the first emoji landed 5.5s later and the last of them half a
// second after that, and in between the catalog listed all ten categories while
// two of them still held nothing. A browser closed anywhere in that window
// leaves behind precisely the catalog ensureEmojiCatalog() below then has to
// drop and fetch again, which is why the login scripts wait here rather than
// leaving the next reaction command to find it.
//
// What is waited for is the catalog going still, not its categories all being
// filled — the stronger of the two, and the cheaper. Stronger because a sync
// stopped part way through a category leaves every category non-empty while
// still being unfinished, which nothing about the catalog's contents can see
// (readEmojiCatalog below says why). Cheaper because a category that is empty
// for good — one the tenant has uploaded no emoji of — settles like any other,
// where waiting for it to fill would sit out the whole budget on every login.
//
// Stillness on its own is not enough, though, because it is only ever asked of
// the catalogs that are already there. A profile that has held another account
// still holds that account's finished catalog, which is still from the first
// poll onwards, so the wait is floored as well as bounded: nothing counts as
// settled until the sync has had long enough to have created its database.
//
// Nothing here raises. A login that reached the chat list has done what it was
// for, and a catalog that is short is repairable later by the commands that
// care about it; all this can do is spare them the trouble, so it says how it
// went and returns either way.
export async function waitForEmojiCatalog(page) {
  console.log('Waiting for the emoji catalog to finish syncing...');
  const started = Date.now();
  const deadline = started + EMOJI_CATALOG_SYNC_TIMEOUT_MS;
  const syncStartedBy = started + EMOJI_CATALOG_SYNC_START_BY_MS;
  let previous = null;
  let unchangedSince = started;
  let everRead = false;
  for (;;) {
    // The window being watched is one a person can close by hand, and a browser
    // that is gone is not a sync to wait for. Left quietly rather than
    // reported: closing the window is the user ending the login, not a fault,
    // and raising into the login's error path would say otherwise.
    if (page.isClosed()) return;

    const catalog = await readEmojiCatalog(page);
    if (catalog) everRead = true;
    // Compared as text because it is only ever asked whether the counts are the
    // ones from the poll before, never how they differ. Sorted first so that
    // what is compared is the counts rather than the order they arrived in:
    // the keys come out in whatever order indexedDB.databases() lists the
    // databases, which nothing specifies, and a profile holding more than one
    // catalog would otherwise be able to restart the settle window on a
    // reshuffle alone — for the whole budget, without a single count moving.
    const counts = catalog && JSON.stringify(Object.entries(catalog.emojiCounts).sort());
    // A catalog that cannot be read yet is the ordinary state of the first
    // seconds after a login — Teams has not created the database — and a
    // catalog that has just grown is one being written to. Neither is stillness,
    // so both start the settle window over instead of ending the wait.
    if (counts !== null && counts === previous) {
      // Both bars have to be cleared, and they answer different questions.
      // The settle window says nothing is writing to the catalogs that are
      // here; the floor says the one this login is waiting for is among them,
      // which on a reused profile the settle window alone cannot tell (the
      // constant says why). Only the floor is measured from the start of the
      // wait — a catalog that goes still late still gets its settle window in
      // full, so this delays a return and never brings one forward.
      const settled = Date.now() - unchangedSince >= EMOJI_CATALOG_SETTLE_MS;
      if (settled && Date.now() >= syncStartedBy) {
        const total = Object.values(catalog.emojiCounts).reduce((sum, count) => sum + count, 0);
        if (!catalog.emptyCategories.length) {
          console.log(`The emoji catalog is complete (${total} emoji).`);
        } else {
          // Said rather than repaired: this is a login, and dropping the
          // catalog it just fetched to fetch it once more is a poor answer to a
          // gap that may not be one. The commands that need the picker check it
          // again for themselves.
          console.log(
            `The emoji catalog has settled at ${total} emoji, but has none for `
            + `${catalog.emptyCategories.join(', ')}. Either the tenant has uploaded no emoji of `
            + 'those kinds, or the sync stopped short; a reaction command will drop the catalog '
            + 'and fetch it again if it finds the gap.'
          );
        }
        return;
      }
    } else {
      previous = counts;
      unchangedSince = Date.now();
    }
    // A catalog that has not been readable once by the time the sync is due to
    // have started is not a slow sync — it is a catalog this code can no longer
    // find, and the rest of the budget cannot change that. Given the shorter
    // deadline of the two so that the day the storage assumptions go stale is
    // reported while the user is still watching, rather than as a silent minute
    // on every login that the "waiting for the sync" line above misreads as the
    // sync being slow.
    if (!everRead && Date.now() >= syncStartedBy) {
      console.log(
        `${EMOJI_CATALOG_SYNC_START_BY_MS / 1000}s on there is still no emoji catalog to read — `
        + 'carrying on. Teams has either not started the sync or has moved the catalog, which is '
        + 'a difference a reaction command reports.'
      );
      return;
    }
    // Only ever a catalog that was read and went on changing: one that was
    // never read has returned above, a good forty seconds earlier.
    if (Date.now() >= deadline) {
      console.log(
        `The emoji catalog was still being written ${EMOJI_CATALOG_SYNC_TIMEOUT_MS / 1000}s on `
        + '— carrying on rather than waiting longer, so it may be left short. A reaction '
        + 'command drops a short catalog and fetches it again.'
      );
      return;
    }
    // Swallowed for the same reason as the check at the top of the loop, which
    // is what a page closed during the pause lands on next.
    await page.waitForTimeout(EMOJI_CATALOG_POLL_MS).catch(() => {});
  }
}

// Makes sure the client has the whole emoji catalog before the picker is asked
// for anything, dropping an incomplete one so that Teams syncs it again.
//
// The catalog lives in IndexedDB inside the persistent profile, and Teams fills
// it in once. A sync that was cut short — the browser closed part way through
// one — leaves behind a catalog the client treats as finished and never returns
// to: the picker lists every category, but the ones the sync never reached
// render as empty headings and no amount of scrolling finds an emoji in them.
// "The emoji is not in the reaction picker" is then perfectly true and says
// nothing about the emoji that was asked for, which is the failure this is here
// to keep a command from reporting.
//
// Deleting the catalog is safe in the way clearing any cache is: it is Teams'
// copy of a list Teams serves, and the reload below is what makes it fetch the
// list again.
export async function ensureEmojiCatalog(page) {
  const before = await readEmojiCatalog(page);
  // Said out loud rather than passed over in silence: a profile whose catalog
  // cannot be read is one where this check is off for good, and the reaction
  // commands' "not in the reaction picker" error names that case as the one
  // thing it can still be. A reader who was told nothing has no way to know
  // they are in it, and will as readily conclude that the check ran and passed.
  if (before === null) {
    console.log(
      "The profile's emoji catalog could not be read, so whether it is complete cannot be told — "
      + 'carrying on. Teams has most likely moved it, in which case an emoji missing from the '
      + 'picker will be reported as missing with nothing done about it.'
    );
    return;
  }
  if (!before.emptyCategories.length) return;

  const missing = before.emptyCategories;
  console.log(
    `The profile's emoji catalog has no emoji for ${missing.length} of its categories `
    + `(${missing.join(', ')}), so an emoji from those cannot be found in the picker. `
    + 'Dropping the catalog and reloading Teams to sync it again...'
  );
  // Every catalog in the profile goes, not just the one the gap was found in:
  // which of them the client is using cannot be told from the outside, and one
  // left behind is one the check keeps finding the same gap in, run after run.
  //
  // The deletion is asked for and not waited on, deliberately: the app is
  // holding the database open, so the request reports "blocked" and stays
  // pending until something closes that connection — which is the reload on the
  // next line. Waiting for it to finish first would be waiting for a reload
  // that has not happened yet.
  await page.evaluate(
    names => { for (const name of names) indexedDB.deleteDatabase(name); },
    await emojiDatabaseNames(page),
  );
  await loadTeams(page);

  // The sync runs by itself from here; all that is left is to give it a moment,
  // since a command that walked the picker before it finished would draw the
  // same wrong conclusion as before.
  const deadline = Date.now() + EMOJI_CATALOG_SYNC_TIMEOUT_MS;
  for (;;) {
    // A catalog that cannot be read means the opposite here of what it meant
    // above. Before the delete it meant there was nothing to repair; after it,
    // it means the database Teams is going to recreate is not there yet — the
    // ordinary state of the first seconds after the reload — so it keeps the
    // loop waiting rather than ending it.
    const catalog = await readEmojiCatalog(page);
    if (catalog && !catalog.emptyCategories.length) {
      console.log('The emoji catalog is complete again.');
      return;
    }
    if (Date.now() >= deadline) {
      // Carried on with rather than raised: the picker is about to be walked
      // anyway, and it — not this — is what decides whether the emoji that was
      // actually asked for is there.
      //
      // Which of the two things went wrong is worth telling apart, because the
      // answers differ. A database still holding exactly the emoji it held
      // before is one we asked to have deleted, still there: a
      // deleteDatabase() stays blocked until every connection to it is closed,
      // and the reload only closes the ones the page itself holds — a service
      // worker's outlives it. Waiting longer would not have helped, so saying
      // "not finished syncing" would send the reader the wrong way. One such
      // database is enough to say so, whatever became of the others.
      const survivor = Object.entries(catalog?.emojiCounts ?? {})
        .find(([name, count]) => before.emojiCounts[name] === count);
      if (survivor) {
        console.log(
          `The emoji catalog still holds the same ${survivor[1]} emoji it did before the `
          + 'reload, so the deletion never went through — something outside the page, such as a '
          + 'service worker, is still holding the database open. Carrying on, but the gap is '
          + 'unrepaired: stopping the browser daemon (node teams-daemon.mjs --stop) and running '
          + 'the command again closes every connection there is.'
        );
      } else if (catalog?.emptyCategories.length) {
        // Which of the two this is cannot be told from here, so neither is
        // asserted: after a reload and a full minute, "still empty" is at least
        // as likely to mean a category with nothing to put in it as a sync that
        // is behind, and the categories are named so the reader can judge.
        console.log(
          `${EMOJI_CATALOG_SYNC_TIMEOUT_MS / 1000}s after the reload the emoji catalog still has `
          + `no emoji for ${catalog.emptyCategories.join(', ')} — carrying on. Either the sync has `
          + 'not got that far, in which case an emoji from those will still be reported as missing '
          + 'from the picker, or they are empty because there is nothing to put in them: a '
          + 'category the tenant has uploaded no emoji of is listed by the catalog and stays empty '
          + 'for good.'
        );
      } else {
        console.log(
          `The emoji catalog could not be read ${EMOJI_CATALOG_SYNC_TIMEOUT_MS / 1000}s after the `
          + 'reload — carrying on, but Teams has not brought it back yet, so whether the gap was '
          + 'repaired is unknown.'
        );
      }
      return;
    }
    await page.waitForTimeout(EMOJI_CATALOG_POLL_MS);
  }
}

// The profile's emoji catalogs, by database name. The check and the repair both
// work from this rather than each recognising a catalog for itself: they have to
// agree on which databases they are talking about, and one substring match is
// easier to keep true to Teams than two copies of it.
function emojiDatabaseNames(page) {
  return page.evaluate(
    fragment => indexedDB.databases().then(dbs => dbs.map(db => db.name).filter(n => n?.includes(fragment))),
    EMOJI_DB_NAME_FRAGMENT,
  );
}

// What the profile's emoji catalogs hold: the titles of the categories with no
// emoji in them, and the number of emoji they hold altogether. Null when no
// catalog can be read at all — none has been created yet, or Teams has moved
// it, in which case there is nothing to conclude and nothing to repair.
//
// The count is carried alongside the titles because it is the only way to tell,
// after a repair, whether the catalog on screen is the new one or the old one
// the delete request never got to remove.
//
// Every catalog in the profile is read, not just one. Teams names the database
// after the signed-in user and the profile directory outlives a re-login, so a
// profile that has ever held a second account carries more than one of them,
// and which one the client is using cannot be told from the name. Reporting a
// category that is empty in any of them drives the repair from the worst copy;
// picking one of them instead would as easily read a stale account's healthy
// catalog and leave the live one broken, silently and for good.
//
// Categories are compared rather than counted because the catalog's size is
// Teams' business and moves with their emoji set, while a category the catalog
// itself names and then has nothing for is wrong however large the rest of it
// is.
//
// `recent` is the one category known to be empty for reasons of its own, and it
// is unlikely to be the only one — a category the tenant has uploaded no emoji
// of is listed here and legitimately holds nothing, and reads from here exactly
// like one a sync missed. There is no telling the two apart short of knowing
// which categories Teams ships as built-in, so such a category costs a repair
// that cannot help it, on every run, for as long as it stays empty. What is
// avoided is claiming otherwise: the message the repair's wait ends with names
// both readings rather than asserting the sync is behind.
//
// That makes a category with nothing in it at all the only gap this recognises,
// and a lower bound on the damage: nothing here knows how many emoji a category
// is supposed to hold, since the metadata is read for the names of the
// categories and nothing else. A sync that stopped in the middle of a category
// rather than between two therefore leaves it looking as finished as any other,
// and an emoji from the part it never reached is still reported as missing from
// the picker with nothing said beforehand. Closing that would take a
// per-category expected count to compare against, if the metadata record turns
// out to carry one.
async function readEmojiCatalog(page) {
  const names = await emojiDatabaseNames(page).catch(() => []);
  if (!names.length) return null;

  return page.evaluate(async ({ names, recentId, metadataStore, emojiStore }) => {
    const readAll = (store) => new Promise((resolve, reject) => {
      const request = store.getAll();
      request.onsuccess = () => resolve(request.result);
      request.onerror = () => reject(request.error);
    });

    // Which categories the emoji store holds a record for, and how many records
    // that is. Walked with a cursor rather than read with getAll(), which would
    // materialise every emoji whole — keywords, shortcodes and all — for the
    // one field on it that is looked at: this runs at the start of every
    // reaction command, and again every couple of seconds throughout a repair,
    // when the client is busy writing that same store back.
    const scanEmoji = (store) => new Promise((resolve, reject) => {
      const filled = new Set();
      let count = 0;
      const request = store.openCursor();
      request.onsuccess = () => {
        const cursor = request.result;
        if (!cursor) return resolve({ filled, count });
        filled.add(cursor.value.categoryId);
        count += 1;
        cursor.continue();
      };
      request.onerror = () => reject(request.error);
    });

    const readCatalog = async (name) => {
      const db = await new Promise((resolve, reject) => {
        const request = indexedDB.open(name);
        request.onsuccess = () => resolve(request.result);
        request.onerror = () => reject(request.error);
        // Only ever reached if the database went away between being listed and
        // being opened, in which case this call has just created an empty one.
        // Nothing can be said about a catalog we made up ourselves.
        request.onupgradeneeded = () => { request.transaction.abort(); resolve(null); };
      }).catch(() => null);
      if (!db) return null;

      try {
        const stores = [metadataStore, emojiStore];
        if (!stores.every(store => [...db.objectStoreNames].includes(store))) return null;

        // Both reads are asked for before either is awaited: an IndexedDB
        // transaction commits as soon as control returns to the event loop with
        // nothing outstanding on it, so a second request issued after awaiting
        // the first would land on a transaction that has already closed.
        const transaction = db.transaction(stores, 'readonly');
        const [metadata, emoji] = stores.map(store => transaction.objectStore(store));
        const [metadataRecords, { filled, count }] = await Promise.all([
          readAll(metadata), scanEmoji(emoji),
        ]);

        // The record that carries the categories is looked for rather than
        // taken to be the first one. getAll() hands records back in key order,
        // so "the first" is the right record only for as long as the store
        // holds nothing besides it, which nothing here guarantees: a
        // schema-version row or a second locale's list appearing alongside it
        // would otherwise read the wrong record, find no categories on it and
        // quietly switch the whole check off.
        const categories = metadataRecords.find(record => Array.isArray(record?.categories))?.categories;
        if (!categories?.length) return null;

        return {
          empty: categories.filter(c => c.id !== recentId && !filled.has(c.id)).map(c => c.title ?? c.id),
          count,
        };
      } finally {
        db.close();
      }
    };

    // A catalog that could not be read drops out here rather than counting as
    // one with nothing missing, so "no catalog at all" stays distinguishable
    // from "every catalog is fine".
    const catalogs = (await Promise.all(names.map(async (name) => {
      const catalog = await readCatalog(name);
      return catalog && { name, ...catalog };
    }))).filter(catalog => catalog !== null);
    if (!catalogs.length) return null;

    return {
      emptyCategories: [...new Set(catalogs.flatMap(catalog => catalog.empty))],
      // Kept per database rather than summed, so that "this one was not
      // deleted" stays askable of each of them. A total says only that
      // something changed somewhere, which in a profile holding several
      // catalogs is exactly what a blocked deletion looks like: the databases
      // no client has open are deleted cleanly and take the total down with
      // them, while the one that matters sits there untouched.
      emojiCounts: Object.fromEntries(catalogs.map(catalog => [catalog.name, catalog.count])),
    };
  }, {
    names,
    recentId: RECENT_EMOJI_CATEGORY_ID,
    metadataStore: EMOJI_METADATA_STORE,
    emojiStore: EMOJI_STORE,
  }).catch(() => null);
}

// Clicks one of the reaction picker's emoji on a message, and leaves the picker
// closed whichever way that goes. This is everything the two commands do to a
// message once they have decided it needs something done to it, and they do it
// in the same eight steps; what differs is only which button they are after and
// what they then expect to happen to the pill, which is what they pass in:
//
//   buttons(picker)  the candidates, narrowed from the open picker
//   notInPicker()    the error to raise when the picker holds none of them
//   stillNeeded()    whether the click is still wanted, asked after the walk
//                    through the picker and before the click itself
//   confirm()        waits for the message to show the change, and says what it
//                    means if it never does
//
// Returns whether a button was clicked — false when stillNeeded() has changed
// its mind in the meantime.
//
// Kept whole here rather than written out in each command because the picker is
// a modal popup and the protocol around it is the delicate part: two copies of
// it would have to be kept in step by hand, and a divergence would leave the
// popup standing over the message pane for the rest of the list.
export async function clickPickerButton(page, message, mid, { buttons, notInPicker, stillNeeded, confirm }) {
  let pickerOpen = false;
  try {
    const picker = await openReactionPicker(page, message, mid);
    pickerOpen = true;

    const button = await findPickerButton(page, buttons(picker));
    if (!button) throw notInPicker();
    // Opening the picker and walking it gave the message plenty of time to
    // finish rendering, so what the caller turns on is read off it once more
    // right before the click. Reacting is a toggle, so a click that is no
    // longer wanted does not amount to nothing — it does the opposite of what
    // was asked for.
    if (!await stillNeeded()) return false;

    await button.click();
    // Clicking an emoji closes the picker, so from here on there is nothing
    // left to dismiss.
    pickerOpen = false;

    await confirm();
    return true;
  } finally {
    // The picker is a modal popup: left open it covers the message pane, and
    // the next message cannot even be hovered through it — one message that
    // went wrong would cost the rest of the list its turn for a reason of its
    // own making. Only the ways out that leave it standing are dismissed, so
    // that a run that went well sends no stray keystroke into the chat. A
    // dismissal that itself fails must not replace the failure that led here.
    if (pickerOpen) await page.keyboard.press('Escape').catch(() => {});
  }
}

// Opens the message's reaction picker and hands it back ready to be searched.
// The picker is a modal popup: from the moment this returns, every way out of
// the caller has to close it, or the message pane stays covered and the next
// message cannot even be hovered.
async function openReactionPicker(page, message, mid) {
  const actions = await openMessageActions(page, message, mid);

  // Forced past the actionability check on purpose: for a message at the top of
  // the pane the toolbar renders under the pinned-message banner, which sits on
  // top of the "More reactions" button and intercepts the click. The button is
  // the right target — it is just visually overlapped — so the receives-events
  // check is the wrong guard here and would only time out.
  await actions.locator('[data-tid="expanded-reactions-picker-entry"]').click({ force: true });

  const picker = page.locator('[data-tid="reaction-picker-root"]');
  try {
    // The picker's frame appears before its emoji do, so wait for the list
    // itself — searching it while it is still empty would find nothing.
    await picker.locator('[data-tid^="emoticon-button-"]').first()
      .waitFor({ state: 'visible', timeout: PICKER_TIMEOUT_MS });
  } catch (err) {
    // The frame is up even though its list never arrived, and this call is
    // about to fail rather than return — so the caller will not know there is
    // anything to dismiss, and it has to be dismissed here.
    await page.keyboard.press('Escape').catch(() => {});
    throw err;
  }
  return picker;
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

// The picker's clickable emoji buttons, which callers narrow to the one they
// are after. The picker keeps hidden copies of its grids in the DOM and those
// cannot be clicked, so only what is actually rendered counts as a match.
export function pickerButtons(picker) {
  return picker.locator('[data-tid^="emoticon-button-"]:visible');
}

// Finds a button in the open picker, scrolling the list until one turns up.
// The caller says what it is looking for; the locator is asked again at every
// step, so it picks up whatever that step has just rendered. Returns the first
// match, or null if the whole list was walked without one.
async function findPickerButton(page, buttons) {
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
