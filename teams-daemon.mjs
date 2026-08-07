import { readFile, rm } from 'node:fs/promises';
import { join } from 'node:path';
import { openTeams, waitForChatList, PROFILE_DIR } from './teams.mjs';
import {
  CDP_ENDPOINT, DAEMON_PORT, IDLE_TIMEOUT_MS, LOG_PATH, clearInfo, connectableDaemon, delay,
  fetchBrowserVersion, isOurBrowser, lastActivity, readInfo, stopDaemon, touchActivity,
  tryCommandLock, writeInfo,
} from './daemon.mjs';

// Usage:
//   nix develop .#playwright --command node teams-daemon.mjs [--headed]
//   nix develop .#playwright --command node teams-daemon.mjs --stop
//   nix develop .#playwright --command node teams-daemon.mjs --status
//
// Holds a signed-in Teams browser open so the other scripts do not each pay for
// a cold boot of the SPA. They start it themselves when they need it, so running
// it by hand is only for debugging, for --status, and for stopping it.
//
// It exits on its own after TEAMS_DAEMON_IDLE minutes without a command, so a
// burst of commands costs one boot and nothing keeps running overnight. Stop it
// explicitly before manual-login.mjs: two browsers on one profile directory
// write over each other's stored session, and nothing at the browser level
// reliably refuses that — so it is refused here instead.

// How often the daemon checks whether it has been idle long enough to exit.
const IDLE_CHECK_INTERVAL_MS = 30_000;
// How long the browser gets to write down the debugging port it bound to.
const PORT_FILE_TIMEOUT_MS = 10_000;

const args = process.argv.slice(2);
const known = ['--stop', '--status', '--headed'];
const unknownFlag = args.find(a => !known.includes(a));
if (unknownFlag) {
  console.log(`Unknown option "${unknownFlag}".`);
  console.log('Usage: node teams-daemon.mjs [--headed] | --stop | --status');
  process.exit(1);
}

try {
  if (args.includes('--stop')) await stop();
  else if (args.includes('--status')) await status();
  else await start({ headless: !args.includes('--headed') });
} catch (err) {
  // This process's output is a log file that a command reads back when it could
  // not start a daemon, so what lands in it has to be the reason rather than a
  // stack trace with the reason somewhere in it.
  console.log(err.message);
  for (let cause = err.cause; cause; cause = cause.cause) console.log(`Caused by: ${cause.message}`);
  process.exit(1);
}

async function start({ headless }) {
  if (CDP_ENDPOINT) {
    console.log(`TEAMS_CDP is set to "${CDP_ENDPOINT}", so the commands attach to that browser and`);
    console.log('there is no daemon to start here.');
    process.exit(1);
  }

  const running = await connectableDaemon();
  if (running) {
    console.log(`The Teams daemon is already running (pid ${running.pid}) on ${running.httpEndpoint}.`);
    return;
  }
  await clearInfo();

  console.log(`--- Teams daemon starting at ${new Date().toISOString()} (pid ${process.pid}) ---`);

  // Chromium writes its debugging port into DevToolsActivePort, next to the
  // profile. Reading it back is the only way to learn the port when it is left
  // to pick a free one, which is the default. It is not proof that the port was
  // bound, though — see readEndpoint — and it can be left over from a previous
  // browser, hence the delete first.
  const portFile = join(PROFILE_DIR, 'DevToolsActivePort');
  await rm(portFile, { force: true });

  const { context, page } = await launch({ headless });

  let endpoint;
  try {
    console.log('Waiting for the chat list...');
    await waitForChatList(page);
    endpoint = await readEndpoint(portFile);
  } catch (err) {
    // Nothing will ever use this browser, and leaving it up would have it hold
    // the profile against the next attempt.
    await context.close().catch(() => {});
    throw err;
  }

  // A daemon that has just started has not been idle, so the clock starts here
  // rather than at whatever a previous daemon left behind.
  const startedAt = Date.now();
  await touchActivity();
  // Written last, once the browser is genuinely usable: the commands take the
  // existence of this record as "there is a daemon to attach to".
  await writeInfo({ pid: process.pid, startedAt: new Date().toISOString(), ...endpoint });
  console.log(`Ready on ${endpoint.httpEndpoint} — attach with node post-message.mjs and friends.`);

  let shuttingDown = false;
  for (const signal of ['SIGINT', 'SIGTERM']) {
    process.on(signal, () => { void shutdown(`on ${signal}`); });
  }

  if (IDLE_TIMEOUT_MS > 0) {
    console.log(`Will exit after ${IDLE_TIMEOUT_MS / 60_000} idle minute(s).`);
    setInterval(() => { void checkIdle(); }, IDLE_CHECK_INTERVAL_MS);
  } else {
    // Nothing else holds the event loop open on purpose, and a permanently
    // connected client shows you as Available for as long as it runs, so this is
    // worth saying out loud.
    console.log('TEAMS_DAEMON_IDLE=0 — staying up until stopped. You will show as Available meanwhile.');
    setInterval(() => {}, IDLE_CHECK_INTERVAL_MS);
  }

  async function checkIdle() {
    if (shuttingDown) return;
    if (Date.now() - await idleSince() < IDLE_TIMEOUT_MS) return;

    // A command that is running, or about to attach, holds the command lock.
    // Taking it before exiting is what makes "the daemon went away underneath
    // me" impossible rather than merely unlikely.
    const releaseLock = await tryCommandLock();
    if (!releaseLock) return;
    if (Date.now() - await idleSince() < IDLE_TIMEOUT_MS) {
      await releaseLock();
      return;
    }
    await shutdown(`after ${IDLE_TIMEOUT_MS / 60_000} idle minute(s)`, releaseLock);
  }

  // Falls back to this daemon's own start time, so that a missing activity file
  // still lets it exit eventually rather than leaving it up — and showing you as
  // Available — indefinitely.
  async function idleSince() {
    return await lastActivity() ?? startedAt;
  }

  async function shutdown(reason, releaseLock = null) {
    if (shuttingDown) return;
    shuttingDown = true;
    console.log(`Shutting down ${reason}.`);
    // The record goes first: a command arriving now has to see "no daemon" and
    // start its own rather than attach to one that is on its way out.
    await clearInfo();
    // Awaited rather than killed, so the persistent profile is flushed. A
    // browser cut off mid-write leaves a stale lock behind in the profile
    // directory that blocks the next start until it is cleared by hand.
    await context.close().catch(err => console.log(`The browser did not close cleanly: ${err.message}`));
    if (releaseLock) await releaseLock();
    console.log(`--- Teams daemon stopped at ${new Date().toISOString()} ---`);
    process.exit(0);
  }
}

// Launches the browser the daemon owns. A launch that fails here is most often
// the profile being unusable — another browser on it, or leftover lock files
// from one that was killed — which Playwright's own message does not say.
async function launch({ headless }) {
  try {
    return await openTeams({ headless, daemon: false, args: [`--remote-debugging-port=${DAEMON_PORT}`] });
  } catch (err) {
    throw new Error(
      `Could not open the browser on the profile "${PROFILE_DIR}". Check that no other browser is `
      + 'using it (a daemon that is already running, or manual-login.mjs).',
      { cause: err }
    );
  }
}

// The debugging endpoint the browser bound to, read from the file it writes at
// startup: the port on the first line, the browser's WebSocket path on the
// second. The path identifies this browser specifically, so attaching through it
// cannot land on a different one that happens to hold the port.
async function readEndpoint(portFile) {
  const deadline = Date.now() + PORT_FILE_TIMEOUT_MS;
  for (;;) {
    const [port, path] = (await readFile(portFile, 'utf8').catch(() => '')).split('\n');
    if (Number(port) > 0 && path?.trim()) {
      const endpoint = {
        port: Number(port),
        httpEndpoint: `http://127.0.0.1:${Number(port)}`,
        wsEndpoint: `ws://127.0.0.1:${Number(port)}${path.trim()}`,
      };
      // The browser writes this file with the port it was asked for even when it
      // could not bind it, and reports the failure nowhere else. Unverified, the
      // daemon would announce itself on a port belonging to somebody else's
      // service and hand every command an endpoint that cannot work.
      if (isOurBrowser(await fetchBrowserVersion(endpoint.httpEndpoint), endpoint.wsEndpoint)) {
        return endpoint;
      }
      throw new Error(
        `Port ${endpoint.port} is not answering as this browser's debugging port — something else `
        + 'is holding it.'
        + (DAEMON_PORT ? ' Leave TEAMS_DAEMON_PORT unset to have a free port picked automatically.' : '')
      );
    }
    if (Date.now() >= deadline) {
      throw new Error(
        `The browser did not open a debugging port (nothing usable in "${portFile}").`
        + (DAEMON_PORT ? ` Port ${DAEMON_PORT} (TEAMS_DAEMON_PORT) may already be in use — leave the `
          + 'variable unset to have a free port picked automatically.' : '')
      );
    }
    await delay(200);
  }
}

async function stop() {
  const result = await stopDaemon();
  if (result.stopped) {
    console.log(`Stopped the Teams daemon (pid ${result.pid}).`);
    return;
  }
  if (result.reason === 'none') console.log('No Teams daemon is running.');
  else if (result.reason === 'stale') console.log('No Teams daemon is running (a leftover record was removed).');
  else {
    console.log(`The daemon (pid ${result.pid}) did not stop when asked — see "${LOG_PATH}".`);
    process.exit(1);
  }
}

async function status() {
  if (CDP_ENDPOINT) {
    console.log(`TEAMS_CDP is set to "${CDP_ENDPOINT}" — the commands attach there, not to a daemon.`);
    return;
  }
  const info = await readInfo();
  if (!info?.pid) {
    console.log('No Teams daemon is running. The next command will start one.');
    return;
  }
  if (!await connectableDaemon()) {
    console.log(`The recorded daemon (pid ${info.pid}) is not answering on ${info.httpEndpoint}.`);
    console.log('The next command will clear the record and start a new one.');
    return;
  }
  const since = await lastActivity();
  console.log(`Running: pid ${info.pid}, ${info.httpEndpoint}, started ${info.startedAt}.`);
  console.log(since
    ? `Last used ${Math.round((Date.now() - since) / 60_000)} minute(s) ago.`
    : 'Not used by any command yet.');
}
