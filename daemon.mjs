import { spawn } from 'node:child_process';
import { openSync, closeSync } from 'node:fs';
import { link, mkdir, readFile, stat, unlink, writeFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';

// The client half of the browser daemon: finding a running one, starting one on
// demand, and serialising commands against the single page it owns. The daemon
// process itself is teams-daemon.mjs.
//
// Booting the Teams SPA dominates the runtime of every command, so one browser
// is kept alive between commands and each command attaches to it over CDP.

const SCRIPT_DIR = dirname(fileURLToPath(import.meta.url));

// Where the daemon's bookkeeping lives: the record of the running daemon, its
// log, the command lock and the last-used timestamp.
export const DAEMON_DIR = process.env.TEAMS_DAEMON_DIR || '.daemon';
export const INFO_PATH = join(DAEMON_DIR, 'daemon.json');
export const LOG_PATH = join(DAEMON_DIR, 'daemon.log');
const LOCK_PATH = join(DAEMON_DIR, 'command.lock');
const ACTIVITY_PATH = join(DAEMON_DIR, 'activity');

// The daemon is used by default; TEAMS_DAEMON=0 restores the old behaviour of a
// fresh browser per command.
export const DAEMON_ENABLED = process.env.TEAMS_DAEMON !== '0';
// The debugging port the daemon's browser listens on. The default of 0 lets
// Chromium pick a free one, which is what we want: the port is then read back
// from the browser rather than guessed, so a fixed port that something else
// already holds cannot silently become the browser we attach to.
export const DAEMON_PORT = numberFromEnv('TEAMS_DAEMON_PORT', 0);
// An externally managed browser to attach to instead. With this set nothing is
// ever spawned or stopped here.
export const CDP_ENDPOINT = process.env.TEAMS_CDP || '';
// How long the daemon stays up with no command using it. A daemon is a signed-in
// Teams client, and one that runs forever shows you as Available forever, so it
// goes away after a burst of activity rather than living until the next reboot.
// 0 keeps it up indefinitely.
export const IDLE_TIMEOUT_MS = numberFromEnv('TEAMS_DAEMON_IDLE', 15) * 60_000;

// How long a command waits for the browser while another command has it. Reading
// a long history takes many minutes, so this is generous — it is a deadlock
// guard, not a queueing policy.
const LOCK_WAIT_MS = 30 * 60_000;
// How long a freshly spawned daemon gets to become connectable. It covers the
// same cold boot a command used to pay for itself, hence the same order of
// magnitude as waitForChatList's timeout.
const STARTUP_TIMEOUT_MS = 180_000;
// How long a stopped daemon gets to close its browser and exit.
const STOP_TIMEOUT_MS = 30_000;
// How large the shared daemon log may grow before it is started over.
const MAX_LOG_BYTES = 256 * 1024;

// The endpoint of a daemon that is running and answering, or null. A record left
// behind by a crashed daemon reads as "none", so the caller starts a new one.
export async function connectableDaemon() {
  const info = await readInfo();
  if (!info?.pid || !isAlive(info.pid)) return null;
  return await isReachable(info) ? info : null;
}

// Returns the CDP endpoint to attach to, starting a daemon if there is none.
// Must be called with the command lock held: the lock is what stops two commands
// that start at the same moment from each spawning a daemon. Nothing below us
// would catch that — two browsers can and will run on one profile directory,
// writing over each other's session — so it has to be prevented here.
export async function ensureDaemon() {
  if (CDP_ENDPOINT) return CDP_ENDPOINT;

  const running = await connectableDaemon();
  if (running) return running.wsEndpoint;

  // Either there never was a daemon, or the one described by the record is gone
  // — it crashed, or its idle timeout fired just before we got the lock. Both
  // are "start one", which is the retry the stale-record case needs.
  await clearInfo();
  return (await startDaemon()).wsEndpoint;
}

// Spawns the daemon detached and waits for it to become connectable. A detached
// child's stderr goes nowhere, so it is pointed at the log file and read back
// here: "it did not come up" is only useful with the reason attached.
async function startDaemon() {
  await mkdir(DAEMON_DIR, { recursive: true });
  // Every daemon appends to the same log, so it is dropped once it outgrows what
  // is useful to read back. Nothing here is worth keeping across that many runs.
  if (await stat(LOG_PATH).then(s => s.size > MAX_LOG_BYTES, () => false)) {
    await unlink(LOG_PATH).catch(() => {});
  }
  const logFd = openSync(LOG_PATH, 'a');
  const child = spawn(process.execPath, [join(SCRIPT_DIR, 'teams-daemon.mjs')], {
    detached: true,
    stdio: ['ignore', logFd, logFd],
  });
  let exitCode = null;
  child.on('exit', (code, signal) => { exitCode = code ?? signal; });
  child.on('spawn', () => closeSync(logFd));
  child.unref();

  console.log('Starting the Teams browser daemon — this first command pays for the cold start...');
  const deadline = Date.now() + STARTUP_TIMEOUT_MS;
  while (Date.now() < deadline) {
    const info = await connectableDaemon();
    if (info) return info;
    if (exitCode !== null) {
      throw new Error(
        `The Teams browser daemon exited (${exitCode}) instead of starting up. Its output:\n${await logTail()}`
      );
    }
    await delay(500);
  }
  throw new Error(
    `The Teams browser daemon did not become reachable within ${STARTUP_TIMEOUT_MS / 1000}s. `
    + `Its output:\n${await logTail()}`
  );
}

// Whether the daemon's browser is actually answering on its debugging port —
// the record alone only says a process is alive, not that it is usable.
async function isReachable(info) {
  return isOurBrowser(await fetchBrowserVersion(info.httpEndpoint), info.wsEndpoint);
}

// What the debugging port says it is, or null if nothing usable answers there.
export async function fetchBrowserVersion(httpEndpoint) {
  if (!httpEndpoint) return null;
  try {
    const response = await fetch(`${httpEndpoint}/json/version`, { signal: AbortSignal.timeout(2000) });
    return response.ok ? await response.json() : null;
  } catch {
    // Not answering, or answering with something that is not JSON.
    return null;
  }
}

// Whether what answered is the specific browser we expect, rather than some
// other service that happens to hold the port. Chromium names its browser
// endpoint after a per-launch uuid, so the path is the identity; the host is
// not compared, since a browser may spell the same address differently than we
// do. This matters because a port that could not be bound is not reported as an
// error anywhere — the browser writes the port down regardless — so "something
// answers" is not enough to conclude the daemon is there.
export function isOurBrowser(version, wsEndpoint) {
  if (!version?.Browser || !version.webSocketDebuggerUrl || !wsEndpoint) return false;
  try {
    return new URL(version.webSocketDebuggerUrl).pathname === new URL(wsEndpoint).pathname;
  } catch {
    return false;
  }
}

// Stops the recorded daemon and forgets it. Reports what it found rather than
// printing anything, so that --stop and the automatic restart below can each
// word it their own way.
export async function stopDaemon() {
  const info = await readInfo();
  if (!info?.pid) return { stopped: false, reason: 'none' };
  if (!isAlive(info.pid)) {
    await clearInfo();
    return { stopped: false, reason: 'stale', pid: info.pid };
  }

  // SIGTERM rather than SIGKILL: the daemon closes its browser on it, which is
  // what flushes the profile.
  try {
    process.kill(info.pid, 'SIGTERM');
  } catch {
    // It exited between the liveness check and here — the outcome we wanted.
  }
  const deadline = Date.now() + STOP_TIMEOUT_MS;
  while (isAlive(info.pid) && Date.now() < deadline) await delay(250);
  if (isAlive(info.pid)) return { stopped: false, reason: 'timeout', pid: info.pid };

  // The daemon clears the record on its way out; this covers one that did not
  // get that far.
  await clearInfo();
  return { stopped: true, pid: info.pid };
}

// The last few lines the daemon wrote, for reporting why a spawn failed.
async function logTail(lines = 30) {
  const log = await readFile(LOG_PATH, 'utf8').catch(() => '');
  return log.trimEnd().split('\n').slice(-lines).join('\n') || `(nothing in "${LOG_PATH}")`;
}

// --- The command lock -------------------------------------------------------
//
// Distinct from the browser profile lock, which is Chromium's own and keeps the
// daemon and manual-login.mjs off the same profile. This one serialises CLI
// invocations against the one shared page, and is held from before the "is a
// daemon there?" check until the command is done — so it also covers starting
// one, and gives the daemon a way to know it must not exit right now.

// Waits for the lock and returns the function that releases it.
export async function acquireCommandLock({ timeoutMs = LOCK_WAIT_MS } = {}) {
  const deadline = Date.now() + timeoutMs;
  let announced = false;
  for (;;) {
    const release = await tryCommandLock();
    if (release) return release;
    if (Date.now() >= deadline) {
      throw new Error(
        `Another teams-cli command has been using the browser for more than ${timeoutMs / 60_000} `
        + `minutes. If none is running, delete "${LOCK_PATH}".`
      );
    }
    if (!announced) {
      console.log('Another teams-cli command is using the browser — waiting for it to finish...');
      announced = true;
    }
    await delay(500);
  }
}

// Takes the lock if it is free, without waiting. Returns the release function or
// null. This is how the daemon makes sure it never exits mid-command.
export async function tryCommandLock() {
  const token = `${process.pid}:${Date.now()}:${Math.random().toString(36).slice(2)}`;
  if (await writeLock(token)) return () => releaseLock(token);
  // Held — unless the holder is gone. A command killed with Ctrl-C leaves its
  // lock behind, and nothing else would ever clear it.
  if (await breakStaleLock() && await writeLock(token)) return () => releaseLock(token);
  return null;
}

// Written to a staging file and then linked into place, rather than created
// empty and filled in afterwards: link() fails with EEXIST atomically, so the
// lock file never exists without its token. A reader that caught it empty could
// not tell who holds it, would read the pid as 0, and would break the lock as
// stale at the very moment it was being taken.
async function writeLock(token) {
  await mkdir(DAEMON_DIR, { recursive: true });
  // One staging name per process, so two of them cannot overwrite each other's
  // content between the write and the link.
  const staging = `${LOCK_PATH}.${process.pid}.staging`;
  try {
    await writeFile(staging, token);
    await link(staging, LOCK_PATH);
    return true;
  } catch (err) {
    if (err.code === 'EEXIST') return false;
    throw err;
  } finally {
    await unlink(staging).catch(() => {});
  }
}

// Removes a lock whose owning process no longer exists. The file is re-read
// right before it is deleted and only removed if it still holds what was
// examined, so a lock taken by someone else in between is not thrown away.
async function breakStaleLock() {
  const held = await readFile(LOCK_PATH, 'utf8').catch(() => null);
  if (held === null) return false;
  // Anything that does not name a pid counts as held: a lock this process cannot
  // attribute is not evidence that its owner is gone, and treating it as stale
  // is exactly how a lock gets taken away from a live command.
  const pid = Number(held.split(':')[0]);
  if (!Number.isInteger(pid) || pid <= 0 || isAlive(pid)) return false;
  const stillHeld = await readFile(LOCK_PATH, 'utf8').catch(() => null);
  if (stillHeld !== held) return false;
  await unlink(LOCK_PATH).catch(() => {});
  return true;
}

// Releasing only removes our own lock: if it was broken as stale and retaken
// while we were still running, it now belongs to someone else.
async function releaseLock(token) {
  const held = await readFile(LOCK_PATH, 'utf8').catch(() => null);
  if (held === token) await unlink(LOCK_PATH).catch(() => {});
}

// --- Activity, so the daemon can tell it is idle ----------------------------

export async function touchActivity() {
  await mkdir(DAEMON_DIR, { recursive: true }).catch(() => {});
  await writeFile(ACTIVITY_PATH, String(Date.now())).catch(() => {});
}

// When a command last used the browser, or null if nothing ever has.
export async function lastActivity() {
  const at = Number(await readFile(ACTIVITY_PATH, 'utf8').catch(() => ''));
  return Number.isFinite(at) && at > 0 ? at : null;
}

// --- The record of the running daemon ---------------------------------------

export async function readInfo() {
  try {
    return JSON.parse(await readFile(INFO_PATH, 'utf8'));
  } catch {
    return null;
  }
}

export async function writeInfo(info) {
  await mkdir(DAEMON_DIR, { recursive: true });
  await writeFile(INFO_PATH, JSON.stringify(info, null, 2) + '\n', 'utf8');
}

export async function clearInfo() {
  await unlink(INFO_PATH).catch(() => {});
}

export function isAlive(pid) {
  if (!Number.isInteger(pid) || pid <= 0) return false;
  try {
    process.kill(pid, 0);
    return true;
  } catch (err) {
    // EPERM means the process exists but belongs to someone else.
    return err.code === 'EPERM';
  }
}

export function delay(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function numberFromEnv(name, fallback) {
  const raw = process.env[name];
  if (raw === undefined || raw.trim() === '') return fallback;
  const value = Number(raw);
  if (!Number.isFinite(value) || value < 0) {
    console.log(`Ignoring ${name}="${raw}" — expected a non-negative number; using ${fallback}.`);
    return fallback;
  }
  return value;
}
