# teams-cli

Small Playwright-based CLI for automating Microsoft Teams (web) — capture a
login session, then post, read and react to chat messages from the command line.

## Requirements

- [Nix](https://nixos.org/) with flakes enabled (provides Node.js, pnpm, and the
  Playwright browsers).

## Setup

Install dependencies:

```bash
nix develop --command pnpm install
```

## Usage

All commands run inside the Playwright dev shell (`.#playwright`), which provides
the browsers.

### 1. Log in

There are two ways to capture a session. Both save the same thing; pick whichever
suits you.

**Manual.** Opens Teams in a visible browser. Log in by hand (including MFA); once
the chat-list appears the session is saved automatically and you can close the window.

```bash
nix develop .#playwright --command node manual-login.mjs
```

**Automatic.** Fills the account, password and MFA method for you and runs
headless. The email and password are read from a git-ignored `.env` file next to
the scripts; copy the template and fill it in:

```bash
cp .env.example .env
# then edit .env:
#   TEAMS_EMAIL=you@example.com
#   TEAMS_PASSWORD=your-password
```

The one-time MFA code cannot be known ahead of time, so once it is sent to your
phone you are asked to type it into the console. The browser closes by itself
once the session is saved.

```bash
nix develop .#playwright --command node auto-login.mjs
```

`.env` holds your credentials, so it is git-ignored — do not commit it. Only
`.env.example` (placeholders) is tracked.

Either way, stop the [browser daemon](#the-browser-daemon) first if it is running
(`node teams-daemon.mjs --stop`) — two browsers on one profile write over each
other's stored session. The login scripts refuse to open a browser rather than
letting that happen.

### 2. Post a message

Posts a message into the chat whose name matches `<chat name>` (partial,
case-insensitive).

```bash
nix develop .#playwright --command node post-message.mjs "<chat name>" "<message>"
```

Add `--dry-run` to type the message into the compose box **without sending** — a
safe way to confirm the correct chat is targeted:

```bash
nix develop .#playwright --command node post-message.mjs "Developers" "Hello" --dry-run
```

### 3. Read recent messages

Reads the recent messages of the chat whose name matches `<chat name>` (partial,
case-insensitive) and writes them to `<output file>` as JSON.

```bash
nix develop .#playwright --command node read-chat-messages.mjs "<chat name>" "<period>" "<output file>"
```

`<period>` is a relative time span ending "now" — a number followed by `m`
(minutes), `h` (hours) or `d` (days):

```bash
nix develop .#playwright --command node read-chat-messages.mjs "Developers" "2d" messages.json
```

Add `--without-reactions-only` to keep just the messages nobody has reacted to
— handy for spotting requests that went unacknowledged. It also runs faster,
since the reaction authors then never have to be read:

```bash
nix develop .#playwright --command node read-chat-messages.mjs "Developers" "2d" messages.json --without-reactions-only
```

The output file holds real names and message bodies, so treat it like the chat
itself and keep it out of the repository. The default `messages.json` and an
`export/` directory are git-ignored for that reason; if you write somewhere
else, make sure that path is ignored too.

The chat history is scrolled back until the start of the period is reached, so
longer periods take longer to read. Each message is written as:

```json
[
  {
    "id": "1785922526738",
    "time": "2026-08-05T09:35:26.738Z",
    "author": "Jane Doe",
    "body": "Hello team, see https://example.com/…",
    "links": [{ "text": "https://example.com/…", "href": "https://example.com/very/long/path" }],
    "reactions": [{ "author": "John Doe", "emoji": "📝" }]
  }
]
```

`id` is the Teams message id and `time` is ISO 8601 (UTC). `body` is the
message text, but Teams truncates long links in it (e.g. `https://…/…`), so
`links` lists each link separately: `text` is the (possibly truncated) label as
it appears in `body`, and `href` is the full, untruncated URL. It is an empty
array when the message has no links. `reactions` is an empty array when nobody
reacted, and `null` when the message had reactions that could not be read — so a
failure is never mistaken for "nobody reacted".
Reactor names are read by hovering the reaction pills — the script never clicks
one, since clicking a pill toggles your own reaction. Be aware that opening a
chat marks its messages as read, which is inherent to reading them through the
web client.

### 4. React to a message

Reacts with `<emoji>` to the message `<message id>` in the chat whose name
matches `<chat name>` (partial, case-insensitive).

```bash
nix develop .#playwright --command node react-to-message.mjs "<chat name>" "<message id>" "<emoji>"
```

The id and the emoji are the ones `read-chat-messages.mjs` reports, so its
output can be fed straight back in:

```bash
nix develop .#playwright --command node react-to-message.mjs "Developers" "1785922526738" "👍"
```

Pass the emoji character itself, not its name. Where several of the picker's
emoji share one character (Teams has both an animated and a plain 👏, for
example), the first one it offers is used.

The chat history is scrolled back until the message is found, so reacting to an
old message takes as long as reading that far back. Reacting is a toggle in
Teams, so a reaction you already left is never clicked again — that would take
it back; such a run reports the existing reaction and changes nothing.

## The browser daemon

Booting the Teams web app — bootstrapping, refreshing tokens, rendering the chat
list — takes far longer than the work any of the commands then do. So the
commands share one browser instead of each launching their own: the first one to
run starts a background daemon holding a signed-in Teams tab, and every command
after it attaches to that tab over CDP and leaves it running.

This needs no setup. Run the commands as documented above; the first one after a
while is as slow as they all used to be, and the rest are not.

The daemon exits by itself after 15 minutes without a command
(`TEAMS_DAEMON_IDLE`), so a burst of commands costs one boot and nothing is left
running overnight. It can also be driven by hand:

```bash
nix develop .#playwright --command node teams-daemon.mjs --status
nix develop .#playwright --command node teams-daemon.mjs --stop
nix develop .#playwright --command node teams-daemon.mjs            # start in the foreground
nix develop .#playwright --command node teams-daemon.mjs --headed   # ...with a visible window, for debugging
```

Worth knowing:

- **You show as Available while it runs.** A connected Teams client is a
  connected Teams client, and other people can see it. That is what the idle
  timeout is for; `TEAMS_DAEMON_IDLE=0` keeps the daemon up until it is stopped.
- **Stop it before logging in.** Two browsers sharing one profile directory
  write over each other's stored session, so `manual-login.mjs` and
  `auto-login.mjs` both refuse to run while the daemon is up. A running daemon
  should mean logging in *less*
  often, though — the
  live tab keeps refreshing its own tokens, subject to whatever
  re-authentication your tenant enforces anyway.
- **Commands run one at a time.** They share a single page, so a second command
  waits for the first to finish rather than driving the same page with it.
- **There is one daemon per working directory.** The profile and the daemon's
  bookkeeping are both relative paths by default, so commands run from another
  directory get their own daemon and their own profile rather than sharing these.
  Set `TEAMS_PROFILE` and `TEAMS_DAEMON_DIR` to absolute paths to share one.
- **`TEAMS_DAEMON=0` turns all of this off**, giving each command its own browser
  and its own cold boot, as before.
- The daemon's log and its bookkeeping live in `$TEAMS_DAEMON_DIR` (default
  `.daemon/`, git-ignored). `.daemon/daemon.log` is where a daemon that failed to
  start explains itself.

## How auth works

A persistent browser profile (`$TEAMS_PROFILE`, default `.profile`) holds
localStorage/cache, but reopening it drops session cookies. Since not all tenants
offer the "Stay signed in?" option, the login scripts (`manual-login.mjs` and
`auto-login.mjs`) also capture the full session (cookies + per-origin
localStorage) to a storageState file (`$TEAMS_AUTH`, default `.auth/user.json`),
which the other scripts restore before navigating.

`teams.mjs` holds what the scripts share — obtaining a Teams page (from the
daemon, or from a browser of the command's own), finding and opening a chat by
name, and scrolling the message pane back through the history — so each script
only contains its own logic. `daemon.mjs` is the client side of the daemon:
finding it, starting it, and serialising commands against it.

## Configuration

| Variable             | Default           | Description                                                     |
| -------------------- | ----------------- | --------------------------------------------------------------- |
| `TEAMS_PROFILE`      | `.profile`        | Persistent browser profile directory                              |
| `TEAMS_AUTH`         | `.auth/user.json` | Playwright storageState (auth) file                               |
| `TEAMS_DAEMON`       | `1`               | Set to `0` to give every command its own browser                  |
| `TEAMS_DAEMON_IDLE`  | `15`              | Whole minutes without a command before the daemon exits; `0` never |
| `TEAMS_DAEMON_PORT`  | `0`               | Debugging port for the daemon's browser; `0` picks a free one     |
| `TEAMS_DAEMON_DIR`   | `.daemon`         | Where the daemon's record, log and command lock are kept          |
| `TEAMS_CDP`          | —                 | Attach to this CDP endpoint instead of managing a daemon          |

`.profile/`, `.auth/` and `.env` all contain login credentials and are
git-ignored — do not commit them.

The debugging port is bound on `127.0.0.1` only, but anything that can reach it
can drive the signed-in browser, so leave `TEAMS_DAEMON_PORT` unset unless you
have a reason to fix it.
