# teams-cli

Small Playwright-based CLI for automating Microsoft Teams (web) — capture a
login session, then post and read chat messages from the command line, and leave
reactions on them or take them back.

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
the chat-list appears the session is saved automatically, and the script tells you
when Teams has finished filling in its emoji catalog and the window can be closed.

```bash
nix develop .#playwright --command node manual-login.mjs
```

**Automatic.** Fills the account, password and MFA method for you and runs
headless. The email and password are read from a git-ignored `.env` file next to
the scripts; copy the template and fill it in:

```bash
cp .env.example .env
chmod 600 .env   # it holds a password in plain text
# then edit .env:
#   TEAMS_EMAIL=you@example.com
#   TEAMS_PASSWORD=your-password
```

The one-time MFA code cannot be known ahead of time, so once it is sent to your
phone you are asked to type it into the console. The browser closes by itself
once the session is saved. Text-message MFA is the only method this script
drives — an account with no phone method registered has to log in manually.

```bash
nix develop .#playwright --command node auto-login.mjs
```

`.env` holds your credentials, so it is git-ignored — do not commit it. Only
`.env.example` (placeholders) is tracked.

Either way, stop the [browser daemon](#the-browser-daemon) first if it is running
(`node teams-daemon.mjs --stop`) — two browsers on one profile write over each
other's stored session. The login scripts refuse to open a browser rather than
letting that happen.

And either way, the chat list appearing is not the end of it: Teams goes on
filling in its [emoji catalog](#4-react-to-a-message) for a few seconds
afterwards, and a browser closed in that gap leaves the catalog half-filled. So
both scripts wait for that sync to settle — saying so while they do — before
`auto-login.mjs` closes the browser or `manual-login.mjs` tells you that you can.

### 2. Post a message

Posts a message into the chat whose name matches `<chat name>` (partial,
case-insensitive).

```bash
nix develop .#playwright --command node post-message.mjs "<chat name>" "<message>"
```

The message is posted as a single message, with its line breaks intact and none
of the composer's typing-time auto-formatting applied — so a message that spans
several lines stays one message, and a line starting `- ` does not become a
bullet:

```bash
nix develop .#playwright --command node post-message.mjs "Developers" "Release notes:
- FE: new chat list
- BE: faster search"
```

A URL in the message does still become a clickable link, as it does in a message
typed by hand. Teams shortens the label of a long one, but the link itself keeps
the whole URL.

Add `--dry-run` to put the message into the compose box **without sending** — a
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

Reacts with `<emoji>` to the messages `<message ids>` in the chat whose name
matches `<chat name>` (partial, case-insensitive).

```bash
nix develop .#playwright --command node react-to-message.mjs "<chat name>" "<message ids>" "<emoji>"
```

The ids and the emoji are the ones `read-chat-messages.mjs` reports, so its
output can be fed straight back in:

```bash
nix develop .#playwright --command node react-to-message.mjs "Developers" "1785922526738" "👍"
```

Several messages can be reacted to at once by passing their ids as one
comma-separated list. Each of them gets the same emoji, and the browser and the
chat are opened once for the whole list rather than once per message, so this is
quicker than a run per message:

```bash
nix develop .#playwright --command node react-to-message.mjs "Developers" "1785922526738,1785922530011,1785922612903" "👍"
```

The ids are worked through newest first, whatever order they are given in, so
one walk back through the history covers the whole list.

A message that cannot be reached does not stop the ones after it: the run works
through the whole list, says what happened to each, and only then fails, naming
the messages it could not react to. A failure that says nothing about the
message — the emoji is not in the reaction picker, for instance — does stop it,
since every id left would only meet the same failure again.

Pass the emoji character itself, not its name. Where several of the picker's
emoji share one character (Teams has both an animated and a plain 👏, for
example), the first one it offers is used.

The emoji the picker offers come from a catalog Teams keeps in the browser
profile and fills in once. A sync that was cut short — the browser closed part
way through it — leaves a catalog the client treats as finished, whose missing
categories render as empty headings in the picker; every emoji in them then
looks as though Teams does not have it. The [login scripts](#1-log-in) wait for
that sync rather than closing the browser on it, so a catalog that is short is
now one from a profile logged in before they did, or from a run that was killed.
Both reaction commands check the catalog before they start all the same, and an
incomplete one is dropped and Teams reloaded so that it is fetched again. That
costs one page load plus the wait for the sync (up to a minute, on the run that
finds it), and says so as it happens.

The chat history is scrolled back until the messages are found, so reacting to
an old message takes as long as reading that far back. That walk reports every
step it takes, and it stops once the pane reaches messages older than the one it
is looking for: an id that belongs to another chat, or to a message that has
been deleted, fails as soon as the history reaches its age rather than after
paging back through the whole conversation. The bound is on the id's age, so a
bad id older than everything in this chat is never reached past and still costs
a walk to the beginning — as does one that is not a plain number, since only
those carry a time to compare. Both are capped by the step and history-page
limits regardless, so neither runs forever.

Reacting is a toggle in Teams, so a reaction you already left is never clicked
again — that would take it back; to take one back on purpose, use
`unreact-to-message.mjs` below.

### 5. Take a reaction back

Removes the `<emoji>` reaction *you* left on the messages `<message ids>`. It
takes the same arguments as `react-to-message.mjs`, including the comma-separated
list of ids, and undoes exactly what that command does:

```bash
nix develop .#playwright --command node unreact-to-message.mjs "<chat name>" "<message ids>" "<emoji>"
```

```bash
nix develop .#playwright --command node unreact-to-message.mjs "Developers" "1785922526738,1785922530011" "👍"
```

Only your own reaction can be removed — Teams offers no way to take back someone
else's. A message you have not reacted to with that emoji is therefore reported
as having nothing to remove, rather than treated as a failure, so re-running the
command is harmless.

The reaction that is taken back is the exact one you left, not merely one that
looks the same: where several of the picker's emoji share a character, clicking
the wrong one of them would add a second reaction instead of removing the first.
If you left more than one of those look-alikes on the same message — which takes
reacting by hand in the client, since `react-to-message.mjs` leaves a message
that already carries your reaction alone — every one of them is taken back, and
the run says how many.

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
  should mean logging in *less* often, though — the live tab keeps refreshing
  its own tokens, subject to whatever re-authentication your tenant enforces
  anyway.
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
daemon, or from a browser of the command's own), the preamble both login scripts
run before they can open a browser, finding and opening a chat by name, scrolling
the message pane back through the history, and everything the two reaction
commands do alike (their arguments, the walk to each message of a list, the hover
toolbar, the emoji catalog check and the emoji picker) — so each script only
contains its own logic.
`daemon.mjs` is the client side of the daemon: finding it, starting it, and
serialising commands against it.

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
