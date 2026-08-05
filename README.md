# teams-cli

Small Playwright-based CLI for automating Microsoft Teams (web) — capture a
login session, then post messages into chats from the command line.

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

Opens Teams in a visible browser. Log in manually (including MFA); once the chat
list appears the session is saved automatically and you can close the window.

```bash
nix develop .#playwright --command node manual-login.mjs
```

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

The chat history is scrolled back until the start of the period is reached, so
longer periods take longer to read. Each message is written as:

```json
[
  {
    "id": "1785922526738",
    "time": "2026-08-05T09:35:26.738Z",
    "author": "Jane Doe",
    "body": "Hello team",
    "reactions": [{ "author": "John Doe", "emoji": "📝" }]
  }
]
```

`id` is the Teams message id, `time` is ISO 8601 (UTC), and `reactions` is an
empty array when nobody reacted. Reactor names are read by hovering the reaction
pills — the script never clicks one, since clicking a pill toggles your own
reaction. Be aware that opening a chat marks its messages as read, which is
inherent to reading them through the web client.

## How auth works

A persistent browser profile (`$TEAMS_PROFILE`, default `.profile`) holds
localStorage/cache, but reopening it drops session cookies. Since not all tenants
offer the "Stay signed in?" option, `manual-login.mjs` also captures the full
session (cookies + per-origin localStorage) to a storageState file
(`$TEAMS_AUTH`, default `.auth/user.json`), which `post-message.mjs` restores
before navigating.

## Configuration

| Variable        | Default            | Description                          |
| --------------- | ------------------ | ------------------------------------ |
| `TEAMS_PROFILE` | `.profile`         | Persistent browser profile directory |
| `TEAMS_AUTH`    | `.auth/user.json`  | Playwright storageState (auth) file   |

Both `.profile/` and `.auth/` contain login credentials and are git-ignored —
do not commit them.
