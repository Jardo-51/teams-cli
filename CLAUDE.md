# teams-cli

Playwright automation for Microsoft Teams (web). Each command is a standalone
`*.mjs` script; `teams.mjs` holds what the commands share and `daemon.mjs` the
shared-browser plumbing. See `README.md` for what the commands do and how to run
them.

## Checks

`nix develop --command pnpm test` runs the tests (`node:test`, files in `test/`)
and `nix develop --command pnpm run check` parses every script; both run in CI on
every push and pull request. Only what needs no browser can be covered that way,
which is what `parsing.mjs` is for — anything pure belongs there rather than
inline in a command. Everything that drives the page is still only proven by
running the command against a live session.

## Gotchas

### Module-level `const` must be declared above the run

Every command script does its work in a top-level `await` block partway down the
file, with its helper functions below. Function declarations hoist, so a helper
can sit anywhere; `const` does not. A constant declared after that block is still
in its temporal dead zone when the run reads it, and the command fails with
`ReferenceError: Cannot access 'X' before initialization` — not on some edge
case, but on the first message it touches.

`node --check` passes on such a file, so only running the script catches it. Keep
every module-level constant in the block at the top of the file, above the
argument parsing, however far that is from the helper that uses it, and when
reviewing an edit check where a newly added `const` sits relative to the run.

### A chat list is not a signed-in session

Teams renders its chat list and the open chat's messages out of an offline cache
whether or not its session still works, so `waitForChatList` passing proves
nothing about being signed in — an expired login reads as a quiet chat and posts
as a message stuck on "Retrying...". Commands wait with `waitForSignedIn`, which
fails with `SessionExpiredError` instead; a new command should too.

To exercise that path without waiting for a login to really expire, launch the
browser with its clock moved a couple of days ahead (`context.clock.setSystemTime`)
and the `login.microsoftonline.com/.../oauth2/v2.0/token` endpoint answering
`400 invalid_grant`: every cached token then reads as expired and none can be
refreshed, which Teams treats the same way. Do it on a copy of the profile, since
a message posted meanwhile stays queued in the profile and goes out on its next
healthy load.
