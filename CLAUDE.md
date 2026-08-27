# teams-cli

Playwright automation for Microsoft Teams (web). Each command is a standalone
`*.mjs` script; `teams.mjs` holds what the commands share and `daemon.mjs` the
shared-browser plumbing. See `README.md` for what the commands do and how to run
them.

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
