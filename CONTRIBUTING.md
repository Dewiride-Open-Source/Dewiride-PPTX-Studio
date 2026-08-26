# Contributing to PPTX Studio

Thank you for considering it. This document is short on ceremony and specific about the few things
that genuinely matter here.

## Sign-off, not a CLA

This project uses the [Developer Certificate of Origin](https://developercertificate.org/). Every
commit needs a `Signed-off-by` line, which `git commit -s` adds for you:

```
Signed-off-by: Your Name <your.email@example.com>
```

That is the whole legal process. There is no CLA to sign and no copyright assignment: you keep your
copyright, and your contribution is licensed under Apache-2.0 like the rest of the tree.

If you write code informed by a copyleft implementation — LibreOffice, ONLYOFFICE, PPTist — say so
in the pull request. It is not automatically disqualifying and it is far cheaper to discuss before
the merge than after. See [`LEGAL.md`](./LEGAL.md).

## Getting set up

Node ≥ 24.11, pnpm ≥ 11.

```sh
pnpm install
pnpm browsers     # one-off; downloads the Chromium the core suite runs in
pnpm check        # everything CI runs
```

`pnpm check` is `layering → format:check → lint → typecheck → build → test`. Run it before opening a
pull request; CI runs the same thing and will not tell you anything new.

Tests and typechecking read package **source**, not `dist`, so you never need to build first.

## Four rules that are not negotiable

These are enforced by CI. If one of them is in your way, that is worth a conversation in an issue —
it is not worth a workaround.

**1. The core never touches Node.** Everything under `packages/` except `cli` runs in a browser tab
and in a Web Worker. No `node:*` imports, no `process`, no `Buffer`, no `"node"` condition in an
exports map. The core test suite runs in real Chromium precisely so that a leak fails loudly instead
of passing under jsdom and breaking for a user.

**2. Dependencies run one way only.** The layer table is
[`tools/layering/layers.ts`](./tools/layering/layers.ts) and `pnpm layering` enforces it. If your
change needs an upward edge, the design is wrong somewhere — open an issue rather than moving a
package in the table quietly.

**3. No top-level await.** One occurrence makes a module asynchronous, and Node's `require(esm)`
bridge then throws `ERR_REQUIRE_ASYNC_MODULE` for every CommonJS consumer of anything that imports
it. Move the await into an async function and call it lazily.

**4. Never synthesize markup you did not read.** This is the architectural heart of the project. If
a part of a `.pptx` was not edited, it is streamed back out byte-for-byte. If an
`mc:AlternateContent` branch is not understood, it is preserved untouched, not rewritten. PowerPoint
emits no diagnostic log when it rejects a file, so "we did not touch it" is the strongest guarantee
available and it is worth a great deal.

## Changesets

Any change that affects a published package needs a changeset:

```sh
pnpm changeset
```

Pick the packages, pick the bump, and write the entry for someone reading the changelog rather than
for someone reading the diff. Changes confined to `tools/`, `docs/` or CI do not need one.

## Commits and pull requests

Conventional-commit prefixes (`feat:`, `fix:`, `docs:`, `refactor:`, `test:`, `chore:`) are
appreciated and not enforced.

A good pull request here says what it does, and what it does **not** do. Given the scope constraints
in [`SCOPE.md`](./SCOPE.md), the second half is often the more useful one.

If you are adding support for a piece of OOXML, please include a fixture that exercises it and a
note about where in the spec (or in `[MS-OE376]`) the behaviour is defined — including the cases
where Office deviates from the standard text, which is where most of the difficulty lives.

## Good first issues

Geometry presets are unusually well-suited to a first contribution: each has an objective visual
test and a number that moves. Look for the `good first issue` label once Phase 2 opens.

## Code of conduct

By participating you agree to abide by the [Code of Conduct](./CODE_OF_CONDUCT.md).
