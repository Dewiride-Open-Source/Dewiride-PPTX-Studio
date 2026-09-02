# `tools/gate1` — the browser round trip, end to end

Gate 1 asks for a live demo that opens a deck with charts, SmartArt, animations, OLE and macros,
re-saves it, and downloads a file that opens in real PowerPoint with no repair prompt. This is the
command that checks it.

```sh
pnpm build
pnpm gate1                        # both save modes, then PowerPoint
pnpm gate1 --no-powerpoint        # everything a machine without Office can do
pnpm gate1 --deck corpus/decks/a26-ole.pptx --headed
```

Exit status: `0` every run passed, `1` one did not, `2` the harness could not run — the third kept
separate for the reason `check-corpus.ts` keeps it separate, that a run which could not answer must
never read as one that answered no.

## What it actually does

1. Serves the repository over loopback with `tools/bench/serve.ts`, cross-origin isolated.
2. Opens [`apps/studio`](../../apps/studio) in the Chromium `pnpm test` already uses. Nothing is
   installed.
3. Presses Save twice through `globalThis.pptxStudio`, the same hook `tools/bench` drives — once
   with no edit, once with the stamp.
4. **Captures Playwright's `download` event**, rather than reading bytes out of the page. This is
   the whole reason the harness exists: a stale `URL.revokeObjectURL` cancels a download, a missing
   `download` attribute navigates instead of saving, and a `Uint8Array` view onto a larger buffer
   writes the wrong bytes. None of the three fails a unit test of the export, and all three sit
   between `exportDeck` and a file on disk.
5. Reads each downloaded file back through `@pptx-studio/writer` — a second opinion, from Node.
6. Asks PowerPoint through COM, with **`OpenAndRepair` off**.

That last flag is the one that matters. It defaults to on, and under automation a repair is silent:
a file PowerPoint would prompt about opens successfully, already repaired, and reports success.
There is deliberately no `--allow-repair` here.

## What it will not do to your machine

- Reads only the repository. The server binds to `127.0.0.1` and serves `apps/`, `packages/` and
  `corpus/`; the only file PowerPoint is pointed at is one this run just wrote.
- Downloads land in `node_modules/.gate1`, which is git-ignored. `--out` moves them.
- Macros cannot run: the oracle sets `AutomationSecurity` to `ForceDisable`, because the property
  defaults to _Low_, which enables all of them.
- If PowerPoint was already running it attaches to that instance and never quits it. It quits one it
  started itself, which is the difference between tidying up and closing your work.

## Why it is not in CI

No hosted runner has Office on it — GitHub's Windows images ship Visual Studio's Office
_development_ workload and not one Office application. The browser half is in CI regardless:
[`apps/studio/src/export.test.ts`](../../apps/studio/src/export.test.ts) runs the same export in
Chromium on every `pnpm test`, against the same committed deck.

See [ADR 0015](../../docs/adr/0015-gate-1-the-browser-round-trip.md) for what was measured and what
is still not claimed.
