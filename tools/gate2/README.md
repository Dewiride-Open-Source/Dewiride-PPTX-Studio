# `tools/gate2` — the silent slide, beside PowerPoint's own

Gate 2 asks for a real deck rendered with gradients, pattern fills, cropped image fills, dashed
strokes, arrowheads, shadows, glows, nested rotated/flipped groups and themed backgrounds — shown
side by side against a reference, with a per-slide diff heatmap. This is the command that checks it.

```sh
pnpm build
pnpm gate2                                   # the committed oracle; no Office needed
pnpm gate2 --capture <dir>                   # PowerPoint's own PNGs, if one was captured
pnpm gate2 --out <dir>
```

It runs as part of `pnpm check`, which `pnpm gate1` cannot: no Office, no network, and the
reference is committed. A gate nothing runs is a sentence again within a sub-phase.

Exit status: `0` the gate holds, `1` it does not, `2` the harness could not run — the third kept
separate for the reason [`check-roundtrip.ts`](../corpus/roundtrip/check-roundtrip.ts) keeps it
separate, that a run which could not answer must never read as one that answered no.

## What it actually does

1. Scans every committed slide's **own part bytes** for the nine features Gate 2 names
   ([`features.ts`](./features.ts)). Detection is on the markup rather than on our model, because a
   gate that asked our parser whether a feature was present would be asking the thing under test.
2. Takes the **union** of the slides that carry them, so nothing is cherry-picked.
3. Renders each one through `@pptx-studio/render-svg` in the same Chromium
   [`tools/fidelity`](../fidelity) rasterises with, and scores it against the oracle sub-phase 3.9
   committed under `corpus/ground-truth/render/fidelity/`.
4. Writes `fidelity/gate2/index.html`: PowerPoint, ours, and the difference, three panels a slide.

## What is gated, and what is only reported

**Gated, with no tolerance in it:** every feature Gate 2 names is carried by at least one committed
slide; every one of those slides draws without the renderer refusing; every one has an oracle to be
compared against. On 2026-09-07 the first of those failed — two of the nine had no corpus slide at
all — and `a44-transforms` and `a45-backgrounds` were written because of it.

**Reported and never gated:** the scores. The two sides differ by a few levels everywhere, so a
pass/fail on the number would need a tolerance, and a tolerance is the thing that gets loosened.
[ADR 0035](../../docs/adr/phase-3-text/0035-the-fidelity-harness.md) settles that; this follows it.

## The reference is PowerPoint, not LibreOffice

The plan names a LibreOffice reference, and this does not use one — recorded as a deviation in
[ADR 0038](../../docs/adr/phase-2-geometry-and-paint/0038-gate-2.md). LibreOffice is a second
implementation with opinions of its own; PowerPoint wrote the format, and 3.9 already commits its
answer for every corpus slide. What LibreOffice would still buy is a reference that runs on a Linux
CI runner, which PowerPoint cannot, and that is carried debt rather than something closed here.

## What it will not do to your machine

Reads the repository and nothing else. No Office automation, no network, no install: the oracle is
committed, so a clone can run this. `--capture` reads PNGs from a directory you name and copies the
ones it uses into the report.
