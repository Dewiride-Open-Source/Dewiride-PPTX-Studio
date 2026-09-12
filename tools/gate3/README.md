# `tools/gate3` — a hundred slides, at any zoom, entirely client-side

Gate 3 asks for a 100-slide deck rendered faithfully at any zoom, entirely client-side. This is the
command that checks it, and it checks it through the page a person would use — `apps/studio`, built,
served over loopback and driven through its automation hook — rather than through the `<img>` path
[`tools/fidelity`](../fidelity) scores by.

```sh
pnpm build
pnpm gate3                                   # the gate; the committed oracle and baseline, no Office
pnpm gate3 --record                          # this machine's first zoom baseline
pnpm gate3 --record --why '<reason>' --expect <n>
pnpm gate3 --deck <path>                     # a measurement of any deck; scores and gates nothing
pnpm gate3 --out <dir> --port <n> --headed
```

It runs as part of `pnpm check`. Exit status: `0` the gate holds, `1` it does not, `2` the harness
could not run — the third kept apart for the reason
[`check-roundtrip.ts`](../corpus/roundtrip/check-roundtrip.ts) keeps it apart. A `--deck` run exits
`0` when it ran, because it reached no verdict.

## What it actually does

1. Serves the repository over loopback ([`tools/bench/serve.ts`](../bench/serve.ts)), opens the
   studio page in the Chromium `tools/fidelity` rasterises with, and asks it to open
   `corpus/decks/a46-hundred-slides.pptx` — the one committed deck with a hundred slides and the one
   PowerPoint's export is committed for at 120, 240, 960, 1920 and 3840 pixels wide.
2. Takes the browser context **offline**. Everything after this — every zoom, every mount, every
   screenshot — has to come from what the page already holds.
3. For each stage zoom (100 %, then 25 %, 200 %, 400 %) and each slide: mounts it, serialises the
   `<svg>`, screenshots the page's own DOM at an integer clip, reduces the screenshot through the same
   `oracleGrid` path PowerPoint's PNGs take, and scores it against PowerPoint's grid at that width.
   Then the strip: every 120-pixel thumbnail, the same way.
4. Writes `fidelity/gate3/index.html`: one column per zoom, every slide worst first, and for the
   three worst at each zoom the page's screenshot beside PowerPoint's grid and the difference.

The stage and the strip are pinned at whole page pixels by an injected stylesheet so a clip is
exact; nothing inside an `<svg>` is touched, and a fractional box is refused rather than rounded.

## What is gated, and what is only reported

**Gated, with no tolerance in it** ([`checks.ts`](./checks.ts)):

- the deck has at least a hundred slides, and every one draws at every zoom and in the strip;
- the SVG at every zoom is byte-identical to the SVG at 100 % once the root's size and the values of
  `stroke-width` and `stroke-dasharray` are masked — the two attributes the stroke rule rounds to
  device pixels (F2, ADR 0054) — and the thumbnail is too, once its id prefix is renamed;
- no page or console error;
- no request the gate does not recognise (the page, its bundle under `/dist/`, the deck), and none
  at all once offline;
- every raster is the one recorded in `corpus/ground-truth/render/fidelity/zoom/expected.<env>.json`,
  by SHA-256, under the same font lock `pnpm fidelity` uses.

**Reported and never gated:** the scores against PowerPoint at each width, and every timing — parse,
first slide, the strip, each mount, each screenshot. The two sides differ by a few levels everywhere,
so a pass/fail on the number would need a tolerance, and a tolerance is the thing that gets loosened
([ADR 0035](../../docs/adr/phase-3-text/0035-the-fidelity-harness.md)).

## `--deck`: a measurement, not a gate

A deck outside the corpus has no oracle and no baseline, so nothing about it can be scored. What the
run does report is what needs no reference: slide count, what did and did not draw, zoom invariance,
the request log, and the timings. It is served from its own directory and read once; nothing is
copied, and nothing is written. The gated run takes no deck argument at all.

## What it will not do to your machine

Reads the repository, and under `--deck` the one file you name. No Office automation, no network, no
install: the oracle and the baseline are committed, so a clone can run this.
