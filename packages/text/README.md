# `@pptx-studio/text`

Text measurement and the line model: how far apart two baselines sit, and how wide a run is (3.2).

Every rule here was measured against Microsoft PowerPoint rather than derived from the standard,
because on the questions that decide a line count the standard either says nothing or says something
PowerPoint does not do. ECMA-376 never mentions a line height at all.

```ts
import { createCanvasMeasurer, lineAdvance, lineTop } from '@pptx-studio/text';

// <a:pPr><a:lnSpc><a:spcPct val="150000"/></a:lnSpc></a:pPr> on an 18pt run
lineAdvance(1800, { kind: 'percent', value: 150000 }); // 32.4pt, not 27
lineTop(2, 32.4); // 64.8 - the third line's top, measured from the frame top

const measurer = createCanvasMeasurer();
measurer.measure('Hamburgefonstiv', { family: 'Arial', sz: 1800 }); // { width: 136.06 }
```

## What was measured

| question                       | answer                                                          | refuted                                              |
| ------------------------------ | --------------------------------------------------------------- | ---------------------------------------------------- |
| the line box at single spacing | a font-independent **1.2 × font size**                          | anything font-derived: 13 faces, 7 sizes, one number |
| `a:spcPct`                     | multiplies the **line box** — 150% at 18pt is 32.4pt            | multiplying the font size (66/463)                   |
| `a:spcPts`                     | used **raw**, no 1.2 anywhere                                   | scaling it by 1.2 too (397/463)                      |
| `@val` precision               | **rounded half up to a whole percent** — `106667` is 107%       | ceiling (365/371), floor (359/371), exact (353/371)  |
| where a block starts           | the **frame top exactly**, no half-leading above the first line | CSS-style centred leading                            |
| a block's height               | `(n − 1) × advance + lastLineHeight`                            | `n × advance`                                        |
| `a:rPr/@spc`                   | **absolute points**, after every character including the last   | between characters only (0/18); em-relative (0/18)   |
| `a:rPr/@kern`                  | a **minimum font size, inclusive**; `0` means never             | a boolean (5/12 and 8/12)                            |

The fixture is [`corpus/ground-truth/text-metrics.json`](../../corpus/ground-truth/text-metrics.json)
and the reasoning is in [ADR 0028](../../docs/adr/0028-measurement-and-the-line-model.md).

## Never `line-height: normal`

A browser resolves `normal` from the font — `hhea.lineGap` on macOS, `OS/2.usWinAscent` on Windows —
so the same markup gives **different line counts on different machines**. PowerPoint asks the font
nothing. That is the single most important thing this package exists to prevent, and it is why
`lineAdvance` takes no font argument at all: there is nowhere to pass one.

## How close measurement gets

Over 364 comparisons on installed faces, Chromium's `measureText` and PowerPoint disagree by a
**median of 0.15%** (p95 0.70%). The outlier is Cambria at 12pt, where PowerPoint hints small text
and Chromium does not: 11.1%.

That number is tracked rather than claimed away — the plan promises "a tracked, improving number,
never pixel-perfection" — and the suite fails if it regresses.

## What is not here yet

The **last** line of a block does not measure the same as the lines above it, and the difference
needs a per-face term that no browser reports and that no measurement here can source. It is
committed as data under `lastLine` in the fixture, with the best model found, its exact score, and
every probe it misses. Finishing it is sub-phase 3.6, where anchoring needs it.

Line breaking is 3.3, autofit 3.4, bullets and fields 3.5.
