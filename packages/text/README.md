# `@pptx-studio/text`

Text measurement, the line model, and line breaking: how far apart two baselines sit, how wide a run
is (3.2), and where a line ends (3.3).

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

## Line breaking is not UAX#14

That is a finding, not a shortcut. The plan called for UAX#14 through the `linebreak` package with
DrawingML tailoring layered on top; experiment T3 refutes the shape of that design. What PowerPoint
does is not a _tailoring_ of UAX#14 but a much smaller rule that UAX#14 is a superset of. Scored over
the same 2212 probes:

| reading                                            | fits          |
| -------------------------------------------------- | ------------- |
| **measured**                                       | **2212/2212** |
| kinsoku applied to every break, as ECMA implies    | 2200          |
| no hanging punctuation                             | 2157          |
| no emergency break (an unbreakable word overflows) | 2085          |
| kinsoku off                                        | 2010          |
| break only after spaces                            | 1898          |
| a faithful UAX#14 reading                          | 1878          |

So there is no Unicode line-break table in this package and no dependency that carries one. A line
may end after a run of spaces, after a hyphen-minus, en dash, em dash or soft hyphen — **only after**,
so the position before an em dash is not an opportunity — between two East Asian characters, and at a
script transition. Nowhere else:

| a break after…            | PowerPoint | UAX#14 |
| ------------------------- | ---------- | ------ |
| `/` in `alpha/beta`       | no         | yes    |
| anything in a URL or path | no         | yes    |
| ZERO WIDTH SPACE          | **no**     | yes    |
| a soft hyphen             | yes        | yes    |
| `-` between digits        | yes        | yes    |
| `,` `.` inside a number   | no         | no     |
| NO-BREAK SPACE, U+2011    | no         | no     |

```ts
import { breakOpportunities, wrapText } from '@pptx-studio/text';

breakOpportunities('aa alpha/beta'); // [3] - the space, and nothing else
breakOpportunities('aa well-known'); // [3, 8]
breakOpportunities('日本語abcです'); // [1, 2, 3, 9, 10] - not inside the Latin word
```

### The three tailoring attributes

`a:pPr/@latinLnBrk` **absent behaves as `false`**, where ECMA-376 gives it a default of `true`.
Following the standard breaks every long word in every deck. `@hangingPunct` absent behaves as
`true`. `@eaLnBrk` is measured **inert** — all three of absent, `1` and `0` lay out identically.

### Kinsoku is built in, and the file usually cannot reach it

PowerPoint's own sets are used **unless** a package both states `p:kinsoku` _and_ sets
`strictFirstAndLastChars="0"`, and then the file's sets replace them entirely — measured with empty
sets, where every position became a break. Stating `p:kinsoku` without the toggle changes nothing,
which matters because that is what almost every real deck does.

The sets are measured one character at a time — 71 may not begin a line, 19 may not end one, six
hang — with six unrestricted controls to show the sweep discriminates. The widely repeated version of
these sets is **wrong in four places**: it lists `¢`, `£`, `¥` and `＃` as restricted, and PowerPoint
restricts none of them while restricting the fullwidth `￡` and `￥` the list omits.

A `p:kinsoku` whose `@lang` is not East Asian makes PowerPoint **refuse the package outright** —
"the file is corrupted and unreadable", with no repair offered.

## What is not here yet

The **last** line of a block does not measure the same as the lines above it, and the difference
needs a per-face term that no browser reports and that no measurement here can source. It is
committed as data under `lastLine` in the fixture, with the best model found, its exact score, and
every probe it misses. Finishing it is sub-phase 3.6, where anchoring needs it.

**Thai over-runs.** PowerPoint segments Thai with a dictionary — it breaks `สวัสดี` from `ชาวโลก`, two
words, and nowhere else. No rule over character classes produces that and no JavaScript library
carries the dictionary, so a Thai line runs past its box rather than breaking. The gap is measured
rather than hidden: the miss count is pinned in the fixture and the suite fails if it grows.

**Arabic breaks at spaces only**, which is correct, but note that Arabic prefix widths are _not
monotone_ — `مرحبا` measures narrower than `مرحب`, one letter shorter, because the fifth letter joins
the fourth into a shorter connected form. Anything measuring prefixes here must not assume otherwise;
a binary search over that sequence returns an arbitrary element of it.

Autofit is 3.4, bullets and fields 3.5, anchoring 3.6.
