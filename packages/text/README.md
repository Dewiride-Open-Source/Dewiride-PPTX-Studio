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
and the reasoning is in [ADR 0028](../../docs/adr/phase-3-text/0028-measurement-and-the-line-model.md).

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

## Autofit

Sub-phase 3.4, measured in experiment T4 — 3521 probes across twenty-eight packages, by driving
`TextFrame2.AutoSize` to force a recomputation and reading the `@fontScale` and `@lnSpcReduction`
PowerPoint wrote into its own file. See
[ADR 0030](../../docs/adr/phase-3-text/0030-autofit.md).

**Two modes, and the plan was right about both.** A stored scale is applied exactly as written —
never recomputed on open, never snapped to a rung, so a `fontScale="45000"` from another producer
stays 45%. A recomputation walks a **fifteen-rung ladder** and takes the first rung that fits.

The ladder's font scales step by 7.5 points of a percent from 100 down to 25, and the reduction is
20% for all of them except the top three. The order is font scale descending, then the smaller
reduction first — **not** the tallest rung that fits, which twenty-two boxes of a 32pt sweep refute,
because rounding can make a smaller font scale the taller rung.

Four things a plausible implementation gets wrong:

- `round(size × fontScale)` is to a **whole point**, so 18pt at 92.5% draws at 17pt — but only when a
  scale applies. A `fontScale` of exactly 100% leaves 10.5pt alone.
- `@lnSpcReduction` is **subtracted from the `a:lnSpc` percentage**, not multiplied into the advance.
  The same number at single spacing, four per cent apart at 150%.
- An exact-point `a:lnSpc` **ignores the reduction entirely**, so such a paragraph cannot be
  autofitted shorter at all.
- `a:spcPts` paragraph spacing does **not** follow the font scale; `a:spcPct` does.
  `spcFirstLastPara` defaults false.

`a:spAutoFit` sets the shape height to the text height plus both insets, times a constant just under
one per cent over — the same on four faces, and it scales the insets too, so it is not a font metric.

**Autofit runs in PowerPoint's layout path.** A COM harness that opens a deck without a window
measures nothing here, and reports it as "PowerPoint never shrinks".

## Bullets, fields and script runs

Sub-phase 3.5, measured in experiment T5 — 6849 probes across 44 packages. The instrument is new and
worth stating: `Slide.Export(path, "EMF")` records PowerPoint's own GDI calls, so a bullet comes back
as **the characters it drew**, with the face, size and colour it drew them in. Nothing is fitted.
See [ADR 0031](../../docs/adr/phase-3-text/0031-bullets-fields-and-script-runs.md).

```ts
import { formatAutonumber, numberParagraphs, symbolBulletChar } from '@pptx-studio/text';

formatAutonumber('romanLcParenBoth', 4); // "(iv)"
formatAutonumber('alphaLcPeriod', 27); // "aa." — not "ba."
symbolBulletChar('•', 'Wingdings'); // U+F095, a filled circle — not U+F022
numberParagraphs(paragraphs); // the numbers, which are in no file
```

### The 41 schemes

The list and its order are PowerPoint's own: its object model was walked from 0 to 47 and it wrote
each accepted style into a package. Four things a careful implementation gets wrong:

| question                 | measured                                     | refuted               |
| ------------------------ | -------------------------------------------- | --------------------- |
| what follows `z`         | `aa`, then `aaa` — repeat, not carry         | bijective (1396/5097) |
| the East Asian zero      | **U+25CB**, not U+3007                       | 5001/5097             |
| the ten-form             | three families, three rules                  | one rule (5069/5097)  |
| circled numbers past ten | Wingdings wraps; Unicode falls back to ASCII | 5053, 5079            |

The counter also **wraps** — at 780 for five alphabets and 392 for Hebrew — Thai carries an
off-by-one nobody else has, and Thai numbering uses 41 of its 44 consonants. None of that is
reachable by a real deck; it is measured because `ST_TextBulletStartAtNum` runs to 32767.

### A symbol bullet goes through the ANSI codepage

`buChar char="•" buFont="Wingdings"` draws **U+F095**. Adding `0xF000` to U+2022 gives U+F022, an
envelope. The rule is the character's **Windows-1252 byte** plus `0xF000`, and `•` is byte 0x95 there
and nothing at all in Latin-1 — so the plausible reading is wrong on exactly the character every deck
uses. A text face is left alone.

### `a:buFont` is inert for a number

An autonumber is drawn in the **first run's** face whatever `buFont` says — and PowerPoint's own UI
writes `buFont="+mj-lt"` on every list it numbers. `a:buClr` and `a:buSzPct` _do_ reach one, so the
three decorations merge as three slots rather than travelling with the kind.

### The numbers are not in the file

A six-item numbered list PowerPoint authored itself carries no `startAt` anywhere. A run is
consecutive paragraphs at one level sharing a scheme and a `startAt`; a **deeper** level passes
through it and a **shallower** one ends it. So `1, 2, [1, 2], 3, 4` — and `1, [1], 2, [1]`, because
the inner list is re-entered. A `startAt="7"` followed by three plain paragraphs renders `7, 1, 2, 3`.

### Where it goes

`max(marL, −indent, bulletLeft + advance)`, with the bullet at `max(0, marL + min(0, indent))`. A
**positive** indent moves nothing, and the bullet is clamped to the frame. A wrapped line is at
`marL` exactly. A picture bullet is 0.7 × the font size tall and keeps its aspect ratio.

### Script runs

Private use to `a:sym`; Hebrew, Arabic, Thai and Devanagari to `a:cs`; kana, Han, Hangul, fullwidth
and CJK punctuation to `a:ea`; everything else — Greek and Cyrillic included — to `a:latin`. An
absent slot falls straight to `a:latin`. `a:cs` is not "everything non-Latin" (8/21), and `a:sym`
really is used (never using it scores 20/21, which looks right on a whole deck and is wrong on every
Wingdings run in it).

### Fields: the plan's design did not survive

`a:fld` was specified as fourteen types formatted with `Intl.DateTimeFormat`. Measured:

- **eighteen** reserved types, including `uaqdatetime1`/`uaqdatetime2` — the Umm al-Qura calendar —
  which no enumeration this project has read mentions;
- **`Intl` reproduces PowerPoint on 25 of 97 readings.** It formats with Windows' own per-locale
  patterns, which is why `datetime12` in German is `"12:56 "`: the pattern ends in an AM/PM
  designator German defines as empty;
- the cached `a:t` is **discarded on open** for every reserved type, in 290 readings, and shown only
  for a type nothing reserves.

So the patterns are a measured table — 172 cells over 14 locales, each verified by re-rendering it
against the string it came from — and `renderField` shows the cached text where a cell is missing,
which is a visibly stale date rather than a confidently wrong one.

## The text frame

Sub-phase 3.6, experiment T6: 601 probes, 21 questions, each fitting exactly one candidate.
[`corpus/ground-truth/frames.json`](../../corpus/ground-truth/frames.json) and
[ADR 0032](../../docs/adr/phase-3-text/0032-anchors-insets-and-vertical-text.md).

### The anchor works inside the inset box

`t` puts the block at `tIns`, `ctr` at `tIns + (H − tIns − bIns − block)/2`, `b` at
`H − bIns − block`. Ignoring the insets scores 120 of 190 — and it scores 120 because it is right
whenever `tIns` equals `bIns`, which is every deck that leaves the default alone. Every anchor probe
is therefore run twice, once at zero insets and once at `tIns=20, bIns=5`; the naive reading is
wrong on all 70 of the second set.

The insets default to **7.2 / 3.6 / 7.2 / 3.6 points, per attribute** — a frame stating only `lIns`
still gets 3.6 above and below. A negative inset is honoured, not clamped. Two opposing insets that
exceed the frame collapse the content box to the **midpoint between them**, which an asymmetric
probe separates from the frame's own centre.

### `just` and `dist` are bottom

In all 190 anchor readings and all 56 line-position comparisons, `anchor="just"` and `anchor="dist"`
place the block exactly where `anchor="b"` does, with the lines inside it at exactly the same
offsets. They stretch nothing. `TextFrame2.VerticalAnchor` reports 0 for both, which is not a member
of `MsoVerticalAnchor` — PowerPoint has no name for what it just read.

A renderer that implements "distribute" from the name will spread the lines of every deck that
carries the value.

### `spcFirstLastPara` reaches the last line

3.4 measured that the flag controls the first paragraph's `spcBef` and the last's `spcAft`. It also
decides how much of the last line's leading the block keeps: off, `(n−1) × advance + lastLine`; on,
`n × advance`. 60 of 60 across five line spacings. At 100% the two are identical, which is why 3.4
could not see it and why `blockHeight` now takes the flag.

### The two axes swap, and the vertical types disagree

`@anchor` positions the block along the axis lines **stack** in; `@algn` positions each line along
the axis text **runs** in. `@vert` swaps them, and the start edge is a property of the value:

| `@vert`          | lines stack | `@anchor` from | `@algn` from | `@`-face | one record per glyph |
| ---------------- | ----------- | -------------- | ------------ | -------- | -------------------- |
| `horz`           | downward    | top            | left         | no       | no                   |
| `vert`           | leftward    | right          | top          | no       | no                   |
| `eaVert`         | leftward    | right          | top          | yes      | no                   |
| `vert270`        | rightward   | left           | bottom       | no       | no                   |
| `mongolianVert`  | rightward   | left           | top          | yes      | no                   |
| `wordArtVert`    | rightward   | left           | top          | no       | yes                  |
| `wordArtVertRtl` | leftward    | right          | top          | no       | yes                  |

**The insets do not rotate with the text.** `lIns` insets from the physical left edge whatever
`@vert` says — 7 of 7, against zero for the reading that turns the frame and its insets together.

### `@vertOverflow` changes what is drawn, never the layout

`ellipsis` **replaces** the whole first line that does not fit with a lone U+2026; it does not
truncate the last line that does. `clip` stops laying out rather than drawing and masking. The
block's height and position are identical in all 18 readings. A line whose bottom lands exactly on
the frame edge is **not** drawn.

### `a:br`, and the empty paragraph

No new bullet, no advanced number, no `spcBef` or `spcAft`, and no first-line indent — the line
after a break starts at `marL`, which the positive-indent probe proves. Its own `a:rPr` does not set
the height of the line it begins; the following run does.

An empty `a:p` is 1.2 × the `a:endParaRPr` size times its own `a:lnSpc`, and `a:endParaRPr` beats
`a:defRPr` when the two disagree — measured both ways round.

### `a:bodyPr` inherits

Every attribute inherits on its own, through the placeholder chain, from any level. A body
placeholder writing a bare `<a:bodyPr/>` is anchored where its master says, not at the schema
default. 24 readings against 12 for both rivals.

## What is not here yet

The **last** line of a block does not measure the same as the lines above it — 3.4 had to settle
that to build a fit test, and it is `0.75 × advance + b × size`, floored at the advance while the
advance is inside `C × size`. `b` and `C` are properties of the typeface: six faces are measured and
committed, `faceLineMetrics()` **throws** for a seventh, and `APPROXIMATE_FACE_METRICS` has to be
named to be used. Chromium's `TextMetrics` cannot supply either — Arial and Verdana report the same
`fontBoundingBoxDescent` while their coefficients differ by ten times that. Sourcing them from the
font itself is 8.1.

**Thai over-runs.** PowerPoint segments Thai with a dictionary — it breaks `สวัสดี` from `ชาวโลก`, two
words, and nowhere else. No rule over character classes produces that and no JavaScript library
carries the dictionary, so a Thai line runs past its box rather than breaking. The gap is measured
rather than hidden: the miss count is pinned in the fixture and the suite fails if it grows.

**Arabic breaks at spaces only**, which is correct, but note that Arabic prefix widths are _not
monotone_ — `مرحبا` measures narrower than `مرحب`, one letter shorter, because the fifth letter joins
the fourth into a shorter connected form. Anything measuring prefixes here must not assume otherwise;
a binary search over that sequence returns an arbitrary element of it.

**Two autonumber schemes throw.** `arabic1Minus` and `arabic2Minus` are real, and PowerPoint shapes
them into glyph ids before drawing — so the EMF that identifies every other scheme holds indices into
a font and no text. The repeat structure is in the fixture; the letters are not, and guessing them
from the Unicode block would be a sequence nothing has ever checked. Reversing a `cmap` is 8.1.

**118 date-pattern cells are missing**, mostly where a locale's calendar is not Gregorian or where
ICU and Windows disagree on a component's spelling. `renderField` shows the cached text for those.

Anchoring is 3.6 — which inherits a smaller question than it was given: not what the last line
measures, but where the space goes when `spcFirstLastPara` is on, since the block grows a `spcBef`
taller without the first line moving off the frame top. Alignment is 3.6's too: T5 shows a centred
paragraph moves the bullet with the text, so the bullet is part of the aligned line, but its probes
carry a marker run that contaminates the width and nothing quantitative is claimed.
