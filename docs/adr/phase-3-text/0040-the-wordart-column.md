# 0040 — The WordArt column, and the lines that stacked the wrong way

- **Status** accepted
- **Sub-phase** 3.6, reopened a second time
- **Experiment** T11 — 67 probes across 9 slides, PowerPoint driven over COM, read
  back through an EMF drawing stream and a bitmap at four pixels to the point
- **Fixture** [`corpus/ground-truth/wordart-columns.json`](../../../corpus/ground-truth/wordart-columns.json)
- **Code** `packages/text/src/frames/frame.ts`, `packages/render-svg/src/text/layout.ts`

## Context

[ADR 0039](0039-upright-glyphs-in-vertical-text.md) left `wordArtVert` and
`wordArtVertRtl` refused, because "their per-glyph column is a font metric
nothing has measured". That was the last thing standing between the corpus and a
complete score: `a10-rtl-cjk-03` carried a `wordArtVertRtl` frame and so was not
drawn at all.

The metric turned out to be measurable in one afternoon, and measuring it turned
up a second bug in the turn code — the same shape as the inset bug 0039 found,
and for the same reason: the fixture had the answer and no test read it.

## The cell is square, and it is seven sixths of the font box

A WordArt frame stacks one character per cell. The cell is **square** — the pitch
down a column is the width of the column, 49 of 49 — and its side is **seven
sixths of the font box**, the box being `usWinAscent + usWinDescent` over the em:

| face            | column    | 7/6 of the box | error       |
| --------------- | --------- | -------------- | ----------- |
| Arial           | 1.3034 em | 1.3032         | **0.0002**  |
| Courier New     | 1.3216    | 1.3218         | **−0.0003** |
| Times New Roman | 1.2919    | 1.2915         | **0.0004**  |
| Verdana         | 1.4179    | 1.4175         | **0.0004**  |
| Segoe UI        | 1.5517    | 1.5517         | **0.0001**  |
| MS Gothic       | 1.1668    | 1.1667         | **0.0001**  |
| Malgun Gothic   | 1.5517    | 1.5517         | **0.0001**  |
| Yu Gothic       | 1.7597    | 1.5015         | 0.2582      |
| SimSun          | 1.2113    | 1.1667         | 0.0446      |

Seven faces at four ten-thousandths of the em, across three sizes each, is a rule
and not a coincidence. The rivals are scored and every one of them misses every
row: the em (MS Gothic's cell is 1.1667, not 1), 1.2 of the em (which is what a
**horizontal** line advances by), the font box itself, and the widest advance in
the run — Courier New's advances are all 0.6 em and its cell is 1.3216.

Inside the cell, two numbers place the glyph, and both are exact on all nine
faces to within a fortieth of the em — finer than the drawing stream writes a
pen. The baseline sits the face's **descent up from the cell's floor**. Across
the cell the glyph's **advance box is centred**. Neither reading needed a
tolerance chosen for it; the fixture records what the pens are quantised to and
the tests are held to that.

The glyph itself is drawn **upright and unturned**, from the plain face. `一` at
96pt inks 81.75 by 5.25 points in a `wordArtVert` frame — the same rectangle, to
the quarter point, as in a horizontal one, 6 of 6 against the bitmap. GDI is
never asked for the `@`-prefixed variant, 53 of 53, which is the whole difference
from `eaVert`: that direction turns the line and stands the glyph back up, and
this one never turns the glyph at all.

### Two faces are recorded rather than ruled

Chromium reports a font box for **Yu Gothic** and **SimSun** that is not the box
PowerPoint measured from — 1.2870 against an implied 1.5083, and 1.0000 against
1.0379. It is the same face: the advances agree. No browser-visible metric closes
the gap; `fontBoundingBox` is text-independent, `line-height: normal` resolves to
the same number, and no family variant (`Yu Gothic UI`, `Yu Gothic Light`,
`NSimSun`, `宋体`) reports the other one. Their columns therefore come out
0.2582 and 0.0446 of the em narrow, and the fixture says so rather than widening
a tolerance until it does not.

This is the same ceiling 0039 hit from the other side, and it is worth stating
plainly: the rule is exact, the _input_ to it is not available in a browser for
every face.

## Two sizes in a column separate what one size cannot

A column set in one size cannot tell a cell that belongs to the run from a cell
that belongs to the column, and the first implementation guessed — it took the
widest cell in the paragraph and used it for every glyph. A mutant that replaced
that with the first run's cell survived the whole suite, which is exactly what a
surviving mutant is for.

Two probes settle it. `sizes-Arial` sets `Wx` at 12pt and `yz` at 28pt in one
paragraph, and its pens are 13.08, 28.68, 61.68, 98.04:

- **each run stacks in a cell of its own size**, 2 of 2 — 15.60 between the two
  12pt glyphs against a 15.638 cell, 36.36 between the two 28pt ones against
  36.489. Neither "the widest run's cell" nor "the first run's cell" fits.
- **every glyph is centred across the column**, which is the widest cell in it, 2
  of 2 — the 12pt `W` sits 12.36pt in, where its own cell would put it 2.16.

So the two axes disagree about which box they follow, and only a mixed column
shows it.

## The lines stacked the wrong way, and that was a bug

Laying `wordArtVert` out meant reading the turn code again, and its lines came
out right to left where PowerPoint's go left to right. `vert-wrap-*` has had the
answer in the fixture since 3.6:

| direction        | `lineLefts`                      |
| ---------------- | -------------------------------- |
| `vert`           | 278400, 256800, 235200, … inward |
| `eaVert`         | 278400, 256800, 235200, … inward |
| `vert270`        | 0, 21600, 43200, … outward       |
| `mongolianVert`  | 0, 21600, 43200, … outward       |
| `wordArtVert`    | 0, 23460, 46920, … outward       |
| `wordArtVertRtl` | 276540, 253080, … inward         |

A quarter turn clockwise sends the first line to the **right**, so `vert` and
`eaVert` are free and `mongolianVert` is not — its stack has to be mirrored
inside the block. It was not, so **`mongolianVert` has been drawing its columns
in reverse order since 3.6**, and `wordArtVert` would have inherited the same
mistake. Nothing caught it because the only multi-line vertical probes are
`vert-wrap-*`, which `frameProbes` holds out, and because reversing a stack
leaves the block's own corner exactly where it was — the one thing the T6 test
compared.

`backwards` in the turn table now mirrors those stacks, and a test lays out every
wrapped probe of every direction and asserts the **step between two lines** has
PowerPoint's sign. That assertion, and not the corner, is what separates the two.

## Every `ST_TextVerticalType` is now drawn

With the two WordArt directions in, the refusal branch had no reachable input
left, so `DRAWN_DIRECTIONS` and the `RENDER_TEXT_UNSUPPORTED` throw beside it are
gone and the turn table is keyed by `VerticalText`, which makes the compiler the
thing that notices if a value is ever added.

## Verification

- **`pnpm fidelity`: 155 slides scored, 0 not drawn**, where 3.9 left 152 scored
  and 3 refused and 0039 left 154 and 1. The corpus mean is 9767 → 9768 basis
  points even though the slide that joined it is a hard one; `a10-rtl-cjk-03`
  scores 9920 against 9945 and 9917 for its two siblings.
- **11 of 11 mutants killed**: the seven sixths, the baseline's sign and its
  edge, the centring and which box it follows, both WordArt turns and
  `mongolianVert`'s, the column's capacity, the per-glyph advance, the per-run
  cell, and the column's width.
- 3256 → **3265 tests**, 91 files, driven from `wordart-columns.json` and
  `frames.json` and never from the implementation.
- `pnpm check` green.

## Deviations from the plan

| plan item                   | disposition | why                                                                                                                      |
| --------------------------- | ----------- | ------------------------------------------------------------------------------------------------------------------------ |
| WordArt renders approximate | **exact**   | The plan allowed an approximation and 0039 refused instead. Neither was necessary: the cell is a fitted rule on 7 faces. |
| WordArt is non-editable     | **kept**    | Nothing here makes a WordArt frame editable; it draws, and 3.6 claimed no more than that.                                |

## Open questions

1. **Chromium's font box is not GDI's on Yu Gothic and SimSun.** It is the same
   disagreement 0039 found in `ideographicBaseline`, and it now costs a visibly
   narrow column rather than a fraction of a glyph. Whether a face's `head` table
   could be read from an embedded font, and whether anything short of that helps
   a face that is merely installed, is unanswered.
2. **The cell is rounded before it is stacked**, by a hundredth of the em at 12pt
   and a two-hundredth at 18 and 28. What it is rounded _to_ is not established —
   no device-unit or whole-pixel reading fits all three sizes — so the residual
   is recorded and the tests are held to it rather than to a rule.
3. **An underline under a WordArt column throws.** Same gap as 0039 left for
   upright East Asian text, and for the same reason: nothing measured whether the
   rule runs beside the column or under each glyph.
4. **`mongolianVert`'s column order was wrong for four sub-phases** and the
   corpus never noticed, because no corpus slide has a wrapping Mongolian frame.
   The probes that caught it are synthetic; a real deck would be better.
