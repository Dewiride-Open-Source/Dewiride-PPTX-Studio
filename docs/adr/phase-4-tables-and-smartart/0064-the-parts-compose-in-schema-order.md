# 0064 — The parts compose in schema order

Date: 2026-09-30
Status: **accepted** — across 10,480 slides of tables in three themes, read cell by cell through COM
and pixel by pixel at 4 px a point, one reading fits every row of every question. ECMA-376 says
which parts reach a cell. The parts compose in the schema's own order, property by property: fills
10,519/10,519, text 49/49, edges 2,693/2,693. A grid edge draws the highest claim of either cell
(86,372/86,372). A direct border belongs to the cell above or left (92/92), and a merged cell owns
only the segments level with its anchor (52/52, 24/24). 33/34 mutants killed; the survivor is
equivalent on every table PowerPoint draws.

**Sub-phase 4.3.** The plan's row says "the 13-layer table cascade". 4.2 settled which style a table
is drawn with (ADR 0063). This settles what each of that style's thirteen parts gives each cell and
each grid edge. It also settles what a cell's own `a:tcPr` and the table's `a:tblPr` do to that, and
where the style sits in the walk that resolves a cell's text. ECMA-376 names the parts and says
little about how they combine. PowerPoint is the only answer that counts, so every rule here was
scored against what PowerPoint reported and drew.

Code: `packages/model/src/{table,errors,resolve/table,resolve/text,types}.ts`, `V034` in
`packages/validate/src/rules/{rules,required}.ts`, `tools/ground-truth/model/tables/cascade/`,
`tools/ground-truth/lib/png.ts`. Fixture: `corpus/ground-truth/table-cascade.json` and its siblings
`table-cascade-ends.json`, `table-cascade-sides.json` and `table-cascade-sweep-{1,2,3}.json`.
Changeset: `.changeset/the-parts-compose-in-schema-order.md`.

---

## C9 — the probes

250 packages, one table per slide, in three themes of our own:

- **Theme A** has a stock-like palette, Georgia and Verdana, and a clrMap that sends `bg1`, `tx1`,
  `bg2` and `tx2` to other slots.
- **Theme B** has a second palette and an identity map.
- **Theme C** is A with a format scheme of gradients, hard shadows and a dashed line.

| group      | decks | slides | what it asks                                                                                     |
| ---------- | ----- | ------ | ------------------------------------------------------------------------------------------------ |
| sweep      | 76    | 4,864  | every built-in, no style and an unknown GUID, 5×5, all 64 flag sets                              |
| ends       | 74    | 4,736  | every built-in on 1×1, 2×2, 1×3 and 3×1, the sixteen end-flag sets                               |
| sides      | 74    | 592    | every built-in in theme B, eight flag sets; each family's plain and Accent 1 styles read in full |
| direct     | 6     | 108    | `a:tcPr` fills and borders: conflicts, partial lines, diagonals, outer edges                     |
| merge      | 4     | 104    | merged cells across parts and bands, and borders written on merged cells                         |
| rtl        | 4     | 20     | `a:tblPr/@rtl`, with borders on either side                                                      |
| tblpr      | 1     | 8      | an `a:tblPr` fill under four styles                                                              |
| background | 1     | 18     | theme C: `tblBg`, an `a:tblPr` fill or effect, the frame larger than the grid                    |
| text       | 9     | 27     | a cell's text through every level of the text cascade, three in placeholders                     |
| control    | 1     | 3      | read first and last in every session                                                             |

`author.ps1` also has PowerPoint set borders, fills and flags itself, twelve operations, so the
analysis knows which cells PowerPoint's own writer puts them in.

Before PowerPoint was asked, `build-deck.ts` checked every pair of readings. For each pair it found
a probe that tells them apart, or searched every table from 1×1 to 6×6 under all 64 flag sets for
one that would. If the search finds a table that separates a pair the probes do not, the build
throws. The 29 pairs no table can separate are recorded in the fixture as indistinguishable and are
reported as one group wherever they both fit. Apart from the band and corner readings below, they
are three order pairs on text, none of them the schema's, and `perEdge` against `visible`.

## The instruments

- **COM**, through a compiled late-bound reader: each cell's fill and first character's colour,
  bold and face. In full decks it also reads each of the six sides and italic, size and bullet.
  A `Borders` object answers every `Item(k)` with the side it was first asked for, so each side is
  read through a fresh one.
- **Pixels.** Every slide is exported as a 3840×2160 PNG, 4 px a point, and the first slide of each
  named deck as a BMP too. `tools/ground-truth/lib/png.ts` decodes the palette PNGs PowerPoint
  writes, and each PNG is held to its BMP twin pixel for pixel. The decoded pixels are cached under
  the hash of the PNG they came from, so a re-exported slide is always decoded afresh.
- **What is compared.** A cell's paint is sampled in its interior, clear of the glyph and the grid.
  Across each grid edge's midpoint, 32 pixels are compared run for run with what each reading's
  lines would paint over the two cells' paints.
- **Partly covered pixels.** A double line's coverage was measured, not assumed: across a 4-pt
  line, `[1 1 1 1 1 0.5 0 0 0 0 0.25 1 1 1 1 1]`. Where a line covers part of a pixel, the pixel
  counts as drawn when it is a blend of that line over what lay beneath it, strictly between the
  two: a whole or a bare pixel never passes for a partial one. How much of a pixel a rasteriser
  covers is 4.4's question. Which pixels are whole, partial or bare stays exact here.
- **Guards, before anything is scored.** Every package opened unrepaired. COM names the style that
  was built and places the table at the origin at the grid's size. Every opaque fill COM reports is
  the paint the interior shows, and every text colour it reports is the glyph's. All 10,480 slides
  passed.

### Sessions

Every session reads the control deck first and last, and the analysis refuses a deck read in a
session that did not. Every read of the control, in every session, must be the same COM and the
same export, byte for byte.

1. The first session read all 248 decks, before the merged-border probes existed.
2. The second read all of them again with those probes. On the 10,370 slides both sessions read,
   the COM readings were byte-identical and the pixels identical, and every finding the new probes
   did not touch scored the same, row for row.
3. A third was begun to read everything a third time, for probes mutation testing asked for. I
   stopped it after 25 decks. A full read costs two hours per added probe, and the second session
   had already answered what reading everything in one session guarded against. Its readings are
   discarded, and the 25 decks whose exports it overwrote were read again.
4. Session `b` read those 25 decks, the four merge decks and the three new text decks: 32 in 25
   minutes.
5. Session `c` read the four merge decks once more, in four minutes, after the band probes.

The fixture holds sessions two, `b` and `c`.

## Which parts reach a cell

| reading           | fill          | text    | edge        |
| ----------------- | ------------- | ------- | ----------- |
| **spec**          | 30,714/30,714 | 770/770 | 2,825/2,825 |
| cornersByPosition | 30,714/30,714 | 0/770   | 784/2,825   |
| cornersOnEither   | 30,714/30,714 | 385/770 | 988/2,825   |
| bandsFromZero     | 5,748/30,714  | 770/770 | 2,083/2,825 |
| rowBandsOverEnds  | 15,374/30,714 | 770/770 | 2,825/2,825 |
| colBandsOverEnds  | 15,358/30,714 | 770/770 | 2,783/2,825 |
| rowBandsOnEnds    | 30,714/30,714 | 770/770 | 2,825/2,825 |
| colBandsOnEnds    | 29,526/30,714 | 770/770 | 2,783/2,825 |

ECMA-376's reading fits all three. An end part reaches its row or column when its flag is on, and
bands are counted past a first row or column and skip the ends. A corner needs both of its flags.
Every reading but one fits one property and fails another. A renderer checked only against fills
would be free to pick the corner reading anybody writes first, and would draw every table with one
flag on and the other off wrong in its text.

The one that fits everything is `rowBandsOnEnds`: row bands also reaching the end rows, counted as
ECMA counts. No drawable table tells it from ECMA's reading. Over every built-in, every table from
1×1 to 6×6 and every flag set, each built-in's `firstRow` and `lastRow` state whatever its row bands
do. The column version is refuted (29,526/30,714). The model implements ECMA's reading.

## The order they compose in

| order         | fill          | text  | edge        |
| ------------- | ------------- | ----- | ----------- |
| **schema**    | 10,519/10,519 | 49/49 | 2,693/2,693 |
| spreadsheetml | 10,023/10,519 | 0/49  | 2,305/2,693 |
| colsOverRows  | 8,923/10,519  | 0/49  | 845/2,693   |
| bandsOverEnds | 1,459/10,519  | 0/49  | 1,842/2,693 |

The schema's element order is the order PowerPoint composes in, lowest first: `wholeTbl`, the
bands, `lastCol`, `firstCol`, `lastRow`, `seCell`, `swCell`, `firstRow`, `neCell`, `nwCell`. It
is not SpreadsheetML's table-style order, which is the rival anyone who has written an Excel reader
reaches for. That rival is right on 95% of fills and wrong wherever a first row meets a last column.

**Per property, never per element.** The highest part that states a fill wins the fill. The
highest that states bold wins bold, whatever else it leaves unsaid (fill 13,172/13,172, text
40,419/40,419, edges 142,979/142,979; the whole-element reading 0 in each).

**The text colour** is the part's explicit colour, never its `a:fontRef`'s (146,299/146,299
against 0). All 160 parts in the built-ins that give a `fontRef` a colour also state an explicit
one. PowerPoint draws nothing but its built-ins (ADR 0063), so a part whose only colour is its
`fontRef`'s is never drawn, and the model reads the explicit colour alone.

**Some properties never disagree.** Every `b` the 74 built-ins state is `on`, and none states `i`.
So which of two applicable parts wins bold can be read in no table PowerPoint draws. The order above
decides it in the model, as it does for every other property.

## Grid edges

| model         | fits          | what it does                                        |
| ------------- | ------------- | --------------------------------------------------- |
| **perEdge**   | 86,372/86,372 | every part of either cell competes for the edge     |
| visible       | 86,372/86,372 | each cell's own side first, then the one that draws |
| heavier       | 86,260/86,372 | each cell's own side first, then the heavier        |
| stacked       | 81,902/86,372 | both sides drawn                                    |
| lowerRight    | 71,218/86,372 | the cell below or right always                      |
| upperLeft     | 66,901/86,372 | the cell above or left always                       |
| tableRelative | 48,033/86,372 | a part's outer edges are the table's                |
| cellRelative  | 38,339/86,372 | a part's edges are each cell's own                  |

`perEdge` and `visible` differ on no drawable table (the search above), and the model implements
`perEdge`. The highest part of either cell that states a line for this edge draws it. A part that
covers both cells offers its `insideH` or `insideV`; a part on one side only offers the side that
faces the edge. The one exception is a merged cell, below.

## Paint

A translucent fill is painted alone, over the style's `tblBg` over the page. Each channel's two terms
are rounded on their own, ties down, before they are summed (25,494/25,494). Rounding the sum
instead fits 20,491. Painting every part's fill in order fits 19,934. Ignoring `tblBg` fits 3,178.

An `a:tblPr` fill replaces `tblBg` under the cells (120/120). Painting it over them fits 60/120, and
ignoring it fits 30/120. Theme C's gradients and shadows under `tblBg` were sampled across and around
the table and are recorded in the fixture, not scored: a gradient's extent and a shadow's offset are
4.4's.

## The default grid

A table whose style resolves to nothing (ADR 0063) draws a 1-pt black line on every edge
(7,680/7,680; `tx1` and none 0). Its text is `tx1` in the theme's minor face, whatever the flags
(3,200/3,200). Taking the text cascade's floor fits 0, and bold on the end rows fits 2,048. This
answers ADR 0063's first open question.

## Direct formatting

- **Fill.** A cell's own `a:tcPr` fill wins over every part (15/15).
- **Which line an edge draws.** Every edge has an owner, whose `a:lnB` or `a:lnR` (or `a:lnT` or
  `a:lnL`) draws the edge whatever the other cell writes (92/92). Owner-then-other fits 89, the
  heavier 68, the later 56 and the style alone 18. For unmerged cells the owner is the cell above
  or to the left, and the cell below or right only on the table's own top and left edges. So an
  `a:lnT` or `a:lnL` inside the table draws nothing. Merged cells refine who owns an edge, below.
- **A partial line replaces the style's line whole** (26/26; merging attribute by attribute fits 0).
  A line with no width draws 9,525 EMU, the shape default, and a line with no fill draws nothing.
- **PowerPoint writes every side alike.** Whenever it sets a border it writes the line on both
  cells, and for a merged cell on every cell along the side (`author.ps1`, twelve operations).

## Merged cells

**Which parts a merged cell takes.** Its anchor's parts, and the end and corner parts of every
position it covers; the bands come from the anchor alone. A row-spanning cell anchored on a
second-band row that covers a first-band row paints as the second band does, not the first.

| reading                                | paint   | edges |
| -------------------------------------- | ------- | ----- |
| **the anchor's, and every end pooled** | 126/126 | 48/48 |
| every position's parts pooled          | 86/126  | 48/48 |
| the anchor's alone                     | 110/126 | 12/48 |
| each position's own                    | 26/126  | 0/48  |

The edges column is scored with end parts offering the edge their kind faces: `firstRow` its
bottom, `lastRow` its top, `firstCol` its right, `lastCol` its left, on either side of an interior
edge. Without that, both pooled readings fit 23/48.

**Who owns an edge beside a merged cell.** Probed with sixteen slides of direct borders on merged
cells in each merge deck, each question scored with the other two held at their winners:

| which segments a merged cell before an edge claims | fits  |
| -------------------------------------------------- | ----- |
| **those level with its anchor**                    | 52/52 |
| its whole side                                     | 40/52 |
| those beside its anchor's own position             | 36/52 |
| each position by its own `a:tc`                    | 12/52 |

| who draws a segment the cell before does not claim     | fits  |
| ------------------------------------------------------ | ----- |
| **the cell after where it is level, else the before**  | 24/24 |
| the cell after, else the cell before on the table edge | 20/24 |
| the cell after, else the style                         | 12/24 |
| the cell after wherever the position is covered        | 8/24  |
| the style                                              | 4/24  |

| how far the cell after reaches on the table's top and left | fits |
| ---------------------------------------------------------- | ---- |
| **its whole side**                                         | 8/8  |
| the segments level with its anchor                         | 0/8  |
| each position by its own `a:tc`                            | 0/8  |

So a merged cell is **level** with a segment of its side when the segment is in its anchor's column
(a horizontal edge) or its anchor's row (a vertical one). The cell before an edge owns a segment
where it is level with it. Otherwise the cell after owns it, where that cell is level with it.
Otherwise, and where no cell comes after, the cell before keeps it. For unmerged cells this is the
owner rule above, because an unmerged cell is always level.

In practice: a row-spanning cell's `a:lnR` draws beside its first row only, and the `a:lnL` of the
cell to its right draws beside the others. A column-spanning cell's `a:lnB` draws under its first
column only, unless it is on the table's bottom edge, where it draws along the whole side. Two tall
cells side by side keep the left one's `a:lnR` beside both rows, because neither is level below the
first. PowerPoint's own writer does exactly what these rules need: for a merged cell's right border
it writes the anchor's `a:lnR` and an `a:lnL` in each cell to the right, and for a wide cell's
bottom an `a:lnT` in each cell below; on the table's bottom edge it writes the anchor's `a:lnB`
alone. A covered cell's own `a:tcPr` is never read, for its lines or its fill.

## Right to left

`a:tblPr/@rtl` mirrors the logical table: every part, band and border is decided on logical columns,
and the drawing is then mirrored (256/256; laying the parts over the visual columns fits 0). Direct
borders too: a cell's `a:lnL` draws on its logical left, its visual right (8/8; on its visual left
0/8).

## A cell's text

Nine decks put a table's text on each rung of the text cascade, 2×6 cells each, with a floor of
Courier New in `hlink` at `p:defaultTextStyle` and `p:otherStyle` where the deck declares them.
What the style gives the text sits **below the cell's own list style** (240/240, against 216 for the
floor and 132 or fewer for the other positions). The walk ends at the master's `p:otherStyle` alone
(66/66). `p:defaultTextStyle` is never read for a table cell (the chain 0/66; other-then-default
60/66). The style reaches every list level (18/18; the first level only 0).

Three of the decks put the table in a content, a table and a title placeholder on a master with no
`p:txStyles`, where PowerPoint substitutes its built-in text styles. A cell takes the built-in
`other` style there, never its frame's bucket (78/78 against 0): a table in a body placeholder is 18
points without bullets, not 28 with them.

So `TextContext.cell` gives a cell its own walk. First its `a:lstStyle`, `origin: 'cell'`, explicit
as a shape's own list style is. Then the style's layer, run properties at every level, `origin:
'tableStyle'`. Then the master's `txStyles.other`. The frame's placeholder chain is not read.

## The model

- `tablePartsAt` and `TABLE_PART_ORDER` are the two findings above; a merged cell pools as measured.
- `tableCellFill`, `tableBackground` and `tableEdgeLine` each resolve one property and return a
  `TableSourced` value naming where it came from: a part, `tblBg`, `tcPr`, `tblPr` or `grid`. A
  property nobody states is `null`, not a default.
- `tableTextLayer` gives the style's bold, italic, faces and colour for a cell, each from the
  highest part that states it, as the run properties the text cascade reads at `origin:
'tableStyle'`; what no part states is left `undefined`.
- `tableEdgeLine` gives each segment to its owner by the level rule, and returns `null` inside a
  merged cell.
- `themedFill` and `themedLine` follow a `fillRef` or `lnRef` into the theme's format scheme and
  return the colour `phClr` takes there.
- `DEFAULT_GRID_LINE` is the default grid's line.
- A grid position or edge the table does not have throws the new `MODEL_TABLE_POSITION`, from every
  one of these, whatever the style.

Nothing is painted here. Drawing the grid, and every coverage question the partly covered pixels
raise, is 4.4.

## The firewall

`V034`, a warning, evidence `measured`. It fires on a cell that writes a border PowerPoint draws
from the cell across the edge on some segment, where that cell writes no line or a different one.
That happens for an `a:lnT` or `a:lnL` wherever the cell before owns the segment, and for a spanning
cell's `a:lnB` or `a:lnR` beyond its anchor's column or row, where the cell after owns it. The file
says one border and PowerPoint draws another: the V030 pattern again. Two lines that both draw
nothing are alike, whether by `a:noFill` or by having no fill. The message names the segment and the
cell whose line PowerPoint draws there. PowerPoint's own writer writes every side alike, so `V034`
never fires on its output.

The test holds it, line by line, to every coloured side an anchor writes in the direct and merge
probes: it fires on exactly the lines whose colour PowerPoint did not draw along their whole
length, once per line and never otherwise, and it is silent on all twelve authored operations.
`a20-tables` writes one such `a:lnL`, and the corpus suite pins that warning beside the other three.

## Verification

- **Tests from the fixture.** `packages/model/src/table-cascade.test.ts` rebuilds every probe's
  table, slide, layout and master from the fixture and resolves it through the model. Across the
  10,435 styled slides it checks all 155,599 cells against COM's fill and text and against the
  painted pixel, and all 384,784 grid edges against the pixels, run for run as the analysis did.
  The counts are derived from the slides' shapes, so a skipped cell fails. It walks every rung of
  the text ladder with its placeholder chain, and holds a cell's own list style to `explicit`.
- **The rivals.** The fixture's 29 findings are held to their shape: every winner fits all its
  rows and every rival fewer. Each of the three rival orders is re-derived through the model and
  misses exactly as many fills as the fixture records. The reversed order, stacked translucent
  fills and the lower cell owning an edge are held to failing.
- `packages/validate/src/validate.test.ts` holds `V034` to the pixels, as above.
- **A review, then a skeptic per finding.** Four read-only agents reviewed the model, `V034`, the C9
  scripts and the tests; a fifth, per dimension, tried to refute each finding. Seventeen survived
  and six were refuted. The ones that changed the record:
  - the table's bottom and right edges beside a merged cell were unmeasured, so they are probed;
  - `V034` missed a spanning cell's `a:lnB` and `a:lnR`;
  - the authored markup was cut off at its first self-closing child, which left the "silent on
    PowerPoint's markup" test with nothing to check;
  - the ladder test built no placeholder chain;
  - the pixel cache was not tied to its export;
  - the rtl deck's direct borders were never scored.
    The rest were errors that named no part, a cell's own list style reading as inherited, and
    comments over the caps.
- **Mutants, 33/34 killed.**
  - **The model, 26/27.** Named by what each mutant does:
    - which parts reach a cell: the first row composed below the columns, row bands counted from
      zero, a corner by position alone, row or column bands reaching the ends;
    - merged cells: a merged cell pooling its bands too or taking its anchor's parts alone, no
      edge by kind, a merged cell's inside drawn, a merged cell claiming its whole side, the cell
      after taking every unclaimed segment, level by the anchor's position;
    - direct formatting and the default grid: the style over a cell's own fill, the `tblPr` fill
      ignored, the cell after always owning the edge, the default grid in `tx1` or at 0.75 pt, its
      text in `dk1`;
    - text: bold or colour from the lowest part, a cell reading `p:defaultTextStyle`, the style
      above the cell's own list style, a cell reading its frame's bucket;
    - the theme: `phClr` dropped from a `fillRef`, an `lnRef` one entry off;
    - positions: `tablePartsAt` accepting any position, the default grid skipping the check.
  - **The firewall, 7/7:** `V034` never firing, content with any facing line, blind to `lnL`, to
    spanning cells, to level, to `lnB` and `lnR`, and comparing lines that draw nothing by markup.
  - **The survivor.** Bold from the lowest part is equivalent: every `b` the built-ins state is
    `on`, so no table PowerPoint draws can tell.
  - **Row bands reaching the end rows.** Under PowerPoint's order that mutant is indistinguishable,
    and it is killed only by the rival-order cross-check, which runs the rival orders through the
    same `tablePartsAt`.
  - **Three survived a first run.** The 0.75-pt grid exposed a partial-pixel check that accepted
    whole pixels. `V034` checking one neighbour only had no probe with a spanning cell under two
    neighbours. A cell reading its frame's bucket rested on a rule that had never been measured.
    Each was killed by a stricter check, a probe and a measurement respectively.
- **`pnpm check`** green before the commit that claims the sub-phase.

## What I had wrong on the way

- My reader asked one `Borders` object for all six sides and got the first side six times. The
  pixels disagreed with COM on every table with different sides, and I first took that for a
  PowerPoint quirk. The reader now takes a fresh object per side.
- I rounded a translucent paint's sum. Rounding each term fits and the sum does not.
- I assumed a cell's text walk ended at `p:otherStyle` then `p:defaultTextStyle`. The full
  combination search found `p:defaultTextStyle` is never read.
- My first direct probes put partial lines on `a:lnT`, which PowerPoint ignores whatever it says.
  They moved to `a:lnB` before the reader reached them.
- My scorer skipped rows where every candidate agreed. Rows where every candidate was wrong
  together hid two misses. Those rows now throw.
- My partial-pixel check accepted a whole pixel as a blend at full coverage, so a 0.75-pt default
  grid passed for the 1-pt one. A mutant found it.
- I gave a merged cell's own line its whole side, then the segments level with its anchor with the
  cell after taking the rest. The table-edge and side-by-side probes the review asked for refuted
  the second: the cell after takes a segment only where it is level itself.
- I pooled every covered position's parts, bands included. The table-edge probe happened to put a
  merge across two bands in the same order no earlier probe had, and four more probes confirmed that
  the bands are the anchor's.
- I required every reading to come from one PowerPoint session, which cost a two-hour read for each
  probe added. Each session now brackets itself with the control, and the analysis holds every
  control read to the first.

## Deviations from the plan

- **`V034`**, not in the plan. The owner rule makes a written border one PowerPoint may not draw.
- **Six fixture files, packed.** One file would be over the corpus's 512 KiB per-file cap, and the
  observations as first written would have taken the corpus to 17.1 MiB, past its 16 MiB total. So
  each deck's codes are raw DEFLATE in base64, each slide's shape is one short string, and a slide
  description many decks share is written once in the main file's `templates`. The observed values
  stay readable. The six files are 787 KB, and the corpus is 15.9 MiB.
- **`MODEL_TABLE_POSITION`**, a new error code: a position outside the grid is the caller's mistake,
  not a malformed attribute.
- **`TableStylePartName`'s comment was false**: it said the schema order is not the order the parts
  compose in. It is, and the comment says so.
- **`tools/ground-truth/lib/png.ts` and a format scheme in `sheet-pptx.ts`.** At 4 px a point a BMP
  per slide is 24 MB, so the reader exports PNG and holds it to BMP twins. Theme C needed a format
  scheme. Without one, `sheet-pptx.ts` writes the same bytes it always wrote.
- **`read.ps1 -Tag -Decks`** reads named decks in a session of their own, bracketed by the control.

## Open questions

1. **The corpus is nearly full.** 15.9 of 16 MiB. 4.4's rendering fixtures will not fit without a
   decision on the cap, which is the owner's.
2. **Vertical double lines.** A vertical 4-pt double line covers its sixth pixel 0.375, not 0.5.
   The blend rule absorbs it and 4.4's rasteriser has to answer it.
3. **`tblBg`'s effects and a gradient's extent.** Theme C's samples are in the fixture and not
   scored.
4. **Dashes in a direct line.** Whether an edge's midpoint lands on a dash is chance, so dashed
   direct lines were not scored.
5. **A covered cell's own `a:tcPr`.** PowerPoint never reads it and its writer leaves it empty.
   Whether PowerPoint keeps one on save was not measured, and `V034` does not warn on one; 4.4,
   which edits merged cells, has to write nothing there.
6. **Diagonals.** No built-in states a `tl2br` or `tr2bl`, so a diagonal is only ever a cell's own
   `a:lnTlToBr` or `a:lnBlToTr`. C9 recorded the pixels of both and did not score them; drawing
   them is 4.4's.
7. **A short row's padded positions.** An `a:lnT` under a position no `a:tc` reached is unmeasured,
   and `V034` stays silent there.
