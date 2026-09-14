# 0056 — The spans are the merge

Date: 2026-09-14
Status: **accepted** — one occupancy rule fits all 76 tables PowerPoint opened as written and the
21 it authored itself, against 52/76 for the flags a person reads first; 11/11 mutants killed and
one shown dead and removed

**Sub-phase 4.1.** The plan's row says "table parse: grid, spans, merges" and nothing else, and the
package had never read past `p:graphicFrame/p:xfrm`: `a:tbl` was reachable only through
`Shape.node`. A DrawingML cell says the same thing twice — `gridSpan`/`rowSpan` on the cell that
grows, `hMerge`/`vMerge` on the cells it grows over — and the schema ties neither pair to the other,
nor a row's cells to the grid's columns. Every reader this project has looked at picks one pair and
hopes. This is the record of asking PowerPoint which pair is normative, what it does when they
disagree, what fixes a row's height and a column's width, and what it writes back.

Code: `packages/model/src/{table,parse/table}.ts`, `Shape.table` in `packages/model/src/types.ts`,
`MODEL_TABLE_ATTR`, `V030` and `V031` in `packages/validate/src/rules/{required,rules}.ts`,
`tools/ground-truth/model/tables/`, `tools/corpus/suites/tables.test.ts`. Fixture:
`corpus/ground-truth/tables.json`. Changeset: `.changeset/the-spans-are-the-merge.md`.

---

## The question, and why the object model answers it

A merge is not a picture. Which cell owns grid position (r, c) is a fact PowerPoint's object model
states directly: `Table.Cell(r, c).Shape` reports a rectangle for every position, and a covered
position reports its **anchor's** — the second cell of a two-column merge answers `216 × 36` at the
first cell's left edge, the fourth row of a two-row merge answers the third row's top and twice its
height. So every candidate rule predicts a rectangle per position from the column widths and row
heights PowerPoint also reports, and one number per position decides. No bitmap, no fitting, no
tolerance beyond a twentieth of a point. That is C5's instrument (ADR 0024) applied to a grid, and
it is the cheapest measurement this directory has made since.

`b04-table`, which PowerPoint wrote, answered first: `Cell(2,3)` reports `Cell(2,2)`'s rectangle
and its text, `Cell(4,1)` reports `Cell(3,1)`'s. Then 82 tables were built to disagree with
themselves.

## C7 — which cells a merge covers

82 probes, one table per package so that a repair names its cause, against PowerPoint 16.0 build
20326: 12 controls written as a hand writes them — a span on the anchor, a flag on each covered
cell — 15 where the span and the flag disagree, 5
where two spans collide, 12 at the grid's edge or with a span of zero, minus one or ninety-nine,
14 with the wrong number of cells or a missing part, 6 lexical forms, 15 about the frame and the
sizes, 3 about the style source. Ten were marked hostile in advance; six were repaired
(`gridcol-no-w`, `tr-no-h`, `tc-body-no-p`, `merge-on`, `span-float`, `frame-missing`) and four
of the ten opened without a word — a row with no cells, a table with no rows, a grid with no
columns, and `w="108pt"`. Nothing was refused. `analyse.ts` throws unless exactly one candidate
fits every row of a question.

| question     | winner                                                   | fits  | the losers                                                                                                    |
| ------------ | -------------------------------------------------------- | ----- | ------------------------------------------------------------------------------------------------------------- |
| occupancy    | spans on the anchor; zero is one, a negative to the edge | 76/76 | below one is one 73/76; magnitude 73/76; span or flag 66/76; span and flag must agree 62/76; flags only 52/76 |
| frame extent | the grid's sums: column widths by row heights            | 76/76 | the `a:ext` of `p:xfrm` 62/76                                                                                 |
| row height   | `@h` is a minimum the content grows                      | 72/72 | `@h` is the height 68/72; the content alone 1/72                                                              |
| column width | `@w`, never less than the side margins plus 2 pt         | 75/75 | `@w` as written 71/75; a fixed 16.4 pt floor 73/75                                                            |

**What the losers would have looked like in the field.** The flags-only reading — the one a person
writes first, because the covered cells are where the merge is visible — misses 24 tables: a
`gridSpan="2"` whose neighbour forgot `hMerge` is drawn merged by PowerPoint and split by the
reader; an `hMerge` with no span to its left is drawn split and read merged; a 2×2 block written
with flags and no spans is four cells in PowerPoint and one in the reader. "Span and flag must
agree" splits every span whose covered cell forgot its flag, 14 of them; "span or flag" merges ten
cells PowerPoint draws apart. The `a:ext` reading draws a table at whatever size the frame last
claimed, which on `frame-narrow` is half the grid; PowerPoint ignores it. A fixed row height clips
three lines into a row written for one.

### The occupancy rule, as measured

- **Each `a:tc` takes the next column of its row**, whatever the spans before it say
  (`row-short-span`: `gridSpan="2"` followed by two plain cells puts the second cell _under_ the
  span, hidden, and leaves the fourth column empty — the Word-style encoding that omits covered
  cells is not read the way it was meant). A short row is padded with empty positions; a fifth
  cell in a four-column grid is dropped (`row-short`, `row-long`, `row-short-span-v`).
- **A position an earlier span claimed is covered whatever its own attributes say.** A covered
  cell's `rowSpan` (`covered-anchor-v`), `gridSpan` (`covered-anchor-h`, `overlap-h`) and text
  (`covered-text`, the `HIDDEN` marker never reported) change nothing, though the text is kept in
  the file.
- **A span stops at the grid's edge or at the first position an earlier span claimed.**
  `span-over-edge-h` (`gridSpan="3"` with two columns left) covers two; `cross` (a `rowSpan` from
  above meeting a `gridSpan` from the left) gives the position to the span that came first in
  reading order and the loser is 1×1; `cross-partial` shows the loser is clamped, not dropped —
  a `gridSpan="4"` meeting a claimed third column covers two. A row span can never meet an earlier
  claim: an anchor above covering those rows would have covered this row too, where the width
  already stopped, so the reader has no vertical clamp to write — the mutant that removed one
  survived and the branch went.
- **Zero is one; a negative or oversized span runs to the edge.** `gridSpan="0"` and `rowSpan="0"`
  are no merge; `gridSpan="-1"` covers the whole row and `rowSpan="-1"` the whole column, exactly
  as `"99"` does; an explicit `"1"` and a `"+2"` read as one and two.
- **`hMerge` and `vMerge` change nothing** — 76 of 76 — and the lexical forms bear it out:
  `hMerge="false"` beside a `gridSpan="2"` is merged, `hMerge="true"` is a flag, `hMerge="on"` is
  a repair.
- **A grid with no column is read as one empty column** of 16.4 pt (208280 EMU: two default
  margins and 2 pt) — the rows’ cells are dropped, not placed in it — **a table with no row as one
  empty row**, and both are written back that way.

### Sizes

`a:tr/@h` is the least a row is drawn at: three 12-pt lines in a 36-pt row draw it at 50.4
(3 × 14.4 + 7.2 of margins), `h="0"` under one line is 21.6 and under an empty run written at 12 pt 28.8 —
PowerPoint lays an empty paragraph out at its 18-pt default —
and a 108-pt row with one line stays 108. A column is never narrower than its cells' side margins
plus 2 pt: `w="0"` draws at 16.4 with the default margins, at 2 with none and at 30.8 with 0.2 in
each side, and a 7.87-pt column is widened the same way. The frame's `a:ext` is bookkeeping: a
frame half or twice the grid's width or height reports the grid's size, and PowerPoint writes the
`a:ext` back as it was (`frame-narrow` resaves `cx="2743200"` under a 432-pt table; the fixture's
`resaved.ext` holds every one). The drawn extent of a table is the sum of its column widths by the sum
of its drawn row heights, at the frame's offset — which is what `Placed.frame` will have to stop
saying in 4.4.

`a:tcPr` reads back as written through `TextFrame2`: `marL="0" marR="182880" marT="0"
marB="91440"` is `[0, 0, 14.4, 7.2]` pt, the defaults are `[7.2, 3.6, 7.2, 3.6]` (91440 and 45720
EMU, the schema's), `anchor="b"`/`"ctr"` are `msoAnchorBottom`/`Middle`, `vert="vert270"` is
`msoTextOrientationUpward`. `w="108pt"`, the universal-measure spelling of `ST_Coordinate`, is a
108-pt column and PowerPoint rewrites it to EMU; the parser accepts the six units the type names and
only `pt` was measured.

## What PowerPoint writes back

Every opened deck was saved as a copy and its `a:tbl` read again, so the fixture also records the
form PowerPoint normalises to — which is what 4.4's writer has to produce and what `V030` guards:

- a missing `hMerge`/`vMerge` under a span is added; a flag with no span is dropped; `"true"`
  becomes `"1"`; an explicit `gridSpan="1"` and a `rowSpan="1"` are dropped;
- a losing span is deleted (`cross`, `overlap-h`, `covered-anchor-*`) and a clamped one rewritten
  as what it covered (`gridSpan="-1"` → `"4"`, `"99"` → `"3"`, `"3"` at the edge → `"2"`);
- a short row is padded with `<a:tc>`, a fifth cell is dropped, an empty row gets four cells, a
  table with no rows one row of `h="0"`, a grid with no columns one `gridCol w="208280"`;
- a block merge takes PowerPoint's own form: `rowSpan` repeated on every horizontally covered cell
  of the anchor's row and `gridSpan` on every vertically covered cell of its column, both flags on
  the interior (`block22`, whose covered cells carry flags but not the repeated spans, comes back
  `rowSpan="2" gridSpan="2"` / `rowSpan="2" hMerge="1"` / `gridSpan="2" vMerge="1"` /
  `hMerge="1" vMerge="1"`);
- a covered cell's text is kept; **an unknown `a:tableStyleId` is discarded** while a built-in one
  the package does not define is kept (`resaved.tblPr.styleId` on `style-unknown` and
  `style-builtin`) — 4.2's problem, recorded here because it is a preservation
  hazard nothing else has measured;
- `a:tr/@h` is written as set, not as drawn (`row-grows` keeps `457200` under a 50.4-pt row; a
  `Rows(2).Height = 10` over COM is saved as `127000` and drawn at 21.6), and `a:ext` is updated
  by some operations and left stale by others (`grow` writes `cy="1554480"`, the 122.4 pt it
  drew; `shrink` keeps `cy="1371600"` under a 93.6-pt table).

### PowerPoint authoring the same thing

`author.ps1` had PowerPoint merge and split its own tables — twelve slides: 1×2, 2×1, 2×2, 2×3, the
south-east corner, a whole row, a whole column, everything, a merge split back, a plain cell split
into two columns, into two rows, into both — and resize nine: a cell grown to three lines, a row
shrunk below its content and grown above it, a column widened and narrowed to 1 pt, the shape
widened, narrowed, taller, shorter. The winner reproduces all 21 grids from PowerPoint's own markup.
Merging everything collapses the grid to one column and one row rather than spanning; a split
adds a `gridCol` (or a `tr`) and puts `gridSpan="2"` (`rowSpan="2"`) on every other row's (column's)
cell; widening the shape scales the columns proportionally; shortening it below the content writes
`h="211667"` and draws 21.6.

## The model

`Shape.table` carries the `a:tbl` of a frame whose `a:graphicData/@uri` is the table URI: `TableProps`
(the seven flags, false when absent — they inherit from nothing — a fill, effects, and
`TableStyleRef`, an id for 4.2 or the inline element for 4.3), `TableColumn`, `TableRow` and
`TableCell` with its four attributes as written, its `@id`, its `TextBody` and `TableCellProps`
(margins as written, `vert`, `anchor`, `anchorCtr`, `horzOverflow`, the six borders as `Line`s, a
fill, `a:headers`). Nothing is defaulted that the file could have said. A span that is not an
`xsd:int`, a flag that is not an `xsd:boolean`, a `gridCol` without `@w` or a `tr` without `@h`
throws `MODEL_TABLE_ATTR` with the part — four of PowerPoint's six repairs; the other two, a cell
body with no `a:p` and a frame with no `p:xfrm`, parse and are `V016`'s and `V017`'s.

`tableGrid(table)` is the rule above: `positions[r][c]` names the owner of every position with its
clamped span and the `a:tc` there (`null` for a padded one), `anchors` lists each owner once in
reading order, `widths` and `heights` are the grid's own. Those few dozen lines are the whole of
what the measurement changed in the package; the rest is a parser in the idiom `parse/text.ts` set.

## The two rules in the firewall

`V030`, a warning, evidence `measured`: a table's rows each hold one `a:tc` per `a:gridCol`, no
written span is larger than what it covers or below one, no flag sits where no span covers, every
covered position carries its flag, and a covered cell carries no span but the one PowerPoint's own
form repeats. PowerPoint opens all of it silently and rewrites it on the next save, which is what a
warning is for: the file says one table and PowerPoint draws another. `V031`, fatal, evidence
`both`: `@w`, `@h`, and the types of the four cell attributes — the repairs. `V016`'s evidence is now
`both` too: a cell body with no `a:p` was measured as a repair. The firewall is thirty-one rules;
the carried-debt entry that said it was still at 1.2's twenty-nine says what it still owes instead.

## What I had wrong about b04-table

I wrote `b04-table`'s manifest entry, its `ROSTER.md` row, its entry in `decks.ts` and the
comment in `build-tier-b.ps1`, and all four said its `ppt/tableStyles.xml` was an empty `a:tblStyleLst`, "so the style
the table names is defined nowhere in the file". It is not: PowerPoint writes the full `a:tblStyle`
of every built-in style a table in the deck uses, and b04's part carries "Medium Style 2 -
Accent 1" in full. The empty part is the stock template before any table is inserted, which is
what `a20-tables` carries. I have corrected all four; 4.2 inherits a measured serialisation of one of
the 74 rather than a negative.

## Verification

- **Tests from the fixture.** `packages/model/src/table.test.ts` (95 tests) parses every one of
  the 76 probes PowerPoint opened as written from the exact `a:tbl` the fixture carries, predicts
  the rectangle every grid position reports under `tableGrid` from the column widths and row
  heights PowerPoint read, and holds it to the one PowerPoint reported — then scores the flags-only
  rule over the same 76 and holds it to the 52 the fixture recorded; rebuilds the 21 authored
  slides from the attributes PowerPoint wrote and holds them the same way, refusing a comparison
  with nothing in it; asserts the four lexical repairs throw `MODEL_TABLE_ATTR` naming the part; holds every
  position's text to the one PowerPoint reported and `a:tcPr` to the margins, anchor and
  orientation its object model read;
  and reads the parse: spans, flags in all four lexical forms, `@id`, the universal measures,
  `a:tblPr`, `a:tcPr`, the invented column and row, the padded row, the dropped cell.
  `packages/validate/src/validate.test.ts` holds `V030` to fire on exactly the probes whose grid
  PowerPoint read back differently — a normalised lexical form, a dropped explicit 1 and the span
  it repeats on a flagged cell being the same grid, not a rewrite — to stay silent on the twelve canonical forms PowerPoint authored, and `V031` to
  the four repairs and silence on the 76. `tools/corpus/suites/tables.test.ts` parses every table
  in the 55 committed decks, as many as the census counted, holds `b04-table` to the grid its
  object model reported and `a20-tables` to its block, its borders, anchors, orientations and fills.
- **The way it will fail.** Not yet by a picture: nothing draws a table until 4.4, and the fixture
  is the only honest check a parser has — the rectangle of every position, on 97 tables.
- **The review of this pass** — four read-only reviewers, one per dimension, two refuters per
  finding — confirmed 25 findings. Three were defects: the invented column for a grid with no
  `a:gridCol` was given the row's first cell, which PowerPoint drops (the fixture's text column,
  never compared until then, is compared now); `V030` let a flag stand in the direction no span
  covers, which PowerPoint drops (`cross`); and a signed EMU was refused by `coordinateOf` where
  `intOf` beside it accepts one. The rest were the record: the row-height question had quietly
  excluded the count family (72/72 now, from 62/62), the `tableStyleId` and stale-`a:ext` claims
  had no fixture row behind them (`resaved.tblPr` and `resaved.ext` carry them now), a candidate
  was missing from the table, "the four it repairs" against six, "as PowerPoint writes them" for
  controls a hand wrote, and the count of the rules still twenty-nine in eleven comments and
  strings.
- **Mutants, 11/11 killed and one dead.** The model 7/7: a grid with no column keeping its cells (2), a negative span read as one (3 tests), a
  covered cell not skipped (39), the horizontal clamp dropped (2), the padding dropped (7), a flag
  read by presence (2), a point of 12000 EMU (1). The firewall 4/4: the missing-flag check dropped,
  the other-direction flag allowed (1),
  the canonical redundancy refused, a float accepted as an int. The vertical clamp survived its
  mutant and was shown unreachable, so it is gone from the model, the rule and the analysis, and
  the analysis emits the identical fixture without it.
- **`pnpm check`** green before the commit that claims the sub-phase.

## Deviations from the plan

- **Two firewall rules, not in the plan.** The measurement answered what PowerPoint rewrites and
  what it repairs, and a rule with measured evidence costs less to add than to carry; 4.4's writer
  will need `V030` to say what "canonical" means.
- **`tools/ground-truth/model/` split into `sheets/` and `tables/`.** C5's five files sat directly
  in the package's directory while every other package's experiments nest one directory each; the
  second experiment made the missing parent visible. Every reference moved with them, including
  the tool paths inside `sheets.json`'s own comment, re-pinned.
- **The frame's `a:ext` is demoted.** `Placed.frame` still comes from `p:xfrm` and is right for
  every other shape; for a table it is what the file last wrote, and 4.4 has to size the frame from
  the grid. Recorded here so it is not rediscovered.

## Open questions

1. **Row growth needs the text engine.** `@h` is a minimum the content grows, and the content's
   height in a cell is `packages/text`'s line model over the cell's inner width — 4.4's, with the
   margins-plus-2-pt floor on a column's width, which the layout must apply before it breaks lines.
2. **The unknown `a:tableStyleId` PowerPoint discards.** `style-unknown` opened clean and was
   saved without its id. Whether the firewall should warn on an id the package does not define is
   4.2's to decide once the 74 built-in GUIDs are known.
3. **Universal measures elsewhere.** `a:gridCol/@w`, `a:tr/@h` and the four `a:tcPr` margins accept
   `108pt` here; `a:off` and `a:ext` still throw on one. Only `@w` was measured, and nothing
   measured says PowerPoint reads a measure on the others.
