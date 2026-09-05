# 0032 — Anchors, insets and vertical text

- **Status** accepted
- **Sub-phase** 3.6
- **Experiment** T6 — 601 probes across 50 packages, PowerPoint driven and read over COM, with an
  EMF export of every slide
- **Fixture** [`corpus/ground-truth/frames.json`](../../../corpus/ground-truth/frames.json)
- **Code** `packages/text/src/frames/frame.ts`, `packages/model/src/parse/body.ts`,
  `packages/model/src/resolve/body.ts`

## Context

The plan says:

> **3.6 — Anchors, insets, vertical text.** `@anchor` t/ctr/b/just/dist, `@anchorCtr`,
> `@vertOverflow`. **Insets are asymmetric by default** — `lIns=91440, tIns=45720, rIns=91440,
bIns=45720`. `vert`/`vert270`/`eaVert` in v1; WordArt/Mongolian render approximately and are
> marked non-editable. **`a:br` is a hard line break, not a paragraph break** — no new bullet
> number, no `spcBef`/`spcAft`, no first-line indent reset. **`a:endParaRPr` sets the height of an
> empty paragraph.** _Verify:_ anchor/inset/vert matrix.

Every claim in it is right. What the measurement adds is the half the plan could not have known:
which box the anchor works inside, what two of the five anchoring values actually do, which axis
each attribute governs once the text is turned, and one thing that belongs to 3.4 and was measured
wrong there.

## The method

`TextRange2.BoundTop` and `.BoundLeft`, per paragraph and per line, plus `Slide.Export(path, "EMF")`
for the questions a bounding box cannot answer.

**Every question is posed as a difference between two probes that differ in one attribute.** That is
the design decision the rest of the experiment rests on, and it exists because step 1 found the
instrument is not uniformly trustworthy: a rectangle 110pt wide holding a centred `Ag CJK` at Arial
18 put its text at `BoundLeft` 24.5pt from the shape's left edge, which makes the advance 61.0pt,
while `BoundWidth` reported 66.0. The two cannot both be right and the position is the one the eye
sees. So the control probe carries the unknown — the advance of a string, the height of a block —
and the difference cancels it. `BoundHeight` **is** trustworthy: two lines of Arial 18 report 43.2,
which is exactly 2 × 1.2 × 18, and the `control` family exists to keep checking that at five line
counts, four sizes and six faces.

Multi-line paragraphs are made with `a:br`, never by wrapping, so no reading here is conditional on
3.3's break model.

### Step 1 paid for itself before a probe was built

`TextFrame2` exposes the frame as an object model, so driving PowerPoint through every orientation,
anchor, margin and column count and then reading the file it saved answers four questions for free:

- **The `MsoTextOrientation` names and the `ST_TextVerticalType` names are cross-wired.**
  `msoTextOrientationVertical` writes `vert="mongolianVert"`;
  `msoTextOrientationHorizontalRotatedFarEast` writes `vert="wordArtVert"`. `wordArtVertRtl` has no
  COM spelling at all — value 7 is "The specified value is out of range".
- **PowerPoint never writes `anchor="just"` or `anchor="dist"`.** `msoAnchorTopBaseline` and
  `msoAnchorBottomBaseLine` both answer "The specified value has been deprecated", so two of the
  five `ST_TextAnchoringType` values are unreachable from any user interface — which is exactly why
  their behaviour is in no documentation and had to be read off the layout.
- **The insets are never written when they hold their default**, and `MarginLeft` reports 7.2 for a
  frame that states nothing — for an autoshape, a text box and both placeholder kinds alike.
- **A rectangle's default is `anchor="ctr"`, a text box's is absent.** The difference is in what
  PowerPoint writes, not in what the format defaults to.

## What was measured

Twenty-one questions, 601 probes, each question fitting exactly one candidate. `analyse.ts` throws
rather than emitting a fixture it cannot fit perfectly, and it did — twice, usefully. See "What went
wrong" below.

| question                 | rows | winner                                                 | best rival |
| ------------------------ | ---: | ------------------------------------------------------ | ---------: |
| `anchor`                 |  190 | inside the inset box, `just` and `dist` as bottom      |        120 |
| `block-height`           |  102 | trimmed last line unless `spcFirstLastPara`            |         78 |
| `just-and-dist-stretch`  |   56 | they stretch nothing                                   |          0 |
| `vert-anchor-axis`       |   45 | the stacking axis, from a per-type start edge          |         33 |
| `block-corner`           |   42 | the frame edge, with no half-leading                   |          0 |
| `bodyPr-inherits`        |   24 | yes, per attribute, from any level                     |         12 |
| `vert-overflow`          |   18 | a line is drawn only if it fits strictly               |         16 |
| `anchor-ctr`             |   18 | centres the block; absent is off                       |         10 |
| `overflow-layout`        |   18 | the block is laid out in full whatever is drawn        |          0 |
| `vertical-face`          |   16 | `eaVert` and `mongolianVert` only                      |         12 |
| `inset-applies`          |   14 | the stated value, or 7.2/3.6 when absent               |         13 |
| `inset-far-edge`         |   14 | the same, read off `algn="r"` and `anchor="b"`         |          2 |
| `empty-paragraph-height` |   13 | 1.2 × the `a:endParaRPr` size, times `lnSpc`           |         11 |
| `column-pitch`           |   12 | `(frame − (n−1)·spcCol) / n`, pitch = width + `spcCol` |          8 |
| `vert-insets`            |    7 | the insets stay on their own physical edges            |          0 |
| `break-bullets`          |    5 | one marker per `a:p`                                   |          2 |
| `vert-algn-direction`    |    5 | the axis start, which `vert270` alone reverses         |          4 |
| `inset-collapse`         |    4 | the midpoint of the two inset edges                    |          2 |
| `break-indent`           |    3 | the line after `a:br` starts at `marL`                 |          2 |
| `break-size`             |    3 | the following run sets the line height, not the `a:br` |          1 |
| `break-spacing`          |    2 | `a:br` pays no `spcBef` or `spcAft`                    |          1 |

### The anchor works inside the inset box

`t` puts the block at `tIns`, `ctr` at `tIns + (H − tIns − bIns − block)/2`, `b` at
`H − bIns − block`. The reading that ignores the insets scores 120 of 190 — and it scores 120
because **it is right whenever `tIns` equals `bIns`**, which is every deck that leaves the default
alone. That degeneracy is why every anchor probe is run twice, once at all-zero insets and once at
`tIns=20, bIns=5`; the second set is what separates the two, and the naive reading is wrong on all
70 of them.

### `just` and `dist` are bottom, and stretch nothing

In all 190 anchor readings and all 56 line-position comparisons, `anchor="just"` and `anchor="dist"`
place the block exactly where `anchor="b"` does, with the lines inside it at exactly the same
offsets. Three paragraphs of two lines — the case where a distribute rule would have somewhere to
put the slack — is identical to bottom anchoring.

`TextFrame2.VerticalAnchor` reports **0** for both, which is not a member of `MsoVerticalAnchor`.
PowerPoint's own object model has no name for what it just read.

The name suggests otherwise, and any renderer that implements "distribute" from the name will spread
the lines of every deck that carries the value. It does not appear that PowerPoint ever writes it —
but files from other producers do.

### `spcFirstLastPara` reaches the last line, which 3.4 got wrong

3.4 measured that the flag controls the first paragraph's `spcBef` and the last's `spcAft`, and that
it defaults false. Both hold. What 3.4 could not see, because its probes were all at single spacing,
is that the flag **also decides how much of the last line's leading the block keeps**:

- off: `(n−1)·advance + lastLineHeight`, the trimmed height 3.4 measured as
  `0.75·advance + b·size` floored at `1.2·size`
- on: `n·advance`

60 of 60, across five line spacings (100%, 125%, 150%, 200%, 300%), three line counts and two sizes.
At 100% the floor makes the two identical, which is exactly why the sweep goes to 300% — and why
this could not have been found in 3.4.

It also confirms 3.4's per-face coefficient independently: at 150% on Arial 18 the measured block is
93.2pt and `0.75 × 32.4 + 0.227618991 × 18 = 28.397` predicts 93.197, on a completely different
instrument from the one that derived it.

`textBodyHeight` in `autofit.ts` is corrected accordingly: it subtracted an advance and added the
trimmed last line unconditionally, and now does so only when the flag is off.

### The two axes swap, and the four vertical types do not agree

`@anchor` positions the block along the axis lines **stack** in; `@algn` positions each line along
the axis text **runs** in. When `@vert` turns the text, the two swap — and which frame edge each
axis starts from is a property of the individual value, not of "being vertical":

| `@vert`          | lines stack | `@anchor` from | `@algn` from | asks for `@face` | one record per glyph |
| ---------------- | ----------- | -------------- | ------------ | ---------------- | -------------------- |
| `horz`           | downward    | top            | left         | no               | no                   |
| `vert`           | leftward    | **right**      | top          | no               | no                   |
| `eaVert`         | leftward    | **right**      | top          | **yes**          | no                   |
| `vert270`        | rightward   | left           | **bottom**   | no               | no                   |
| `mongolianVert`  | rightward   | left           | top          | **yes**          | no                   |
| `wordArtVert`    | rightward   | left           | top          | no               | **yes**              |
| `wordArtVertRtl` | leftward    | **right**      | top          | no               | **yes**              |

45 readings for the axis mapping; the reading that keeps the anchor vertical scores 21.

Three of these are readings off the drawing stream rather than inferences from a box.
`LOGFONTW.lfFaceName` comes back as `@MS Gothic` for `eaVert` and `mongolianVert` and as plain
`MS Gothic` for `vert` and `vert270` — GDI's convention for the vertical variant of a face, which
separates "the line is rotated" from "the glyphs are upright" without a pixel being examined. And
`wordArtVert` draws `W|x|y|z` as four separate `ExtTextOutW` calls where `vert` draws `Wxyz` as one.

**`lfEscapement` is always zero.** PowerPoint rotates with a world transform, not with GDI's
escapement, so the instrument this experiment was designed around does not work and the `@` prefix
and the record count are what carry the answer instead.

**The insets do not rotate with the text.** `lIns` insets from the physical left edge whatever
`@vert` says, in all seven readings; the alternative — that the frame and its insets turn together —
scores zero. So under `vert` the block starts `rIns` from the right and `tIns` from the top, because
those are the edges its two axes begin at.

`@upright` changed nothing in any of the 32 probes that varied it.

### `@anchorCtr` centres the block along the line axis

Not the same as `algn="ctr"`, and the difference only shows when the lines are of different widths:
alignment centres each line in the box independently, while `@anchorCtr` puts the **widest** line's
box in the middle and leaves the others aligned inside it. An absent attribute behaves as `0`.

Under rotation it follows the line axis with everything else — `anchorCtr` on a `vert` frame moves
the block vertically.

### `@vertOverflow` changes what is drawn, never the layout

- `overflow`, and an absent attribute: everything is drawn, overflowing the shape.
- `ellipsis`: the whole first line that does not fit is **replaced** by a lone U+2026. The drawing
  stream holds `Line one`, `Line two`, `…` — not `Line two…`.
- `clip`: PowerPoint stops laying out. The records for the lines that do not fit are simply absent;
  there is no draw-then-mask.

The block's height and the anchor's placement of it are identical in all 18 readings. Overflow is a
decision about drawing, taken after layout.

**A line whose bottom lands exactly on the frame edge is not drawn.** A box exactly two line boxes
tall (43.2pt at Arial 18) shows one line; at 44pt it shows two. The comparison is strict.

`horzOverflow="clip"` is not distinguishable in the drawing stream — the full string is handed to
GDI in all three cases — so it is presumably a clip region, and this experiment does not claim it.

### `a:br`, confirmed on all four counts

One `•` for a two-line paragraph and two for two paragraphs; `1.` once however many `a:br` follow
it; the same height as a bare two-line paragraph where two paragraphs cost `spcBef + spcAft`; and
the line after the break starts at `marL`, not at `marL + indent`. The last is proved by the
positive-indent probe: `marL=20, indent=+36` puts line one at 56 and line two at 20, so the break is
a wrapped line and not a first line.

One thing the plan does not mention: **the `a:rPr` on the `a:br` itself does not set the height of
the line it begins.** A break carrying `sz="800"` and one carrying `sz="4000"` produce the same
43.2pt block; the following run governs.

### An empty paragraph

1.2 × the `a:endParaRPr` size, multiplied by the paragraph's own `a:lnSpc`. When `a:endParaRPr` and
`a:pPr/a:defRPr` disagree, `a:endParaRPr` wins — measured both ways round, which is the only way to
tell "it was read" from "they happened to agree". A paragraph whose single run holds the empty
string behaves the same as one with no runs at all.

A trailing empty paragraph occupies its height but is **not** reported by `Paragraphs().Count`.

### Columns

Width is `(content − (n−1)·spcCol) / n` and the pitch is width + `spcCol`; `rtlCol` fills from the
right. `numCol="0"`, `numCol="17"` and a negative `spcCol` are each a package PowerPoint **repairs**
rather than reads, measured one probe to a package — so `ST_TextColumnCount`'s 1..16 is enforced by
the application and not only by the schema. Those three are `validate` rules for 1.2.

### `a:bodyPr` inherits

Nothing in this repository had asked, and it decides where the text goes on almost every real slide:
a body placeholder on a slide writes a bare `<a:bodyPr/>` while the master's states an anchor.

**Every attribute inherits on its own, through the placeholder chain, from any level.** Twenty-four
readings — six properties (`anchor`, `lIns`, `anchorCtr`, `wrap`, `vert`, `numCol`) each stated at
exactly one of slide, layout and master, plus a control stating it nowhere — and the resolved frame
is identical whichever level states it. The reading that treats a bare `<a:bodyPr/>` as the schema
default scores 12 of 24; so does the one that stops at the layout.

## What went wrong

**The scorer refused a fixture twice, and was right both times.**

The first refusal was `block-height`: at single spacing `n × 1.2 × size` and
`(n−1)·advance + trimmedLast` are the same function, so the control family could not separate them
and two candidates fitted 42 of 42. The fix was to merge that question into the one whose probes can
separate them, not to delete the rival.

The second was `vert-axes`, at 15 of 45 — because the model was being scored on the axis that
carries the over-reported advance. Splitting it into the stacking axis, where the extent is the line
box and exact, and a direction question read as a sign, made both exact.

**The mutation sweep found a real bug.** Twenty-eight mutants, and the one that survived the first
run — turning `>` into `>=` in the overflow test — survived because no probe had a line landing
exactly on the boundary. Measuring it showed the implementation was the wrong one of the two: a box
exactly two line boxes tall draws one line. The code changed, six probes were added, and the second
sweep killed 28 of 28.

**The fixture was packed coarser than the measurement.** Positions were stored in eighths of a
point, which is the grid `BoundLeft` is quantised to — but a stated 7.2pt inset is exact, and
eighths report it as 7.25. Repacked in thousandths.

## Consequences

- `packages/text/src/frames/frame.ts` holds the box model: `contentBox`, `frameAxes`,
  `anchorFraction`, `offsetAlong`, `blockOrigin`, `columnBox`, `emptyParagraphHeight`, `drawnLines`.
- `blockHeight` in `line-model.ts` takes `spcFirstLastPara`, and `textBodyHeight` in `autofit.ts`
  passes it through. 3.4's fit test was wrong for any body with the flag on and line spacing above
  100%.
- `packages/model` parses `a:bodyPr` into `BodyProps` — every field `T | undefined` — and resolves
  it per attribute up the placeholder chain.
- `MAX_COLUMNS`, and typed refusals for the three shapes PowerPoint repairs.

## Deviations from the plan

- **`just` and `dist` are not distribute rules.** The plan lists all five anchoring values as though
  each has a behaviour; two of them are bottom.
- **WordArt and Mongolian are not "approximate".** Their axes are measured exactly like the other
  five; what is not measured is the per-glyph cell of the two `wordArt` types, whose column width is
  1.30 × the size for Arial and 1.32 for Courier New — a font metric, so 8.1's data.
- **`packages/model` still has no `@pptx-studio/text` dependency.** Adding one would change
  `package.json`, which needs the user's say-so, so the inset defaults are restated in
  `resolve/body.ts` exactly as `line-model.ts` restates `Spacing`. This is the second sub-phase in
  which that has happened and it is worth deciding.

## Open questions

1. **`horzOverflow="clip"`** draws the same records as `overflow`. Presumably a clip region; not
   claimed here.
2. **`a:bodyPr/@rot`** at an angle that is not a multiple of 90° makes PowerPoint render the text as
   outlines — no text records at all — so the drawing stream says nothing about it. `rot="90"` and
   `rot="270"` swap the bounding box like `vert`; `rot` beyond 360° is taken modulo. What a
   renderer should do with `rot="15"` is unmeasured.
3. **`@upright`** changed nothing in 32 probes. Either it needs a shape rotation to show, or it is
   inert; the probes held the shape unrotated, so this is a gap and not a finding.
4. **The line axis extent.** `BoundWidth` overstates an advance by about 5pt, so every question here
   is posed as a difference. What that 5pt is — a trailing side bearing, a caret allowance — is
   unknown, and 3.8 will need it to place a caret.
5. **`@compatLnSpc` and `@forceAA`** are parsed and preserved and nothing measured them.
6. **A block wider than its column** was not probed: the column family used short paragraphs.
7. **`vertOverflow` with vertical text** — the two were never crossed.
