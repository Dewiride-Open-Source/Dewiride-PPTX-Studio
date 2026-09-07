# ADR 0034 — Text in both renderers

**Sub-phase 3.8.** Status: accepted. Adds `packages/render-svg/src/text/`,
`packages/render-dom/src/text/`, `packages/text/src/lines/baseline.ts` and
`packages/text/src/runs/decoration.ts`. Fifteen questions: fourteen fit exactly one candidate
reading, and the fifteenth refutes both of the cheap implementations a renderer would reach for.
The plan's own description of this sub-phase is wrong in one place, and wrong in a way that draws
half the flipped text on a slide upside down.

Code: `packages/render-svg/src/text/{resolve.ts, layout.ts, emit.ts, draw.ts}`,
`packages/render-dom/src/text/layer.ts`, `packages/text/src/lines/baseline.ts`,
`packages/text/src/runs/decoration.ts` and its generated
`packages/text/src/runs/face-rules.gen.ts`.
Measurement: `corpus/ground-truth/text-rendering.json`, from
`tools/ground-truth/render/text/{probes.ts, build-deck.ts, read.ps1, measure-in-browser.ts,
analyse.ts, write-tables.ts, verify-render.ts}`.

---

## What this sub-phase had to decide

Sub-phases 3.1 to 3.7 settled where a _line_ goes: the ten-source cascade, the 1.2 line box,
breaking, autofit, bullets, the frame, and which typeface a run is actually drawn in. None of them
asked where the glyphs sit inside that line, which way a mirrored shape's text faces, or what an
underline is a fraction of. Those are the drawing questions, and a renderer cannot avoid any of
them.

The plan gives three instructions: an HTML text layer "counter-flipped", real `<text>`/`<tspan>` in
`render-svg` and never `foreignObject`, and one layout engine under both so they cannot drift. Its
verification is that "the two renderers agree on line boxes".

So T8 asked PowerPoint. **160 probes across 7 packages**, every one opened without a repair, read
three ways: `TextRange2`'s own line boxes through the object model, an EMF export of every slide
for the world transform, the baseline origin and the per-character advances GDI was handed, and a
1920 × 1080 bitmap wherever the question is about ink that nothing records.

### The instruments

`Slide.Export(path, "EMF")` is the one that matters. `EMR_EXTTEXTOUTW` carries `ptlReference` under
a text alignment of `TA_BASELINE`, so the baseline origin is a number rather than a guess; the
world transform in force is the angle the block is drawn at, and its determinant says whether the
glyphs are mirrored; and `offDx` holds the advance of every character. The logical unit turns out
to be a six-hundredth of an inch — checked rather than assumed, because `lfHeight` is the em size
in those units and 160 probes are 160 chances for it to be something else.

Two limits are worth stating. **Off a quadrant PowerPoint draws the glyphs as filled paths**, so a
30° probe has no text record at all and its angle has to be fitted from the ink. And **PowerPoint
reports no baseline** through the object model, which is why the baseline question is an EMF
question and the line-box question is a COM one.

Two mistakes in the rig are worth recording because both produced confident wrong answers before
they were found. The first run put a 600 × 240pt frame at the top left of the slide; rotated 90° it
hangs off the page, and four probes came back with no ink at all, which reads as "PowerPoint drew
nothing" rather than "the export clipped it". The frame for the two turn questions is now 400 × 160
centred on the slide, whose half-diagonal fits at every angle. The second was in the EMF reader:
without tracking `EMR_SAVEDC`/`EMR_RESTOREDC` the world transform accumulates across drawing calls,
so every text record after the first reported an origin twice too far from the page corner.

---

## The findings

| question            | winner                            | rows  | runners-up                                                                                       |
| ------------------- | --------------------------------- | ----- | ------------------------------------------------------------------------------------------------ |
| `flip`              | flipH ignored, flipV turns 180°   | 20/20 | both flips ignored 10, both flips turn 10, flipV ignored 10, glyphs mirrored 5                   |
| `rotation`          | `@upright` drops the shape's turn | 19/19 | always the sum 15, shape rotation only 11, body rotation only 10, never turns 7                  |
| `upright-box`       | extents swap on a quadrant        | 5/5   | never swapped 3, the true rotated box 4, always swapped 2                                        |
| `line-box`          | the largest run sets it           | 36/36 | the first run 34, the last run 34, the smallest 32                                               |
| `baseline`          | the face's own share of the box   | 36/36 | 3.4's autofit coefficient 24, a fixed fraction 12, CSS half-leading 9, the ascent 1, ink share 0 |
| `line-advance`      | one of this line's boxes          | 9/9   | the next line's box 8, the em with no leading 0                                                  |
| `align`             | anchored without trailing spaces  | 21/21 | the whole line 19, the stretched width 19, everything left 15                                    |
| `justify`           | `dist` stretches the last line    | 21/21 | every justify spares the last 19, every justify stretches 15, nothing stretches 15               |
| `trailing-space`    | the spaces hang past the edge     | 6/6   | the spaces count 2                                                                               |
| `decoration-scale`  | proportional to the em            | 18/18 | a fixed length 6                                                                                 |
| `decoration-face`   | **none of these**                 | 16/16 | the browser already agrees 0, one offset for every face 4, the browser's thickness 0             |
| `decoration-extent` | the rule stops at the last glyph  | 3/3   | the rule covers the spaces 0                                                                     |
| `baseline-shift`    | per cent of the size it asked for | 4/4   | per cent of the drawn size 0, per cent of the line box 0, the size is not reduced 0              |
| `caps`              | small caps at four fifths         | 3/3   | three quarters 2, the size does not change 2                                                     |
| `runs`              | the line is shaped whole          | 6/6   | each run shaped alone 4                                                                          |

### A flip is not a counter-flip, and that is the plan's own mistake

The plan says the HTML text layer is "counter-flipped (PowerPoint does not mirror text on flipped
shapes)". The first half of the parenthesis is right and the instruction is not. Sixteen quadrant
probes read the world transform directly and four more at 30° read the ink's centre:

- **`flipH` does nothing at all.** The transform, the origin and the ink are identical to the
  unflipped shape at every angle. Not counter-mirrored in place — ignored, so a left-aligned line
  still starts at the left.
- **`flipV` turns the whole block 180° about the frame centre.** Exactly: the unflipped origin
  reflected through the centre, to the unit, at 0°, 90°, 180° and 270°.
- **The two together are a half turn**, because the second is the first composed with a mirror the
  text ignores.
- **The determinant is +1 on all twenty.** The glyphs are never mirrored.

A renderer that counter-flips both axes — the plan's instruction taken literally — is right on ten
of the twenty rows and draws the other ten the wrong way up. `packages/render-svg/src/text/emit.ts`
therefore puts the text in a **sibling** group of the shape's own, because the shape's group carries
the `scale(-1)` that a `flipH` asks for.

### The baseline is the face's own share of the line box

This is the number a renderer cannot guess, and every implementation guesses it differently.
Measured across eight faces and four sizes, the baseline sits at

```
drop = lineHeight × ascent / (ascent + descent)
```

with the ascent and descent the browser's own `fontBoundingBox` values. 36 of 36 to within one EMF
unit — 0.12pt — and the rivals are the interesting part:

- **CSS half-leading**, which is what every browser does to an inline box and what a renderer
  written by hand inherits for free, fits **9**. It is out by 1.4pt at 54pt on Courier New, a fifth
  of a line.
- **A face-independent fraction of the box** fits 12. There is no such constant: the share runs from
  0.735 on Courier New to 0.827 on Verdana.
- **3.4's per-face autofit coefficient** fits 24, and where it fails is instructive. `b` for five of
  the six faces 3.4 measured is exactly `1.2 × descent / (ascent + descent)` — the same rule seen
  from the other end. It fails on Courier New alone, whose 0.300 is 3.4's own floor kicking in
  rather than the face's share of 0.318.

The face box has to be read at a size where the browser has stopped quantising. Chromium rounds
`fontBoundingBoxAscent` to whole pixels, so at 100px Arial's descent share reads 0.2250 where it is
0.2277 — half a point out at 54pt. `FACE_BOX_PX` is 1000.

### Neither cheap way of drawing an underline survives

This is the one question with no winner, and the refutation is what decides the implementation.

- **Chromium's own `text-decoration` is 0 of 16.** It draws every one of the eight faces at 0.05em
  below the baseline and 0.1em thick. PowerPoint's offsets run from 0.084em on Tahoma to 0.235em on
  Courier New — a spread of 18pt at 120pt — and its thicknesses vary too.
- **One offset for every face is 4 of 16**, and the four it fits are coincidences among the strikes.
- CSS has **no way to position a strikethrough at all**.

So both rules are drawn as geometry, from a generated per-face table, in both renderers. The derived
shapes hold at 18, 40 and 66 points: a `heavy` rule is half again as thick and starts one plain
thickness higher, a `dbl` is two half-thickness rails at those same two edges, and a `dblStrike` is
two plain rails one thickness either side of where the single one goes. A rule stops at the last
glyph — three probes, and the trailing spaces are simply not underlined.

### A line is shaped whole, and a run boundary splits only the drawing call

`AVAWA` as one run and as `AV` + `AWA` produce byte-identical advance arrays, and the kern between
the `V` and the `A` that straddles the boundary is present in both. A colour change does split the
`ExtTextOut` call into two records — and the advances still do not change.

That decides the emitters. A `<tspan>` carrying its own `x` restarts shaping and loses exactly that
kern, so `emit.ts` places the `<text>` once and lets the pieces flow; justification is `word-spacing`
for the same reason. The HTML layer does the same thing with spans in one line box.

### The rest

- **`@upright="1"`** drops the shape's rotation and lays the text out in the frame's extents,
  **swapped when that rotation snaps to 90° or 270°** — 2.10's quadrant-snap rule turning up again
  on a different question. Five probes, and the true rotated bounding box fits four of them.
- **`a:bodyPr/@rot` adds to the shape's rotation.** 30° + 45° reads 75° off the ink, where 45° and
  30° are 40pt away in the bounding box.
- **The line box is 1.2 × the largest size on the line**, in any order, and every run on the line
  shares one baseline. Four mixed-size probes; "the first run decides" fits two of them.
- **Alignment anchors the line without the spaces it ends with**, which hang past the edge. This is
  not an edge case: line one of every wrapped paragraph ends in the space it broke at, so a renderer
  that counts it centres every wrapped paragraph half a space to the right.
- **`dist` stretches the last line of a paragraph and `just`, `justLow` and `thaiDist` do not.**
- **`a:rPr/@baseline` raises a run by that percentage of the size it asked for**, and draws it at two
  thirds of that size. `cap="small"` uppercases the lowercase at four fifths and leaves what was
  already capital at full size.

---

## What was built

`packages/render-svg/src/text/` is four files and the seam between them is deliberate.

`resolve.ts` walks 3.1's cascade once and hands over concrete values. `layout.ts` turns those into
positioned lines — the turn, the box, the insets, the columns, breaking, the baseline, the
alignment, the rules — and takes its measurer and face-box probe as arguments, which is the only
way a test can drive the whole pass from literals. `emit.ts` turns a laid-out block into SVG.
`draw.ts` is the seam `shape.ts` uses.

`packages/render-dom/src/text/layer.ts` reads the identical `TextBlock` and builds an HTML layer:
one absolutely-positioned div per shape, one per line, a span per piece, and the same `PieceRule`
rectangles the SVG draws. The line boxes cannot drift because there is one set of them.

### Where the HTML layer is not exact, and why

A browser puts an inline baseline a half-leading plus an ascent below the line box, from font
metrics it has **rounded to whole pixels**. `TextLine.strutPt` is the `line-height` that solves the
first for PowerPoint's share, and it lands the baseline within one CSS pixel and no closer. A second
attempt — a line box of zero height, whose baseline would be its own top edge — moved the problem
rather than solving it: with `line-height: 0` on the pieces the line box's top sits at
`baseline − (ascent − descent)/2`, which is the same rounded metric one step further away.

So the HTML layer is within a point and the SVG emitter is exact, which is the right way round:
thumbnails, PNG and PDF go through the SVG, and the HTML layer exists for selection, IME and screen
readers.

---

## Verification

`tools/ground-truth/render/text/verify-render.ts` closes the loop the other way from every other
sub-phase. It loads the same seven `.pptx` files through `opc` and `model` in a real Chromium, lays
each probe's text out with `render-svg`, and compares against what PowerPoint itself reported — the
line boxes through the object model, and the colour of every drawing call through the EMF:

```
text render: 415/423 line measurement(s) within 0.1pt across 132 unturned probe(s),
             28 turned one(s) left to the EMF
             4/4 run colour(s) match the colour PowerPoint drew
             8 anchored line(s) inside the measurer's own disagreement, worst 0.441pt
```

The eight are all the left edge of a centred or right-aligned line, which is the column less the
line's own advance and therefore carries 3.2's measured browser disagreement whole — a median of
0.15% of the string, which on a 128pt line is a fifth of a point. Nothing in this sub-phase changes
that number; it simply arrives where it was always going to.

This matters more than the unit tests, and for the reason 2.10 recorded: 39 unit tests passed there
while every gradient painted a flat block. The layout tests here all drive the pass from literals,
so until this ran, `resolve.ts` had never been asked a question by a real package — and it was
wrong twice over. See "What went wrong".

`packages/render-dom/src/text/layer.test.ts` carries the plan's own verification: five shapes, laid
out once, emitted as SVG and mounted as HTML, with the browser's rendered baseline read back through
a zero-height inline block. The SVG's `<text y>` is the layout's baseline exactly; the HTML's is
within a point.

**Mutation sweep: 49 mutants, 49 killed.** The first pass wrote 48 and left six standing; two of
those had targets that should not have existed and were deleted rather than tested around, one was
provably equivalent, and three were missing assertions. Five more were added for the colour path
after the verifier found it broken. The survivors are recorded under "What went wrong" below.

`pnpm check` green: structure, layering, 492 references, 92 corpus entries, format, lint, typecheck,
13 builds, round trip 52/52, 25 pkg:qa, 87 test files, 3060 tests.

---

## Deviations from the plan

1. **The text layer is not counter-flipped.** Measured 20 of 20: `flipH` is ignored and `flipV` is a
   half turn. The plan's instruction would draw ten of those twenty probes upside down.
2. **The rules are drawn, not delegated to `text-decoration`.** Chromium's own decoration is 0 of 16
   against PowerPoint and CSS cannot place a strikethrough at all.
3. **`@pptx-studio/render-svg` gained a dependency on `@pptx-studio/text`**, which changes
   `package.json` and `pnpm-lock.yaml`. Asked for and approved before it was made, under hard rule 2.
4. **`render-dom` still mounts one `<svg>` per slide** rather than the plan's per-shape DOM nest, as
   2.10 already decided. The text layer is a sibling of that `<svg>`, which is what bet 4 asks for
   and what the flip finding requires anyway.
5. **Four of the seven `ST_TextVerticalType` values throw.** `horz`, `vert` and `vert270` are drawn;
   `eaVert`, `mongolianVert`, `wordArtVert` and `wordArtVertRtl` need the `@`-prefixed vertical face
   or per-character stacking, and drawing them approximately would be a plausible wrong answer where
   a typed error is an honest one.
6. **`tools/ground-truth/render/` became a parent of two experiments**, `transforms/` (C6, 2.10) and
   `text/` (T8), because a second experiment in one flat directory would have put two `probes.ts`
   beside each other. Every reference moved in the same commit.

---

## What went wrong

**The line-spacing units, twice.** `LineSpacing.value` is the raw attribute — thousandths of a
percent, or hundredths of a point — and `lineAdvance` quantises it. `resolve.ts` quantised it first,
so every line came out zero points tall, and five tests failed with the same `expected 0` before the
cause was obvious. The fix is that the resolver stores what the file says and one function converts.

**`read.ps1 -Only` silently shrank the readings.** Re-reading one deck rewrote the whole readings
file with that deck alone, and the next analysis ran on a seventh of the data. It now carries the
decks it skipped forward. This cost one confused run and would have cost a wrong fixture.

**Every coloured run drew black, and two separate bugs did it.** The verifier caught both, and only
because it was extended to compare the colour of each drawing call against `EMR_SETTEXTCOLOR`.

The first was an unchecked cast. `colorOf` took the run's fill as `unknown` and narrowed it with
`as { kind?: string }` — but `Fill`'s discriminant is `type`, not `kind`, so the branch never
matched and every run resolved to no colour at all. The cast is what hid it: typing the parameter as
`Fill` made the compiler say so in one line. This is the engineering standard's "unchecked cast"
rule earning its place.

The second was underneath it. `Rgba` channels are **0..1**, and both emitters treated them as bytes
— `Math.round(0.75).toString(16)` is `"01"` — so even with the fill resolved, a red run would have
come out `#010000`. Both now go through `toHexColor` and `toCss`, which `paint` already had and
tests.

Neither was reachable from a unit test, because every layout test builds `ResolvedRun` literals with
`color: null`. There is now a test that asks the resolver a question about markup, and one that
asserts the hex each emitter writes.

**Six mutants survived the first sweep**, and each says something:

- _measure an empty string for the face box._ `fontBoundingBoxAscent` does not depend on the string,
  so this is a genuinely equivalent mutant; the probe measures `Hxg` because a reader should not have
  to know that. Removed from the sweep with the reason recorded here.
- _a double underline at full thickness_, _a justified line already too wide_, _a superscript that
  goes down._ Three missing assertions, all added.
- _floor the quadrant instead of rounding._ The layout had restated 2.10's quadrant-snap rule instead
  of importing `swapsExtents`, which is measured there over eighteen angles and tested. Deleted the
  duplicate; the mutant went with it.
- _count the trailing spaces in the anchored width._ The subtraction was dead: `wrapText`'s
  `measuredEnd` already ends before the space run a line broke at, which 3.3 established and this
  had duplicated. Deleted, and the comment now cites 3.3 rather than restating it.

### Half the corpus threw, and `verify-render` could not have caught it

Every probe deck T8 authors names a typeface on every run, so the seven decks the verifier reads
never asked what happens when none does. Half the corpus does: `a08-bullets` reaches its ninth
outline level under a master that declares five, and no source below that level names a face. The
resolver returned the empty string, `cssFamily` refused to quote it, and the slide did not draw.

The answer was already measured and had been left in the fixture. Six of T1's probes reach a cascade
that names no typeface - `font-silent` was built for exactly this question - and all six came back
as the theme's **minor** face. Two more say the built-in styles name one too: a title placeholder at
level 1 under a master with no `p:txStyles` drew in the major face, a body placeholder in the minor.
3.1 recorded the sizes those probes produced and not the faces, so the table shipped with four
columns where the measurement had five.

So `BuiltinLevel` gained a `typeface`, derived in `write-tables.ts` from the same probe rows rather
than typed in, and `resolveLatinTypeface` joined `resolveSize` as a resolver that cannot return
`undefined`. Past the first title level PowerPoint reported no face at all, which is recorded as
`null` and falls to the floor - the one part of this that is a choice rather than a reading, and it
is open question 10.

Two lessons, both about where the check was pointed. A verifier that reads only the decks its own
experiment authored is testing the experiment. And a resolver that hands a renderer an empty string
has already lost: the throw belongs where the cascade runs out, naming the run, not three files
later at a CSS shorthand.

### The layout read `wrapText`'s indices in the wrong units

`breakOpportunities` says what it returns in its own header - "positions are code point indices" -
and `wrapText` works over `toCodePoints(input.text)`. This layout took the `start` and `end` it got
back and used them to index the joined string, whose offsets it had built from `cell.text.length`.
Those two spaces are the same string right up until a character outside the basic multilingual plane,
and then they diverge by one per astral character.

Found by the first sweep of 3.9's harness, on `a40-unicode` slide 2, as
`URIError: URI malformed` - `encodeURIComponent` refusing a lone surrogate. The emitted markup held
`\ud83d` with no low half, and Chromium's own `DOMParser` rejected the document: **`render-svg` was
emitting text that is not well-formed XML.** Reduced to six characters it is
`wrapText({ text: '𝕏𝕏 abc', … })` returning `{ start: 0, end: 3 }`, which is `'𝕏𝕏 '` by code point
and `'𝕏\ud835'` by code unit.

Malformed output was the loud half. The quiet half is that `rangeMeasurer` measured the wrong
substring, so **every line break in a paragraph containing an astral character was decided on a
width that belonged to different text** - and `render-dom` builds on the same node tree, so both
renderers had it.

The fix is that a `Cell` now carries its code points and every offset and slice in the pass counts
in them, which is the space the breaker already worked in. The lesson is narrower than the last
one: an index is a measurement in units, and a function that returns one should be read for which
units before its result is used as an offset. The four call sites had no test that could tell the
two apart because every one of them ran on Latin text.

---

## Open questions

1. **`@upright` was never probed together with `flipV` or with `a:bodyPr/@rot`.** The code treats
   `upright` as dropping the whole `a:xfrm` turn, flip included, and as leaving `a:bodyPr/@rot` in
   place. Both are the natural reading of the rows that were measured and neither is measured.
2. **The quadrant snap in `textTurn` has no probe between 45° and 90°.** The layout imports 2.10's
   `swapsExtents`, which is measured there over eighteen angles for group children; that it is the
   same rule for `@upright` is an assumption resting on five probes at 0, 30, 90, 180 and 270.
3. **`dotted`, `dash` and `wavy` underlines have no measured geometry.** PowerPoint paints them as a
   pattern, so the bitmap holds no solid rail to find. They are drawn at the single rule's position
   with a dash array, and the wave's amplitude is invented.
4. **The strike thickness is not the face's `yStrikeoutSize`.** Arial's is 150/2048 = 0.0732em and
   PowerPoint drew 0.050em. The table records what was drawn; where the number comes from is unknown,
   and 8.1's `sfnt` reader could settle it.
5. **Only eight faces have measured rule metrics.** Everything else gets Arial's through
   `APPROXIMATE_FACE_RULES`, which is wrong by 0.13em on a monospace face. 8.1 reads `post` directly
   and makes the table unnecessary.
6. **Bullets are resolved but not drawn.** 3.5 measured them and `bulletLayout` exists; the layout
   pass does not call it yet, so a bulleted paragraph draws its text at the right indent with no
   bullet in front of it.
7. **The autofit path is view-only.** `a:normAutofit/@fontScale` is applied verbatim, which is 3.4's
   view mode; the edit mode that re-runs the ladder belongs to 6.4.
8. **A run's `a:ln` and `a:effectLst` are parsed and ignored.** Outlined and shadowed text draws as
   plain fill.
9. **A built-in title level past the first names no typeface**, so a hand-written master with a
   level-2 title draws in the minor face. PowerPoint's object model reports no face for those rows,
   which is why the table records `null`; what it actually paints is unmeasured, and a bitmap of that
   slide would settle it.
10. **`packages/validate` has no rule for an `@u` or `@strike` outside its simple type**, nor for an
    `a:bodyPr/@rot` outside `ST_Angle`.
11. **The HTML layer's baseline is within one CSS pixel and not exact**, because Chromium rounds a
    font's ascent before laying a line out. Whether a `transform` on each line could close it without
    a per-line DOM measurement is untried.
