# 0054 — A hundred slides, at any zoom

Date: 2026-09-13
Status: **accepted** — Gate 3 holds: 100 slides × 5 widths and a 2x display through the page,
offline; 28/28 mutants killed in the first pass, 23/23 in the second, 13/13 in the third and
11/11 in the fourth

**Sub-phase 3.11.** Gate 3 asks for "a 100-slide deck rendered faithfully at any zoom, entirely
client-side". None of its four claims was true or checkable when this began: the largest committed
deck had eleven slides, the page drew no pictures and no `w="0"` lines, its stage was a constant 900
pixels wide with no zoom control, and nothing measured what the page did or where it fetched from.
This is the record of making each claim a measurement, and of what the measurements found.

Code: `tools/ground-truth/render/zoom/`, `tools/ground-truth/render/snap/`, `tools/gate3/`, `apps/studio/src/slides/`,
`tools/corpus/tiers/a-generated/decks/deck/a46-hundred-slides/`, `tools/fidelity/{oracle,record,fixtures}.ts`,
`tools/fidelity/metric/{grid,deck}.ts`, `packages/paint/src/lines/line.ts`,
`packages/render-svg/src/{paint,shape,slide,image/media}.ts`, `packages/render-svg/src/text/`. Fixtures:
`corpus/ground-truth/zoom.json`, `corpus/ground-truth/snap.json`, `corpus/decks/a46-hundred-slides.pptx`
and its oracle at five widths, `corpus/ground-truth/render/fidelity/zoom/expected.<env>.json`.

---

## The sentence, as four measurements

| claim                | what now checks it                                                                                                                                                                                                                                                                                                                                                                                  |
| -------------------- | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| a 100-slide deck     | `a46-hundred-slides`, generated, 100 slides, in `corpus/decks/`; the oracle holds PowerPoint's export of every slide at 120, 240, 960, 1920 and 3840 pixels wide                                                                                                                                                                                                                                    |
| rendered faithfully  | every slide scored through the **page's own DOM** against that export, at every width; `pnpm fidelity` scores it through the `<img>` path as well                                                                                                                                                                                                                                                   |
| at any zoom          | experiment F2 asked PowerPoint what its export does at seven widths that a linear scale does not, and F3 where it puts an edge on the device grid at fourteen — where the rule holds, and where it stops; both rules are in the renderer; the gate asserts the SVG at every zoom is the SVG at 100 % apart from what those rules own, and that a 2x display's 100 % is the 200 % markup to the byte |
| entirely client-side | the gate fetches the deck once, takes the browser context offline, and logs every request; a request it does not recognise, or any request after going offline, fails it                                                                                                                                                                                                                            |

`pnpm gate3` runs all of it in `pnpm check` and CI from committed fixtures, needs no Office, and
exits 0 only when every count is zero. The scores are reported and gate nothing, for the reason ADR
0035 gives.

## F2 — what zoom does to PowerPoint's own picture

Seven export widths (120, 240, 480, 960, 1200, 1920, 3840), 38 probes on seven slides, each slide
exported twice as BMP and compared with itself; 266 cases against PowerPoint 16.0 build 20326.
Every family had its candidate readings enumerated — including the ones anyone would write by hand —
and scored at 0.15 px, the marker family at a 1-px box tolerance because a head's height is read
from an ink box rather than a profile; `analyse.ts` throws unless exactly one model per family fits
every case, and `assertSeparable()` refuses a pair of models no width can tell apart. The plan
named six widths; 1200 was added so a six-point tile lands between pixels, and it is the width that
separates O4 from O1.

| family          | winner                                                                     | fits  | closest loser                                       |
| --------------- | -------------------------------------------------------------------------- | ----- | --------------------------------------------------- |
| hairline        | **H1** one device pixel at every width                                     | 42/42 | H4 half a point, never under a pixel — 36/42        |
| thin stroke     | **T4** rounded to whole pixels, never under one                            | 56/56 | T5 rounded up — 48/56; T2 true width ≥ 1 px — 46/56 |
| border          | **T4**                                                                     | 14/14 | T2, T3 — 10/14                                      |
| border edge     | **O4** wholly outside, the whole-pixel stroke snapped to the grid, half up | 8/8   | O1 wholly outside, unsnapped — 7/8                  |
| dashed hairline | **D1** the period of the drawn pixel width; a zero width has no dashes     | 28/28 | D4 solid at every width — 21/28                     |
| marker          | **M8** a ten-point head at or under two points, else five whole-pixel pens | 49/49 | M5, M7 — 44/49                                      |
| pattern         | **P1** the tile is six points at every width                               | 6/6   | P2 six points snapped to pixels — 5/6               |

What the losers would have looked like in the field: H4 is the reading the `w-quarter` clue in
`lines.json` suggested and it is wrong at 120 and 240 px, where half a point is under a pixel; T2 is
what an SVG renderer does by default and it is wrong for every width between one and two pixels;
D4 — "a dashed hairline is solid" — is right 21 times and wrong exactly where a dash is wider than
the pixel it is drawn with.

Three families needed no rule. **Text is linear**: every run's extent at every width is within 1 px +
2 % of a linear scale of the 960 export (worst 4 px on a 3840-pixel slide), and nothing is dropped
at 120 px — a 6-pt Arial run is still there at 0.75 px per point. **Gradients are identical** from
480 px up (worst 1 level of 255; 5 levels at 120, where there are only 49 distinct levels to
have). **The frame is stretched**, never letterboxed, at every width, and 3840 wide is not a
ceiling: `Slide.Export` produced every width asked for.

**M8, the arrowhead pen**, is in the renderer: a head on a line at or under two points is drawn
from a two-point pen, above two points from the whole-pixel stroke, and never from under a device
pixel — `markerPen` in `paint`, 49/49 against the fixture at its 1-px tolerance. What finishing it
found is under "Line ends", below.

**O4's half-up snap** was, in the first two versions of this record, carried rather than
implemented: O4 fits 8/8 and O1 — wholly outside, unsnapped — 7/8, and the one case that separates
them is `border-1pt` at 1200 px. One case is an anecdote. What made it worth carrying was the
company it had: every whole-pixel stroke in the thin family from 240 px up has a profile peak of
exactly 1.0 — one crisp row, or two, or five, never a stroke straddling two rows at half coverage —
while a stroke centred on an integer device coordinate in our SVG straddles, and Chromium
antialiases it over two rows. Which way PowerPoint snaps, on which coordinate, for odd and even
widths, on a diagonal and on a curve, is experiment F3, below; O4 turned out to be one face of a
rule that is not the border's at all.

**Pattern coverage under 480 px** is PowerPoint's own artefact: `horz` at 480 px exported with no
ink at all, and at 120 and 240 px with a 3-px period where six points is 0.75 and 1.5 px. P1 is
scored from 960 up, where the period is six points to 0.02 px, and the fixture keeps the low widths
under `patternBelow` rather than pretending a rule fits them.

## F3 — where an edge lands on the device grid

**The question.** F2 said every whole-pixel stroke is crisp from 240 px up and said nothing about
where; a crisp row can be the row above the coordinate or the row below, and an even pen and an odd
one cannot both be centred on a grid line and be crisp. `tools/ground-truth/render/snap/` draws
lines, rectangle outlines, an L-shaped custom outline, filled rectangles, pictures with and without
a border, a stroke aligned inside its frame, flat line ends, and the shapes that are not rectangles
— an ellipse, a rounded rectangle, a triangle, a line rising two points over two hundred, a
rectangle turned a quarter — at up to four sub-pixel offsets, 0, ¼, ½ and ¾ of a point, which at
960 px are the four quarters of a pixel and at 1200 px four others; the strokes, fills, pictures,
borders, inset strokes and ends see all four, the shapes that are not rectangles three, the L and
the half-pixel pens two. Ten slides, exported twice at fourteen widths, 3388 cases, none excluded.
The widths are F2's seven and seven more chosen across three things a width can be — a whole
number of dots per inch, a whole-pixel height, an eighth of a pixel per point (which is always a
whole dpi): a whole dpi without a whole height (1000), a whole height without a whole dpi (1008,
1184), both without an eighth (1040, 1120), an eighth without a whole height (120, 1320), and none
of the three (1100). The rule is read from the six
widths that are an eighth of a pixel per point with a whole-pixel height, 240 to 3840, and every
other width is scored against it. The instrument is different from F2's: a model here predicts the
coverage profile across the edge, row by row, so an antialiased reading and a snapped one are
scored on the same numbers, and a profile is what the file keeps; a row's scale is the rounded
height's, a column's the width's. Tolerance 0.15 of a row; two models are separable when they
differ by 0.2 somewhere.

**What it found.** One rule fits every axis-aligned stroke, 852 of 852, on every kind of shape
that carries one:

> The stroke's centre rounds half up to the device grid. The pen is the width rounded half up to
> whole pixels, never under one — except that an exact half pixel rounded down at 1200 wide, below.
> An odd pen then sits on pixel centres, half a pixel past the rounded coordinate; an even pen sits
> on the grid line. Either way every edge of the pen is a pixel boundary, and the stroke is crisp
> whatever the offset was.

The readings it refutes, with their scores on the stroke family's 336: the true width antialiased
where it lies, 80; the whole-pixel pen antialiased where it lies, which is what this renderer drew
from the stroke rule until now, 100; the pen's top edge rounded half up without moving the centre — SR,
which is what `shape-rendering="crispEdges"` does on its own — 256; the same with the centre
rounded first, SG, 328, missing exactly the eight cases below. Fill edges and picture edges round
half up, 192 of 192, so a filled rectangle a quarter pixel down starts whole on the row it is
in, not on a 75 % row, and three quarters down starts on the next. A flat line end follows its pen onto pixel centres — the end column of an
odd pen is half covered — and a one-pixel pen reaches a quarter pixel further at each end, 144 of 144. A picture border is not a snapped stroke: the frame edge rounds half up and the pen sits half
the _true_ width outside it, antialiased, 48 of 48, which is what O4 saw at its one width. A
slanted line is snapped at its endpoints and antialiased between them, 24 of 24 — the endpoints
move, the line does not straddle where it happens to cross a row. A triangle's base and a
rectangle turned 90° snap like any line. A curve does not: an ellipse's top is the stroke rule's
band a quarter pixel down and its side an eighth right while its bottom is exactly the rule's, and
a rounded rectangle's straight edges lose a quarter pixel across rows and an eighth across
columns from their leading edge once the pen is wider than one — 72 of 72 each with that bias
written into the model, 35 and 48 without it. The
bias is the curve rasteriser's and the renderer does not chase it.

**An `algn="in"` stroke is neither centred nor inset.** Nothing had measured one. Its band is
the device pen with its outer edge one pixel outside the rounded frame edge — half a pixel for an
odd pen — and the rest inside, and a pen narrower than a pixel slides inward by what it was
widened by: 48 of 48 (IN). So a 1-pt inset stroke at 100 % straddles its frame edge, half a row
each side, where a centred 1-pt stroke is crisp; a 2-pt one at 400 % is one pixel outside and
seven inside. The alignment ignored — the stroke rule centred on the frame — coincides with IN
wherever the two constructions land the same, the 1-pt pen at 480 and 1920 and the 2-pt pen at
240, 960 and 1200, four of eight at every width but 3840 and 20 of 48 in all; the reading anyone
writes first, the true width inside the frame, 1 of 48; the border family's readings mirrored
inward, 0 to 10.

**The exact half pixel, settled where the page draws.** A pen of exactly _k_ + ½ pixels — 6, 10
and 14 points at 240, 1.5 to 6.5 at 960, 0.75 to 2.75 at 1920 — rounds _up_ and draws crisp, 28
of 28; at 1200 the same half rounds _down_ and draws astride the pixel centre, 16 of 16; and no
integer-EMU width is a half pixel at 3840, where a half would need an eighth of a point. The
fixture carries the three readings whole: up everywhere, 236 of 252; down everywhere, 224; half
to even, 230; the rule with 1200's exception, 252 of 252. 1200 is not alone, and it is not a
rounding: at 1040 a 6.5-px pen is drawn six wide _on a pixel centre_ — the width rounded down and
the parity up — and at 1120 a 3.5-px pen is drawn _four_ wide on a pixel centre, the width up and
the parity down; at 1120 a frame edge on an exact half, 479.5, rounds down where 960 rounds every
such edge up, while a picture edge on a half, 304.5, rounds up there as everywhere. The scales
where nothing of this happens are the ones whose arithmetic is exact: the powers of two, which are
every zoom the page offers. No arithmetic on the width reproduces
the rest — every route through points, twips, EMU, inches, 96-dpi pixels or single precision
gives the half exactly — and the renderer rounds up.

**Where the rule stops.** Each of the seven other widths is scored against the rule, family by
family, and mapped where a mapping could explain it:

| width | dpi  | true height → exported | eighth | the rule                                       | what it did instead                                                                                                                                                                                 |
| ----- | ---- | ---------------------- | ------ | ---------------------------------------------- | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| 1040  | 78   | 585                    | no     | 239 of 242                                     | the two 6.5-px ties, the 1-pt L step astride its row                                                                                                                                                |
| 1120  | 84   | 630                    | no     | 227 of 242                                     | the eight 3.5-px strokes, a frame edge at 479.5, two L steps, the 2-pt border crisp one pen outside                                                                                                 |
| 1000  | 75   | 562.5 → 563            | no     | columns hold; rows 0 of 42 ties, 8 of 16 fills | every partial row a quarter (275 of 292); the rule at 1000/960, rows moved by the stretch rounded _up_ to a quarter pixel: strokes 28/28, fills 8/8, borders 8/8, inset 8/8, ties 40/42, 137 of 150 |
| 1320  | 99   | 742.5 → 743            | yes    | columns hold; rows the same                    | quarters (283 of 300); the same mapping 140 of 150, every stroke, tie, fill, border and inset                                                                                                       |
| 120   | 9    | 67.5 → 68              | yes    | columns hold; rows the same                    | quarters (296 of 322); the mapping 108 of 150 — a stroke gains a quarter-row above it                                                                                                               |
| 1008  | 75.6 | 567                    | no     | no row; 9 of 28 columns, 3 of 8 fill edges     | quarters at the chance rate (78 of 487); two- and three-row transitions on every edge; an idealised render at 75 or 76 dpi box-resampled fits 122 of 242 at best                                    |
| 1184  | 88.8 | 666                    | no     | no row; 6 of 28, 4 of 8                        | the same, 62 of 242 at best                                                                                                                                                                         |
| 1100  | 82.5 | 618.75 → 619           | no     | no row; 12 of 28, 3 of 8                       | the same, 49 of 242 at best                                                                                                                                                                         |

So every case fits the rule on the six widths that are an eighth of a pixel per point with a
whole-pixel height; on the other two whole-dpi, whole-height widths it is 239 and 227 of 242, and
what leaves it is an exact half of a pen drawn one way and placed the other, a frame edge on an
exact half, the custom L's inner step, and at 1120 a 2-pt border a sixth of a row from BT. A
fractional height does not stretch the snapped render: its columns obey the rule and its rows are
the rule's rows moved by the stretch in quarter-pixel steps, as a rasteriser sampling four times a
row and rounding up would leave them — the curves and the picture edges do not follow, and at 9
dpi a stroke gains a quarter-row nothing here explains. A fractional dots per inch is a different
picture altogether: no row fits the rule, a column does only where the resampling's phase happens
to land it, partial coverage lands on a quarter at the 15 % a uniform antialiased edge does by
chance, and a crisp edge spans two or three rows, which is what a render at some other size
resampled through a smooth kernel looks like; a render at the whole dpi below or above — an
idealised scale, since neither is an export with whole sides — resampled through a box filter fits
half the cases at 1008 and a quarter elsewhere, so the source and the kernel are not identified. `a41-a4` at 960 px wide is 88.6 dots per inch: the
antialiased edge the third pass read by hand on its first slide, 255, 255, 212, 80, 68, 68 down
one column, is the third signature, and the basis point that slide gives up to the snap rule is
the oracle's resampling, not a rule the renderer is missing. The page's zooms are 9, 18, 72, 144
and 288 dots per inch; the 120-px strip is the one with a fractional height, and its rows are the
quarter-step rows above.

## The rules in the renderer

`deviceStrokeWidth(widthEmu, pxPerPt)` in `packages/paint`: `max(1, round(w · s))` device pixels, a
zero width one pixel. `render-svg` applies it when `RenderOptions.width` names a device — the page
always does, the CLI does, and a caller that gives no width gets the true width and a hairline as
`stroke-width="1" vector-effect="non-scaling-stroke"`. A dashed hairline is drawn solid (D1). The
device scale has a floor, `MIN_PX_PER_PT = 0.05`: below a twentieth of a pixel to the point
`deviceStrokeWidth` throws `LINE_DEVICE_SCALE`, the page's fit zoom stops there, and the widest one
device pixel a stroke is ever rounded up to is twenty points.

**Zoom is a mount at that width, never a resize.** This is the plan's "measured exception": the SVG
for 25 % is not the SVG for 400 % with another `width`, because `stroke-width` and
`stroke-dasharray` are rounded to the device pixels of the width the slide was mounted at. The
architecture claim "zoom is a transform, never a re-layout" survives in the half that matters: the
text is laid out once, at one point size, and the gate proves the line breaks and every other byte
of the markup are the same at every zoom. `MountedSlide.resize()` is gone.

The stroke rule moved 29 of the 31 corpus slides it touched toward PowerPoint (mean 9768 → 9770
bp over the 154 slides then scored) and made none worse.

### The grid rule

`devicePen(widthEmu, pxPerPt)` in `packages/paint` returns the pen and, when it is odd, `shift`:
half a device pixel in EMU. `render-svg` applies F3 when a width names a device and every segment
of every path of the shape is horizontal or vertical on the slide — a preset rectangle, a line with
no rise, a rectilinear `custGeom`, any of them turned by a multiple of 90° or flipped — and does
nothing to a curve, a diagonal or a turned rectangle, which stay antialiased where they lie. A
rectilinear path carries `shape-rendering="crispEdges"`, under which Skia fills the pixels whose
rows round half up from the edges, which is F3's fill rule and, for an even pen, F3's stroke rule.
An odd pen is moved half a device pixel down and right first, as a `transform="translate(h h)"`
on the path itself in the shape's own coordinates — `(h, −h)` at 90°, `(−h, −h)` at 180°, the
signs flipped with a flip — after which Skia's rounding lands it on pixel centres exactly as the
export does. Where a shape has both a fill and an odd pen the fill is its own path with no
translate, so its edges keep the fill rule and the pen alone moves; an even pen shares the fill's
path. A pen of an exact half pixel rounds up, as every power-of-two export does.

A rectangular picture's border is drawn as F3 read it: the device pen on the frame outset by half
the true width, over the picture, with no clip at all — `crispEdges` when the true width is a
whole number of device pixels, where it is exactly BT, and antialiased otherwise, where it is BT
up to the frame edge's rounding, a number no zoom-invariant markup can carry. Which of the two a
border gets is the width's, not the zoom's: a 1-pt border is crisp at 100, 200 and 400 % and
antialiased at 25 % and in the strip; the 0.75-pt default is antialiased at 100 % too, the regime
F3 measured with its 1- and 2-pt borders at 240 and 480. That crispness is the pen's, so the gate's mask owns
it as it owns `stroke-width`: a path marked `data-band` may be crisp at one zoom and not at
another, and no other path may. The band's inner half overlaps the picture at a small zoom, as
the export's does. A picture that is not a rectangle keeps the double-width band clipped to the
outside, and Blink applies a `<clipPath>` antialiased whatever its child asks — the render test
puts `crispEdges` on a clip's path and reads a 75 % row at its quarter-pixel edge, where the same
edge on a crisp rectangle is whole. An `algn="in"` stroke keeps its clipped band too, and that
band is measured wrong by a known amount: the export's pen sits one device pixel outside the
rounded frame edge and ours sits wholly inside where the frame lies, so ours is up to a pixel and
a half too far in at every zoom, a device-pixel constant the markup cannot express and the render
test holds to the formula. What the border construction moved: the twelve bordered-picture slides
of a46 2 to 4 bp each at 100 % (9971 → 9975 on the first) and nothing else; the gate's 100 %
column 9956 → 9957 and the strip 9749 → 9753, where PowerPoint's own 120-px export agrees with
its 960 at 9742; 25, 200 and 400 % unchanged. On the Linux runner (run 34760265774) the same:
100 % 9930 → 9931, the strip 9736 → 9741, the rest and the corpus mean of 9829 unchanged.

A named `height` now stretches the slide into its box, `preserveAspectRatio="none"`, which is
what F2 measured the export doing and what the page and both harnesses had been asking for by
passing a rounded height; with only a width the aspect holds. On a 16:9 deck at the page's zooms
the box is exact and nothing stretches; on A4 at 960 px it is 0.4 px over the height, which is
where PowerPoint's own pixels are.

What it moved: 86 of 254 fidelity rasters, 75 slides toward PowerPoint, one — `a41-a4-01`,
above — one basis point away, and the corpus mean 9841 → 9844 bp. By deck: a07-text-cascade
9891 → 9919, a10-rtl-cjk 9941 → 9966, a11-autofit 9921 → 9945, a12-masters 9973 → 9992,
a45-backgrounds 9967 → 9984, a08-bullets 9962 → 9977, a09-fields 9976 → 9988, a06-lines 9969 → 9979. The gate's columns: 25 % 9878 → 9897 and the strip 9727 → 9749, where a pixel is scarce and a
straddled stroke costs most; 100 % 9954 → 9956; 200 % 9960 → 9961; 400 % and the 2x display within
a point. The mask owns the translate as it owns `stroke-width`: a crisp path's translate may
appear, disappear and change between zooms; whether a path is crisp may not, and a translate on
anything else is a shape moving.

### Line ends

Finishing M8 found that no renderer had ever drawn a line end. ADR 0023 measured the six heads —
sizes, anchors, the stealth's notch, the open arrow's stroked V — into `marker-table.ts` and
`markerGeometry`, and said "nothing is painted; 2.10 builds the renderers"; 2.10 did not build this
one, Gate 2 counted arrowheads in the corpus rather than in the raster, and the first version of
this ADR wrote that `markerGeometry` "still scales from the nominal width" as if something called
it. Nothing did. The a06-lines slides and a46's thirteen connector slides were scored without their
heads, and a 6-pixel head at 960 px is under one cell of the metric, so the score never said so
(+2 bp on a06-02 once drawn, unchanged elsewhere).

`render-svg` now draws a head as an SVG `<marker>` in user units on the open ends of a stroked
path — `marker-start` where the first subpath is open, `marker-end` where the last is; a closed
rectangle with a `tailEnd` gets none — sized by `markerGeometry` from `markerPen`, painted as the
stroke is, `orient="auto-start-reverse"` at the start so both tips point outward. Two details the
review of this pass caught. A paint server is in the line's user space and a marker has its own, so
a gradient- or pattern-stroked line's head is painted with `context-stroke` rather than the
server's `url(#…)`, which would have sampled the ramp from the marker's back corner. And the open
arrow is a V stroked with the pen, so its vertex sits half a pen back with a round join: C4 measured
the arrow's inked tip on the endpoint with no overshoot and its silhouette a pen longer and wider
than nominal (`lines.json`), which is what a stroked V centred on the endpoint would fail by half a
pen. The raster tests read the head's height at its back the way F2 read PowerPoint's — 10 rows at
960 px where the export drew 9, 5 at 240 where it drew 4, both inside the fixture's tolerance,
centred on the line to a pixel — and the column past an arrow's tip for ink, and find none. Because
the pen is a device measurement, a marker's numbers differ between zooms exactly as `stroke-width`
does, and Gate 3's mask owns them the same way: a marker's `markerWidth`, `markerHeight`, `refX`,
`refY` and the numbers of its outline may change with the zoom; the outline's commands, its
orientation, its paint and which end it is on may not.

## The deck, and the limits it found

`a46-hundred-slides`: Tier A, generated, LCG-seeded, 100 slides of ten kinds — cover 1, agenda 1,
section header 8, bullets 25, two-column 13, picture with caption 12, shape grid 13, quote 13, bar
diagram 13, closing 1 — with a `sldNum` field on every content slide (90), two procedural PNGs shared
by twelve slides' relationships, a master text-style override, `normAutofit` on every fourth bullets
slide, and a shape grid of twelve presets under solid, gradient, pattern fills, dashed lines, a
connector with a triangle head, an outer shadow and a glow, rotations, one `flipH` and one nested
group turned 20°. 164,143 bytes deflated, 222 entries; the census literal locks 803 shapes, 746
preset geometries, 182 gradient fills, 35 pattern fills, 39 connectors, 90 fields, 13 shadows and 13
glows. No date field: `a09-fields` already carries the calendar drift ADR 0038 Q3 records.

Its claim, in a43's terms, was the limits nothing had ever hit. Each one hit:

| limit                                        | what happened                                                                                                                                                           |
| -------------------------------------------- | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| the grid set's one-byte slide count          | `encodeGridSet` silently truncated at 256; it now throws above 255, and sets shard by bytes at 480 KiB (`shardGridSets`) — a46 at 960 px is three files                 |
| `CAPS.perFile` (512 KiB) on one deck's grids | the same sharding; `oracle.json` records name their file                                                                                                                |
| `CAPS.total` (12 MiB)                        | 8.0 MiB + a46 0.16 + 1.17 at 960 + 4.68 at four widths = 13.9 MiB; raised to 16 MiB, the user's decision                                                                |
| the harness's per-slide deck reload          | `drawSlide` loaded the package once per slide — quadratic on a46; it now caches the loaded deck per URL for the run                                                     |
| a Linux baseline that could not grow         | `--record` is refused in CI and `--bootstrap` once a baseline exists; the stale `expected.linux-x64.json` is deleted and the dispatch job records both baselines afresh |
| `sldNum` resolved per slide                  | 1..100, correct in every export and every render                                                                                                                        |
| `<defs>` ids across 101 mounted SVGs         | `idPrefix` per mount — `thumb<i>`, `slide<i>` — and the gate renames one to the other to compare them                                                                   |
| one measurer per mount                       | one `TextEngine` per deck, passed to every mount; the strip went from one `OffscreenCanvas` per slide to one per deck                                                   |

Two defects surfaced only because a hundred slides of ordinary content went past the renderer at
once, and neither was in the plan:

- **Every bordered picture was blank.** The clipped outline band's `clip-path` was on the same path
  as the fill, so the fill was clipped to the band. Fill and band are two paths now; a test draws a
  one-pixel red PNG behind a border and reads the pixel. +54 bp on the corpus mean.
- **No bullet was drawn.** The model resolved `a:buChar`, `a:buAutoNum` and `a:buBlip` and
  `packages/text` laid them out, and `render-svg` never emitted one. It does now, through
  `drawableBullet`, `numberParagraphs` and `bulletLayout`; 49 slides improved, none worsened, +7 bp.
  One slide is refused as a consequence: `a08-bullets-02` uses `arabic1Minus`, whose glyphs T5 could
  not read back (PowerPoint shapes them to glyph ids), so `TEXT_AUTONUMBER_UNMEASURED` is thrown
  rather than a guess drawn. Recorded as vanished from the baseline.

## The page

`apps/studio/src/slides/`: a strip of every slide at 120 px, mounted in frame-sized chunks after
the stage has painted; a stage at 25, 100, 200, 400 % or fit; one text engine per deck; a media
resolver over the package (`mediaFromStore`, shared with the CLI and the harness); a delegated click
listener instead of one per mount; and an automation hook — `openFromUrl`, `show`, `stageSvg`,
`thumbSvg`, `stageShapes`, `stageBox`, `thumbBox`, `stripDone` — that reports timings and never
asserts them.

The strip's first chunk waits for a frame before it mounts anything, so the stage — mounted
synchronously beside it — paints first. The first version mounted its first twelve milliseconds of
thumbnails inside `createStrip()`, before the stage had been asked for anything, while its comment
said the opposite; the test that catches it counts the strip's `<svg>`s the instant the strip is
created and expects none.

**The display's own pixels.** `RenderOptions.devicePixelRatio` multiplies the named width for the
stroke rule alone: at 100 % on a 2x display a slide is 960 CSS pixels wide and its strokes are
rounded to the 1920 device pixels it is drawn with, which is the markup of 200 %. The page passes
`window.devicePixelRatio` into every mount and watches it through `matchMedia` — a window dragged
to another display re-mounts the stage at the same zoom and redraws the strip. Without this every
whole-pixel stroke on a 2x display was two device pixels wide, and the sub-phase had said so as an
open question rather than measuring it.

The first drive of a46 took **18.6 s** to draw the strip while every line of JavaScript accounted
for 0.45 s. Chromium's trace put it in two `RasterTask`s of nine seconds each, and bisecting one
glowed hexagon in a 240-pixel SVG found the whole cost in one place:

| the shape, in the emitted filter      |  raster |
| ------------------------------------- | ------: |
| as emitted: rotated, gradient, dashed | 9486 ms |
| the same with no rotation             |   28 ms |
| solid, no stroke, rotated             | 8431 ms |
| solid, no stroke, flat                |   31 ms |

`feMorphology operator="dilate"` under a `rotate()` transform is pathological in Chromium **when the
filter's user space is EMU**: the cost rises with the radius in user units (radius 1: 0.5 s; 2857:
6 s; 28575: 9 s; 285750: 15 s) and vanishes when the same filter is written in a coarser unit
(÷10: 8.3 s; ÷100: 30 ms; ÷12700: 32 ms). So an effect filter is now written in **points**, inside a
`scale(12700)` / `scale(1/12700)` pair around the shape's paths — the second written as a string,
because `num()`'s three decimals round it to zero. Fifteen effect slides moved within the noise
floor (largest −6 bp, every `maxD` unchanged) and the strip fell to **761 ms**.

| a46 through the page, win32                  |                     measured |
| -------------------------------------------- | ---------------------------: |
| package parsed and model loaded              |                        95 ms |
| first slide painted                          |                       243 ms |
| 100 thumbnails mounted                       | 761 ms, 430 ms on the thread |
| JS heap after the strip                      |                        22 MB |
| one stage mount, any zoom                    |                      ~1.5 ms |
| one screenshot at 240 / 960 / 1920 / 3840 px |       41 / 65 / 127 / 360 ms |

## What the gate found

`pnpm gate3` on win32-x64, 2 m 11 s: 100 slides, 400 stage renders and 100 thumbnails, all drawn,
all invariant, two requests for the deck and none after going offline, 500 rasters recorded.

| zoom          | px wide | mean bp, first pass | with the grid rule | with the border on the frame | worst slide                 |
| ------------- | ------: | ------------------: | -----------------: | ---------------------------: | --------------------------- |
| 100 %         |     960 |                9954 |               9956 |                         9957 | a46-hundred-slides-40, 9819 |
| 25 %          |     240 |                9878 |               9897 |                         9897 | a46-hundred-slides-24, 9692 |
| 200 %         |    1920 |                9960 |               9961 |                         9961 | a46-hundred-slides-40, 9799 |
| 400 %         |    3840 |                9964 |               9964 |                         9964 | a46-hundred-slides-40, 9799 |
| strip, 12.5 % |     120 |                9727 |               9749 |                         9753 | a46-hundred-slides-24, 9243 |

The strip is the honest column: at 120 px a cell is one pixel, text is under a pixel tall and every
antialiasing decision shows. Before the grid rule its worst slide was 88, a bar diagram whose labels
PowerPoint drops into grey at that size and ours keeps legible; with it, slide 24. The stage columns climb with width because the cell grows with
it, as ADR 0022 predicted for anything that scales.

The Linux runner said the same thing 15–27 bp lower, which is Carlito standing in for Calibri
(CI run 34717006014, HeadlessChrome 151): 9928 / 9861 / 9933 / 9936 by zoom and 9708 for the strip,
the same three worst slides, `pnpm fidelity` at 9826 over the corpus; the page parsed a46 in 105 ms,
painted the first slide at 190 ms and the strip at 480 ms, in 14 MB of heap. Both platforms hold the
gate and both baselines are committed. With the grid rule the runner re-recorded both (CI run
34754279761): 9930 / 9881 / 9935 / 9936 by zoom and 9736 for the strip, 9935 on the 2x display,
`pnpm fidelity` at 9829 — the 25 % column and the strip up 20 and 28 bp, as on win32.

**The 2x display**, measured. After the zooms the gate switches the page to a device pixel ratio of
2 over CDP without telling it, asks for every slide at 100 % again, and reads the strip the page
redrew on its own reaction. On win32 and on the Linux runner: 100 of 100 stage slides are the 200 %
markup to the byte and 100 of 100 thumbnails the 25 % markup — the two stroke attributes and the
markers included, since 2 px/pt is 2 px/pt however it is reached. The rasters are Chromium's and
not the page's, and they are not identical: 7 of 100 match the 200 % raster byte for byte and the
rest differ by up to 5 levels of 255 (mean 9999 bp against 200 %), which is how Skia treats a 2x
device matrix against a 2x transform. Scored against PowerPoint's 1920-px export as a sixth column
the 2x stage is 9960 bp, the 200 % column's own number, worst slide 40 at 9798 (9961 with the grid
rule, as the 200 % column is). So the markup is
gated and the raster reported. The strip is the half that proves the page noticed: nothing in the
pass asks it to redraw, and a thumbnail it refuses on that redraw is counted as not drawn. One
mechanism worth writing down: Playwright re-applies the context's device metrics on every
`page.screenshot`, which silently put the ratio back to 1 — the pass captures over the same CDP
session it set the ratio on.

**PowerPoint against itself**, the self-check the plan asked for, is a column of `pnpm gate3` now
(`oracleSelfBp`, "PowerPoint vs its 960" in the report). Its own 120, 240, 1920 and 3840-px exports
of a46, reduced to the same grids and scored against its 960 export, agree at 9742 / 9879 / 9977 /
9977 bp; the page's columns at those widths were 9727 / 9878 / 9960 / 9964 in the first pass and
are 9753 / 9897 / 9961 / 9964 with the grid rule and the border on the frame. At 25 % and in the strip the page is now 18 and 11
bp _closer_ to PowerPoint's export than PowerPoint's own 960 export is, which is the calibration the
zoom columns needed and a bound they have passed.

Two facts the gate established about the page that the plan had assumed. First, an inline `<svg>`
root is **not** pixel-snapped by Blink: `getBoundingClientRect()` put the stage at y = 817.1875, and
a Playwright clip is viewport-relative unless `fullPage`, and then only as wide as the viewport. The
gate pins the stage and the strip at page (0, 0) with an injected stylesheet, hides the one it is
not shooting, and refuses a fractional box rather than rounding one. Second, only `stroke-width`
and `stroke-dasharray` differed between zooms across every slide measured — the line-end markers
joined them once heads were drawn, and a crisp path's half-pixel translate once the grid rule was —
and the outline band's containment rectangle, whose reach was two drawn widths. Its reach is now
`2 × (nominal width + one device pixel at the floor)`, zoom-invariant, after measuring that a reach
past ~600 pt changes how Skia antialiases the clip (−10 bp on every bordered picture; identical up to
600 pt).

`pnpm fidelity` over 254 slides: corpus mean **9841 bp** in the first pass and 9844 with the grid
rule, a46 mean 9955 then 9957, minimum 9819; noise floor 4 bp. The report's by-deck table puts the five lowest decks where the plan has not reached —
a23-smartart 7255, b04-table 7653, b06-smartart 8571, a22-chartex 8801, a37-mce 8823. The request
log the gate keeps starts after the harness's own injections: the page's load and the harness's
probe fetch are before it, the deck fetch and everything after are in it.

### The real deck

One deck outside the corpus was measured through `pnpm gate3 --deck`, unscored — there is no oracle
for it and it is not a fixture — handed over by name and read once from its own directory:
100 slides, every one drawn at all four zooms and in the strip, zoom-invariant, no page error,
two requests for the deck and none after going offline; parsed in 160 ms, first slide at 300 ms,
100 thumbnails at 659 ms (314 ms on the thread), 28 MB of heap. It is the first deck not written by
this repository to go through the page end to end, and it needed nothing fixed.

## Verification

- **Tests from the fixture.** `line.test.ts` reproduces every F2 stroke probe at every width —
  hairline, thin, border and dash-quarter within 0.15 px, the 21 dashed-hairline rows solid, the 49
  marker rows within their 1-px tolerance through `markerPen` — and asserts each losing hairline,
  thin and marker model misses at least one; the text, gradient, pattern and border-edge rows are
  held by `tools/ground-truth/fixtures.test.ts`; `render.test.ts` renders the hairline slide at 240 and 3840 through the `<img>` path
  and profiles it, and reads a triangle head's height and centre the same way; the text test
  rasterises a run at 960 and 3840 and holds its ink to 4× within F2's tolerance; the effect test
  derives the emitted radius from `paint`'s own graph and reads the glow's pixels; the band test
  compares the clip at 960 and 240. The first version of this section claimed the text test before
  it existed.
- **The way it will fail.** The page in Chromium against PowerPoint's own pixels at five widths,
  offline; and a page that resizes instead of re-mounting on zoom is caught at the first slide
  (`FID_SVG_INTRINSIC_SIZE`: the stage was still at the fit size).
- **Tests from F3.** `tools/ground-truth/fixtures.test.ts` re-derives the grid rule from
  `corpus/ground-truth/snap.json` in its own words and holds it to 852 of 852 strokes (732
  before the fourth pass widened the tie family), the fill
  and picture edges to 192, the ends to 144, the borders to 48, the slant to 24 and both curve
  biases to 72, asserts that the fixture names the one perfect reading per family, and pins every
  rival it re-derives — the antialiased reading, the top-rounded reading, half-to-even, a pen
  wholly outside the frame, a slant snapped at its read, the curves without their bias — to the
  score the fixture recorded for it; `packages/paint/src/lines/line.test.ts` holds `devicePen`
  to the rows of 328 of the 336 stroke cases and counts the eight 1200-px ties it departs from, and
  to the 960 and 480 ties' widths; `packages/render-svg/src/render.test.ts` checks the markup —
  crisp and shifted, a fill under an odd pen split from it, crisp alone, neither, the six turns
  and flips against the frame's own inverse, the L and the chevron, the band under its clip, no
  device, the stretch — reads a clip path's edge to show Blink antialiases it whatever its child
  asks, and then draws F3's own 28 horizontal stroke probes at 960 and 1920 through the `<img>`
  path and holds every row to PowerPoint's within the fixture's tolerance, 56 of 56, where the
  same markup with the snap stripped from it misses 25 of 28; `tools/gate3/gate3.test.ts` holds
  the mask to the translate a crisp path carries and nothing else.
- **Tests from F3's fourth pass.** `tools/ground-truth/fixtures.test.ts` holds the widened
  fixture — 852 strokes, the 44 ties by width with up-everywhere, down-everywhere and half-to-even
  pinned to the fixture's scores, the inset stroke's 48 with the alignment ignored and the true
  width inside pinned, the 1040 and 1120 miss lists by name with the 3.5-px pen drawn four wide,
  the 6.5-px pen six wide and the 479.5 edge rounded down each re-read from the profiles, the
  fractional heights' columns to the rule and their rows to the quarter-step mapping exactly as
  the fixture scored it and above the exact stretch, and the fractional dots per inch to no reading
  and no quarters; `packages/paint/src/lines/line.test.ts` holds `devicePen` to the 28 halves
  at 240, 960 and 1920 and counts the 16 at 1200 it departs from;
  `packages/render-svg/src/render.test.ts` checks the border band's markup — on the outset
  frame, after the picture, crisp at 960 and 1920 and not at 240, no clip, a rounded picture
  still clipped, and a custom geometry a rectangle only as one closed walk round the frame's
  corners either way — then draws F3's eight border probes through the `<img>` path and holds
  every row to PowerPoint's at 960 and, at 240, all but the four whose frame edge is 102.5 px,
  those held to the half pixel, and draws the eight inset probes and holds the gap between our
  band's centre and the export's to the formula; `tools/gate3/gate3.test.ts` holds the mask to
  a band's crispness and no other path's, and to a band's translate, crisp or not.
- **Mutants, 11/11 killed in the fourth pass.** The renderer 7/7 (the frame outset by the whole
  width, the band never crisp, always crisp, the clipped construction kept for a rectangle, the
  border under the picture, any picture taken for a rectangle, the frame's corners in any order);
  the pen 1/1 (a half rounded down); the mask 3/3 (the band rule dropped, every path's crispness
  stripped, the band rule after the translate rule so a crisp band could move unseen).
- **The review of this pass** — six read-only reviewers, one per dimension, a refuter per finding
  — confirmed 57 findings and refuted 1. Two were defects: a custom-geometry picture that touched
  the frame's four corners in any order, closed or not, was given the rectangular border; and the
  mask stripped a crisp band's translate before it stripped the band's crispness, so a crisp band
  that moved between zooms was no break. One was my arithmetic: `centrePt × scale` with the
  inexact 1040 scale put a 474.5 a hair under the half and called PowerPoint's correct row a miss,
  which is why 1040 reads 239 of 242 and not 238. The rest were the record overstating the rule's
  domain, a "twelve widths" left from the design's first draft, the fixture's writer not counting
  a record's key against the width, and comment caps; every one is fixed.
- **Mutants, 13/13 killed in the third pass.** The pen 3/3 (every pen shifted, none shifted, the
  width rounded down); the renderer 8/8 (a turned rectangle still crisp, curves crisp, the shift
  not turned, the fill not crisp, the box never stretched, any edge counted as aligned, a fill
  never split from its pen, a fill split with no device); the mask 2/2 (the translate kept, the
  translate stripped from any path). One more — the shift applied to a clipped band — survived
  because the band never spreads the shift at all, so the clause it removed was redundant; the
  clause is gone.
- **Mutants, 23/23 killed in the second pass.** The marker pen 3/3 (a one-point floor, no device
  floor, the nominal width above two points); the line ends 7/7 (every path open, the head not
  reversed, every head centred, head and tail swapped, the head beside the line, a gradient head
  keeping the server's url, the arrow's vertex on the endpoint); the ratio 4/4 (ignored by the
  renderer, ignored by the stage's re-show, the watcher listening once, the floor dropped); the
  strip 3/3 (drawn before the frame, waiting one frame, a cancel that does nothing); the by-deck
  rollup, `mergeOracle`, and the gate's masks and its ratio clause 6/6. The report's per-slide
  table has no test and its column fix was checked by reading the report.
- **Mutants, 28/28 killed.** F2's renderer rule 9/9 (among them width ≤ 0 → null, no
  `vector-effect`, a 0.75-pt minimum, the minimum on every width, `ceil` for `round`, the dash
  array from width 0, the marker from width 0, `w·s` unconditionally); the device floor 1/1
  (0.05 → 0.01); bordered picture 1/1; bullets 8/8; the effect filter 3/3 (radius left in EMU, the
  inverse scale through `num()`, no inverse group); the band reach 1/1 (the device pixel dropped
  from it); the harness 4/4 (a mask that blanks every `width`, a rename without its boundary, the
  hundred-slides clause dropped, cross-origin counted as static); the page 1/1 (resize on zoom).
- **Behaviour-preserving refactors** — the media resolver, `geometryOf`'s cell, the oracle loader,
  `record.ts`'s `Baseline` — left all 155 raster digests of the day unchanged.
- `pnpm gate1` green in both save modes with PowerPoint opening the result without repair, and the
  bench smoke over its generated 20-slide deck (census 0.166 s); the plan asked for both and the
  first version of this record did not say they had run.
- **An independent audit** of the plan against the tree — 37 read-only agents, one per plan
  section and one refuter per reported gap — found 20 gaps that survived refutation and 12 that
  did not. Every survivor is fixed above or recorded below.
- `pnpm check` green: structure, layering, legal, references, corpus, format, build, lint,
  typecheck, round trip 55/55, fidelity, Gate 2, Gate 3, package QA, tests.

## What the release said

Pull request #28 merged on CI run 34717677018; release run 34717979793 published
`@pptx-studio/paint` 0.2.0, `@pptx-studio/render-svg` 0.3.0, `@pptx-studio/render-dom` 0.2.0 and
`@pptx-studio/cli` 0.3.1 over OIDC, with `model` bumped for the dependency, and pushed the version
commit to `main`. The canary was green before the release (run 34717716541) and red 38 seconds
after it — `npm install` asked for `@pptx-studio/cli@0.3.1` at 20:48:08 and the registry had
published it at 20:47:30 and was not yet serving it — so it opened issue #29 as ADR 0053 designed,
and the next run (34718207672) was green and closed it. Replication lag, not the release; a canary
dispatched straight after a publish should expect one red run.

Pull request #31 merged on CI run 34745569128 (main's run 34745831906 green on the merge). The
first release dispatch refused itself — no completed CI run for the merge commit yet — which is the
gate ADR 0053 built working as designed; the second, run 34746098190, published
`@pptx-studio/paint` 0.3.0, `@pptx-studio/render-svg` 0.4.0, `@pptx-studio/render-dom` 0.3.0,
`@pptx-studio/cli` 0.3.2 and `@pptx-studio/model` 0.1.3 for the dependency, and pushed the version
commit to `main`. The canary was green before the release (run 34745835364) and, dispatched once
the registry served the new CLI rather than 38 seconds after the publish, green after it (run 34746231070) with no issue opened.

Pull request #33 merged on CI run 34754633226 (main's run 34754916075 green on the merge
`be85847`); release run 34755115512 published `@pptx-studio/paint` 0.4.0,
`@pptx-studio/render-svg` 0.5.0, `@pptx-studio/render-dom` 0.3.1, `@pptx-studio/cli` 0.3.3 and
`@pptx-studio/model` 0.1.4 for the dependency, version commit `7a09f71`. The canary was green
before it (run 34754940327). After it I waited for the registry to serve the new `render-dom`
and `cli` and dispatched; the registry was still serving `model` 0.1.3, the consumer install
asked for `^0.1.4`, and run 34755263913 went red and opened issue #34 — the lag the last
paragraph describes, on a package I had not waited for. Once every package resolved, run
34755351825 was green and closed the issue.

Pull request #36 merged on CI run 34763025354 (main's run 34763313914 green on the merge
`4c970b7`); release run 34763589307 published `@pptx-studio/render-svg` 0.6.0,
`@pptx-studio/render-dom` 0.3.2 and `@pptx-studio/cli` 0.3.4, version commit `9625cca`. The
canary was green before it (run 34763062307); after it I waited until `npm view` served all three
new versions and dispatched, and run 34763792565 was green with no issue opened.

## Deviations from the plan

- **Zoom is a mount, not `resize()`.** The plan named `MountedSlide.resize()` and said a
  device-pixel rule, if measured, would be the one exception. It was measured (T4, H1), so the
  exception is the design and `resize()` is deleted rather than kept beside it.
- **`--extend` was dropped.** An append-only record mode for a baseline that cannot otherwise grow
  in CI was designed and not built: deleting the stale Linux baseline and letting the dispatch job
  bootstrap both files is one mechanism instead of two, and `record-linux-baselines` records
  whichever is missing.
- **A record run exits 0.** `pnpm fidelity --record` wrote the baseline and then threw over the
  digests it had just recorded. It reports what it recorded now, and `pnpm gate3` was written the
  same way.
- **Five corpus count locks** in `tools/corpus/{lexical,suites}` were still at fifty-four decks when
  a46 landed; the commit that added the deck left them red. Fixed with the page.
- **The effect filter's unit, the band's reach, bullets and bordered pictures** were not in the
  plan; each is above.
- **PowerPoint, not LibreOffice**, as ADR 0038 decided, and the gate runs from committed grids.
- The CI job `record-fidelity-baseline` that ADR 0041 names is now `record-linux-baselines`.
- **Seven export widths, not the plan's six**: 1200 px, above.
- **The marker family is scored at 1 px**, not 0.15, above; the first version of this record said
  0.15 for every family.
- **M8 was carried and then built**, and O4 was carried, read honestly, and then measured: the
  third pass ran F3 and built the grid rule, above. The stroke rule's exact-half rounding is the
  960 export's and not the 1200 export's, and a curve's quarter-pixel bias is recorded and not
  drawn.
- **A named height stretches.** The plan and the first two versions of this record had the root
  keep its aspect in a named box; F2 had measured the export stretching into its box and the page
  was passing a rounded height. The renderer now does what the page asked for.
- **`sourceNote` per fixture claim** was planned, not built, and the manifest said "PowerPoint 365
  rendered the corpus deck" over our own raster digests and over 2.12's `blips.json`. Every claim
  now says where its bytes came from.
- **The by-deck table** in the fidelity report and **`mergeOracle` as a pure, tested function**
  were planned and arrived with the second pass; the per-deck means the first version quoted were
  computed outside the harness.
- **The strip's first chunk** ran before the stage, above; it now waits two frames, since one
  frame's callback runs before that frame paints.
- **A display ratio under the floor.** A thumbnail on a display at a third of a pixel per CSS
  pixel would have asked the renderer for a twenty-fourth of a pixel to the point and been refused;
  the page rounds at the floor's ratio when the display's is under it (`deviceRatioAt`).
- **The review of this pass** — six read-only reviewers, one per dimension, a refuter per finding
  — confirmed 42 findings and refuted 5. The two renderer defects and the report's overwritten
  column are above; the rest were wording, and every one is fixed.
- The plan's `findings.textLinearWithinPx` does not exist; the fixture's `textTolerance` is prose,
  and the test states the tolerance itself.
- **The fourth pass ran F3 again, at fourteen widths,** to close the three things the third pass
  left on the plan row: the half pixel at the widths no probe reached, a scale that is not a
  quarter, and a clipped band under Blink's clip. Its design was wrong once on the way: I had
  taken "an eighth of a pixel per point" for the thing that decided snapping, and 1000 px — not an
  eighth, 75 dpi — behaved like 120; 1040 and 1120 were added to reach the cell the first design
  had no width in, and they showed that a whole dpi with a whole height is not the rule's domain
  either, only the powers of two are. The fixture is regenerated,
  not adapted; it is written so that a record stays on one line under prettier, since fourteen
  widths of profiles would not otherwise fit the corpus cap.

## Open questions

1. **The strip's 2x raster.** The strip now stretches into its 120 × 68 box as the export does, so
   the quarter-pixel letterbox is gone; what remains is that a 68-px thumbnail is 136 device pixels
   tall on a 2x display and the 240-px oracle is 135, which is why the 2x pass compares the strip's
   markup and not its raster.
2. **Pattern fills under 480 px** in PowerPoint's own export (`patternBelow`): whether the thinning
   is the export's or the screen's is not known.
3. **The exact half pixel off the powers of two.** At 1040, 1120 and 1200 the width and the parity
   of a _k_ + ½ pen round apart, and not the same way apart; at 1120 an outline's frame edge on a
   half, 479.5, rounds down while a picture edge on a half, 304.5, rounds up. The fixture keeps
   every case and no arithmetic reproduces them. The page never draws at
   such a scale; a renderer for a fit zoom would.
4. **The screen at a fit zoom.** The export at a fractional dots per inch is a resampled picture
   of something, not a render at that scale, so it cannot say whether PowerPoint's own screen
   snaps at 63 %. Only a capture of the window can, and no experiment here has one.
5. **The curve rasteriser's quarter pixel**, and a flat end's half column, are measured and not
   drawn. The export ends an odd pen on a half-covered column (a quarter further for a one-pixel
   pen); `crispEdges` ends ours on a whole column, half a pixel from it. Curves stay antialiased
   where they lie.
6. **An `algn="in"` pen a device pixel outside its frame.** Measured, and drawn up to a pixel and
   a half too far in, because the constant is a device pixel and the markup is the same at every
   zoom. A masked attribute on the band — as the border's crispness now is — could carry it; one
   deck in the corpus has such a stroke.
7. **The clip-reach edge.** Skia antialiases the outline band's clip differently once its reach
   passes roughly 600 pt; a picture that is not a rectangle with a border wider than ~280 pt would
   cross it at any zoom. No corpus deck has one.
8. **A stroke at 9 dpi.** At 120 px a horizontal stroke gains a quarter-row above the rows the
   quarter-step mapping predicts, on 16 of its 28 strokes and 8 of its 42 ties, and 1000 and 1320 do not. The strip's
   column is bounded by it.
9. **Effect filters at DPR > 1** are written in points, which is one CSS pixel at 100 %. The 2x stage
   pass took 27 s for 100 slides where the 200 % column's screenshots took 13 s; the extra is not
   measured apart and is nothing like the seconds a rotated filter cost, but whether the pathology
   returns at other ratios is not measured.
10. **Line ends nobody measured.** A head on a freeform with several open subpaths, dashes running
    under a head, a round cap poking past a triangle's tip on a thick line, the caps of an open
    arrow's arms, and where along a gradient-stroked line its head samples the ramp: the renderer
    does the plain thing for each and no fixture says whether PowerPoint does.
11. **A second real deck.** One deck outside the corpus has been through the page. That is an
    anecdote, not a rule.
