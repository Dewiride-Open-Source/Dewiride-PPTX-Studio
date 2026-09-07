# 0022 — Gradients and pattern fills

Status: accepted
Sub-phase: 2.7
Date: 2026-09-03, amended 2026-09-07
Fixture: `corpus/ground-truth/fills.json`
Supersedes nothing. Amends the approved plan's 2.7 in three places, listed below.

## Context

2.6 settled what a DrawingML _colour_ resolves to. A fill is a field of colours,
and almost nothing about how DrawingML lays that field out is written down in a
form you could implement from. `a:lin/@ang` has a zero direction and a sense the
standard never states. `@scaled` gets one sentence that does not say what the
transform is. `a:fillToRect` describes a rectangle SVG has no primitive for. The
54 pattern tiles are not in the standard at all — only their names are.

The plan proposed sourcing the tiles from Mono libgdiplus and predicted a "1.875
gamma ramp pre-sampled to 33 stops". Downloading libgdiplus needs a permission
this session did not have, and the gamma claim is exactly the kind of thing 2.6
overturned twice. So both were measured.

## The experiment, in two halves

**PowerPoint authored first.** Before any probe was written, PowerPoint was asked
to _make_ fills through its own object model and the resulting `.pptx` files were
read. That cost one script and settled three things:

- `Fill.Patterned(i)` over the whole `MsoPatternType` range is accepted for
  exactly **54** values, and PowerPoint writes each name into `a:pattFill/@prst`
  itself. The enumeration is therefore PowerPoint's, not a transcription.
- A two-colour gradient is written with **two** stops. Not 33. Whatever the
  pre-sampled ramp is, the authoring UI does not do it — so if it is real it
  belongs to the renderer, and only a bitmap can say. The largest stop list
  PowerPoint writes anywhere in its 24 preset gradients is ten.
- PowerPoint writes `a:gsLst` **out of `@pos` order**. Its own from-centre
  variant is `pos="50000"`, `pos="0"`, `pos="100000"`.

**Then C3.** 182 probes in 34 packages: `tools/ground-truth/paint/fills/probes.ts`,
`tools/ground-truth/paint/fills/build-deck.ts`, `tools/ground-truth/paint/fills/read.ps1`, `tools/ground-truth/paint/fills/analyse.ts`. Ramps sampled at 481
points across a full-width strip; two-dimensional probes on a 13 × 13 lattice;
pattern tiles read at 1280 × 720, which over a 13.333-inch slide is exactly 96
DPI and the only resolution at which a tile is 1:1 with its own pixels; and
profile corners at 1:1 over 257 pixels, because a 481-point strip across a
1800-pixel shape cannot resolve a feature one pixel wide.

Fifteen are hostile and each is alone in its package, so a refusal names itself —
the rule C2 paid for. Thirty-eight came from the second pass; see the section
below.

## Decisions

**1. A gradient with two distinct colours and a stop at each end is not
interpolated linearly.** Black to white paints `BABABA` at its midpoint, not
`808080`. Everything else is linear in sRGB, measured at worst 0.8 of a byte over
1920 samples against 59 for the curve.

**2. The rule is two _colours_, not two _stops_.** `[0 white][50% black][100%
white]` has three stops, two colours, and is curved. `[0 black][50% 808080][100%
white]` has the same three positions, three colours, and is linear. Two stops at
20% and 80% are linear. Those three facts together are exactly the gradient GDI+
can express as a two-colour `LinearGradientBrush` with a blend array — whose
factors must start at 0 and end at 1, and which is the only one of the two brush
constructors that applies gamma correction. A mechanism, not a coincidence.

This matters more than it sounds: variant 3 and variant 4 of _every_ gradient in
PowerPoint's own gallery are three stops in two colours. "Two stops" would have
been the wrong rule for a very large number of real decks.

**3. The blend space is a plain 2.2 power** — not the sRGB piecewise transfer,
not linear light. One weight curve applied in that space reproduces five
independently measured colour pairs to within one byte per channel.

**4. The curve ships as measurement, not as a formula.** No closed form fits: a
power law, `1 - (1 - t)^γ`, smoothstep, smootherstep, a logistic power, and the
sRGB and linear-light transfers were all fitted over 1920 samples and the best is
eight bytes out at its worst. `1 - (1 - t)^1.875` is the closest
single-parameter fit through the middle of the range — the plan's constant is not
nothing — but it misses the dark end by thirteen.

**5. Fifteen pre-sampled stops, adaptively placed, fitted in both directions.**
SVG interpolates stop colours linearly in sRGB and `color-interpolation` is not
implemented for gradients in any shipping browser, so pre-sampling is the only
way to draw the curve. Fifteen knots hold both measured ramps to within 1.9 of 255. The knots must be fitted to the ascending _and_ descending ramp: a set
fitted to one is within 1.2 forwards and thirteen backwards.

**6. `gradientColorAt` is defined as interpolation between the emitted stops.**
Not as the weight curve evaluated directly, which is nineteen bytes away from
PowerPoint on the steep end. The definition that matters is the one that agrees
with what a browser will paint.

**7. Alpha is linear in the factor, colour is not.** A white ramp from alpha 0 to
alpha 100% over red paints `FF4040` a quarter of the way — a linear quarter, not
the curve's 0.142. The same split GDI+ makes.

**8. Stops are sorted by `@pos`; the parse keeps document order.** PowerPoint's
object model hands them back unsorted, so both facts are real and a writer that
sorts on save has edited a part nobody asked it to touch.

**9. `@ang` is clockwise from +x with y down.** Six angles on a square each
measured within a tenth of a degree of the value written.

**10. `@scaled="1"` bends the angle _towards_ the long axis.** The written angle
is measured in a space where the shape is a unit square, and what stretches is
the _bands_, not the arrow — so the direction of steepest change is the normal to
the stretched bands: `atan2(w·sin a, h·cos a)`. On a 3:1 shape a written 45°
measures 71.4° against 71.6° predicted; the inverted reading, which is what a
careful derivation from GDI+'s `isAngleScalable` produces, predicts 18.4°. Three
shapes, one formula.

**11. The ramp spans the shape's full projection onto that direction**, centred
on the shape's centre, so the extreme corner is exactly the first stop. At 45° on
a square the two off-axis corners read 0.50; on a 3:1 shape with `scaled="0"`
they read 0.75 and 0.25 against 0.751 and 0.249 predicted.

**12. `gradientUnits="objectBoundingBox"` is the wrong SVG construct for either
value of `@scaled`.** It reproduces the bend but also shears the ramp, putting
the corner of a 3:1 shape at 0.21 where PowerPoint puts it at 0.01. Compute the
vector and emit `userSpaceOnUse`.

**13. Path gradients do not reverse stop order.** Stop 0 sits at the focus and
the last stop at the edge — the same way round as SVG's `radialGradient`. The
plan says they reverse; measured, they do not.

**14. `a:fillToRect` behaves as a point at its centre.** A focus rectangle
covering the middle half of a shape paints byte-for-byte what a point focus at
the centre paints, over all 169 samples. There is no flat region to model.

**14a. The two ways of writing nothing are two different pictures.** An
**all-zero** `fillToRect` is the top-left corner. An **absent** `a:fillToRect`
is the **centre** — `pathdef-norect-rect` fits the centre at rms 0.45 and the
corner at 72.6. So the element is nullable in the model rather than defaulted:
collapsing the two spellings moves the light of every gradient that omits it.

**14b. An absent `@path` paints the box ramp.** `pathdef-nopath` fits `box` at
0.45 and the circle at 24.8. The schema makes the attribute optional and names no
default; this is the measured one.

**15. `path="circle"` is a circle, reaching its last stop at the corner.** On a
3:1 shape the top edge's midpoint reads 0.30 and the left edge's 0.94, the ratio
of their distances in shape units; an ellipse fitted to the box would read 1.00
for both. `path="rect"` is Chebyshev with each axis normalised to the distance
from the focus to the edge it faces, which also reproduces the corner focus
PowerPoint writes for itself. `path="shape"` follows the geometry's own outline —
identical to `rect` on a rectangle, elliptical on an ellipse.

**15a. An off-centre focus does not move the circle. It is SVG's focal radial.**
The circle the last stop lands on is the _shape's_ — centred, with the
half-diagonal for a radius — whatever the focus says; the focus only says where
the ramp starts, exactly as SVG's `fx`/`fy` do. The contour through a point is
the circle grown from the focus towards that outer circle, so the ramp position
is the smaller root of

```
|P − F − t(C − F)| = tR      C = the shape centre, R = hypot(w, h) / 2
```

Seventeen probes over eight foci and three aspect ratios fit this to rms 0.53 and
two bytes at worst. The reading 2.7 shipped first — concentric circles about the
focus, scaled to the farthest corner — is out by up to 49, and the same
construction in a space where the shape is a unit square by up to 39. That last
one only shows on a shape that is not square, which is why the wide and tall
probes exist.

Two consequences worth stating. The renderer emits this as one
`<radialGradient>` with `fx`/`fy` and needs no approximation at all. And `a` in
that quadratic is zero exactly when the focus lands on the outer circle — which
a corner focus on a square shape does — so it is solved in the stable form; the
textbook one divides by almost nothing and paints the shape inside out.

**16. The 54 tiles are eight pixels square at 96 DPI: 6 pt, 76200 EMU.** The same
shape exported at 640, 1280, 1920 and 2560 pixels wide gives periods of 4, 8, 12
and 16, so the tile is a fixed _physical_ size — neither a fraction of the shape
nor a fixed number of device pixels. In SVG that is `patternUnits="userSpaceOnUse"`
with a size in slide units.

**17. Tiles are stored as coverage, not bits.** Fifty-one are one-bit; `dnDiag`,
`upDiag` and `diagCross` are drawn as antialiased diagonal lines and contain
greys. `lgGrid` and `cross` measure byte-identical — the only such pair, and
consistent with the GDI+ heritage where `LargeGrid` aliases `Cross`.

**18. A missing `a:fgClr` paints black and a missing `a:bgClr` paints white**,
regardless of the theme. Both are optional and PowerPoint accepts either absent.

## What building it found

### The plan's three claims, scored

| claim                                                                        | verdict                                                                                                                                                                                                                                                                                  |
| ---------------------------------------------------------------------------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| "the 1.875 gamma ramp pre-sampled to 33 stops for two-stop 0/100% gradients" | **half right, and the useful half.** There _is_ a curve, it _is_ special-cased, and pre-sampling _is_ the only way to draw it. But the case is two _colours_ rather than two stops, 1.875 is a fit rather than the mechanism, and fifteen adaptive stops beat thirty-three uniform ones. |
| "Path gradients reverse stop order"                                          | **wrong.** Stop 0 is at the focus, as in SVG.                                                                                                                                                                                                                                            |
| "54 pattern tiles codegen'd from Mono libgdiplus (MIT), not Wine (LGPL)"     | **54 is right; the source is not needed.** PowerPoint enumerates the names and paints the pixels, so the tiles are measured and the licence question does not arise.                                                                                                                     |

### PowerPoint refuses seven things and accepts eight

Each hostile probe was alone in its package, so each verdict is attributed rather
than inferred.

| refused (opened only with repair)         | accepted                                     |
| ----------------------------------------- | -------------------------------------------- |
| `a:lin/@ang` negative                     | `a:gs/@pos` spelled `"50%"`                  |
| `a:lin/@ang` past 21600000                | `a:gsLst` element absent entirely            |
| `a:gs/@pos` = 150000                      | stops with no `a:lin` and no `a:path`        |
| `a:gs/@pos` = −50000                      | two stops at the same position               |
| a gradient with **one** stop              | an inverted `a:fillToRect`                   |
| an **empty** `a:gsLst`                    | `a:pattFill` with no `@prst`                 |
| `a:pattFill/@prst` not in the enumeration | `a:pattFill` missing `a:fgClr`, or `a:bgClr` |

The pair worth staring at is rows 5 and 6 against row 2: an empty `a:gsLst` is
refused and a missing one is fine.

### PowerPoint softens every corner in a ramp profile, by a fixed number of pixels

Wherever the profile has a corner — a stop boundary, or the apex of a V — the
rasteriser rounds it off. It is a distance in **device pixels**, not a share of
the ramp, and the probes settle that without needing a model of the rounding at
all: the corner is drawn with arms of 1, 2, 4 and 8% of the shape, and each deck
is exported at 1920 and at 3840. That makes three pairs which are the same corner
in device pixels and different fractions of the ramp — `arm1` at 3840 against
`arm2` at 1920, and so on — and all three pairs agree **to within one byte**.

The magnitude is under a pixel: at 21 bytes per pixel the apex falls about ten
bytes short of where the two straight arms meet, and by 3 bytes/px it is one.
Most of even that is the apex not landing on a pixel centre. Still recorded
rather than modelled — it is sub-pixel and an SVG stop list cannot express it —
but the shape of the answer is now known rather than guessed.

One observation that is not part of the rule: on the steepest probe, where each
arm is 1% of the shape, the two exports disagree by up to 18 bytes **at the outer
stops** while agreeing at the corner. That is PowerPoint's export resampling a
very short ramp, not a property of gradients.

### `a:tileRect` and `@flip` are very nearly decorative

A `tileRect` smaller than the shape does make a _linear_ gradient repeat, and the
repeat is mirrored whatever `@flip` says. A `tileRect` on a _path_ gradient does
nothing measurable: PowerPoint's own from-corner fill carries
`tileRect l="-100000" t="-100000"`, and with it and without it all 169 samples
agree to within one byte. So the attribute PowerPoint writes for itself is one it
does not read.

**`@flip` does nothing at all**, and the first pass could not have said so. Its
four flip probes were a _centred_ path tile, which is its own mirror image on
both axes, so their agreement was evidence of nothing. Repeated with a tile that
has an off-centre focus, and again with a 30° linear ramp — both asymmetric on
both axes, and a mirrored copy of the linear one differs from itself by 155 bytes
— all four modes still agree to within one byte. The attribute is inert on a
gradient fill. Note this is `a:gradFill/@flip`; `a:tile/@flip` on a _blip_ fill
is a different attribute and does mirror alternate tiles (ADR 0036).

### A gradient on a slide background is laid out over the slide

A background has no shape box, so the ramp needs a rectangle from somewhere, and
the slide is the only candidate in sight. Four decks carry the same fill on their
background and on a shape in one quadrant; predicting the background over the
slide's own rectangle lands within a byte, for a horizontal ramp, a vertical one,
a centred circle and an off-centre one. This confirms what `render-svg` already
did rather than changing it.

### The measurement had to be fixed twice, by its own evidence

Two bugs in the analysis were caught by the data rather than by review, and both
are worth recording because the same trap is waiting for anyone who repeats this.

_Foreground by majority._ Choosing the tile's foreground as the minority colour
inverts `pct70`, `pct75`, `pct80` and `pct90` — and an inverted `pct75` is
bit-identical to `pct20`. The collision is what gave it away. The foreground has
to come from the markup.

_Precision by direction._ Recovering the blend weight from the black-to-white
ramp means inverting a 2.2 power, and near white one byte of an eight-bit export
is 0.009 of weight — so a weight read from that ramp predicts a _descending_
channel fourteen bytes out at the far end. That is a fact about the measurement,
not about PowerPoint. Each half of the range is now read from the ramp that
resolves it.

### The mutations, and the two gaps they found

Sixteen mutations, each applied by literal replacement to a `diff -r`-verified
baseline, type-checked and run. Failures out of 42:

| mutation                                                  | killed |
| --------------------------------------------------------- | -----: |
| the curve blended in sRGB, with no gamma at all           |      6 |
| the curve emitted as two stops instead of pre-sampled     |      6 |
| the curve blended in linear light                         |      5 |
| stops taken in document order                             |      4 |
| the curve applied without requiring a stop at each end    |      4 |
| path gradients reversed, stop 0 at the edge               |      4 |
| the curve gated on two _stops_ rather than two colours    |      3 |
| `@scaled` bending the other way                           |      2 |
| `@scaled` ignored entirely                                |      2 |
| `path="rect"` measured Euclidean rather than Chebyshev    |      2 |
| alpha following the curve rather than the factor          |      1 |
| the ramp spanning the diagonal rather than the projection |      1 |
| `path="circle"` as an ellipse fitted to the box           |      1 |
| the tile sized one point per pixel                        |      1 |
| tiles thresholded to one bit                              |      1 |
| a missing `a:fgClr` defaulting to white                   |      1 |

Two of those killed nothing on the first run. Alpha following the curve was
invisible because the alpha probe has a backdrop and was excluded from the
sweep; thresholding the antialiased tiles was invisible because the only
assertion about them read the generated table rather than what `resolvePattern`
returns. Both gaps were real and both are now closed — which is the argument for
running the mutations at all rather than trusting a suite that is green.

## Consequences

- `packages/paint` gains `fill.ts`, `gradient.ts`, `pattern.ts` and two generated
  tables. `@pptx-studio/geometry` is still not a dependency, so `path="shape"`
  falls back to the bounding rectangle; 2.10 can supply the outline.
- A renderer emits `userSpaceOnUse` gradients with computed endpoints, and
  `patternUnits="userSpaceOnUse"` with a 6 pt tile.
- `svgStops` emits colour and opacity separately rather than folding alpha into
  `rgba()`, because SVG interpolates them separately and premultiplied.
- The `Fill` union has a place for `a:blipFill` and `a:grpFill` and nothing behind
  either.

## What is not done

- **Nothing is painted.** This produces numbers and stop lists.
- `a:blipFill` is unassigned in the plan; Gate 2 asks for cropped image fills and
  no sub-phase owns them. Flagged for the plan rather than absorbed here.
- `a:grpFill` needs the enclosing group's rectangle: a child must show the
  _slice_ of the group's gradient under it, not a copy. 2.10.
- Slide and master backgrounds carry `a:gradFill` with no shape at all. The
  extent is presumably the slide; unmeasured.

## The second pass, 2026-09-07

Five of the six open questions below were left open when 2.7 was accepted. They
are now closed, by 38 further probes in 11 further packages — decisions 14a, 14b
and 15a above, and the two "what building it found" sections on softening and on
`@flip`. The experiment is the same four files, extended; the fixture is the same
file, regrown from 144 probes to 182.

Two things are worth separating out.

**The off-centre focus was a real defect, not a gap.** 2.7 shipped the natural
generalisation of its verified rule and flagged it `centred: false`. That reading
is wrong by up to 49 bytes, and it took eight foci to see it — the single probe
2.7 had fits four different constructions about equally badly, which is why it
looked like an open question rather than a bug. One sample is an anecdote even
when it disagrees with you.

**Re-running the first 144 probes reproduced them exactly.** Both generated
tables — `gradient-ramp.ts` and `pattern-tiles.ts` — regenerate byte-identically
from the new fixture, and the seven hostile packages PowerPoint repairs are the
same seven. That is the reproducibility check nobody asked for and it is worth
having.

One defect in the harness was found and fixed on the way: `read.ps1` wrote
`@(@(1920, 1080))`, which PowerShell unrolls, so every deck had been exported a
second time at width 1080. Nothing read those files, so no measurement was
affected.

## Open questions

1. **What the corner rounding actually is.** It is now known to be a fixed
   distance in device pixels rather than a share of the ramp, and to be under a
   pixel. Which filter it is remains open: box widths from 0 to 4 px all fit the
   six gentler probes at rms 0.4 to 0.7, and none fits the steepest.
2. **`path="shape"` on anything but a rectangle.** The outline case needs
   `@pptx-studio/geometry`, which `@pptx-studio/paint` deliberately does not
   depend on, so both renderers still draw an ellipse's `shape` ramp as a box.
   Two probes measure it and neither is asserted against.
3. **`path="rect"` in SVG.** Still an inscribed ellipse rather than concentric
   rectangles; SVG has no primitive for the Chebyshev metric.
4. Carried from 2.6, unchanged: what decides PowerPoint's rounding at an exact
   half; whether PowerPoint for the web resolves `a:sysClr` per viewer; whether a
   `clrMapOvr` reaches a colour inside a theme style; whether `a:scrgbClr` is
   clamped before or after its transforms; what a `schemeClr` naming an undefined
   slot does.

## Verification

- `pnpm check` green, both passes. At the second: structure, layering, docs,
  corpus, format, lint, typecheck, build, round trip 52/52, fidelity, package QA,
  3157 tests across 89 files.
- `packages/paint/src/fills/fill.test.ts` — 78 tests, every one anchored to the
  fixture. The load-bearing ones: every ramp in the fixture reproduced to within
  two bytes with two named exceptions; every angle probe predicted from the shape
  and the two attributes; the 54 tiles re-derived from the fixture so the shipped
  table cannot drift from the measurement; and **every path probe predicted
  sample by sample** — 26 of them, rms under 2 bytes each.
- `tools/ground-truth/paint/fills/analyse.ts` refuses to write a fixture it
  cannot fit. It throws if the named model misses any probe of its kind, and it
  throws again if any _other_ model fits them all, because then the probes did
  not separate the candidates and the fixture would be recording a coincidence.
- Mutation sweep, second pass: **9 of 10 killed**. The survivor takes the far
  root of the focal quadratic instead of the near one, and is equivalent — the
  focus lies inside the outer circle, so the two roots straddle zero and exactly
  one survives the filter. It is commented as such at the point it is made.
- All eleven new packages opened with **no repair**, including all four that omit
  something the schema makes optional. The seven hostile packages that repair are
  the same seven as in the first pass.
