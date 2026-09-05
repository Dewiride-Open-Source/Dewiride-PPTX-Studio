# 0023 — Strokes and effects

Status: accepted
Sub-phase: 2.8
Date: 2026-09-04
Fixture: `corpus/ground-truth/lines.json`
Supersedes nothing. Amends the approved plan's 2.8 in two places and adds one
element the plan does not mention.

## Context

2.7 settled how a shape's interior is painted. This one is about its edge, and
about the things drawn around it — and almost none of it is written down in a
form you could implement from.

ECMA-376 names eleven preset dashes and gives no arrays. It names six arrowheads
and gives no geometry. It gives `ST_LineEndCap` a default that Office does not
use. It says `a:miter/@lim` limits the miter and does not say against what, or
what the limit is when the attribute is absent. And `@blurRad` is "the blur
radius" with no statement of what a radius is — which is the one number an SVG
filter needs, because `feGaussianBlur` takes a standard deviation.

The plan proposed taking arrowhead geometry from **LibreOffice's vertex tables**.
LibreOffice is MPL-2.0, a file-level copyleft this Apache-2.0 tree cannot take —
the same objection the plan itself raises in 4.2 about `predefined-table-styles.cxx`,
and the same shape of objection 2.7 had to libgdiplus. So the geometry was
measured instead, which needs no permission and no licence and is more faithful
anyway: what a renderer needs is what PowerPoint paints.

## The experiment, in two halves

**PowerPoint authored first.** The move that made 2.7 cheap, repeated. Before any
probe existed, PowerPoint was asked to _make_ strokes and effects through its own
object model (`tools/ground-truth/paint/lines/author.ps1`) and the files it saved were
read. That cost one script and settled six things:

- `Line.DashStyle` reaches all **eleven** `ST_PresetLineDashVal` names. Two of the
  twelve object-model values differ **only by `@cap`**: `msoLineSquareDot` is
  `val="dot" cap="sq"` and `msoLineRoundDot` is `val="dot" cap="rnd"`. In
  PowerPoint's own gallery, the cap _is_ the difference between two entries.
- Apart from those two, PowerPoint writes **no `@cap` at all**. So whatever the
  default is, it governs nearly every line in nearly every deck.
- A line the user has not styled carries **no `a:ln` element whatsoever**; its
  stroke comes from `p:style/a:lnRef idx="2"` into the theme's `a:lnStyleLst`.
  "What is the default line width" is a theme question before it is a schema one.
- Arrowheads are exactly **six types × three widths × three lengths**. The seventh
  type index is refused.
- The 43 legacy shadow presets save as **three different elements** — `a:prstShdw`
  (20 of them), `a:outerShdw` (13) and `a:innerShdw` (9).
- `a:reflection` has thirteen attributes and PowerPoint's nine presets name all of
  them, so none has to be guessed.

**Then C4.** 213 probes across 47 packages: `tools/ground-truth/paint/lines/probes.ts`,
`tools/ground-truth/paint/lines/build-deck.ts`, `tools/ground-truth/paint/lines/read.ps1`, `tools/ground-truth/paint/lines/analyse.ts`. The dash and
arrowhead packages are exported at 3840 × 2160 — four pixels to the point — and
the rest at 1920 × 1080.

Fifteen of the 213 are hostile and each is alone in its package, so a refusal
names itself. **Thirteen of the fifteen were refused**, against seven of fifteen
in 2.7: the stroke vocabulary is policed far more strictly than the fill one.

## The measurement's own unit: crossings, not pixels

2.7 committed whole sample strips because a gradient's answer _is_ a curve. A
stroke's answer usually is not — where a dash begins and ends is a set of
positions. An antialiased edge crosses half coverage somewhere _inside_ one
pixel, and interpolating between the two straddling samples locates it to about a
tenth of one. So the fixture stores sub-pixel **crossings**, in points, and the
raw row only where the answer really is a shape.

That is what makes a four-width dash segment measurable to better than half a
percent, and it is why all 34 numbers in the dash table came out within 0.02 of a
whole number.

## Decisions

**1. The eleven dash arrays are exactly the ones everyone quotes.** For once the
folklore is right, and it is right three ways over: measured from the crossings,
cross-checked against an explicit `a:custDash` of each quoted array painting an
identical row, and shown to scale with `@w` at 3, 6 and 24 points.

| preset    | array | preset          | array       |
| --------- | ----- | --------------- | ----------- |
| `solid`   | —     | `dashDot`       | 4 3 1 3     |
| `dot`     | 1 3   | `sysDashDot`    | 3 1 1 1     |
| `sysDot`  | 1 1   | `lgDashDot`     | 8 3 1 3     |
| `dash`    | 4 3   | `lgDashDotDot`  | 8 3 1 3 1 3 |
| `sysDash` | 3 1   | `sysDashDotDot` | 3 1 1 1 1 1 |
| `lgDash`  | 8 3   |                 |             |

In multiples of the stroke width, which is the only unit they can be quoted in:
`a:custDash/a:ds/@d` is a percentage _of the line width_, so a width multiple is
the only form in which a dash can be written back into a file.

**2. `a:ln/@cap` defaults to `flat`, not to the `sq` ECMA-376 specifies.** A 200pt
line at 24pt wide measured 300..500 with no `@cap` and with `cap="flat"`, and
288..512 with `cap="sq"` — half a stroke width of overshoot at each end. The plan
predicted this and the plan is right.

What the plan does not say is how much it matters. PowerPoint writes `@cap` for
two of its twelve dash styles and for nothing else, so this default is not an
edge case behind an unusual attribute; it is the state of essentially every line
in every deck.

**3. A cap does not lengthen a dash — PowerPoint compensates, and so must we.**
This is the finding the plan gestured at with "cap compensation" and is the one
with the most direct consequence for the renderer.

`dot` is `[1, 3]`. With `cap="rnd"` it still paints ink of **exactly one width**
and a gap of **exactly three** — the same as `cap="flat"`, to within 0.003 of a
width. So PowerPoint shortens the geometric segment and lets the cap add the
length back.

SVG does not. `stroke-dasharray` describes the path the cap is then added to, so
copying the array across unchanged with `stroke-linecap: round` makes every dot
twice as long as it should be. `dashArray()` therefore shortens each dash by one
width and lengthens each gap by one whenever the cap is not flat. `dot` becomes
`[0, 4w]`, which a round cap renders as the disc it is meant to be, and the
period survives — which it must, or the pattern drifts along the line.

**4. The join defaults to `round`, and the miter limit to 8.** ECMA-376 states
neither. A stroked corner with no join element and one with `a:round` both reach
exactly half a stroke width past the corner; a bevel reaches 0.33.

The limit was bracketed with eight corners of increasing sharpness. Miter ratios
of 2, 4, 5, 6, 7 and **7.91** all painted a full miter; **8.01**, 8.11, 12 and 20
all fell back to a bevel. And `@lim` is measured against the **whole** stroke
width, not half of it: on a corner of ratio 5, `lim="400000"` clips and
`lim="600000"` does not.

Both numbers differ from SVG's, whose defaults are `miter` and `4`. So
`stroke-linejoin` and `stroke-miterlimit` are always written out explicitly —
an omitted attribute would be a wrong attribute twice over, and every corner
between four and eight widths would render differently.

**5. `@cmpd` subdivides `@w`; it does not add to it.** All five forms sum to
exactly one stroke width:

| `@cmpd`     | rails and gaps, outside first |
| ----------- | ----------------------------- |
| `sng`       | 1                             |
| `dbl`       | ⅓ · ⅓ · ⅓                     |
| `thickThin` | 0.6 · 0.2 · 0.2               |
| `thinThick` | 0.2 · 0.2 · 0.6               |
| `tri`       | ⅙ · ⅙ · ⅓ · ⅙ · ⅙             |

A `dbl` at 36pt is two 12pt rails with a 12pt gap spanning 36pt in total — not
two 36pt rails, and not a 72pt band. The table is stored as whole **sixtieths**
rather than as decimals, because every value is a third, a fifth, a sixth or a
tenth and sixtieths carry all four exactly; `compoundRails` divides by sixty
last. A rounded decimal left a 36000 EMU double stroke six thousandths narrow,
which a test caught.

The plan says to "treat `@cmpd` as `sng` but preserve the attribute". Preserving
it is right and unchanged. Treating it as `sng` is now a choice rather than a
necessity, and 2.10 can make it with the table in hand.

**6. The six arrowheads, measured rather than transcribed.** Sizes are the same
three for both `@len` and `@w`: **sm = 2, med = 3, lg = 5** stroke widths,
confirmed on all 54 combinations and at stroke widths of 4, 8 and 16 points.

Three findings inside that:

- **A diamond and an oval are centred on the line's endpoint.** The other three
  put their tip on it. Getting this wrong displaces a large head by half its
  length, and the `ah-both` probe measured it directly: a large oval head on a
  16pt line reached 40pt _before_ the line's own start.
- **A stealth shares its silhouette with a triangle.** Only how much of that
  silhouette is inked separates them — 0.84 against 1.00 — and the notch is a
  quarter of the head's length.
- **An open `arrow` is a stroked V, not a filled outline**, which is why its
  measured silhouette is wider than its nominal width.

The outlines were recovered by fitting three families to the measured per-column
silhouette rather than by thresholding it. That is not fussiness: an ellipse's
silhouette reaches the shaft's own height well inside its back edge, so a
threshold reports every oval about nine percent short, and the first run of the
fit did exactly that.

**7. Every blurred edge in DrawingML is one operator with one constant.**
`a:outerShdw/@blurRad`, `a:blur/@rad`, `a:glow/@rad` and `a:softEdge/@rad` are
four attributes on four elements, and fourteen measurements across all four give:

> **σ = r / 3**, with the measured ratio never leaving 0.3300 … 0.3338.

The edge itself moves for two of the four: a glow grows the shape and a soft edge
shrinks it, both by **0.9 r**, while a shadow and a plain blur leave it alone. And
a glow's `r` is **half** the radius it declares — with that halving, a glow's two
numbers are the same two constants as everything else, and without it they look
like two more to remember.

Neither "Gaussian" nor "isotropic" is assumed. An error function fitted to the
measured edge left a worst residual of **0.84/255** across every radius, and the
ramp across a horizontal edge matched the ramp across a vertical one to **0/255**
over 75 samples. Comparing two ramps byte for byte tests the data; fitting two
sigmas and comparing them tests the model as much as the data.

**8. `dir` is zero along +x and turns toward +y.** Measured at all eight multiples
of 45° and cross-checked against `Shape.Shadow.OffsetX`/`OffsetY`. On a slide
that is clockwise, because y runs down the page.

**9. A shadow's `@algn` names the point that stays put, and defaults to `b`.**
`sx="200000"` doubled the box's width and left its horizontal centre; `sy="200000"`
doubled its height and left its **bottom** where it was; `sy="-100000"` mirrored
it about that same bottom edge; `kx="1200000"` widened the box by exactly
`h · tan 20°`. All four are consistent with one anchor, and with `@algn` absent
that anchor is bottom-centre — which is, unusually for this sub-phase, exactly
what the schema says.

**10. An inner shadow darkens the edge `dir` points at.** The naive derivation
says the opposite: offset the shape's own alpha along `dir`, invert, clip, and the
_left_ edge goes dark at `dir=0`. Measurement says the right edge does, and both
directions were probed because one reading is not a rule.

**11. The glow is painted over the outer shadow.** `a:effectLst` is an
`xsd:sequence`, so a file cannot express an order — which means the painting
order is a fact about Office and had to be measured. A shape carrying a green
glow and a red shadow paints green over red where the two overlap. Note that this
is the _reverse_ of the order the schema lists them in.

## What building it found

### The plan's three claims, scored

| claim                                                                    | verdict                                                                                                                                      |
| ------------------------------------------------------------------------ | -------------------------------------------------------------------------------------------------------------------------------------------- |
| "`a:ln/@cap` defaults to flat/butt in Office, not `square` as ECMA says" | **right**, and more consequential than stated                                                                                                |
| "11 preset dash arrays as line-width multiples with cap compensation"    | **right** on all three counts; the arrays are the quoted ones and the compensation is real and now quantified                                |
| "Arrowhead markers from LibreOffice's vertex tables"                     | **declined.** MPL-2.0 is a file-level copyleft this tree cannot take — the plan's own objection in 4.2. Measured instead, at no licence cost |

### `a:prstShdw` is not in the plan and is in real files

Twenty of PowerPoint's forty-three legacy shadow presets save as `a:prstShdw`,
an element the plan for this sub-phase does not mention and that no renderer this
project has read implements. Nothing documents what any of the twenty look like.

They are not all plain offset shadows. Against a 70 × 60pt block thrown 60pt
right, the twenty produce visible boxes ranging from 60 × 60 to 220.5 × 90 — some
are squashed to half height, some stretched to three times the width, some
skewed. What each paints is recorded in the fixture under `presetShadows`;
modelling them is 2.10's problem, but they can no longer be a surprise.

### Thirteen of fifteen hostile probes refused

| refused                                                         | accepted                                |
| --------------------------------------------------------------- | --------------------------------------- |
| negative `@w`; `@w` past the 1584pt maximum                     | an **empty** `a:custDash`               |
| a `prstDash/@val` outside the eleven                            | a `custDash` segment of **zero** length |
| a negative `custDash` segment                                   |                                         |
| `@cap`, `@cmpd`, `@algn` outside their enumerations             |                                         |
| a line-end `@type` outside the six, or a size outside sm/med/lg |                                         |
| a negative `a:miter/@lim`, `@blurRad` or `a:glow/@rad`          |                                         |
| `a:softEdge` with no `@rad`                                     |                                         |

The contrast with 2.7 is the finding: seven of fifteen refused there, thirteen of
fifteen here. Every enumeration in the stroke vocabulary is validated. The two
that get through are both `custDash` degeneracies, so a renderer has to cope with
an empty dash list and a zero-length segment and with nothing else.

### The measurement had to be fixed six times, by its own evidence

Worth recording, because each was a silently plausible number rather than a crash:

1. **A single-element array of pairs is not one.** PowerShell unrolls
   `@(@(1920, 1080))` back to `@(1920, 1080)`, so every non-fine deck exported a
   1920-wide and a 1080-wide bitmap instead of one 1920 × 1080. `,@(...)` is the
   only spelling that survives.
2. **A shadow behind an opaque shape measures nothing.** The first blur block used
   `dist=0` and read empty slide. Every shadow probe now throws its shadow clear
   first.
3. **A window holding both sides of a blurred slab is not an edge**, and the
   single-erf fit ran away to the top of its sweep rather than failing.
4. **The compound probes measured their neighbours.** A 36pt compound stroke
   reaches 22pt outside its own rectangle, and a 40pt gap between shapes put one
   shape's right rail inside the next one's read window — an extra rail on every
   row. Re-spaced to 320pt.
5. **A clamped read window does not start where it asked to.** Three sections
   recovered the sample origin by arithmetic on the requested window rather than
   from the sampler, and read one soft edge ten points to the left of where it is.
   The origin is now stored with the samples.
6. **A fit whose sample count varies rewards the candidate that sweeps in more
   easy samples.** Scoring each arrowhead candidate only over its own length and
   dividing by that count made every long candidate look better, and every narrow
   head came back about seventy percent too long.

### The mutations, and the gap one of them found

Twenty-four mutations against a `diff -r`-verified baseline, all killed. Failures
out of 66 tests:

| mutation                                                                                                                                              | tests failing |
| ----------------------------------------------------------------------------------------------------------------------------------------------------- | ------------- |
| blur σ = r/2                                                                                                                                          | 5             |
| edge growth 1.0 instead of 0.9                                                                                                                        | 4             |
| compound not subdivided                                                                                                                               | 4             |
| no cap compensation; compensation of half a width; marker sizes 1/2/3                                                                                 | 3 each        |
| shadow anchor `ctr`                                                                                                                                   | 3             |
| compensation on a flat cap; cap default `sq`; join default `miter`; miter limit 4                                                                     | 2 each        |
| compound rails at full width; every marker tip-anchored; glow at full radius                                                                          | 2 each        |
| default width 12700; inset stroke straddling; `dir` inverted; inner shadow offset negated                                                             | 1 each        |
| stealth with no notch; open arrow filled; soft edge with no erosion; inner shadow reserving outside room; dash table unrotated; glow under the shadow | 1 each        |

**One killed nothing on the first run:** painting the glow _under_ the outer
shadow. The test asserted the merge had three layers ending in the shape, which
is true whichever order the other two go in. It now traces each merge input back
through the graph to the effect it came from and asserts
`['outerShdw', 'glow', 'shape']`, plus a second test that the order is the same
whichever way the file lists them. That is the argument for running mutations
against a suite that is already green.

## Consequences

- `@pptx-studio/paint` gains `line.ts`, `effect.ts` and three generated tables
  (`dash-table.ts`, `compound-table.ts`, `marker-table.ts`). None is
  hand-written, and `line.test.ts` re-derives all three from the fixture — so a
  measured constant cannot quietly become somebody's memory of one.
- Five typed error codes: `LINE_DASH_UNKNOWN`, `LINE_COMPOUND_UNKNOWN`,
  `LINE_END_UNKNOWN`, `LINE_END_SIZE`, `EFFECT_NEGATIVE_RADIUS`. Every one names
  something PowerPoint itself refuses, so each is a caller mistake rather than a
  file in the wild.
- `effectFilter` returns a typed primitive list rather than markup, for the same
  reason `svgStops` does in 2.7: this package has no DOM, and the two renderers in
  2.10 build their nodes differently.
- `tools/ground-truth/paint/lines/author.ps1` joins `tools/ground-truth/paint/fills/author.ps1` as a committed
  experiment. Two sub-phases running, asking the format's author in writing has
  been the highest-yield step in both.

## What is not done

- **Nothing is painted.** 2.10 builds the renderers.
- **`a:prstShdw` is recorded, not modelled.** Twenty measured boxes are in the
  fixture; turning them into twenty transforms is 2.10's call.
- **`a:reflection` is a type and a measurement, not a renderer.** Its thirteen
  attributes are captured and one preset's profile is on record.
- **`algn="in"` is measured, not emulated.** The plan proposes clip-and-double-stroke
  and the measurement confirms an inset stroke sits wholly inside the geometry
  (within one pixel of antialiasing). Building the clip is 2.10.
- **The `arrow` marker's stroke width is described, not fitted.** It is a stroked V
  and its silhouette was measured; the exact stroke it is drawn with is not
  separated out.

## Open questions

1. **Is the 0.9 in the glow and soft-edge offset exactly nine tenths?** Five soft
   edge radii put it at 0.875, 0.9375, 0.906 and 0.902 and three glow radii at
   0.938 of the half-radius. Every one is within half a pixel of 0.9 at the
   resolution it was measured at, and none is within half a pixel of 1 — but 0.9
   is an odd constant and something rounder may be hiding behind it.
2. **What does `a:softEdge` do at a corner?** Only straight edges were measured.
3. **Does `@rotWithShape` on a glow or a soft edge mean anything?** Only the
   shadow's was measured, where it plainly does.
4. **What are the twenty `prstShdw` presets, as transforms?** Twenty boxes are
   recorded; the transforms behind them are not derived.
5. **Does a compound stroke's dash apply to the band or to each rail?** A probe
   exists (`cmpd-dbl-dash`) and its row is in the fixture; the analysis does not
   yet interpret it.
6. **Is the miter limit exactly 8, or is 8 a coincidence of this corner set?** It
   is bracketed between 7.91 and 8.01 on one family of corners.

Carried from 2.7: the off-centre path focus; `a:path` with no `@path`; whether
corner softening is a fixed 1.6px; whether a gradient slide background uses the
slide as its extent; whether `@flip` does anything at all.

## Verification

`pnpm check` green end to end: layering, corpus (82 entries), format, lint,
typecheck, build, round trip 52/52, pkg QA, and **68 test files / 1909 tests**
— `line.test.ts` contributing 66.

Twenty-four mutations, all killed. The probe decks were generated by this
repository and opened in the real PowerPoint 365 on this machine; nothing else on
the machine was read, and nothing was downloaded or installed.
