# 0020 — Adjust handles, and the inverse of a shape

Date: 2026-09-02
Status: **accepted** — 243 handles across 120 presets, every axis of every one of them dragged
through its full range at five sizes and asked to land where it was sent
Sub-phase: 2.5 — adjust handles

---

## Context

2.1 read the preset definitions, 2.2 turned them into numbers, 2.3 resolved the arcs and 2.4
emitted the paths. All four run one way: a definition and a size produce a drawing.

A handle runs the other way. The user grabs a yellow dot, moves it, and that point has to become a
new value for one of the shape's adjust guides — written into `a:prstGeom/a:avLst` and read back by
PowerPoint. Nothing before this sub-phase inverts anything.

The plan called for "analytic inversion of the dominant `*/ w adjN 100000` forms, bisection
fallback for everything else". What was built is one method rather than two, for reasons the plan
could not have known and this record has to state as a change rather than as the plan executed.

---

## Decisions

### An axis is a scalar readout of the handle's position

This is the decision everything else hangs off, and it is what lets one routine serve all 287
adjustable axes.

`a:ahXY` names its axes by coordinate — `gdRefX` moves the handle's x, `gdRefY` its y. `a:ahPolar`
names them by radius and angle, and gives no origin to measure either from. So there are four
scalar readouts of one position — `x`, `y`, `r`, `ang` — and an axis is one readout paired with one
adjust guide. Dragging is solving `readout(v) = target` for `v`, and the four cases differ only in
which scalar is read.

The readout matters more than the solver. Read a polar handle's angle axis as its `pos.y` instead
of as an angle and the numbers collapse: 58 of the classifications turn non-monotone rather than 5,
and `mathNotEqual`'s angle handle becomes literally constant, because its `pos.y` is the built-in
`t` and never moves at all. Twelve shapes stop being invertible. Choosing the right scalar is the
whole of the problem; the root-finder is the easy half.

### The polar origin is the shape centre, and the presets say so

ECMA gives `a:ahPolar` no origin — the element carries `gdRefR`, `minR`, `maxR`, `gdRefAng`,
`minAng`, `maxAng` and a Cartesian `a:pos`, and never says what the radius and angle are measured
from. VML's predecessor element had an explicit polar centre; DrawingML dropped it.

The preset data answers it anyway, because ten of the eighteen angle axes state where their handle
goes twice over. Sweep `arc`, `blockArc`, `chord`, `pie`, `circularArrow` or `leftCircularArrow`
through 49 angles at six aspect ratios — 200x100, 100x100, 100x400, 914400x685800, 40x900 and 3x1 —
and compute `atan2(pos.y - h/2, pos.x - w/2)`. It reproduces the guide named by `gdRefAng` to within
**2.2e-8 of a 60000th of a degree**, which is one unit in the last place of the arctangent and not a
modelling error.

At 200x100 the locus is an ellipse: the handle's distance from the centre runs from 50 to 100 across
the sweep. The angle still lands exactly. That is 2.3's finding restated from the other side — a
DrawingML angle names a ray out of the centre, not a parameter of an ellipse — and it is what makes
the identity evidence about the _origin_ rather than about a circle.

Both open implementations that support dragging take a polar angle straight from an arctangent about
the shape centre, which is the same answer reached from the implementer's side.

### The readout is always inverted, and never read off the pointer

The ray identity is not a licence to skip the inverse. The other eight angle axes are not on their
own ray: `circularArrow`'s innermost handle is out by up to 156 degrees, its stub handle by a
constant 19.04, and `mathNotEqual`'s by 171. Setting the guide to the pointer's angle would be
exactly right for ten axes and visibly wrong for eight.

Inverting is right for all eighteen and costs one extra step of arithmetic on the ten. ONLYOFFICE
carries a shape-specific branch for `mathNotEqual` that corrects it by half a turn, which is the
same anomaly seen from the other end — a special case where a general method needs none.

### One method, because false position _is_ the analytic answer

Sample the readout at 33 points across the axis's range, take the sub-interval that straddles the
target, and refine with false position plus the Illinois correction.

The map from an adjust value to a readout, composed through the whole guide chain, is **affine at
every size on 243 of the 287 axes** and monotone on 282. That is the shape the plan saw. But it does
not need two methods and it must not look at the formula text:

- Affinity is a property of the whole composed chain, not of any one `fmla`. Pattern-matching the
  multiply-divide form would have found some of the 243 and missed the rest — and would have had to
  be kept in step with `applyOperator` for ever.
- On an affine map the first false-position step lands on the root **exactly**. The analytic case is
  therefore recovered as a special case of the general one rather than as a second code path.

The prior art corroborates negatively. LibreOffice hand-derives an inverse per preset, dispatching
on shape name, with Newton's method for four ellipse-inset handles and a catch-all that assumes the
guide is the coordinate over the width or height times 100000 — which is wrong for the `ss`-scaled
and 200000-scaled forms, and its own notes say the approach cannot work for a custom geometry.
ONLYOFFICE evaluates at the guide's minimum and maximum, fits one linear coefficient, and never
checks the result — exact for the affine majority and silently wrong for the other 44. Apache POI
models a handle as an empty marker interface. python-pptx exposes values with no clamping and no
knowledge that `ahLst` exists.

### A scan, not an endpoint bracket, and ties break towards where the shape already is

Five of the 1435 axis-by-size combinations genuinely fold back. `curvedLeftArrow` and
`curvedRightArrow` at 100x100 run their handle down 0.615 units a step for thirty-nine steps and
back up 0.179 on the fortieth; `circularArrow`, `leftCircularArrow` and `leftRightCircularArrow`
swing their innermost handle 147 degrees backwards once at 40x900, when it crosses the centre of a
shape 40 units wide.

A fold means more than one adjust value puts the handle under the pointer. Bracketing the endpoints
misses a root inside a fold; taking the first root found teleports the handle to the far branch when
the pointer crosses the turn. Nearest-to-current is the only rule under which a small movement of
the mouse makes a small change to the written value, which is what "the handle follows the pointer"
means.

The folds are **inside** the declared range, so clamping does not make the inverse single-valued.

### The two axes of a handle are solved one after another, angle first

For an `a:ahXY` the order changes nothing, and that was measured rather than assumed: across all 220
of them, sweeping the guide named by `gdRefY` through its whole range never moves `pos.x`.

The four two-axis polar handles are not decoupled. `blockArc`'s handle sits at `iwd2*cos(u)`,
`ihd2*sin(u)` for the unskewed angle `u`, so the ray out of the centre is exactly the angle guide and
does not depend on the radius guide at all — while the radius depends on both. Solving them against
the same pre-drag state puts the handle **21% of the shape** away from the pointer at 200x100 and
14.8% at 100x400. Solving the angle first and the radius against the answer lands it within 0.0002%.

On a square shape `iwd2` equals `ihd2` and the whole error vanishes, which is why a square-only test
would never have found it. It was not found by testing; it came out of the research sweep, and the
measurement above was taken to check the claim.

### `minX` and `maxX` bound the guide, in the guide's own units

This knowingly diverges from the ISO prose, which describes them as positions — "the maximum
horizontal position that is allowed for this adjustment handle". Microsoft's own preset data cannot
be read that way. `arc` declares `maxAng="21599999"`, one unit short of a full turn in 60000ths of a
degree and not a coordinate at any size. `donut` declares `maxR="50000"` against its own
`pin 0 adj 50000`, while the handle's distance from the centre _decreases_ as the guide grows.

So the divergence is not from the spec to us; it is between Microsoft's shipped preset data and
Microsoft's own spec prose, and we follow the data.

92 of the 287 axes take at least one bound from a guide rather than a literal, so the bounds are
resolved at the shape's size like everything else.

### An absent bound freezes the axis

ISO says it per attribute: if `maxX` is omitted "it is assumed that this adjust handle cannot move in
the x direction. That is the maxX and minX are equal." It never says equal to what, so the value the
shape holds is the only answer available.

This is the opposite of the obvious reading and unreachable from the presets — all 243 records state
both bounds or declare no `gdRef` at all. It is reachable from a `custGeom`, where freeing an axis
the author froze would let a drag write a value the file says cannot change.

A bound of `2147483647` written out in full is a different thing: a stated bound that happens to be
the widest `ST_Coordinate` there is, which is how the four callout families say "anywhere". Honoured
as stated.

### A declared range that excludes the shape's own value is widened to admit it

A range that excludes the value the shape actually holds is not a range for that value, so the
smallest correction is to admit it. 16 of the 1435 axis-and-size combinations, in two different ways.

`mathDivide` ships `adj2="5880"` while its own `maxAdj2` works out to 2930, and `leftRightRibbon`
ships `adj2="50000"` against a `maxAdj2` of 46875. Both already clamp themselves with `pin`, so the
shape draws identically either way. Widening is what stops merely touching a handle from rewriting a
deck's own number to a different one that draws the same — which in a package whose first principle
is preserving what it did not edit is the whole point.

### The value written is an integer; the value searched is not

`ST_AdjCoordinate`'s lexical space is a guide name or an integer, and `a:prstGeom`'s only child is
`a:avLst`, so what lands in the file is `<a:gd name="adj2" fmla="val 2930"/>`. Rounding happens once,
after the clamp, and is pulled back inside the bounds afterwards because a bound may be fractional —
`trapezoid`'s `maxAdj` is `50000*w/ss`, or 66666.67 on a 914400x685800 shape.

The search must not be integral, and making it so is what surfaced the evaluator bug below.

---

## What building it found

### The evaluator could not be asked for a fractional adjust value

`evaluateGuides`'s ergonomic override form — a plain `Record<string, number>` — serialised each
number with `String()` into a `val` formula and read it back through `ST_AdjCoordinate`'s **integer**
grammar. So `{ adj: 16667.5 }` threw `FMLA_OPERAND`, and so did `{ adj: 1e21 }`, because `String()`
gives `"1e+21"`.

Latent since 2.2 and reachable from its own public API. 2.5 made it unavoidable, because the search
evaluates the geometry at fractional values on the way to the answer. A number now stays a number:
the record form sets the value directly instead of round-tripping through a grammar that is a fact
about the markup rather than about the numbers a caller holds.

### `star24` and `star32` state their handle bound in the wrong units

Five of the seven stars bound `adj` at `50000` — fifty per cent, in the adjust's own units, the same
number their own `pin 0 adj 50000` uses. `star24` and `star32` write `maxY="ssd2"`, half the shorter
side, which is a coordinate.

The consequence is unit-dependent, which is the part worth recording: `ssd2` scales with the shape,
so on a 100-unit shape the declared bound is 50 against a guide running to 50000 — a two-thousandth
of the real travel — and on the same shape in EMU it is 342900, nearly seven times the real travel,
most of it a flat region above the shape's own clamp. Widening keeps the handle moving in both, and
differently in each. A package that never decides the caller's unit should not have a handle whose
behaviour depends on it, so this is an experiment rather than a fix.

### A committed preset does invert its `pin` bounds, at a legal size

2.2 recorded that `pin` is an ordered chain and not `max(x, min(y, z))`, that the two differ when the
bounds are inverted, and that Apache POI rewrote the operator into the min/max form — but also that
"nothing in the preset data inverts its bounds — every use is of the form `pin 0 adjN <constant>`".

All three parts of that last sentence are wrong. Across the 196 uses in the 187 presets the first
operand is non-zero 20 times, the second is not an adjust value twice, and the third is a computed
guide 84 times. And the bounds do cross: `mathDivide.a3` is `pin 1000 adj3 maxAdj3`, and on a shape
one unit wide and a thousand tall `maxAdj3` works out to 36.745 against a floor of 1000. The chain
returns 36.745 and POI's rewrite would return 1000.

One case in 1568 evaluations at eight aspect ratios, found only because 2.5 had to know whether a
search range could invert. `formula.ts` and `formula.test.ts` now carry it.

### Two more statements in `evaluate.ts` were narrower than the schema

`ST_AdjCoordinate` is a union with `ST_Coordinate`, which admits a universal measure like `2in` as
well as a plain integer, so "a guide name or an integer, and nothing else" describes what this
module reads rather than what the markup allows. And `ST_GeomGuideName` is reported to be an
unrestricted token, so a guide literally named `50000` is expressible — which makes "a name in scope
wins" a decision rather than a consequence of the two spaces being disjoint.

Both are now stated as scoped facts with the gap named. No one on this project has read the schema
itself; the research that raised these reached only Microsoft's reproduction of ISO 29500-1's first
edition, and that limit is carried into the open questions rather than papered over.

### The bug the tests did not find

The polar ordering defect above was live, passing 34 committed tests, and would have shipped. The
round-trip test varies one axis at a time, which is exactly the shape of check that cannot see a
coupling between two — and both of the sizes where it would have shown are non-square, while the
obvious test size is not.

Recorded because it is the second time in three sub-phases that a defect has hidden behind a square
shape. 2.4 found that only `cloud` could separate the two arc orders because every other scaled arc
sits at a cardinal angle; this one hides for the same reason.

### The tests were checked by breaking the module

| Mutation                                                     | Tests killed |
| ------------------------------------------------------------ | -----------: |
| the polar origin moved from the shape centre to the top-left |           10 |
| an angle axis read as `pos.y` rather than as an angle        |            8 |
| false position keeps the wrong end of the bracket            |            3 |
| a value that already satisfies the drag is solved for anyway |            3 |
| the fold tie-break replaced by first-bracket-found           |            1 |
| an absent bound read as unbounded rather than frozen         |            1 |
| the two polar axes solved against the pre-drag state         |            1 |
| an inert axis solved rather than left where it is            |            1 |

The third row is not hypothetical: keeping the wrong end of the bracket is the bug this module
shipped with for its first hour. False position with the ends confused still converges on the
straddling interval it was handed and still looks plausible in the debugger — it destroys the
bracket two probes later and then returns the midpoint of an interval the root has left. The three
tests that catch it are the ones that assert a specific answer rather than a small residual.

The last four rows kill one test each because one is all that is needed: each guards a decision only
one situation can distinguish, and the test is that situation.

---

## Consequences

`pnpm check` green.

- `resolveHandles` answers what a selection UI asks: where each handle is drawn, which adjust value
  each axis writes, the range it may take and what it holds now. `dragHandle` answers what an edit
  asks: given a point, the adjust values to write. `invertAxis`, `readHandleAxis`, `handleCentre` and
  `roundAdjust` are the parts worth using alone.
- Deliberately **not** part of `resolveGeometry`. A slide draws every shape and edits one, so the
  handles of the shapes nobody selected are work nobody asked for.
- **`packages/geometry` is 294 tests**, up from 254.

---

## What is not done

- **No handle has been compared against PowerPoint.** Every check here is against the presets' own
  arithmetic, as in 2.3 and 2.4. A strong internal oracle; not an external one — and unlike the two
  before it, this sub-phase decides behaviour a user will feel directly.
- **Nothing is painted, and nothing is dragged.** This produces numbers; the pointer, the hit target
  and the preview are 2.11 and Phase 5.
- **Bounding boxes.** Carried from 2.3 and 2.4, still outstanding.

### Open questions, each wanting an experiment in the manner of 0.7

- **Does Office enforce `ahLst` bounds at all?** Author a `custGeom` whose handle declares
  `maxX="10000"` while its `gdLst` clamps the same adjust at 50000, drag the handle to the far edge
  and read back the `avLst`. This settles `star24` and `star32`, and it settles whether widening is
  right or merely harmless.
- **What is the polar origin, to Office?** Author a `custGeom` with an `a:ahPolar` whose `pos` is
  nowhere near the centre, drag it, and read back the angle. The presets agree with the centre on ten
  of eighteen axes; nothing establishes what Office does with the other eight.
- **Does Office write a value outside `[minX, maxX]`?** Same experiment, reading the file rather than
  the screen. If it clamps, our clamp matches; if it does not, preservation may argue for keeping
  the deck's number.
- **Is a universal measure legal in an `a:gd` operand?** Author a `custGeom` with `minX="2in"` and see
  whether PowerPoint opens it. Two sub-questions ride on it: whether `ST_AdjCoordinate` really admits
  one, and whether Office implements it.
- Carried from 2.4: a path declaring `@w` but not `@h`, and a guide referenced inside a scaled path.
- Carried from 2.3: whether Office clamps `swAng` rather than normalising it, what it does with a
  zero radius mid-path, and whether a negative radius is ever accepted.
