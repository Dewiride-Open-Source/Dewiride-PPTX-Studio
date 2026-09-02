# 0017 — The formula evaluator

Date: 2026-09-02
Status: **accepted** — 17 operators, 42 built-in guides, every preset evaluated at six sizes
Sub-phase: 2.2 — formula evaluator

---

## Context

2.1 transcoded the 187 preset definitions and deliberately evaluated none of them. Every operand is
still the string the source wrote, because a formula's operands are guide names, built-in names and
literals mixed together and nothing can separate them without a shape size and an `avLst`.

This sub-phase supplies both and runs the program: the seventeen `a:gd/@fmla` operators, the guides
every shape has without declaring them, and the merge of a deck's adjust values over a preset's
defaults. What comes out is a map of names to numbers.

Nothing is drawn. The arithmetic stops one step short of geometry, which is 2.3 and 2.4.

---

## Decisions

### The operator table was derived from the data, not recalled

An operator with its operands transposed does not throw. It draws a shape that is quietly the wrong
shape, and the failure surfaces weeks later as "the circular arrows look odd". So the table was
established against the 187 presets, which are a corpus of 3922 worked examples, rather than written
from memory and spot-checked.

Two scans over the committed buckets did most of the work — deliberately as a _second_ reader of the
encoding, not by importing the shipped decoder, so a bug in that decoder could not make the scan
agree with it.

`blockArc` turned out to pin six operators at once:

```text
stAng  pin 0 adj1 21599999     the adjust is what is clamped, so pin is (min, value, max)
sw12   +- sw11 21600000 0      21600000 is one full turn
iswAng +- 0 0 swAng            the negation idiom, so the third operand is subtracted
swAng  ?: sw11 sw11 sw12       the ternary, used to normalise a sweep angle
wt1    sin wd2 stAng           the multiply by the first operand is real
dx1    cat2 wd2 ht1 wt1        the ellipse unskew
```

The last of those is the one that could not have been guessed. With the operand order below, `dx1`
and `dy1` come out as `(w/2)·cos u` and `(h/2)·sin u` for `u = atan2(wt1, ht1)` — which is exactly
the unskew this plan's 2.3 entry names, `t = atan2(wR·sin θ, hR·cos θ)`. Two documents written
independently agreeing on the same formula.

### `mod` is the vector modulus, and arity is what settles it

Not by reading the name, which is what makes this the classic error. **All 48 uses take three
operands**, and a binary remainder cannot. The names then confirm it: `cornerTabs.md = mod w h 0`
and `mathMultiply.dl = mod w h 0` are the shape's diagonal, `gear6.lAD1 = mod xAD1 yAD1 0` is a
segment length.

The third operand is the literal zero in all 48, so every preset use degenerates to a plain
hypotenuse. It is still computed in three dimensions, because custom geometry in a user's deck is
under no such restriction.

### The built-in guides are derived from the naming rule, not from a copied list

`wdN` is the width over N, `cdN` a full turn over N, `McdN` is M of those. Applying the rule to a
set of divisors is shorter than an enumeration and harder to get individually wrong.

It is also the only version that survives contact with the data. The presets reference `wd3`,
`wd12`, `hd3`, `hd10`, `ssd16` and `cd3`, and published enumerations of "the built-in guides" — the
plan's own entry says 44 of them — do not all include those. An enumeration transcribed from one of
them would be missing a guide that `rtTriangle` and `curvedLeftArrow` need.

`builtins.test.ts` closes the loop from the other side. It holds the list of every name the 187
presets actually reference, read off the committed buckets, and asserts the table covers all of it —
and it recomputes that list from the buckets in the same test, so the literal cannot go stale. Two
sources; a disagreement is a finding.

### Adjust values merge by name

Not by position and not by index. **Eleven distinct adjust names** appear across the presets: `adj`,
`adj1` through `adj8`, and `hf` and `vf` on the polygons and stars. A positional merge would quietly
assign a deck's first adjust to `hf` on a pentagon.

An override naming an adjust the preset does not declare is still evaluated and kept — a deck that
writes one is saying something, and dropping it silently is worse than carrying a value nothing
reads. The exception is a name colliding with a built-in, where the built-in wins: letting a deck
redefine `w` would turn the shape's own width into a suggestion.

### One forward pass, and that is measured rather than assumed

Built-ins, then `avLst`, then `gdLst` in document order, each formula seeing only what was defined
before it.

Across all 3622 computed guides there are **zero forward references and zero self references**. 2.1
asserted the ordering in a doc comment; this sub-phase measured it, and then made the measurement
permanent: an operand naming a later guide throws `FMLA_OPERAND`, and the whole-corpus sweep would
fail rather than silently building a shape on an undefined guide.

### Division by zero returns the IEEE result and does not throw

This is the decision that needed a number rather than an opinion, so one was taken: **212 of the
1204 division formulas divide by a guide rather than by a literal**, 36 shapes divide by `ss`, nine
by `w`, seven by `h`.

A shape with zero height is completely ordinary — a flat line, a shape being dragged through zero —
so `x / 0` is reachable at legal sizes rather than being a corrupt-file case. Throwing would mean a
shape that vanishes from a slide mid-drag and takes an error path with it.

So the IEEE result is kept and `nonFiniteGuides` lets a caller ask. What a renderer should do about
it is 2.10's decision, made with a renderer in front of it. The reachability is pinned as a number —
49 presets affected at zero width, 49 at zero height, 55 when both collapse — so a change to the
arithmetic shows up as a number moving.

### Everything is a double, and nothing is truncated

`circularArrow` chains more than 200 guides deep. Rounding each intermediate to an integer would
accumulate visible error by the end of it. Integers appear in the source data, not in the
arithmetic.

### The four-operand `+-` is accepted only when the fourth operand is zero

2.1 reported eight formulas giving the add-subtract operator four operands and carried them through
without repair. 2.2 can say something stronger than "carried through": they are **provably** benign.

All eight have the same shape — `circularArrow.xB = +- xH 0 dxB 0` — and the fourth operand is the
literal `0` every time. `x + y − z` ignoring a trailing zero, `x + y − z + w` and `x + y − z − w`
all give the same number when `w` is zero. Every reading agrees, so there is nothing to get wrong.

A non-zero fourth operand is a different case, one where the readings genuinely diverge, and the
honest response there is to stop. Accepting the benign case while refusing the ambiguous one is the
whole reason not to simply widen the arity to four.

### `sqrt` follows Office, not ECMA, and takes the absolute value first

The one place this evaluator knowingly departs from the standard.

Microsoft's implementer notes for ECMA-376 (MS-OI29500) record four deviations in this section, and
this is the consequential one: the standard defines `sqrt x` as the square root of x, and **Office
computes the square root of its absolute value**. Apache POI implements the standard and returns
`NaN`; LibreOffice ends at zero.

Following the standard would be the worse failure. `NaN` is contagious — every guide downstream
becomes `NaN`, and a shape whose points are `NaN` does not draw a wrong shape, it silently draws
nothing at all. A shape that vanishes with no error is the hardest class of bug this project has.

Whether that is reachable from a legal adjust value was measured rather than argued, by sweeping
every sqrt-using preset across its adjust range and checking the sign of each radicand. **It is
reachable**: five guides in three shapes go negative — `curvedDownArrow.q5` and `.q11`,
`curvedUpArrow.q5` and `.q11`, and `leftCircularArrow.u9`. With ECMA's plain `sqrt` those three
shapes draw nothing at ordinary adjust values.

### `at2` guards the origin explicitly

`Math.atan2(0, -0)` is π, not zero. A guide chain reaches a negative zero trivially — `*/ 0 1 -1` is
enough — and without the guard a zero-length vector comes back as a half turn, which draws a
plausible shape pointing the wrong way. Written out rather than left to IEEE.

### `pin` is an ordered chain, and that is not `max(x, min(y, z))`

The two agree whenever the bounds are the right way round and diverge when they are inverted:
`pin 10 20 3` gives 3 as a chain and 10 as a min/max composition. Apache POI rewrote this operator
into the min/max form in 5.2.5, so the divergence is live in a shipped implementation rather than
theoretical. ECMA describes the chain, so the chain is what this is. No preset inverts its bounds;
hand-authored custom geometry can.

### How the research was run, and what it changed

The operator table was derived from the corpus first, then checked against a fan-out of six
independent derivations — ECMA prose and prior-art implementations for both the operators and the
guides, plus one pass each on angle conventions and numeric edge cases — reconciled, and the top
three disagreements put through adversarial verification. Sixteen disputed points came back.

It corroborated everything the corpus had already settled: `mod`, `pin`'s operand order, `cat2` and
`sat2`, the strict `>` in the ternary, `at2` being un-normalised, and the trig unit. It changed
three things, all recorded above — `sqrt`, the `at2` origin guard, and the `pin` divergence — of
which only `sqrt` alters what gets drawn.

It also produced a caution worth keeping: **no derivation read the ECMA PDF itself.** All spec text
came from mirrors, and two mirrors visibly disagree on minor details. One derivation reported a
claim about adjust clamping that traced to a search summariser's inference rather than to any
primary source, and was caught only because the other five did not repeat it. Nothing in this ADR
that changes behaviour rests on that class of evidence.

---

## What building it found

### `at2` is not normalised, and three guides depend on that

`circularArrow` contains the same idiom three times:

```text
ist0   at2 sdxC sdyC
ist1   +- ist0 21600000 0
istAng ?: ist0 ist0 ist1
```

That is the caller wrapping a possibly-negative angle into a positive one by hand. It is dead code
if `at2` normalises — so the operator returns the raw `arctan(y/x)` range, and normalising it here
would silently disable the idiom and reverse the sweep direction of every circular arrow.

This was the highest-risk open question going in, and the preset data answered it without needing
either the specification or PowerPoint.

### The presets need 34 built-in guides, not 44

Every operand in every position across all 187 presets resolves to a literal, a declared guide, or
one of exactly 34 names. The closure is complete: **zero** operands reference anything else.

The table ships 42, because the missing eight are the rest of the families the rule generates — the
eighths of a circle, `ls`, a few divisors — and custom geometry in a user's deck may name them even
though no preset does.

### The `±2147483647` sentinels never reach the arithmetic

184 occurrences, and every one is an adjust-handle minimum or maximum. None appears in a formula, a
connection site, a text rectangle or a path command. So they are "unbounded" markers for the handle
UI in 2.5 and there is no integer-overflow question to answer here.

### `rect` is an oracle

`rect` has no guides at all, and its four connection sites must be the midpoints of its edges at
angles 270°, 180°, 90° and 0°. That is a fact about rectangles, not about this evaluator, and it
also fixes the angle convention: **zero points right, and because y increases downward a positive
angle turns clockwise**, so `cd4` points at the bottom of the shape. `triangle` says the same thing
independently — its apex site is `3cd4` and its base corners are `cd4`.

---

## Consequences

`pnpm check` green.

- **17 operators**, arity-checked; **42 built-in guides**; adjust values merged by name.
- **`pnpm test` is 62 files, 1570 tests**, up from 59 and 1495.
- Three independent oracles, none of which is derived from the evaluator:
  - `triangle` at 200×100 computed by hand — guides, path corners, text rectangle, and the apex
    sitting exactly at `hc`, which is what makes it isosceles.
  - `rect`'s connection sites as the four edge midpoints.
  - **`blockArc`'s computed points lying on the ellipse they are an arc of**, checked across three
    sizes × six start angles × three inset values. This holds for any angle rather than at a spot,
    and it is what would catch `cat2` and `sat2` with their operands transposed — that mistake moves
    the start of the arc from 3 o'clock to 6 o'clock and the invariant survives it, but the two
    explicit clock-position tests do not.
  - `hexagon.dy1` coming out as exactly half the height, which pins the angle **unit**: the identity
    only holds if `3600000` is read as 60 degrees.
- All 187 presets evaluate at six sizes with every guide finite, and without throwing — which is
  simultaneously the proof that a single forward pass suffices.

---

## What is not done

- **Nothing is drawn.** No arc is unskewed and no path is emitted. `arcTo`'s angles are still raw.
- **No preset has been compared against PowerPoint.** Internal consistency is established; agreement
  with the renderer everyone else has is not, and cannot be until 2.4 emits a path. The cheapest
  external oracle available then is PowerPoint's own SVG export of a shape, which is its evaluation
  of the same preset at a known size.
- **The ternary's boundary at exactly zero is undetermined by the data.** Every use in the corpus is
  an angle normalisation of the form `?: a a (a + 21600000)`, and at `a == 0` both branches name the
  same angle, so the presets cannot separate `> 0` from `>= 0`. Implemented as strict and pinned by
  a test, so a future change to it is deliberate rather than incidental.
- **`ahLst` ranges are not enforced during evaluation.** An adjust outside its handle's declared
  range is evaluated as given; the presets clamp what they care about themselves, with `pin`. Whether
  the range is also a hard constraint is 2.5's question, where there is a handle to drag.
- **Whether an unlisted built-in should resolve by rule.** A `custGeom` naming `wd7` throws today.
  Widening that is a decision to take when a real deck demands it, rather than pre-emptively.

### Open questions the research raised and could not settle

Each needs an experiment against the PowerPoint on this machine, in the manner of 0.7. None blocks
2.3.

- **What `w` and `h` are bound to inside a path that declares `a:path/@w` and `@h`** — the shape
  extent, or the path space. Implementations differ and the spec does not say. This evaluator takes
  the size from its caller, so the question belongs to 2.4, which is where a path space first
  exists; the factoring is deliberate rather than an evasion.
- **Divide by zero.** POI returned IEEE infinity before 5.2.5 and returns 0 after it. This keeps the
  IEEE result, which is the reversible choice — a caller can still substitute zero, where collapsing
  to zero here would destroy the information. What PowerPoint does is untested.
- **Whether an adjust outside its `ahLst` range is clamped by PowerPoint.** Test: one `roundRect`
  with `adj` at 90000, well past the nominal maximum, beside one at 50000. Identical rendering means
  it clamps.
- **Whether a deck may shadow a built-in** by declaring a `gd` named `ss` or `hc`. This lets the
  built-in win; POI's lazy fallback would let the deck win.
- **`tan` at the 90° pole**, reachable only from hand-authored custom geometry.
