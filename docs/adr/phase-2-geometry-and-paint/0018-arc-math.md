# 0018 — Arc math

Date: 2026-09-02
Status: **accepted** — 393 arcs across 63 presets, resolved and checked against the presets' own answers
Sub-phase: 2.3 — arc math

---

## Context

`a:arcTo` is the one drawing command in DrawingML that is not what it looks like. It names four
things — two radii and two angles — and none of the three a path normally needs: no centre, no
destination point, and no statement that its angles mean what they appear to mean.

2.2 left them raw. This sub-phase resolves one arc against the point a path has already reached and
answers where it starts, where it ends, what ellipse it lies on and which way round it goes.

Nothing is drawn. Turning that into path data is 2.4.

---

## Decisions

### The unskew formula is taken from the presets, not from a specification

The plan states it: `t = atan2(wR·sin θ, hR·cos θ)`. That is correct, and the reason to believe it is
not that the plan says so. **Four presets state it themselves, in guide arithmetic.** `pie` computes
its own arc start point before drawing the arc to it:

```text
wt1 = sin wd2 stAng        ->  wd2 * sin(stAng)
ht1 = cos hd2 stAng        ->  hd2 * cos(stAng)
dx1 = cat2 wd2 ht1 wt1     ->  wd2 * cos(atan2(wt1, ht1))
dy1 = sat2 hd2 ht1 wt1     ->  hd2 * sin(atan2(wt1, ht1))
x1  = hc + dx1,  y1 = vc + dy1
```

which is `centre + (wR·cos t, hR·sin t)` with `t = atan2(wR·sin θ, hR·cos θ)`, spelled out. `chord`,
`arc` and `blockArc` do the same.

That matters more than a corroborated formula. It means the file **states the answer twice** — once
as guide arithmetic reaching a point, once as an `arcTo` that must reach the same point — so this
module can be checked against a number it had no hand in producing.

`stAng` is therefore the angle of a **ray from the centre**, not an ellipse parameter. The two agree
on a circle and at the four cardinal angles and nowhere else, which is exactly why the mistake
survives casual testing: `ellipse`, `donut`, `roundRect` and every rounded corner in the corpus look
perfect without the step.

### The arc is placed by its start point

`arcTo` has no centre, so the centre is whatever puts the point at `stAng` where the pen already is:

```text
centre = current − (wR·cos t₀, hR·sin t₀)
```

An arc therefore cannot be read on its own. Across the 187 presets there are 393 arcs and **not one
of them opens a path**, so the current point always exists — measured, and held by a test, because
an arc as a path's first command would be a case with no defined answer.

### A negative sweep is preserved, not normalised

81 of the 393 arcs carry a literal negative `swAng`. `donut` is why it matters: its outer ring is
four arcs of +90° and its inner ring four of −90°. The rings wind in opposite directions, and under
the nonzero fill rule that opposite winding **is** the hole. Take the absolute value, or rewrite −90
as +270, and the doughnut fills in solid — and the endpoint is identical either way, so nothing
downstream notices.

### The unskew is lifted, not wrapped

`atan2` alone folds every angle into a half-open turn, which throws away the winding an arc needs. So
a whole number of turns is removed before the `atan2` and added back after. The result is continuous
and strictly increasing in the angle across every turn boundary, which lets the parametric sweep be
read off as a plain subtraction — including for sweeps past half a turn, and including a sweep that
crosses the branch cut.

Without the lift, a sweep from 170° to 190° comes out as −340° instead of +20°: the right endpoint,
the wrong way round the ellipse, and a filled shape that is inside out.

### `swAng` is clamped to ±360°, not normalised into it

Clamping and normalising differ visibly: 400° clamped is 360° and ends where it began; normalised it
is 40° and ends somewhere else, taking every following command with it. Office clamps. No preset
reaches the case — the largest swing in the corpus is exactly one turn — so this follows the
documented behaviour and is pinned by a test rather than by evidence from a file. Flagged below as
wanting an experiment.

### A whole turn is split into two half turns

SVG's elliptical arc is defined by its two endpoints, so it cannot express a closed loop: an `A`
command whose destination equals its origin is not drawn as a circle, it is **not drawn at all**.
20 of the corpus's arcs are exactly that. `smileyFace` would lose its outline and both eyes; `sun`,
`actionButtonInformation`, `actionButtonHelp`, `cloudCallout`, `mathDivide` and `funnel`'s hole would
lose a piece each.

The split is at the antipode, giving two pieces of exactly half a turn, so `large-arc-flag` is false
on both and the ambiguity at exactly π does not arise. Anything short of a whole turn stays one
segment.

### Radii are used as absolute values

A negative radius describes the same curve as its positive twin — it mirrors the parameterisation,
not the ellipse — and SVG requires the absolute value, so taking it here keeps the two agreeing. No
preset produces a negative radius at any size, which a test holds to, so this is a robustness
decision rather than a corpus-driven one.

### A zero radius is degenerate, and degenerate means nothing to draw

`arcSegments` returns an empty list and the pen does not move. This is not an error case: a square
corner is written as an arc of radius zero rather than left out, so `round2DiagRect` and
`round2SameRect` reach it at their **default** adjust values, and drawing nothing there is the
correct answer. Non-finite operands are carried through the same way, consistent with 2.2's decision
to return IEEE results rather than throw.

---

## What building it found

### `actionButtonSound` has no arcs at all

The plan names eight presets as the ones that exercise every branch: `moon`, `blockArc`, `pie`,
`arc`, `chord`, `circularArrow`, `smileyFace`, `actionButtonSound`. The last of those is a speaker
icon drawn entirely with `lnTo` — **zero `arcTo` commands**. It cannot verify anything here.

Substituted `donut`, `ellipse` and `funnel`, each covering a branch none of the other seven reach:
opposite winding, the inscribed-ellipse identity, and the only reverse whole turn in the corpus.

### The presets state the answer more than once, in three different ways

- **`blockArc` states it four times.** Its outer arc runs `stAng → istAng` on the outer ellipse and
  must end on the guides `x3,y3`; the inner arc runs back and must end on `x4,y4`. Neither of those
  two points is used by the path at all — they exist only to bound the text rectangle — so they are
  an entirely separate calculation of where these arcs have to finish.
- **`circularArrow` round-trips `at2` through the unskew.** It builds its arc start point
  parametrically, as `(rw2·cos p, rh2·sin p)`, then takes that point's **ray** angle with `at2` and
  hands it to `arcTo` as `stAng`. The file is asking for the parameter to be recovered from the ray.
  The unskew is exactly that inverse; an implementation treating `stAng` as a parameter centres the
  arc somewhere else and the arrowhead detaches from the band. Also in `leftCircularArrow`,
  `leftRightCircularArrow` and the nine `bD` guides of `gear9`.
- **`funnel` says it a third way**, in polar form: `n1 = wd2·hd4 / sqrt(hd4²cos² + wd2²sin²)`, which
  is the polar equation of the same ellipse, evaluated along the ray. Different arithmetic, same
  point, same required centre.

### One arc in five is misplaced without the unskew

At 320×90 — an ordinary non-square size — **84 of the 389 non-degenerate arcs** sit at a different
parametric angle than the naive reading gives, across 16 of the 63 shapes. The rest are circular, or
start on a cardinal angle where the two agree.

That number is the case for the module. It also explains why the bug is durable: four fifths of the
corpus is unaffected, and the affected fifth is not obviously wrong so much as subtly reshaped.

### The tests were checked by breaking the module

Three mutations, each a plausible implementation:

| Mutation                          | Tests killed |
| --------------------------------- | -----------: |
| unskew replaced by the identity   |           32 |
| `swAng` passed through `Math.abs` |           14 |
| the whole-turn split removed      |            5 |

The first is the informative one. It kills tests across nine shapes through four separate oracle
families — and `ellipse`, `donut` and `smileyFace` still pass, which is the trap stated as a
measurement rather than as a warning.

---

## Consequences

`pnpm check` green.

- `arcGeometry` resolves one arc against a point; `arcSegments` breaks it into pieces SVG can draw;
  `unskewAngle`, `clampSweep`, `arcEnd` and `arcPointAt` are the parts worth using on their own.
- **`packages/geometry` is 214 tests**, up from 95; the repository is 63 files and 1689 tests, up
  from 62 and 1570.
- The corpus-wide checks run every one of the 393 arcs at four sizes: every start lands exactly on
  the point the pen was already at, every sampled point lies on its own ellipse to within 1e-9, and
  nothing comes back non-finite.

---

## What is not done

- **Nothing is drawn.** `arcSegments` returns geometry, not path data. Formatting it into a `d`
  string is 2.4, and so is the path walker that supplies each arc's current point — the one in
  `arc.test.ts` is test scaffolding and is deliberately not exported.
- **No arc has been compared against PowerPoint.** Every check here is against the presets' own
  arithmetic or against a fact about ellipses. That is a strong internal oracle and it is not an
  external one.
- **Arc bounding boxes.** An axis-aligned ellipse arc's extrema are its endpoints plus whichever
  cardinal parameters fall inside the range. Cheap, and 2.4 or 2.5 is where something needs them.

### Open questions, each wanting an experiment in the manner of 0.7

- **Whether Office really clamps `swAng` rather than normalising it.** Test: a `custGeom` with
  `swAng="24000000"` (400°) followed by a `lnTo`. If the line starts where the arc began, it clamps.
  Nothing in the preset corpus can answer this.
- **What Office does with a zero radius mid-path** — treat the arc as absent, as here, or as a line
  to the computed endpoint. The two agree for the preset cases because the endpoint is the start.
- **Whether a negative radius is ever accepted** and, if so, whether it mirrors. Unreachable from the
  presets; only hand-authored `custGeom` could produce one.
