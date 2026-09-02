# `@pptx-studio/geometry`

DrawingML geometry: the **187 preset shape definitions** (2.1), the **formula evaluator** that turns
one of them plus a size into numbers (2.2), the **arc math** that says where each of the 393
`a:arcTo` commands in them actually goes (2.3), and the **path emitter** that turns all of it into
SVG path data (2.4). `a:prstGeom` and `a:custGeom` go through the same code, because they are the
same thing written two ways.

Nothing is _painted_ — a `d` string is not a fill, a stroke or an effect. That is 2.6 to 2.8.

```ts
import { getPreset } from '@pptx-studio/geometry';

getPreset('triangle');
// {
//   name: 'triangle',
//   avLst: [{ name: 'adj', fmla: ['val', '50000'] }],
//   gdLst: [{ name: 'a', fmla: ['pin', '0', 'adj', '100000'] }, ...],
//   ahLst: [{ kind: 'xy', gdRefX: 'adj', minX: '0', maxX: '100000', gdRefY: null, ... }],
//   cxnLst: [{ ang: '3cd4', pos: { x: 'x2', y: 't' } }, ...],
//   rect: { l: 'x1', t: 'vc', r: 'x3', b: 'b' },
//   pathLst: [{ w: 0, h: 0, fill: 'norm', stroke: true, extrusionOk: true, commands: [...] }],
// }
```

## Evaluating one

```ts
import { evaluateGuides, getPreset, resolvePoint } from '@pptx-studio/geometry';

const triangle = getPreset('triangle')!;
const guides = evaluateGuides(triangle, { w: 200, h: 100 }, { adjust: { adj: 50000 } });

guides.get('x2'); // 100 — the apex, at the horizontal centre
resolvePoint(triangle.cxnLst[1].pos, guides); // { x: 50, y: 50 }
```

`evaluateGuides` seeds the built-in guides from the size, merges the deck's adjust values over the
preset's defaults **by name**, then runs the `gdLst` in document order. What comes back is one map
holding all three kinds, because an operand elsewhere in the shape may name any of them and cannot
tell them apart.

A few things that are easy to get wrong, and are pinned by tests rather than by hope:

- **`mod` is the vector modulus** `√(x²+y²+z²)`, not a remainder. Arity settles it — all 48 uses
  pass three operands — and the names confirm it: `cornerTabs.md = mod w h 0` is the diagonal.
- **`pin` clamps its middle operand**: `pin min value max`.
- **`at2` returns `arctan(y/x)` and is not normalised.** It can be negative, and `circularArrow`
  normalises it by hand three times over. Wrapping it here silently reverses those arcs.
- **Angles are 60000ths of a degree**, a full turn is `21600000`, and because y increases downward a
  positive angle turns **clockwise**: `cd4` (90°) points at the bottom of the shape.
- **`+-` accepts a fourth operand only when it is exactly zero** — see the anomalies below.
- **`sqrt` takes the absolute value first** — `sqrt(|x|)`, which is what Office does and not what
  ECMA says. Apache POI implements the standard and returns `NaN`, which is contagious: the shape
  then draws nothing rather than drawing wrong. It is reachable — five guides across
  `curvedDownArrow`, `curvedUpArrow` and `leftCircularArrow` take a negative radicand at ordinary
  adjust values.
- **`pin` is an ordered chain, not `max(x, min(y, z))`.** They differ on inverted bounds:
  `pin 10 20 3` is 3, not 10.

### Adjust values merge by name

Not by position. Eleven distinct adjust names appear across the presets — `adj`, `adj1`…`adj8`, and
`hf`/`vf` on the polygons and stars — so a positional merge would assign a deck's first adjust to
`hf` on a pentagon. An override for a name the preset does not declare is kept; one that collides
with a built-in is ignored, because letting a deck redefine `w` would make the shape's own width a
suggestion.

### Division by zero is reachable, and does not throw

212 of the 1204 division formulas divide by a _guide_ rather than a literal, and 36 shapes divide by
`ss`. A shape with zero height is ordinary — a flat line, or a shape mid-drag — so `x / 0` happens
at legal sizes. Guides keep the IEEE result and `nonFiniteGuides()` reports which ones went
non-finite; 49 of the 187 presets are affected at zero width, 55 when both dimensions collapse.
Throwing there would mean a shape that vanishes mid-drag and takes an error path with it. What a
renderer should do about it is 2.10's decision, taken with a renderer in front of it.

## Arcs

`a:arcTo` names two radii and two angles, and none of the three things a path command normally
carries: no centre, no destination, and no hint that its angles are not what they look like.

```ts
import { arcGeometry, arcSegments } from '@pptx-studio/geometry';

// blockArc's outer arc, at 200 x 120 with the default adjusts.
const arc = { wR: 100, hR: 60, stAng: 10800000, swAng: 21600000 };
const g = arcGeometry({ x: 0, y: 60 }, arc);
g.centre; // { x: 100, y: 60 } - derived from the start point, not given
g.end; //    { x: 0, y: 60 }   - a whole turn, so back where it started
arcSegments({ x: 0, y: 60 }, arc).length; // 2, because SVG cannot draw a closed arc
```

Three things to know before using it:

- **The angles are rays, not ellipse parameters.** `stAng` is the angle of a ray from the centre;
  the parameter that reaches the same point is `atan2(wR·sin θ, hR·cos θ)`. They agree on a circle
  and at the four cardinal angles and nowhere else, so an implementation that skips the step looks
  right on `ellipse`, `donut` and every rounded corner. At an ordinary non-square size **84 of the
  389 non-degenerate arcs in the corpus land somewhere else without it.**
- **A negative `swAng` really does run backwards.** `donut`'s outer ring is four arcs of +90° and its
  inner ring four of −90°; the opposite winding is the only reason the hole is a hole. Both windings
  finish at the same point, so normalising −90° to +270° is invisible until something is filled.
- **A whole turn has to be split.** An SVG `A` whose destination equals its origin draws nothing at
  all, and 20 of the corpus's arcs are exactly that — `smileyFace`'s outline and both eyes among
  them. `arcSegments` returns two half turns for those and one segment for everything else.

`swAng` is clamped to ±360°, which is what Office does and is not the same as normalising into that
range: 400° clamped ends where it began, normalised it ends 40° along and drags every following
command with it. No preset reaches the case.

A zero radius is reported as `degenerate` and draws nothing, without moving the pen. That is a real
case rather than a guard: `round2DiagRect` and `round2SameRect` write their two square corners as
arcs of radius zero, at their default adjust values.

## Drawing one

```ts
import { getPreset, resolveGeometry } from '@pptx-studio/geometry';

const resolved = resolveGeometry(getPreset('flowChartDisplay')!, { w: 120, h: 60 });
resolved.paths.length; // 1 — one ResolvedPath per a:path, never merged
resolved.paths[0].d; // 'M0 30L20 0L100 0A20 30 0 0 1 100 60L20 60Z'
resolved.paths[0].fill; // 'norm'
resolved.textRect; // { l: 20, t: 0, r: 100, b: 60 }
```

`resolveGeometry` takes a `Geometry`, which is the preset type and the `custGeom` type both — an
`a:custGeom` carries the same six children and differs only in having no name, so `custGeom({...})`
builds one and everything downstream is shared.

Four things worth knowing:

- **One `<path>` per `a:path`.** 63 of the presets declare more than one, and **60 have paths that
  disagree about fill or stroke** — `smileyFace` is a filled head with an unfilled mouth, each
  action button a filled plate with an unfilled glyph. Merging a shape's paths into one `d` cannot
  render any of those 60.
- **`a:path/@w` and `@h` of 0 mean no path space**, and 270 of the 320 paths are that case, so
  reading them as a divisor without the zero check divides by zero on five paths in six. The 50 that
  do declare one use nine different spaces — `flowChartCollate` is drawn with the coordinates 0, 1
  and 2 in a two-unit square. The axes are independent.
- **An arc is resolved in path space, and the figure is scaled afterwards.** The other order agrees
  on circles, at cardinal angles and at square sizes, which covers every preset but `cloud` — whose
  ring of eleven arcs closes to 0.32% either way at 100×100, and tears open to 5.5% the other way at
  200×100.
- **`d` is empty when something is not finite**, and `finite` says so. A `d` holding a `NaN` is
  invalid rather than merely wrong: a browser drops the whole path silently. At zero width 11
  presets reach a non-finite number inside a path, at zero height 16, and 19 when both collapse. The
  numbers stay in `segments`.

`close` maps straight to `Z`, which has the same meaning — the next command starts a new subpath at
the closed one's first point. 35 presets rely on that and carry on drawing after a `close`.

## The definitions themselves are still strings

`'*'`, `'w'`, `'a'`, `'200000'` — a formula's operands are guide names, built-in names and literals
mixed together, and **nothing can tell them apart without a shape size and an `avLst` to resolve
against**. Parsing `200000` into a number at transcode time would mean deciding at build time which
operands are numbers. `evaluateGuides` is where that decision is finally taken, with the information
needed to take it.

Three things are worth knowing before you read a `Geometry` yourself rather than resolving one:

- **`path.w` and `path.h` of 0 mean no path-space scaling**, not a degenerate path — see
  [Drawing one](#drawing-one).
- **`arcTo` carries no destination point.** The end is computed from the two radii and the two
  angles, against the point the previous command left behind — see [Arcs](#arcs).
- **`rect` is `null` for five shapes** — `line`, `lineInv`, `chartPlus`, `chartStar`, `chartX` —
  and for any `custGeom` that omits `a:rect`. Defaulting a missing text rectangle to the shape
  bounds is a policy, so the package does not pick one for you.

## Where the data comes from

Transcoded at build time from Apache POI's `presetShapeDefinitions.xml` (Apache-2.0). The file is
**not** redistributed here; only the generated buckets are, and each records the SHA-256 of the
input it came from. See [`NOTICE`](../../NOTICE) and
[`tools/geometry-codegen`](../../tools/geometry-codegen).

Two things in that source are outside ECMA's grammar and are **carried through rather than
repaired**, because a transcoder that quietly corrected its input would be making semantic decisions
that belong to the evaluator:

| What                                      | Where                                                            |
| ----------------------------------------- | ---------------------------------------------------------------- |
| Eight `+-` formulas with a fourth operand | the circular-arrow family — `OVER_LONG_FORMULAS`                 |
| Six formulas containing a double space    | `heptagon`, `pentagon`, `star5`, `star7`, `leftArrow`, `upArrow` |

2.2 established that the first of those is **provably** benign rather than probably benign. All
eight have the same shape — `circularArrow.xB = +- xH 0 dxB 0` — and the fourth operand is the
literal `0` every time, so `x + y − z` ignoring it, `x + y − z + w` and `x + y − z − w` all give the
same number. Every reading agrees, so there is nothing to get wrong. The evaluator accepts a fourth
operand when it is exactly zero and refuses it otherwise, because a non-zero one is a case where the
readings genuinely diverge and the honest response to that is to stop.

There is a third, which the reader cannot detect generically and which is pinned by a test instead:
`pie`'s text rectangle reads `l="il" t="ir" r="it" b="ib"`, with top and right transposed relative to
the guides' own names. It is the only one of the 29 shapes using that guide set written that way.

## Size, and the six buckets

The buckets are encoded strings rather than object literals: a faithful object literal of 539 KB of
XML comes out **larger** than the XML, because `{ kind: 'lnTo', to: { x: 'x2', y: 't' } }` spends 34
characters saying what `L x2 t` says in 6. All 187 shapes are 119 619 characters of encoded text,
decoded lazily one shape at a time. The grammar is documented in [`src/decode.ts`](src/decode.ts),
which is its only reader.

Built and gzipped — which is what a browser actually downloads:

| Bucket      | Presets |   gzip |
| ----------- | ------: | -----: |
| `arrows`    |      31 | 6.1 KB |
| `misc`      |      47 | 6.9 KB |
| `basic`     |      43 | 5.8 KB |
| `stars`     |      21 | 5.5 KB |
| `callouts`  |      16 | 2.1 KB |
| `flowchart` |      29 | 1.8 KB |
| the decoder |       — | 3.2 KB |

`getPreset` pulls in all six buckets — about **33 KB gzipped** — because a general renderer cannot
know which shapes a deck holds until it opens one. A consumer that does know can import one:

```ts
import { FLOWCHART } from '@pptx-studio/geometry/presets/flowchart.gen.js';
FLOWCHART.get('flowChartDecision');
```

and pay 1.8 KB plus the decoder instead of 33 KB. **Bucket membership is a packaging decision, not a
fact about geometry** — `basic`, `arrows`, `callouts`, `flowchart`, `stars`, `misc` follow
PowerPoint's own galleries where a name makes that mechanical — so a direct bucket import should pin
a version. `getPreset` never moves.

## Regenerating

The input is not in the repository and nothing here downloads it. See
[`tools/geometry-codegen/README.md`](../../tools/geometry-codegen/README.md).
