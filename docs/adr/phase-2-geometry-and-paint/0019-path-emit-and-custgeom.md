# 0019 — Path emit, and `custGeom` as the same thing

Date: 2026-09-02
Status: **accepted** — 320 paths across 187 presets, every one of them rewritten into a path space
and asked to come out the same
Sub-phase: 2.4 — path emit + `custGeom` unification

---

## Context

2.1 read the preset definitions, 2.2 turned them into numbers and 2.3 resolved the arcs. What was
still missing was the thing a renderer actually asks for: given a shape and a size, what do I draw.

Two questions had to be answered together. The first is what a path is made of, which is mostly
bookkeeping until `a:path/@w` turns up. The second is whether `a:custGeom` is a separate feature or
the same one written differently — and the answer decides whether this package has one code path or
two.

---

## Decisions

### One `ResolvedPath` per `a:path`

Not one per shape. The 187 presets hold 320 paths and 63 declare more than one, which on its own
would not force anything — the paths could be concatenated into a single `d`.

What forces it is that **60 of the 187 have paths that disagree about fill or stroke**. 96 paths
are `fill="none"` and 95 of those are stroked. `smileyFace` is a filled head with an unfilled
mouth; each of the 17 action buttons is a filled plate with an unfilled glyph on it. Concatenating
those loses a distinction the shape's own definition draws, and no amount of care downstream gets
it back.

### `@w` and `@h` of zero mean there is no path space

Not a degenerate path, and not a divisor. **270 of the 320 paths are unscaled**, so reading these
as a divisor without the zero check is a division by zero on five paths in six.

The 50 that do declare one use nine different spaces — 1x1, 2x2, 5x5, 6x6, 8x8, 10x10, 20x20,
21600x21600 and 43200x43200. `flowChartCollate` is genuinely drawn with the coordinates 0, 1 and 2
in a two-unit square.

The two axes are independent, so a zero `@w` beside a non-zero `@h` scales y and not x. That is
what two separate attributes mean; it is not reachable from the presets, where all 50 scaled paths
declare both and every one of them is square.

### A guide inside a scaled path is in path space

This one is a decision rather than a finding, and the measurement behind that word is worth stating
precisely. Across all 50 scaled paths, **every operand in an x, y, `wR` or `hR` position is an
integer literal**. The only names that appear anywhere inside them are `cd2`, `cd4` and `3cd4`, and
they appear only in `stAng` and `swAng` positions, where they are angles and belong to no coordinate
system at all.

So the corpus cannot answer it, and the reading taken is the one that keeps two steps separate:
resolving an operand produces a number, and the path is what says which coordinate system a number
in it is in. A guide is not exempt from the path's own declaration.

### An arc is resolved in path space, and then the figure is scaled

Scaling an ellipse arc by different factors on the two axes gives another ellipse arc, so two
orders were available: resolve the arc in path space and scale the result, or scale the radii first
and resolve in shape space. They agree on a circle, at cardinal angles, and whenever the shape is
square — which between them covers every preset but one.

`cloud` is the exception, and it settles it. Its outline is a ring of eleven arcs at angles like
-190.5 degrees, chained end to end in a square 43200 path space. The ring is meant to close and
closes to within 0.32%, which is the residue of the integer angles the file was authored with.
Resolved in path space that residue is an affine image of a fixed figure and stays where it is.
Resolved in shape space the ring tears open:

| size      | path space | shape space |       |
| --------- | ---------: | ----------: | ----- |
| 100 x 100 |      0.32% |       0.32% | agree |
| 200 x 100 |      0.18% |       5.46% | 31x   |
| 400 x 100 |      0.11% |       3.89% | 34x   |
| 100 x 800 |      0.31% |       4.95% | 16x   |

A 5% gap in the outline of a cloud is a visible tear, so the arc is resolved against the path's own
coordinate system and the whole figure is scaled after — the same order everything else in the path
gets. A path is a drawing in its own space, and it is the space that is stretched rather than the
drawing rules.

The radii scale per axis and the SVG flags survive, because an axis-aligned scale maps an
axis-aligned ellipse arc onto another one over the same parameter range. A reflection on exactly one
axis would flip the sweep flag; that is handled, and it is unreachable from a slide, where mirroring
is `flipH` on the transform rather than a negative size.

### `close` closes a subpath and maps straight to `Z`

35 presets have a `close` that is not the last command of its path and carry on drawing afterwards;
51 paths have no `close` at all; one path reaches eleven subpaths. SVG's `Z` behaves identically —
the command after it starts a new subpath at the closed one's first point — so nothing has to be
synthesised.

### A path that cannot be drawn emits nothing rather than `NaN`

`d` is the empty string whenever any operand, scale factor or coordinate is not finite, and
`finite` says so. A `d` holding a `NaN` is not merely wrong, it is invalid: a browser drops the
whole path and reports nothing, which is the worst of both outcomes.

This is a real case rather than a corrupt-file one. At zero width **11 presets** reach a non-finite
number inside a path, at zero height **16**, and **19** when both dimensions collapse. The numbers
stay in `segments` for a caller that wants to see what happened, which keeps this consistent with
2.2's decision to report IEEE results rather than throw.

### `Geometry` is one type, and `PresetShape` is the special case with a name

`a:custGeom` carries `avLst`, `gdLst`, `ahLst`, `cxnLst`, `rect` and `pathLst` — the same six
children with the same meanings. So a custom geometry is a preset definition written inline in the
shape instead of being looked up by name, and the only difference the type carries is that
`a:prstGeom` has a name and `a:custGeom` does not.

`evaluateGuides` and `resolveGeometry` both take `Geometry`. `PresetShape extends Geometry` with
`name: string`. There is no second code path to keep in step, which is the point.

### The `d` format

Rounded to four decimals with trailing zeros dropped, no space after the command letter, and never
a negative zero. Rounding is what makes two paths comparable for equality, which is what the
conformance test needs; the negative zero is what would otherwise make two identical paths compare
unequal. `precision` is an option for a caller that wants something else.

---

## What building it found

### Three presets state the path-space rule in shape-space arithmetic

The same trick that checked 2.3: a guide computed in shape space that has to agree with a
coordinate written in path space.

- **`flowChartDisplay`** is drawn in a 6 x 6 space with a flat top edge from x=1 to x=5. Its text
  rectangle is `l="wd6" r="x2"` with `x2 = w*5/6` — the two ends of that edge, in shape space.
- **`flowChartMagneticDisk`** has `y3 = h*5/6` where the path writes `L 6,5`, and a text rectangle
  whose top is `hd3` — exactly twice the vertical radius of a cap arc written as one unit in six.
- **`flowChartPunchedTape`** has `y2 = h*9/10` against `L 20,18` in a 20-unit space, and `hd5`
  against a scallop radius of 2.

### `flowChartTerminator` states it a fourth time, and only an arc could have produced the number

Its text rectangle is `il = w*1018/21600`, `it = h*3163/21600` and their mirrors. Those are the
corners of the rectangle inscribed at 45 degrees round the end caps: `3163 = 10800 - 10800/sqrt(2)`
and `1018 = 3475 - 3475/sqrt(2)`, to within the rounding of the integers the file holds. All four
corners land on the emitted caps' ellipses to within 2e-4, at every size.

### Only two shapes in the corpus can tell the two arc orders apart

`cloud` and `cloudCallout`, and only because their arcs sit at non-cardinal angles inside a scaled
path. That is **44 of the 393 arcs**. Every other scaled arc is at `cd2`, `cd4` or `3cd4`, where
the two readings agree exactly — the same shape of invisibility 2.3 found in the unskew, and the
reason the mutation below kills so little.

### `close` followed by `arcTo` never happens

Every DrawingML command except `arcTo` names its own destination, so where a `close` leaves the pen
is only observable if an `arcTo` follows it — and across all 320 paths that never occurs. The
corpus therefore cannot check this behaviour at all, and the only witness is a synthetic path. That
is stated as a test rather than left as a gap.

### A nameless geometry was printing a bare colon

`resolveOperand` built its message with an inline join that produced `": operand ..."` when the
site had neither a shape name nor a guide name — unreachable while every geometry was a named
preset, and reachable the moment `custGeom` existed. Both call sites now share `siteLabel`.

### The tests were checked by breaking the module

| Mutation                                        | Tests killed |
| ----------------------------------------------- | -----------: |
| `@w` used as a divisor without the zero check   |           11 |
| radii scaled first, arc resolved in shape space |            3 |
| arc radii not scaled at all                     |            5 |
| `close` leaves the pen where it is              |            1 |

The second and fourth are the informative ones. Scaling the radii first survives the whole
conformance suite and all four flowchart oracles — only `cloud` notices, exactly as the corpus
census predicts. `close` scores one because one is all the corpus can offer.

---

## Consequences

`pnpm check` green.

- `resolveGeometry` answers the whole question a renderer asks: one `ResolvedPath` per `a:path`
  with segments, a `d` string and its own fill and stroke, plus the text rectangle and the
  connection sites. `resolvePath`, `pathScale` and `pathData` are the parts worth using alone.
- **`packages/geometry` is 254 tests**, up from 214; the repository is 64 files and 1729 tests, up
  from 63 and 1689.
- The conformance test rewrites every preset into an eight-times path space — a legal `custGeom`
  expressible in `fmla` alone — at four sizes, and requires the emitted string to be **identical**,
  not merely close: 1280 path comparisons, with the segment positions agreeing to better than 1e-9
  as a separate assertion.

---

## What is not done

- **Nothing is painted.** A `d` string is not a fill, a stroke, a gradient, a dash or an effect.
  2.6 to 2.8.
- **No path has been compared against PowerPoint.** Every check here is against the presets' own
  arithmetic or against a fact about ellipses, as in 2.3. A strong internal oracle; not an external
  one.
- **Bounding boxes.** Still wanted, still cheap for arcs, and now also needed for the two cubic and
  quadratic cases. 2.5 or 2.11 is where something asks.
- **`a:custGeom` is not read from XML here.** The type is shared and the evaluator is shared; the
  reader that turns an `a:custGeom` element into a `Geometry` belongs with the other part parsers
  in 2.9.

### Open questions, each wanting an experiment in the manner of 0.7

- **A path declaring `@w` but not `@h`.** Independent axes is what the attributes say and no file in
  the corpus tries it. Test: a `custGeom` with `w="100" h="0"` and a diagonal line.
- **A guide referenced inside a scaled path.** The decision above says path space. Test: a
  `custGeom` with `w="100" h="100"` on a 200-wide shape whose path draws `lnTo` at a guide equal to
  `w`. If it lands at 200 the guide was in path space and doubled; if at 100, shape space.
- Carried from 2.3: whether Office clamps `swAng` rather than normalising it, what it does with a
  zero radius mid-path, and whether a negative radius is ever accepted.
