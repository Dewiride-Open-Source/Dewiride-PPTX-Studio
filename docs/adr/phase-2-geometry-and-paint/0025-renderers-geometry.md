# ADR 0025 — Renderers, geometry only

**Sub-phase 2.10.** Status: accepted. Supersedes nothing; **corrects one line of the approved plan's
5.3** and adds two findings the plan does not mention.

Experiment: **C6**, `tools/ground-truth/{author-transforms.ps1, transforms.ts,
build-transform-deck.ts, read-transforms.ps1, analyse-transforms.ts}`.
Fixture: `corpus/ground-truth/transforms.json` — 65 probes across 15 packages, 54 bitmap samples,
plus two decks PowerPoint authored itself.

---

## The question, and why it could be asked in writing

`a:xfrm` carries `@rot`, `@flipH` and `@flipV` and says nothing whatever about the order they
compose in. The two candidate readings are not equivalent: for any reflection `F`,

```
F R(t) F = R(-t)
```

so `R(t) F` and `F R(t)` differ by the sign of the angle. That is what makes the question answerable
without measuring a single pixel. Take a shape PowerPoint has already rotated 30 degrees and ask it
to mirror the shape. The mirrored figure is fixed — it is whatever the user sees — and PowerPoint has
to write an `a:xfrm` that produces it. Under flip-then-rotate it must write `-30`; under
rotate-then-flip it must write `+30`. Its writer and its renderer agree with each other by
construction, so **what it writes is the order it renders in**.

It wrote `rot="19800000" flipH="1"`. Minus thirty. Again at 45, 120 and 200 degrees, and again on
the vertical axis. Both flips together left the angle unchanged, which is the same fact from the
other side — `flipH flipV` is a half turn and rotations commute.

`author-transforms.ps1` is 250 lines and settled the sub-phase's central question, plus four others,
before a probe existed. Third sub-phase running that the cheapest thing in the directory was the
highest-yield: **ask the format's author in writing, first.**

It also settled the measurement method for everything else. A leaf shape inside a group has no slide
position of its own; the group's transform gives it one, and `Shape.Left` reports that position.
Measured against PowerPoint's own output: a child of a group rotated 30 degrees and mirrored reported
`L=321.2435302734375`, and the identical number appears in the file as `x="4079793"` once the group
is ungrouped. So the object model is the renderer's own arithmetic to four decimal places, and 46 of
C6's 65 probes need no bitmap at all.

---

## Findings

### 1. Flip happens before rotate — confirmed three ways

Once from what PowerPoint wrote (above); once from the _positions_ of a mirrored group's children,
where the two models put a child 80 points apart and the readings picked flip-first; and once by
scoring, where the flip-first map got 46 of 46 composed positions and the rotate-first map 45 — the
one it missed being the only probe with a group that is both turned and mirrored, which is the only
case where the two can disagree.

Composition, in closed form. Writing a turn as `R(a) F`:

```
R(a) F₁ · R(b) F₂  =  R(a + s₁·b) · (F₁F₂),    s₁ = −1 when F₁ mirrors exactly one axis
```

So a group's rotation and its child's **subtract** under a mirrored group. Scored against 13
orientation probes: this rule 13, rotate-first 9, summing the angles 9.

### 2. The child coordinate map, and its three zeros

```
x_slide = off.x + (x_child − chOff.x) × (ext.cx / chExt.cx)
```

| candidate            | right |
| -------------------- | ----- |
| the above            | 46/46 |
| ignoring `ext/chExt` | 36/46 |
| ignoring `chOff`     | 7/46  |

Dropping the scale is the dangerous one — right on three quarters of the probes and on every group
nobody has resized. **A `chExt` of zero means the axis is not scaled**; an `ext` of zero really is
zero; and a group with **no `a:chOff`/`a:chExt` at all** uses `chOff = 0`, not `chOff = off`. A
negative `chExt` is a broken file, and PowerPoint repairs it.

### 3. A rotated child in a non-uniformly scaled group: a quadrant snap

The true image is a parallelogram and `a:xfrm` has nowhere to put a shear, so PowerPoint
approximates. Measured at eighteen angles: it keeps the centre and the angle, and **exchanges which
scale factor reaches which extent** when the child's angle rounds to a quarter or three quarters of a
turn.

| rule                         | right |
| ---------------------------- | ----- |
| quadrant, rounding halves up | 18/18 |
| `\|sin\| ≥ \|cos\|`          | 17/18 |
| `\|sin\| > \|cos\|`          | 14/18 |

At 45 degrees it swaps and at 135 it does not, so the tie goes upward both times — which is why the
`sin`/`cos` forms, the ones anyone writes first, are wrong at one end or the other. Confirmed in the
pixels: three bitmap samples placed to be inside a sheared parallelogram only, inside a plainly
scaled rectangle only, and inside either, and all three agree with the frame the object model
reports.

### 4. A group's own fill is never painted

A `p:grpSp` with a solid fill, over a rectangle no child covers, shows the slide behind it. The
group's fill exists only to be asked for.

### 5. `a:grpFill` is a slice, over a rectangle nobody writes down

The plan says `a:grpFill` "inherits from the enclosing group" and does not say over what. It is a
slice rather than a copy — two congruent children at opposite ends of one group read red and blue
where a copy would read the same purple twice — and the rectangle is **the bounding box of everything
the group contains**, not the group's own `off`/`ext`.

Solved rather than guessed. Every probe carries a reference shape with the identical gradient over a
known rectangle, drawn behind everything, so a sampled colour turns back into a fraction along the
ramp; two fractions solve for the span. Four arrangements, including a vertical ramp and three levels
of nesting:

| arrangement                                   | measured span | content bounds | group declares |
| --------------------------------------------- | ------------- | -------------- | -------------- |
| two children at 150..250 and 550..650         | 150..650      | 150..650       | 100..700       |
| one child at 300..400                         | 300.8..400.8  | 300..400       | 100..700       |
| three groups deep, leaf at 300..400           | 300.8..400.8  | 300..400       | 100..700       |
| vertical ramp, children at 130..190, 320..380 | 129.9..381.0  | 130..380       | 100..400       |

A sibling that does not use `grpFill` still sets the extent, and so does one placed entirely outside
the group's declared rectangle. A `grpFill` that reaches no group paints nothing — measured on a
top-level shape, over an explicit `a:noFill`, and over a group that states no fill at all.

### 6. A group does not scale the strokes inside it

A 4pt outline inside a group stretched to double width reports 4pt after the group is composed away,
and the pixel row a doubled stroke would cover is white.

### 7. `@showMasterSp` hides everything inherited, not the master's contribution

Its name says master; it means _above_.

| slide | layout | master's shape | layout's shape |
| ----- | ------ | -------------- | -------------- |
| —     | —      | shown          | shown          |
| `0`   | —      | hidden         | **hidden**     |
| —     | `0`    | hidden         | shown          |
| `0`   | `0`    | hidden         | hidden         |

One recursive definition covers all four rows. Sub-phase 7.9 sells this as a "hide layout graphics"
toggle, which turns out to be the name the behaviour actually has. An unmatched master or layout
placeholder is never drawn on a slide.

This closes one of 2.9's open questions.

---

## Decisions

### The two renderers share a node tree, not just a layout

The plan says "a second string-only SVG emitter over the same layout engine". Sharing a layout is not
enough on its own: emitting a `<linearGradient>` twice, once as text and once through
`createElementNS`, is two chances to write the same gradient wrong in two different ways. So the
shared surface goes one level lower — `render-svg` builds a tiny immutable node tree and serialises
it, `render-dom` instantiates the identical tree, and the suite asserts the browser's own
`XMLSerializer` returns the string emitter's output character for character.

### User units are EMU

The viewBox is the slide in English Metric Units. Every number the model and `paint` hand over is
already in them — a stroke width, a pattern tile's six points, a blur radius — and rescaling at the
boundary is a conversion that has to be right in a dozen places instead of none.

### Geometry moves into the model

2.9 recorded a preset's _name_, which says what a shape is and is not enough to draw it: `a:avLst` is
where a rounded rectangle's corner radius lives, and a preset drawn with its defaults when the file
overrode them is a visibly wrong shape. `Shape.geometry` now carries the parsed `a:prstGeom` or
`a:custGeom`, and `@pptx-studio/geometry` becomes a dependency of `model`.

It belongs there rather than in a renderer because **geometry inherits**: a slide placeholder that
states no `a:xfrm` usually states no `a:prstGeom` either and takes both from its layout. Reading it
off `shape.node` in a renderer would put half of one answer in the model and the other half somewhere
that cannot see the chain.

### The top level is not wrapped in the `p:spTree`'s own transform

Every deck writes `p:grpSpPr/a:xfrm` as four zeros, and the measured rules make four zeros the
identity anyway — a `chExt` of zero is no scaling and a `chOff` of zero is no offset. The two
readings agree on every file that exists; placing the top level directly is the one that also agrees
on a file that does not.

---

## What this changes in the plan

**5.3 says: "on group resize, bake the scale into each child's `off`/`ext` and reset
`chOff`/`chExt`".** PowerPoint does the opposite. Asked to double a group's width it changed the
group's `ext` alone and left `chExt` and every child byte-identical:

```xml
<!-- the group, fresh -->              <!-- after Width × 2, Height × 0.5 -->
<a:ext cx="5080000" cy="2794000"/>     <a:ext cx="10160000" cy="1397000"/>
<a:chExt cx="5080000" cy="2794000"/>   <a:chExt cx="5080000" cy="2794000"/>
<!-- both children: unchanged -->      <!-- both children: unchanged -->
```

Baking is what `Ungroup` does, and that is a different operation. 5.3's underlying worry is real — a
rotated child in a non-uniformly scaled group genuinely cannot be expressed — but the answer is
finding 3 above, not baking on every resize. Rewriting `ext` alone is also what makes a group resize
one edit instead of _n_, and what keeps a group's children round-tripping byte-for-byte.

**Two additions for 1.2's firewall**, neither in the plan's appendix:

- `a:chExt` with a negative extent — PowerPoint repairs the file.
- (from 2.9, restated) a master's placeholder vocabulary, and one theme part per master.

---

## Open questions

1. **A `path="rect"` or `path="shape"` gradient** is emitted as an ellipse. SVG has no primitive for
   a Chebyshev ramp; concentric rectangles inside a `clipPath` would be exact and are a structural
   change to how a fill is emitted.
2. **`a:grpFill` on a rotated child.** The fill box is carried in slide coordinates and translated
   into the child's local space, which is exact for an unrotated child and approximate otherwise. Not
   measured — no probe rotated a `grpFill` child.
3. **The content bounds of a rotated leaf.** The union uses unrotated frames. Whether PowerPoint uses
   the ink bounds instead is unmeasured.
4. **A group rotated inside a non-uniformly scaled group.** The extent swap applies to the group's
   own frame, and its children then scale against the swapped extents. Consistent, unmeasured.
5. **The tile phase of a pattern under `a:grpFill`.** `grpFill` reaches a pattern — measured, ink
   density 0.71 against a reference's 0.78 — but whether the tiling is continuous with the group's is
   not.
6. **Which colour map an inherited master shape resolves against** when the slide carries an
   `a:overrideClrMapping`. The renderer uses the shape's own sheet.
7. **Whether a layout shows its master's placeholders.** Only slides were probed.

Carried from 2.8 and earlier: whether 0.9 is exactly nine tenths; what `a:softEdge` does at a corner;
`@rotWithShape` on glow and soft edge; the twenty `prstShdw` presets as transforms; whether a
compound stroke's dash applies to the band or to each rail; whether the miter limit is exactly 8; the
off-centre path focus; `a:path` with no `@path`; corner softening; a gradient slide background's
extent; `@flip`.

**Scope gap, re-flagged:** `a:blipFill` remains unassigned in the plan though Gate 2 asks for cropped
image fills. It has a place in the `Fill` union, a place in the renderer's `switch`, and nothing
behind either.
