# @pptx-studio/render-svg

A slide to SVG: transforms, geometry, fills, strokes and effects, as a string.

```ts
import { PartStore } from '@pptx-studio/opc';
import { loadDocument } from '@pptx-studio/model';
import { renderSlide } from '@pptx-studio/render-svg';

const store = PartStore.open(bytes);
const deck = loadDocument(store);
const svg = renderSlide(deck.slides[0], deck.slideSize, { width: 1920 });
```

**A string, and only a string.** Nothing here touches the DOM, so the same call
works in a tab, in a Web Worker and in Node. Text is the one part that needs to
measure something: pass `text: false` for geometry alone, pass a `measurer` of
your own, or let it default to `OffscreenCanvas`. `@pptx-studio/cli` supplies a
measurer that reads the font's own tables, which is how `pptx-studio render`
draws text in Node with no browser at all — see
[ADR 0042](../../docs/adr/phase-3-text/0042-rendering-without-a-browser.md).

Pictures need an `r:embed` resolved against the rels of the part the fill was
written in, which only the caller can do; pass `media` and they are embedded as
`data:` URIs so the document stands alone.

Sub-phase 2.10. Everything below was measured against Microsoft PowerPoint rather than argued from
the standard; the fixture is `corpus/ground-truth/transforms.json`, the experiment is C6 in
`tools/ground-truth/`, and the reasoning is `docs/adr/phase-2-geometry-and-paint/0025-renderers-geometry.md`.

Almost everything hard about drawing a slide is a transform, and almost every transform rule in
DrawingML is either unwritten or written misleadingly. This package is where those rules live.

## A flip happens before a rotation

`a:xfrm` carries `@rot`, `@flipH` and `@flipV` and says nothing about the order they compose in. The
two readings are not equivalent — for any reflection `F`, `F R(t) F = R(-t)`, so `R(t) F` and
`F R(t)` differ by the **sign of the angle**.

PowerPoint answered it in writing. Take a shape already rotated 30 degrees, ask it to mirror the
shape, and read back what it saved:

```xml
<!-- before: rotated 30 degrees -->
<a:xfrm rot="1800000"><a:off …/><a:ext …/></a:xfrm>

<!-- after Flip Horizontal -->
<a:xfrm rot="19800000" flipH="1"><a:off …/><a:ext …/></a:xfrm>
```

`19800000` is **minus** thirty degrees. Only a renderer that mirrors first has to negate the angle to
reproduce a mirrored figure; one that rotates first would have written `+30`. Same answer at 45, 120
and 200 degrees, on both axes, and again in the positions of a mirrored group's children — where the
two models put a child 80 points apart and PowerPoint picked the flip-first one.

Both flips at once left the angle **unchanged**, which is the same fact from the other side:
`flipH flipV` is a half turn, and rotations commute.

So the emitted transform is

```
translate(centre) rotate(deg) scale(±1 ±1) translate(-half)
```

which SVG applies right to left, putting the mirror innermost.

## A group's children are written in their own coordinate system

```
x_slide = off.x + (x_child − chOff.x) × (ext.cx / chExt.cx)
```

with the two axes independent. Scored over 46 measured positions:

| rule                 | right |
| -------------------- | ----- |
| **the above**        | 46/46 |
| ignoring `ext/chExt` | 36/46 |
| ignoring `chOff`     | 7/46  |

Dropping the scale factor is the one to fear. It is right on three quarters of the probes, and it is
right on every group nobody has resized — which is every group until a user drags one. PowerPoint's
own resize is what separates them: it changes the group's `ext` and leaves both `chExt` and every
child untouched, so a fresh group and a stretched one differ in exactly one number.

Three zeros, and they do not mean the same thing:

| written                   | means                                           |
| ------------------------- | ----------------------------------------------- |
| `chExt` of 0 on an axis   | that axis is **not scaled**                     |
| `ext` of 0 on an axis     | genuinely zero — the children have no width     |
| no `chOff`/`chExt` at all | `chOff = 0`, no scaling — **not** `chOff = off` |

A negative `chExt` is a broken file: PowerPoint repairs it on open.

## A rotated child in a group that scales the axes differently

There is no honest answer. The true image of a rotated rectangle under `diag(sx, sy)` is a
parallelogram, and `a:xfrm` holds an offset, an extent, an angle and two mirrors — there is nowhere
to put a shear. So PowerPoint approximates, and what it does is snap the child's angle to the
nearest quadrant and, when that quadrant is a quarter or three quarters of a turn, hand the child's
width the group's **vertical** factor and its height the horizontal one.

Measured at eighteen angles. The rounding is the part worth having:

| angle              | swapped |
| ------------------ | ------- |
| 0, 15, 30, 44      | no      |
| 45, 46, 60, 75, 90 | **yes** |
| 135, 180           | no      |

At 45 it swaps and at 135 it does not, so the tie goes upward both times — and `|sin| > |cos|`, the
form anyone writes first, is wrong at one end or the other. The centre stays where a plain scaling
would put it, and the angle is carried through untouched.

## A group's own fill is never painted

Measured: a `p:grpSp` with a solid fill, over a rectangle no child covers, shows the slide behind it.
A group's fill exists only to be asked for.

## `a:grpFill` is a slice, of a rectangle nobody writes down

Not `a:noFill`, and not a copy either. A `grpFill` child shows the enclosing group's fill laid out
over **the bounding box of everything that group contains** — not over the group's own `off`/`ext`,
which is what the file makes obvious and what is wrong.

A group declaring `off.x=100 ext.cx=600`, holding two children at 150..250 and 550..650, lays its
ramp over exactly **150..650**. Solved from the bitmap on four arrangements, including a vertical
ramp and three levels of nesting. A sibling that does not use `grpFill` still counts; a sibling
outside the group's declared rectangle still counts. A `grpFill` that reaches no group paints
nothing.

## A group does not scale the strokes inside it

A 4pt outline in a group stretched to double width is still 4pt — confirmed through the object model
and again in the pixels.

## `@showMasterSp` hides everything inherited

Its name says master and it means _above_. On a slide it hides the layout's furniture as well as the
master's; on a layout it hides the master's:

| slide `showMasterSp` | layout `showMasterSp` | master | layout | slide |
| -------------------- | --------------------- | ------ | ------ | ----- |
| absent               | absent                | shown  | shown  | shown |
| `0`                  | absent                | hidden | hidden | shown |
| absent               | `0`                   | hidden | shown  | shown |
| `0`                  | `0`                   | hidden | hidden | shown |

Placeholders are not part of it: a master placeholder and a layout placeholder that no slide shape
matched were invisible in all four cases.

## Using it

```ts
import { renderSlide, layoutSlide } from '@pptx-studio/render-svg';

const svg = renderSlide(slide, { cx: 12192000, cy: 6858000 }, { width: 960, height: 540 });
```

User units are EMU, which is what `p:sldSz` says and what every number the model and `paint` hand
over is already in — a stroke width, a pattern tile's six points, a blur radius. `layoutSlide`
returns the placed tree on its own, for a caller that wants to hit-test or draw it some other way;
`@pptx-studio/render-dom` mounts the identical node tree as live elements.

## What it does not do yet

- **`a:blipFill`** paints nothing; `p:graphicFrame` draws nothing. Phases 4 and 9.
- **A compound stroke** draws as a single rail. The five forms are rails at signed offsets from the
  geometry and SVG cannot offset a path; `compoundRails` in `paint` has the measured table for when a
  renderer can use one.
- **A `path="rect"` or `path="shape"` gradient** is emitted as an ellipse rather than as concentric
  rectangles. SVG has no primitive for a Chebyshev ramp, and an ellipse inscribed in the shape is
  much closer than a circle. Named as an open question rather than as a solved one.

## Text

Real `<text>` and `<tspan>`, never `foreignObject` — which taints a canvas and is dropped inside an
`<img>` by Safari. One `<text>` per line and one `<tspan>` per run, and no `x` on the tspans: a
line is shaped as one string and a run boundary splits only the drawing call, so a tspan carrying
its own position would lose the kern that crosses it.

The text of a shape is a **sibling** of its geometry, not a child. `flipH` mirrors the outline and
leaves the glyphs alone, so text cannot sit inside the group that carries the mirror; `flipV` turns
the whole block half a revolution about the frame centre. Measured 20 of 20 in experiment T8, where
counter-flipping both axes — which is what the plan assumed — fits 10.

The baseline sits at the typeface's own ascent-to-descent share of the 1.2 line box, which is not
the CSS half-leading model every browser implements: that fits 9 of 36 and is out by a fifth of a
line on Courier New. Underlines and strikethroughs are drawn as geometry rather than asked for with
`text-decoration`, because Chromium's own decoration agrees with PowerPoint on none of eight faces
and CSS cannot place a strikethrough at all.

`corpus/ground-truth/text-rendering.json`, and
[ADR 0034](../../docs/adr/phase-3-text/0034-text-in-both-renderers.md).

Each run's `font-family` is whatever `TextOptions.cssFamilyFor` answers for it, and by default the
run's own family, quoted: this renderer measures and draws through one name, so the two agree by
construction. A caller that measures elsewhere passes a hook that names the face it measured in
first — `@pptx-studio/cli` does, out of font tables — so its markup says what was drawn rather than
only what was asked for.
[ADR 0052](../../docs/adr/phase-3-text/0052-the-box-rounds-half-up-and-the-markup-names-its-face.md).
