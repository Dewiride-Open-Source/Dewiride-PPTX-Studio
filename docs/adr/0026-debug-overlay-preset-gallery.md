# ADR 0026 — Debug overlay and preset gallery

**Sub-phase 2.11.** Status: accepted. Closes Gate 2. Supersedes nothing; fixes one defect in 2.10's
`render-dom` and records one correction to a claim made in 2.10's own ADR.

Code: `packages/render-svg/src/overlay.ts`, `packages/render-dom/src/overlay.ts`,
`apps/studio/src/{gallery.ts, slides.ts}`, `apps/studio/gallery.html`.

---

## What this sub-phase had to decide

The plan asks for "per selected shape: resolved path with per-subpath colouring, the `a:rect` text
rectangle, `cxnLst` sites, `ahLst` handles with ranges, and a live `gdLst` dump", verified by "a
187-preset gallery with draggable handles". Almost all of the machinery already existed — 2.3 and
2.4 resolve the geometry, 2.5 inverts the handles — so the questions here are about placement, not
about DrawingML.

### The overlay is built in `render-svg`, not in `render-dom`

The plan's package layout puts `overlay` under `render-dom`. Everything the debug overlay draws is a
**fact about the geometry** — where the paths went, where the text rectangle is, how far a handle can
travel — and 2.10's decision was that the two renderers share a node tree rather than only a layout,
precisely so they cannot disagree about what is on screen. An overlay built separately in the DOM
renderer would be a second opinion about where a handle is, and the one place that must never happen
is the chrome a user drags.

So `shapeOverlay` returns an `SvgNode` tree like everything else, `render-dom` instantiates it with
the same `createNode` it uses for shapes, and the suite asserts the two agree character for
character. `render-dom` adds exactly four things a node tree cannot have: an element to mount into,
hit-testing, the drag, and the inverse transform.

### A handle's range is drawn as its locus, not as its bounds

The obvious reading of "handles with ranges" is a line from the position at `min` to the position at
`max`. That is wrong for every polar handle: `blockArc`'s angle axis traces an ellipse, so a
straight chord between the endpoints leaves the shape entirely and describes a range the handle
cannot reach. The overlay therefore samples the position at `LOCUS_SAMPLES` points across the axis,
through `resolveHandles` — the same evaluator that draws the shape — and draws the curve it actually
follows. The test asserts non-collinearity on `blockArc` so a future simplification to a chord fails.

### A drag writes `a:avLst`, and the shape is rebuilt from it

`dragHandle` returns `{ guideName: value }`, already rounded to the integers a file holds, for the
guides that handle controls and no others. The gallery merges that into the `a:avLst`, rebuilds the
shape from the resulting XML and re-resolves the handles. So the number in the panel is exactly what
a save would write, and a handle confined to a curve does not follow the pointer off it — which is
correct, and is what PowerPoint does.

`mergeAdjust` replaces a guide in place and appends one the shape had left at the preset default,
because `a:avLst` is an ordered list and a shape with three adjusts whose second one moves must
still write three, in order.

---

## The defect this found in 2.10

**`render-dom` and `render-svg` were emitting different numbers.** `createNode` wrote attribute
values with `String(value)`; the string serializer writes them through `num`, which rounds to three
decimals and normalises `-0`. So the two renderers disagreed:

```text
render-svg   x="117158.343"
render-dom   x="117158.34312"
```

Every coordinate in a slide render happens to be a whole number of EMU, which is why 2.10's suite —
including its character-for-character comparison of the two renderers — passed while this was
untrue. The overlay's text rectangle is the first fractional coordinate the package ever emitted, and
it failed on the first run. `createNode` now uses `num`.

This is the second time in two sub-phases that the thing which caught a real bug was a comparison
against something outside the package, rather than an assertion written alongside the code.

## The correction to 2.10's ADR

2.10 wrote that a handle with no `gdRef` "cannot be dragged, and 2.5 found twelve shapes with an
axis that is not invertible either", conflating two different facts. Measured: **no preset has a
handle with no axis at all.** All 120 presets that declare an `a:ahLst` declare at least one
`gdRef` on every handle. The guard against an axis-less handle is kept because an `a:custGeom` in a
real file may have one; the claim about presets was wrong and is corrected in the source.

---

## What the browser was asked, rather than told

`framePoint` is the arithmetic form of the `transform` attribute `frameTransform` emits, and a
renderer that draws with one and hit-tests with the other puts its handles in the wrong place on
exactly the rotated shapes where nobody would suspect the renderer. So the test builds the real
attribute, hands it to Chromium, and compares against the browser's own `getCTM`.

The bound is **one EMU**, not equality, and that is a measurement rather than a fudge: an `SVGMatrix`
is single precision, so at a slide coordinate near 1.2 million EMU a float32 mantissa is worth about
a tenth of an EMU. The observed disagreement is 0.019. One EMU is a hundred-thousandth of a point.

The same test computes what a rotate-before-mirror renderer would say and asserts the browser
refutes it, by more than a shape's width.

---

## What the mutation sweep changed

Twenty mutations, and the first run killed only thirteen. The seven survivors were not noise:

- **Five were untested guards.** The non-finite checks on the text rectangle, the connection sites
  and the handles were added in response to a `NaN` in the markup, and the case that produced it —
  `moon` at zero width — turned out not to exercise any of them. Measured across the 187 presets at a
  collapsed extent: 14 have a non-finite text rectangle at zero width, 7 a non-finite connection
  site, and 39 a handle that cannot be placed. There is now a sweep over all three degenerate sizes
  that asserts each guard drops exactly the non-finite ones, and asserts the counts are non-zero so
  it cannot quietly stop testing anything.
- **Two were real assertion gaps**: `mergeAdjust` rounding on the _replace_ path (the test only
  covered the append path), and the widened-bound colour.
- **One was a weakness in the code, not the tests.** `unmount` removed four listeners in four
  statements, so removing three and forgetting one was possible — and only partly detectable, which
  is why the mutation survived a test that was otherwise correct. The four are now registered with a
  single `AbortController` signal, so the failure mode does not exist.

Second run: **18 of 20 killed.** The two survivors are equivalent mutants and can be shown to be:
drawing a non-finite path changes nothing because such a path always carries an empty `d` and is
dropped a line later; and `polyline` can never receive a non-finite point because `locusOf` skips
those samples before it is called.

## Gate 2

`apps/studio` now draws the deck that is dropped on it: a slide strip, a stage, click-to-select on
any shape, and the debug overlay on the selection. Rendering runs on the main thread while the census
stays in the Worker — a census answers a question about bytes and returns JSON, while a rendered
slide is a DOM tree and a `Sheet` chain carrying raw byte offsets is not worth serialising across
`postMessage`. Sub-phase 12.1's virtualization and an `OffscreenCanvas` raster path are where that
changes.

**What Gate 2 asks for and does not get here:** the plan's gate names "a side-by-side against a
LibreOffice reference with a per-slide diff heatmap". That is the fidelity harness, which the plan
itself schedules at 3.9, and the container, `fonts.conf` and scoring it needs do not exist yet.
Standing in its place is something stronger for this phase: 2.10's `verify-render.ts` compares our
raster against **PowerPoint's own** at 54 measured points across 65 probe slides, and the whole-image
agreement is recorded. LibreOffice is a second opinion; PowerPoint is the answer. The heatmap arrives
with 3.9.

Text is absent by design. A `p:txBody` renders as its outline, because measure-equals-render needs
the text engine of 3.2–3.4 and half of one is worse than none.

---

## Open questions

1. **Selection enters groups immediately.** The innermost `data-shape` wins, so clicking inside a
   group selects the leaf rather than the group. Deliberate group entry and exit is 5.2.
2. **A locus is sampled at a fixed 25 points.** Enough to read, not enough to be a hit-test target;
   dragging near a curved locus is pointer-driven and does not snap to it.
3. **The overlay draws no chrome for a shape whose geometry is entirely non-finite** — a zero-extent
   shape reports its broken paths and nothing else. Whether an editor should show a zero-sized shape
   some minimum affordance is a UI question, not a geometry one.
4. **`data-axes` is empty for an axis-less handle and the drag ignores it.** Unreachable from any
   preset; unmeasured against a real `custGeom`.

Carried from 2.10: the `path="rect"` gradient drawn as an ellipse; `a:grpFill` on a rotated child;
the content bounds of a rotated leaf; a group rotated inside a non-uniformly scaled group; the tile
phase of a pattern under `grpFill`; which colour map an inherited master shape resolves against;
whether a layout shows its master's placeholders.

**Scope gap, re-flagged for the third time:** `a:blipFill` is unassigned in the plan though Gate 2
asks for cropped image fills. It has a place in the `Fill` union, a place in the renderer's `switch`,
and nothing behind either. A picture on a dropped deck therefore renders as its outline.
