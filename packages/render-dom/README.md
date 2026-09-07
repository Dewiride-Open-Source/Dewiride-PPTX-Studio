# @pptx-studio/render-dom

The live renderer: one `<svg>` per slide, built as real elements.

Sub-phase 2.10. Every fact about DrawingML this package draws on lives in
[`@pptx-studio/render-svg`](../render-svg#readme) and was measured against Microsoft PowerPoint in
experiment C6.

## One tree, two renderers

The plan asks for two renderers "over the same layout engine", so that a thumbnail and the editor
cannot disagree about what is on the screen. Sharing a _layout_ is not enough for that on its own:
emitting a `<linearGradient>` twice, once as text and once through `createElementNS`, is two chances
to write the same gradient wrong in two different ways.

So the shared surface goes a level deeper. `render-svg` builds a node tree and serialises it; this
builds the identical tree and instantiates it. The suite asserts that the browser's own
`XMLSerializer` over the mounted nodes returns the string emitter's output character for character.

## The namespace

`document.createElement('path')` makes an HTML element that happens to be spelled `path`, and it
draws nothing at all — silently, with no error anywhere. Every element here goes through
`createElementNS`. Attributes are the mirror image: they belong to no namespace and go through plain
`setAttribute`.

The suite runs in real Chromium rather than jsdom, which is the repository's standing choice and
earns it here: the test calls `getTotalLength()` on the mounted path, which an HTML element cannot
answer.

## Zoom is a transform, never a re-layout

The mounted `<svg>` carries a `viewBox` in EMU and is sized in CSS pixels, so scaling it is one
attribute and no arithmetic — `resize()` writes `width` and `height` and nothing else moves. The
plan's rule is the reason: font metrics are not linear in point size, so a renderer that re-lays out
per zoom level changes its line breaks as the user zooms.

## Using it

```ts
import { mountSlide } from '@pptx-studio/render-dom';

const mounted = mountSlide(host, slide, { cx: 12192000, cy: 6858000 }, { width: 960, height: 540 });
mounted.element(shape.cNvPrId); // the element drawing one shape, for hit-testing
mounted.resize(1280, 720);
mounted.unmount();
```

`document` may be passed explicitly, because this package has to work in a Web Worker with no global
one — and because a test that has to stub a global is a test that can pass for the wrong reason.

## What it does not do yet

Text (3.8), the overlay canvas and adjust-handle chrome (2.11), selection and gestures (Phase 5),
and slide virtualization (12.1). The host is emptied on mount; anything a caller wants beside the
slide is a sibling of the host, not a child of it.

## The text layer

`mountTextLayer` builds an HTML layer over the slide from the identical `TextBlock`
`@pptx-studio/render-svg` emits: one div per shape, one per line, a span per piece, and the same
rule rectangles. Real text nodes, so selection, an IME and a screen reader all work on them.

There is one set of line boxes and both renderers read it, which is what stops them drifting. The
one place the HTML layer is not exact is the baseline: a browser puts an inline baseline a
half-leading plus an ascent below the line box, from font metrics it has rounded to whole pixels,
so the `line-height` that reconciles the two lands within one CSS pixel and no closer. The SVG
emitter, where a baseline is a coordinate, is exact — which is the right way round, since that is
what thumbnails and export go through.
