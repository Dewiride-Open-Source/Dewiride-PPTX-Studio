# @pptx-studio/render-svg

## 0.6.0

### Minor Changes

- 6ceb245: A rectangular picture's border is drawn as PowerPoint's export draws it.

  At a named width the border is the device pen on the frame outset by half the true width, over
  the picture, with no clip: crisp when the width is a whole number of device pixels and antialiased
  otherwise (F3, `snap.json`: 48 of 48 borders). The path carries `data-band="out"`, and a zoom mask
  may let its crispness change with the pen as it lets `stroke-width` change. A picture that is not
  a rectangle keeps its double-width band clipped to the outside. `StrokePaint` gains `whole`, whether
  the pen is the true width in whole device pixels.

  The rule's own domain is now measured: every case fits it on the six widths that are an eighth of
  a pixel per point with a whole-pixel height, 239 and 227 of 242 at the two other whole-dpi,
  whole-height widths where an exact half rounds its width and its parity apart, an exact half
  rounds up at every power-of-two scale — which is every zoom the page has — and an `algn="in"`
  stroke, measured for the first time, sits a device pixel outside its rounded frame edge; the
  renderer's clipped band for it is up to a pixel and a half too far in, and the record says so.

## 0.5.0

### Minor Changes

- da4dc86: Edges land where PowerPoint's export puts them on the device grid.

  `devicePen` in `paint` gives a stroke its whole-pixel pen and, when that pen is odd, the half
  device pixel it sits past the rounded coordinate (F3, `snap.json`: 732 of 732 axis-aligned
  strokes). At a named width `render-svg` draws every rectilinear path — a rectangle, a horizontal or
  vertical line, a rectilinear `custGeom`, turned by a quarter or flipped — with
  `shape-rendering="crispEdges"`, and moves an odd pen half a device pixel down and right, so a
  whole-pixel stroke is one crisp row at every offset and a fill edge starts on the row the export
  starts it on. Curves, diagonals and turned rectangles stay antialiased where they lie; a clipped
  band stays antialiased under its clip.

  A named `height` now stretches the slide into its box, as the export stretches into its own
  whole-pixel box; with only a width the aspect holds. `render-dom` mounts the same tree.

### Patch Changes

- Updated dependencies [da4dc86]
  - @pptx-studio/paint@0.4.0
  - @pptx-studio/model@0.1.4

## 0.4.0

### Minor Changes

- 33a6f9d: Line ends are drawn, and strokes round to the display's own pixels.

  `render-svg` draws the six `a:headEnd`/`a:tailEnd` heads as markers on the open ends of a stroked
  path, sized from the pen PowerPoint's export uses at every width: two points at or under a
  two-point line, the whole-pixel stroke above it, never under a device pixel (`markerPen` in `paint`,
  49 of 49 F2 cases). No renderer had drawn a line end before.

  `RenderOptions.devicePixelRatio` multiplies the named width for the stroke rule alone, so a slide at
  100 % on a 2x display carries the strokes of 200 % at the same CSS size; `render-dom` passes it
  through. The studio page passes the display's ratio and re-mounts when the window moves to another
  display.

### Patch Changes

- Updated dependencies [33a6f9d]
  - @pptx-studio/paint@0.3.0
  - @pptx-studio/model@0.1.3

## 0.3.0

### Minor Changes

- 2f84efc: Strokes are drawn at the width the slide is drawn at, so a zoom is a fresh mount and never a resize.

  PowerPoint's export, asked at seven widths from 120 to 3840 pixels, draws a stroke at its width
  rounded to whole device pixels and never under one, and a zero width as one pixel — 42/42 hairlines
  and 56/56 thin strokes, against every hand-written alternative (F2, `zoom.json`). `paint` gains
  `deviceStrokeWidth(widthEmu, pxPerPt)` and the floor it holds, `MIN_PX_PER_PT`. When
  `RenderOptions.width` names a device, `render-svg` rounds `stroke-width` and `stroke-dasharray` to
  its pixels and draws a dashed hairline solid; without one, the true width, and a hairline as
  `stroke-width="1" vector-effect="non-scaling-stroke"`. `MountedSlide.resize()` is gone from
  `render-dom`: mount again at the new width. The text layout is the same at every width.

  Bullets are drawn. `render-svg` lays out `a:buChar`, `a:buAutoNum` and `a:buBlip` through
  `@pptx-studio/text` and emits them in both renderers; an autonumber scheme the measurement could not
  read (`arabic1Minus`, `arabic2Minus`) throws `TEXT_AUTONUMBER_UNMEASURED` rather than drawing a
  guess. A bordered picture keeps its picture: the clipped outline band is its own path. An effect
  filter is written in points inside a `scale(12700)` pair — a rotated glow in an EMU-unit filter
  cost Chromium nine seconds a shape, and thirty milliseconds in points. `mediaFromStore(store)` is
  the one media resolver, shared by the CLI, the harness and the page. ADR 0054.

### Patch Changes

- Updated dependencies [2f84efc]
  - @pptx-studio/paint@0.2.0
  - @pptx-studio/model@0.1.2

## 0.2.0

### Minor Changes

- 55f3073: `TextOptions.cssFamilyFor` decides what a run's `font-family` says.

  `renderSlide` wrote the run's own family, quoted, which is right for a renderer that measures and
  draws through the same name and wrong for one that measured elsewhere. The hook is called once per
  piece with the run's family, weight and style; it defaults to the quoted family, so a caller that
  passes nothing emits byte-identical markup. `fontStack` takes the face a run was drawn in as a
  second argument and leads with it, deduplicated as before.

### Patch Changes

- Updated dependencies [55f3073]
  - @pptx-studio/text@0.2.0

## 0.1.1

### Patch Changes

- 76b505c: Published with a provenance attestation. The first release went out over a token rather than OIDC, and npm attests automatically only on the OIDC path, so the registry recorded no attestation and `npm audit signatures` could not verify these packages. Nothing else about them has changed.
- Updated dependencies [76b505c]
  - @pptx-studio/geometry@0.1.1
  - @pptx-studio/model@0.1.1
  - @pptx-studio/paint@0.1.1
  - @pptx-studio/text@0.1.1

## 0.1.0

### Minor Changes

- 7d905f4: First release.

  `@pptx-studio/render-svg` turns a slide into an SVG string, and `pptx-studio render`
  does it from the command line with no browser and no LibreOffice — text included,
  measured out of the font's own tables. The other nine packages are what those two
  are built on and are published because they have to be, not because their surface
  is settled.

  Everything here is pre-1.0 and the API will change. What will not change without a
  very good reason is what the writer does to a package it was not asked to edit: a
  part nobody touched is re-emitted byte for byte, across all 54 decks in the corpus.

### Patch Changes

- Updated dependencies [7d905f4]
  - @pptx-studio/geometry@0.1.0
  - @pptx-studio/model@0.1.0
  - @pptx-studio/paint@0.1.0
  - @pptx-studio/text@0.1.0
