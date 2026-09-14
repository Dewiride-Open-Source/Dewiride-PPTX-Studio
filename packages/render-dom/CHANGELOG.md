# @pptx-studio/render-dom

## 0.3.3

### Patch Changes

- e8ba475: A table is parsed, and its grid is the one PowerPoint draws.

  `Shape.table` carries the `a:tbl` of a `p:graphicFrame`: the `a:tblPr` flags and style source, every
  `a:gridCol`, every `a:tr` with its cells, and each cell's spans, body and `a:tcPr` as written.
  `tableGrid` resolves the merges by the rule C7 measured (`tables.json`, 76 of 76 tables PowerPoint
  opened as written and 21 it authored itself): each `a:tc` takes the next column, the spans on the
  anchor are the whole story, a covered position is covered whatever its own attributes say, a span
  stops at the edge or at the first claimed position, and `hMerge`/`vMerge` change nothing. `a:tr/@h`
  is a minimum, a column is never narrower than its cells' side margins plus 2 pt, and the frame's
  `a:ext` says nothing about the drawn size. A table attribute of the wrong type throws
  `MODEL_TABLE_ATTR`.

  `@pptx-studio/validate` gains `V030`, a warning on a table whose rows, spans and flags disagree —
  PowerPoint opens it silently and rewrites it on the next save — and `V031`, fatal on the missing or
  mistyped table attributes PowerPoint repairs. The downstream packages carry the new rules and the
  new field; nothing they draw changes.

- Updated dependencies [e8ba475]
  - @pptx-studio/model@0.2.0
  - @pptx-studio/render-svg@0.6.1

## 0.3.2

### Patch Changes

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

- Updated dependencies [6ceb245]
  - @pptx-studio/render-svg@0.6.0

## 0.3.1

### Patch Changes

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

- Updated dependencies [da4dc86]
  - @pptx-studio/paint@0.4.0
  - @pptx-studio/render-svg@0.5.0
  - @pptx-studio/model@0.1.4

## 0.3.0

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
  - @pptx-studio/render-svg@0.4.0
  - @pptx-studio/model@0.1.3

## 0.2.0

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
  - @pptx-studio/render-svg@0.3.0
  - @pptx-studio/model@0.1.2

## 0.1.1

### Patch Changes

- 55f3073: The help text spells `this platform's own font directories` with its apostrophe, and the
  `render-dom` tarball carries `CHANGELOG.md` like the other eleven.
- Updated dependencies [55f3073]
  - @pptx-studio/render-svg@0.2.0

## 0.1.0

### Minor Changes

- 63d04e0: `mountSlide` and `mountOverlay` can be installed.

  The live DOM renderer was the twelfth package and the only one never published. It sat at `0.0.0`
  in changesets' `ignore`, so `@pptx-studio/render-dom` did not exist on the registry at all and the
  two mount entry points were reachable only from a checkout of this repository.

  That was not a decision about the code. A trusted publisher is a setting on an npm package and npm
  has no way to attach one to a name the registry does not hold, so the first publish of a new name
  cannot go out over OIDC — the release would have reached it, failed to authenticate, and stopped
  partway. The name now exists, its trusted publisher points at this repository's `release.yml`, and
  the package is a normal member of the release.

  Nothing about the package itself changed. `shapeOverlay`, the debug overlay, stays in
  `render-svg` where it always was.
