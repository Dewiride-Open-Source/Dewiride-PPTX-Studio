# @pptx-studio/paint

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

## 0.1.1

### Patch Changes

- 76b505c: Published with a provenance attestation. The first release went out over a token rather than OIDC, and npm attests automatically only on the OIDC path, so the registry recorded no attestation and `npm audit signatures` could not verify these packages. Nothing else about them has changed.

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
