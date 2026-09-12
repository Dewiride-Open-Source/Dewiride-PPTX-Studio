# @pptx-studio/render-dom

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
