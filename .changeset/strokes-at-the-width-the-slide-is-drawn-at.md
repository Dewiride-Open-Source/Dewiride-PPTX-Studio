---
'@pptx-studio/paint': minor
'@pptx-studio/render-svg': minor
'@pptx-studio/render-dom': minor
'@pptx-studio/cli': patch
---

Strokes are drawn at the width the slide is drawn at, so a zoom is a fresh mount and never a resize.

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
