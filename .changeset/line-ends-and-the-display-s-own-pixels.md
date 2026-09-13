---
'@pptx-studio/paint': minor
'@pptx-studio/render-svg': minor
'@pptx-studio/render-dom': minor
'@pptx-studio/cli': patch
---

Line ends are drawn, and strokes round to the display's own pixels.

`render-svg` draws the six `a:headEnd`/`a:tailEnd` heads as markers on the open ends of a stroked
path, sized from the pen PowerPoint's export uses at every width: two points at or under a
two-point line, the whole-pixel stroke above it, never under a device pixel (`markerPen` in `paint`,
49 of 49 F2 cases). No renderer had drawn a line end before.

`RenderOptions.devicePixelRatio` multiplies the named width for the stroke rule alone, so a slide at
100 % on a 2x display carries the strokes of 200 % at the same CSS size; `render-dom` passes it
through. The studio page passes the display's ratio and re-mounts when the window moves to another
display.
