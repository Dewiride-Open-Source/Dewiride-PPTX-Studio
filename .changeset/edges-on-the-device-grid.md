---
'@pptx-studio/paint': minor
'@pptx-studio/render-svg': minor
'@pptx-studio/render-dom': patch
'@pptx-studio/cli': patch
---

Edges land where PowerPoint's export puts them on the device grid.

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
