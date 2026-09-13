---
'@pptx-studio/render-svg': minor
'@pptx-studio/render-dom': patch
'@pptx-studio/cli': patch
---

A rectangular picture's border is drawn as PowerPoint's export draws it.

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
