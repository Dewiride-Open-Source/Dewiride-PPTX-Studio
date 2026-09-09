---
'@pptx-studio/cli': patch
---

The ideographic baseline comes out of the font's `BASE` table.

`pptx-studio render` placed every upright East Asian glyph on the face's box descent,
which is right only for a face carrying no `BASE` table. Yu Gothic reports a box descent
of 0.302 em and an ideographic baseline of 0.1201 em, so a vertical `eaVert` frame was
drawn 0.18 em out of place on it — and the browser renderer, which asks Chromium, was
never wrong about the same face.

`sfnt.ts` now reads the `BASE` horizontal axis' `ideo` coordinate under the `DFLT`
script, and `FaceMetrics` carries it as `ideographic`. Where there is none the box
descent stands, and a coordinate outside the em answers zero, which is what
`@pptx-studio/text`'s own probe does with an unusable one.

The rule is measured, not assumed: four probe fonts whose `DFLT`, `latn` and `hani`
coordinates all disagree settle it at 30 of 30 in `corpus/ground-truth/font-metrics.json`.
Chromium reads `DFLT` and no other script — a `BASE` table naming only `latn`, `hani`
and `kana` is ignored whole — and it reads the same coordinate for a Latin run as for a
Han one, so the answer is a property of the face rather than of the text.

A coordinate outside the em is handed over as written rather than clamped — Chromium
reported 1.2 em for a font that says so — so both engines see the same unusable number
and both answer zero.
