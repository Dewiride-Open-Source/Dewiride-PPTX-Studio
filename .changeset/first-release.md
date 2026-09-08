---
'@pptx-studio/census': minor
'@pptx-studio/cli': minor
'@pptx-studio/geometry': minor
'@pptx-studio/model': minor
'@pptx-studio/opc': minor
'@pptx-studio/paint': minor
'@pptx-studio/render-svg': minor
'@pptx-studio/text': minor
'@pptx-studio/validate': minor
'@pptx-studio/writer': minor
'@pptx-studio/xml': minor
---

First release.

`@pptx-studio/render-svg` turns a slide into an SVG string, and `pptx-studio render`
does it from the command line with no browser and no LibreOffice — text included,
measured out of the font's own tables. The other nine packages are what those two
are built on and are published because they have to be, not because their surface
is settled.

Everything here is pre-1.0 and the API will change. What will not change without a
very good reason is what the writer does to a package it was not asked to edit: a
part nobody touched is re-emitted byte for byte, across all 54 decks in the corpus.
