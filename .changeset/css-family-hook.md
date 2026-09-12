---
'@pptx-studio/render-svg': minor
'@pptx-studio/text': minor
---

`TextOptions.cssFamilyFor` decides what a run's `font-family` says.

`renderSlide` wrote the run's own family, quoted, which is right for a renderer that measures and
draws through the same name and wrong for one that measured elsewhere. The hook is called once per
piece with the run's family, weight and style; it defaults to the quoted family, so a caller that
passes nothing emits byte-identical markup. `fontStack` takes the face a run was drawn in as a
second argument and leads with it, deduplicated as before.
