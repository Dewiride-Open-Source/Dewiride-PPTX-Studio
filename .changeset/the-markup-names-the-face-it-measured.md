---
'@pptx-studio/cli': minor
---

`pptx-studio render` writes the face it measured a run in into the markup, and reads the face box
the way the browser reports it.

The SVG said `font-family="Calibri Light"` and nothing else, so a viewer picked its own face while
the line breaks and origins had been computed for whatever stood in — the one place the two
renderers could still disagree about what was drawn, open since ADR 0044. Every run now carries a
list: the face this machine drew it in, then the face the deck asked for, then the substitution
table's entry, then Calibri and Carlito, then `sans-serif`. `Calibri Light` drawn in Carlito reads
`"Carlito", "Calibri Light", "Calibri", sans-serif`; a face the machine has leads its own list.

`FontMeasurer` gains `cssFamilyFor`, answered from the resolved-face cache the measurer already
filled, so it cannot name a face other than the one that measured. A face whose name cannot be
written into a CSS shorthand is no longer claimed by the index, as a blank name already was not.

The face box is rounded half up to a whole pixel at the size the browser reads it, and the
ideographic baseline of a face with no `BASE` table is that rounded descent: T13 on a 2048 and a
2000 em, 28 of 28 and 36 of 36, where the exact fraction scores 25 and 34. ADR 0052.
