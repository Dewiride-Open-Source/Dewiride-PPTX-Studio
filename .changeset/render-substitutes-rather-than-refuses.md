---
'@pptx-studio/cli': patch
---

`renderDeck` substitutes a typeface nothing stands in for rather than refusing to draw.

A deck naming `Calibri Light` on a Linux server holding 53 perfectly good faces threw
`CLI_NO_FACE` and rendered nothing. The chain ran exact family, the substitution table,
Calibri, Carlito — and then gave up, which is the one place the font guard's design was
not applied: `FaceUse` carries `asked`, `drawn` and `substituted` so the answer can be
"Calibri Light was drawn in DejaVu Sans" rather than an exception.

Two steps now follow the measured chain. The shorter forms of the asked-for name, so
`Calibri Light` draws in Calibri and `Segoe UI Semilight` in Segoe UI where the machine
has them — DrawingML has no weight axis, so a style word is part of the typeface name.
Then, past every named candidate, the first family by name draws the run. Both are
reported as substitutions under the name the deck asked for, and neither displaces the
measured chain: ADR 0033 scored PowerPoint's own last resort at 22 of 22 and a guess must
not override a measurement.

The choice is by family name, not by the order `readdir` returned, and the font walk now
sorts its entries — so two machines holding the same faces draw the same picture, where
before the answer was the filesystem's own.

`CLI_NO_FACE` still throws, and now means exactly one thing: no font was found at all,
which is nothing to draw with rather than an unusual typeface.
