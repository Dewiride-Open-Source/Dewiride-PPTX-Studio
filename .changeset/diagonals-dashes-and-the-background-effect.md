---
'@pptx-studio/model': minor
'@pptx-studio/validate': minor
---

`tableBackground` now returns the table background's two layers, `{ fill, effects }`, each taken
from `a:tblPr` where it writes one and from the style's `tblBg` otherwise. An empty `a:effectLst` in
`a:tblPr` removes the style's effect, as PowerPoint draws it, and `themedEffects` follows an
`effectRef` into the theme. `tableCellDiagonals` gives a cell's diagonals as drawn: the anchor's
own `a:lnTlToBr` and `a:lnBlToTr` across its whole span, a covered position's never, and swapped in a
right-to-left table.

validate's `V034` also warns on every line a covered cell writes, which PowerPoint keeps on save
and never draws, and on a border written under a position no `a:tc` reached.
