---
'@pptx-studio/model': minor
'@pptx-studio/validate': minor
---

A table cell's fill, borders and text now come from its style the way PowerPoint draws them.
`tablePartsAt` names the parts that reach a grid position, and `TABLE_PART_ORDER` is the order they
compose in. `tableCellFill`, `tableBackground` and `tableEdgeLine` resolve one property each and
name where it came from: a part, `tblBg`, `a:tcPr`, `a:tblPr` or the default grid. Merged cells and
the owner of every edge segment follow what PowerPoint drew. `tableTextLayer` gives the style's
bold, italic, faces and colour as run properties for the text cascade. `DEFAULT_GRID_LINE` is the 1-pt black line a table with no
built-in style draws. `TextContext.cell` puts a cell's list style and its style's text layer into
the text cascade, and `Origin` gains `cell` and `tableStyle`. A position or edge a table does not
have throws the new `MODEL_TABLE_POSITION`.

validate adds `V034`, a warning on a table cell that writes a border PowerPoint draws from the cell
across the edge instead, where that cell writes none or a different one.
