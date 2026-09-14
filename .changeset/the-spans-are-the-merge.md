---
'@pptx-studio/model': minor
'@pptx-studio/validate': minor
'@pptx-studio/render-svg': patch
'@pptx-studio/render-dom': patch
'@pptx-studio/writer': patch
'@pptx-studio/cli': patch
---

A table is parsed, and its grid is the one PowerPoint draws.

`Shape.table` carries the `a:tbl` of a `p:graphicFrame`: the `a:tblPr` flags and style source, every
`a:gridCol`, every `a:tr` with its cells, and each cell's spans, body and `a:tcPr` as written.
`tableGrid` resolves the merges by the rule C7 measured (`tables.json`, 76 of 76 tables PowerPoint
opened as written and 21 it authored itself): each `a:tc` takes the next column, the spans on the
anchor are the whole story, a covered position is covered whatever its own attributes say, a span
stops at the edge or at the first claimed position, and `hMerge`/`vMerge` change nothing. `a:tr/@h`
is a minimum, a column is never narrower than its cells' side margins plus 2 pt, and the frame's
`a:ext` says nothing about the drawn size. A table attribute of the wrong type throws
`MODEL_TABLE_ATTR`.

`@pptx-studio/validate` gains `V030`, a warning on a table whose rows, spans and flags disagree —
PowerPoint opens it silently and rewrites it on the next save — and `V031`, fatal on the missing or
mistyped table attributes PowerPoint repairs. The downstream packages carry the new rules and the
new field; nothing they draw changes.
