---
'@pptx-studio/model': minor
'@pptx-studio/validate': minor
---

A table's style is the one PowerPoint draws. `BUILTIN_TABLE_STYLES` is PowerPoint's own Table
Styles gallery: 74 styles, each with its GUID, its name and the `a:tblStyle` PowerPoint writes for
it. `tableStyleOf(table)` returns the built-in a table's GUID names, in any case, or `null` for
PowerPoint's default 1-pt black grid; nothing in `ppt/tableStyles.xml` changes which. `builtinTableStyle`
and `parseTableStyle` read a `CT_TableStyle` into its thirteen parts. An inline `TableStyleRef` now
carries its `id`, a style GUID without braces or with padding throws `MODEL_TABLE_ATTR`, and
`MODEL_ERROR_CODES` lists every code, including the new `MODEL_TABLE_STYLE`.

validate adds `V032`, a warning on a table that names no built-in style, and `V033`, refusing the
table-style forms PowerPoint repairs.
