# @pptx-studio/validate

## 0.3.0

### Minor Changes

- ae2e566: A table's style is the one PowerPoint draws. `BUILTIN_TABLE_STYLES` is PowerPoint's own Table
  Styles gallery: 74 styles, each with its GUID, its name and the `a:tblStyle` PowerPoint writes for
  it. `tableStyleOf(table)` returns the built-in a table's GUID names, in any case, or `null` for
  PowerPoint's default 1-pt black grid; nothing in `ppt/tableStyles.xml` changes which. `builtinTableStyle`
  and `parseTableStyle` read a `CT_TableStyle` into its thirteen parts. An inline `TableStyleRef` now
  carries its `id`, a style GUID without braces or with padding throws `MODEL_TABLE_ATTR`, and
  `MODEL_ERROR_CODES` lists every code, including the new `MODEL_TABLE_STYLE`.

  validate adds `V032`, a warning on a table that names no built-in style, and `V033`, refusing the
  table-style forms PowerPoint repairs.

## 0.2.1

### Patch Changes

- b58dd19: Built with tsdown 0.23. The declarations take a new shape, and the API is unchanged.

  - Every value is exported where it is declared (`export declare function …`), not in a trailing
    `export { … }`.
  - The types are exported in one `export type { … }`.
  - The JavaScript is byte-for-byte what 0.22 emitted.
  - In geometry, opc, validate and xml, the source maps name fewer symbols.

- Updated dependencies [b58dd19]
  - @pptx-studio/opc@0.1.3
  - @pptx-studio/xml@0.1.2

## 0.2.0

### Minor Changes

- e8ba475: A table is parsed, and its grid is the one PowerPoint draws.

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

## 0.1.2

### Patch Changes

- 88cdb45: V028 no longer mistakes the extent of a transform for an extension.

  The rule holds `mc:AlternateContent` and every `ext` inside an `extLst` opaque, by local name so
  that vocabularies never read are covered too. It matched any element spelled `ext`, and the
  extent inside `a:xfrm` is spelled `ext` as well - so a placeholder given a transform of its own,
  which is what PowerPoint writes when one is dragged, was refused as a rebuilt extension. An `ext`
  is now an extension only under an `extLst`.

## 0.1.1

### Patch Changes

- 76b505c: Published with a provenance attestation. The first release went out over a token rather than OIDC, and npm attests automatically only on the OIDC path, so the registry recorded no attestation and `npm audit signatures` could not verify these packages. Nothing else about them has changed.
- Updated dependencies [76b505c]
  - @pptx-studio/opc@0.1.1
  - @pptx-studio/xml@0.1.1

## 0.1.0

### Minor Changes

- 7d905f4: First release.

  `@pptx-studio/render-svg` turns a slide into an SVG string, and `pptx-studio render`
  does it from the command line with no browser and no LibreOffice — text included,
  measured out of the font's own tables. The other nine packages are what those two
  are built on and are published because they have to be, not because their surface
  is settled.

  Everything here is pre-1.0 and the API will change. What will not change without a
  very good reason is what the writer does to a package it was not asked to edit: a
  part nobody touched is re-emitted byte for byte, across all 54 decks in the corpus.

### Patch Changes

- Updated dependencies [7d905f4]
  - @pptx-studio/opc@0.1.0
  - @pptx-studio/xml@0.1.0
