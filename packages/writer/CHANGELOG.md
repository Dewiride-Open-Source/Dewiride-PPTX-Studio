# @pptx-studio/writer

## 0.1.2

### Patch Changes

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

- Updated dependencies [e8ba475]
  - @pptx-studio/validate@0.2.0

## 0.1.1

### Patch Changes

- 76b505c: Published with a provenance attestation. The first release went out over a token rather than OIDC, and npm attests automatically only on the OIDC path, so the registry recorded no attestation and `npm audit signatures` could not verify these packages. Nothing else about them has changed.
- Updated dependencies [76b505c]
  - @pptx-studio/opc@0.1.1
  - @pptx-studio/validate@0.1.1
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
  - @pptx-studio/validate@0.1.0
  - @pptx-studio/xml@0.1.0
