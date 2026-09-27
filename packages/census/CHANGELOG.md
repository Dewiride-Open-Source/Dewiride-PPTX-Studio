# @pptx-studio/census

## 0.1.2

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
