# @pptx-studio/render-svg

## 0.2.0

### Minor Changes

- 55f3073: `TextOptions.cssFamilyFor` decides what a run's `font-family` says.

  `renderSlide` wrote the run's own family, quoted, which is right for a renderer that measures and
  draws through the same name and wrong for one that measured elsewhere. The hook is called once per
  piece with the run's family, weight and style; it defaults to the quoted family, so a caller that
  passes nothing emits byte-identical markup. `fontStack` takes the face a run was drawn in as a
  second argument and leads with it, deduplicated as before.

### Patch Changes

- Updated dependencies [55f3073]
  - @pptx-studio/text@0.2.0

## 0.1.1

### Patch Changes

- 76b505c: Published with a provenance attestation. The first release went out over a token rather than OIDC, and npm attests automatically only on the OIDC path, so the registry recorded no attestation and `npm audit signatures` could not verify these packages. Nothing else about them has changed.
- Updated dependencies [76b505c]
  - @pptx-studio/geometry@0.1.1
  - @pptx-studio/model@0.1.1
  - @pptx-studio/paint@0.1.1
  - @pptx-studio/text@0.1.1

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
  - @pptx-studio/geometry@0.1.0
  - @pptx-studio/model@0.1.0
  - @pptx-studio/paint@0.1.0
  - @pptx-studio/text@0.1.0
