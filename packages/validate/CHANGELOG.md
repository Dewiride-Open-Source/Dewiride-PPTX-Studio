# @pptx-studio/validate

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
