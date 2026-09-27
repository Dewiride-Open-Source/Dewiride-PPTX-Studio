# @pptx-studio/text

## 0.2.1

### Patch Changes

- cbc60d4: Linear time on hostile input, and a command oracle that cannot be split by the path it is handed.

  - **opc:** `sourcePartNameForRels` scans a name once, where it used to backtrack quadratically on a
    name holding many `/_rels/` segments (CodeQL `js/polynomial-redos`).

    It now also maps a name containing U+2028 or U+2029, which `isRelationshipPartName` already
    accepted. Before, `PartStore.danglingRelationships()`, documented never to throw, threw on such a
    deck, and census threw with it.

  - **text:** `decoratedStretches` finds the end of the rule with one backward scan. A run of spaces
    followed by a glyph no longer costs quadratic time. As before, only U+0020 is dropped; a tab or a
    no-break space still carries the rule.
  - **render-svg:** the same fix for the rules a slide draws. One unwrapped line of 2¹⁷ spaces took
    19 seconds to lay out and now takes milliseconds.
  - **cli:** `bisect --oracle command` puts the candidate path in `PPTX_STUDIO_CANDIDATE` and replaces
    `{}` with a quoted reference to it: `"$PPTX_STUDIO_CANDIDATE"`, or `"%PPTX_STUDIO_CANDIDATE%"` under
    cmd.exe.

    A temporary directory whose name holds a space, `&`, `$` or `%` is no longer split or expanded by
    the shell. **Breaking:** `{}` arrives quoted, so a command that quoted it itself must stop doing
    so.

## 0.2.0

### Minor Changes

- 55f3073: `TextOptions.cssFamilyFor` decides what a run's `font-family` says.

  `renderSlide` wrote the run's own family, quoted, which is right for a renderer that measures and
  draws through the same name and wrong for one that measured elsewhere. The hook is called once per
  piece with the run's family, weight and style; it defaults to the quoted family, so a caller that
  passes nothing emits byte-identical markup. `fontStack` takes the face a run was drawn in as a
  second argument and leads with it, deduplicated as before.

## 0.1.1

### Patch Changes

- 76b505c: Published with a provenance attestation. The first release went out over a token rather than OIDC, and npm attests automatically only on the OIDC path, so the registry recorded no attestation and `npm audit signatures` could not verify these packages. Nothing else about them has changed.

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
