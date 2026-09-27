# @pptx-studio/opc

## 0.1.3

### Patch Changes

- b58dd19: Built with tsdown 0.23. The declarations take a new shape, and the API is unchanged.

  - Every value is exported where it is declared (`export declare function …`), not in a trailing
    `export { … }`.
  - The types are exported in one `export type { … }`.
  - The JavaScript is byte-for-byte what 0.22 emitted.
  - In geometry, opc, validate and xml, the source maps name fewer symbols.

## 0.1.2

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
