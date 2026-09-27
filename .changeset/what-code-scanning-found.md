---
'@pptx-studio/opc': patch
'@pptx-studio/text': patch
'@pptx-studio/render-svg': patch
'@pptx-studio/cli': minor
---

Linear time on hostile input, and a command oracle that cannot be split by the path it is handed.

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
