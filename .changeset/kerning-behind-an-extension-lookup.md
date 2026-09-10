---
'@pptx-studio/cli': patch
---

Kerning is read through a type 9 GPOS Extension lookup.

`pptx-studio render` measured every face whose `kern` feature sits behind an
`ExtensionPosFormat1` subtable as having no kerning at all — no error, no diagnostic, just
lines about a twelfth too wide, which moves every break and misfires autofit. Lato is such
a face, in all eighteen of the styles a Linux runner carries, and so is most of what a
modern font compiler emits: an extension is how a lookup reaches a subtable past the 64 KB
a 16-bit offset can address, and compilers wrap unconditionally rather than only when they
must.

`sfnt.ts` read `if (u16(r, lookup) !== 2) continue`, so lookup type 9 was skipped along
with mark attachment and cursive positioning, which genuinely do not move an advance. It
now follows a type 9 lookup one level, to the type its `ExtensionPosFormat1` names at the
32-bit offset from its own start, and reads the `PairPos` there. Exactly one level: the
specification forbids an extension targeting another, so nothing recurses, and a subtable
naming any other type is skipped as before.

The rule is measured, not assumed. A twelfth probe font wraps the same `PairPos` the
eleventh carries plainly, and Chromium drew the two identically at all six sizes and on
all six strings — so the browser honours the wrapper and the reader was wrong rather than
strict. Following one extension fits 72 of 72 in
`corpus/ground-truth/font-metrics.json`, alone at the top; the reading that skipped it
fits 66, exactly the six rows of the wrapped probe.
