# Legal basis and clean-room record

This file exists so that the licensing posture of PPTX Studio is auditable rather than assumed.
It is maintained as work happens, not written at release time.

## Why Apache-2.0 and not MIT

PPTX Studio reimplements a patent-encumbered file format family. Apache-2.0 §3 grants an explicit,
irrevocable patent licence from every contributor to every user; MIT grants none at all. For a
project whose entire surface area is format reimplementation, that difference is the licence
choice.

## Reimplementation basis

The file formats this project reads and writes are specified by:

- **ECMA-376, _Office Open XML File Formats_** (Editions 1–5) — Parts 1 (Fundamentals and Markup
  Language Reference), 2 (Open Packaging Conventions), 3 (Markup Compatibility and Extensibility)
  and 4 (Transitional Migration Features).
- **ISO/IEC 29500** — the ISO/IEC publication of the same material.
- **`[MS-OE376]`**, _Office Implementation Information for ECMA-376 Standards Support_, for the
  documented places where Microsoft Office's behaviour differs from the standard text.
- **`[MS-PPTX]`** and the other Microsoft Open Specifications documents for extension namespaces
  (`p14`, `p15`, `a14`, `am3d`, `asvg`, `adec`, …) that appear inside `mc:AlternateContent` and
  `extLst`.

Microsoft's **Open Specification Promise** (`[MS-DEVCENTLP]`) is a covenant not to assert
Microsoft's necessary patent claims against implementations of the covered specifications, of which
the Office Open XML file formats are one. Two limits on that promise are worth stating plainly
because they shape how this repository is run:

1. It covers **conforming** implementations, and only to the extent they conform. Where we
   deliberately deviate from the standard text to match Office's actual behaviour, we say so in a
   comment next to the deviation and cite `[MS-OE376]` where it documents that behaviour.
2. It is a promise from **Microsoft**, about **Microsoft's** patents. It says nothing about any
   third party's. That is the gap Apache-2.0 §3 closes for contributions to this repository.

Nothing here is legal advice, and none of it is a substitute for your own counsel.

## Clean-room record

Reading a permissively-licensed implementation to understand a format is normal engineering.
Reading a copyleft implementation and then writing something structurally identical is not. This
table records which category each project we consult falls into, and it is updated **before** the
code that consults it is written.

| Project                          | Licence       | How we use it                                                                                                                                                                                                                                                                                                                                       |
| -------------------------------- | ------------- | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Apache POI                       | Apache-2.0    | **Transliteration permitted.** Preset shape definitions are transcoded at build time (2.1); attribution in `NOTICE`.                                                                                                                                                                                                                                |
| python-pptx                      | MIT           | **Transliteration permitted.** Consulted for part/relationship structure.                                                                                                                                                                                                                                                                           |
| PptxGenJS                        | MIT           | **Transliteration permitted.**                                                                                                                                                                                                                                                                                                                      |
| pptx-automizer                   | MIT           | **Transliteration permitted.**                                                                                                                                                                                                                                                                                                                      |
| Mono libgdiplus                  | MIT           | **Transliteration permitted.** Pattern tile geometry (2.7). Deliberately used in preference to Wine, which is LGPL.                                                                                                                                                                                                                                 |
| `fontello/ttf2eot`               | MIT           | **Fork permitted.** Patched EOT writer (8.3).                                                                                                                                                                                                                                                                                                       |
| LibreOffice / `sw`, `oox`, `svx` | MPL-2.0       | **Read for understanding only.** MPL-2.0 is file-level copyleft: any file derived from it would have to stay MPL-2.0 inside an Apache-2.0 tree. Behaviour learned from it is re-derived independently and the independent derivation is recorded. Applies specifically to the 74 built-in table styles (4.2) and the arrowhead vertex tables (2.8). |
| ONLYOFFICE `sdkjs`               | AGPL-3.0-only | **Not read.** Listed here so the boundary is explicit.                                                                                                                                                                                                                                                                                              |
| PPTist                           | AGPL-3.0-only | **Not read.**                                                                                                                                                                                                                                                                                                                                       |

If you contribute a change informed by an MPL, GPL or AGPL codebase, say so in the pull request.
It is not automatically disqualifying, and it is much cheaper to discuss before the merge.

## Corpus licensing

Fixture decks carry third-party copyright and, frequently, licensed embedded fonts. Every entry in
`corpus/manifest.json` must carry a `license` field, and CI fails without one (sub-phase 1.1).
Accepted sources are limited to:

- decks authored by this project,
- CC0 / CC-BY / public-domain material with the attribution recorded,
- public-domain government publications,
- synthetic feature probes generated by our own tooling.

Anything else is committed only as a redacted derivative together with the SHA-256 of the original,
so a report can be reproduced without redistributing the source deck.

## Third-party redistribution

See `NOTICE`. Metric-compatible substitute fonts live in a separate package,
`@pptx-studio/fonts-metric-compat`, under SIL OFL 1.1 — kept out of this tree so that the licence
boundary is a package boundary and OFL's reserved-font-name rules never touch an Apache-2.0
manifest.

**Aptos** is not redistributable and its EULA forbids format conversion as well as redistribution.
It is never bundled and never converted; where a deck asks for it we use the locally installed face
if there is one and otherwise fall back to Carlito, which is what Microsoft's own Aptos → Calibri
fallback amounts to. The substitution is shown in the UI rather than hidden.
