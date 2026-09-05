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

### The ECMA-376 schemas are read, not redistributed

`packages/xml/src/edit/schema-order.gen.ts` is generated from the Transitional XML schemas published as
an electronic insert to **ECMA-376 Part 4, 5th edition (December 2016)**, distributed by Ecma
International as `OfficeOpenXML-XMLSchema-Transitional.zip`.

**The schemas themselves are not in this repository.** `tools/schema-codegen/generate.ts` reads them
from a directory named on the command line, and what is committed is the child-ordering facts
extracted from them — which elements a given element admits, and in what order — together with the
SHA-256 of every input file, so a contributor can verify they are regenerating from the same
material. That keeps the licence question about Ecma's own files from arising at all, and it is the
same posture as the corpus: cite the source, commit the derivation.

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
every `corpus/**/manifest.json` must carry a `license` field, and CI fails without one. Accepted
sources are limited to:

- decks authored by this project,
- CC0 / CC-BY / public-domain material with the attribution recorded,
- public-domain government publications,
- synthetic feature probes generated by our own tooling.

Anything else is committed only as a redacted derivative together with the SHA-256 of the original,
so a report can be reproduced without redistributing the source deck.

### How that is enforced

`pnpm corpus` (`tools/corpus/`) runs second in `pnpm check`, after layering and **before**
formatting — a run that stops on whitespace is a run that never reports that a fixture has no
licence. It reads every `corpus/**/manifest.json` and asserts, among other rules:

- `C001` a `license` from a closed set. `NONE` and `NOASSERTION` are rejected **by name**: an entry
  whose licence is unknown is an entry that cannot be redistributed, and accepting SPDX's spellings
  for "nobody established this" would make the rule a formality.
- `C003` the licence is one the **source** may claim. This is the rule that does the legal work.
  `C001` is satisfied by typing six characters; `C003` is what says a CC-BY deck somebody else wrote
  is not ours to relicense as Apache-2.0.
- `C004` a complete attribution block for anything third-party, per CC BY 4.0 §3(a)(1).
- `C008` **every file under `corpus/` is claimed by exactly one manifest.** This is the only rule
  that fires in the situation that actually occurs — a deck dropped into the tree without anyone
  touching a manifest. It enumerates with `readdirSync` rather than `git ls-files` specifically so
  that it sees untracked files, and so a local `pnpm check` catches the deck before it is committed.
- `C010` the committed bytes hash to what the manifest says. Without it a `license` field describes
  whatever happens to be at that path on the day of the audit.
- `C012` size caps, so that a fixture too large to review is a recipe and a hash rather than bytes.
- `C015` a redacted derivative names its original's SHA-256, and that hash is not its own.

The full rule table is `tools/corpus/manifest/schema.ts`. If that file and this section ever disagree, this
section is the one that is wrong, because that file is the one a contributor runs into.

### Sub-phase 0.7's fixtures, and what was deliberately left out

`corpus/ground-truth/` is written to that rule already, ahead of CI enforcing it. Two entries are
worth spelling out because they are the awkward cases:

- **`eot-headers.json`** records the EOT header of 76 `ppt/fonts/*.fntdata` parts that PowerPoint
  wrote — flags, version, weight, `fsType`, PANOSE, family name. The decks themselves are **not**
  committed and neither is a single byte of font data: those are Microsoft's fonts. What is here is
  a measurement _about_ them, in the same sense that a table of file sizes is not a copy of the
  files.
- **`fonts/probe.ttf`** and its EOT wrappers are a font this project authored from nothing, released
  CC0-1.0, existing so that experiment B could test whether PowerPoint renders an embedded font
  without wrapping somebody else's. Its family name is `ProbeAlpha`, which is not a reserved name of
  anything.

### Sub-phase 0.8's benchmark decks

`corpus/bench/` holds recipes and hashes, not decks. The largest deck the
benchmark reads is 188 MiB, which has no business in a git history, so what is
committed is the recipe that produces it, its exact size and its SHA-256 — the
generator is deterministic, so anyone can rebuild the file and check they have
the same one.

Everything in those decks is generated by `tools/bench/deck.ts`: the markup, the
theme, and the images, which are noise rather than photographs. The one thing
that is not generated on the spot is the embedded font, and that is the CC0
`ProbeAlpha` this project authored in sub-phase 0.7. Nothing anyone else owns is
inside a file this project hands around.

### MicroType Express, if Phase 8 needs it

Sub-phase 0.7 established that PowerPoint compresses every embedded font with MicroType Express.
The only free LZCOMP implementation is **libeot** (MPL-2.0). Two notes for whoever picks this up:

- MPL-2.0 is file-level copyleft. It travels with the files, not with the program, so it belongs in
  its own package with its own licence boundary — the same reason `fonts-metric-compat` is a
  separate package.
- **Monotype's MicroType Express Licensing Statement** grants a royalty-free, irrevocable,
  sublicensable licence to the patents and to the sample code. That is what makes this approachable
  at all; without it the patent posture would rule out both using libeot and porting the W3C sample.

Neither has been acted on. This section exists so the question is not rediscovered from scratch.

## Third-party redistribution

See `NOTICE`. Metric-compatible substitute fonts live in a separate package,
`@pptx-studio/fonts-metric-compat`, under SIL OFL 1.1 — kept out of this tree so that the licence
boundary is a package boundary and OFL's reserved-font-name rules never touch an Apache-2.0
manifest.

**Aptos** is not redistributable and its EULA forbids format conversion as well as redistribution.
It is never bundled and never converted; where a deck asks for it we use the locally installed face
if there is one and otherwise fall back to Carlito, which is what Microsoft's own Aptos → Calibri
fallback amounts to. The substitution is shown in the UI rather than hidden.
