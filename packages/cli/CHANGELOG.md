# @pptx-studio/cli

## 0.2.0

### Minor Changes

- 725ad6d: `renderDeck` takes only what it draws.

  Its options were the `render` verb's own, so a caller that wanted SVG strings passed
  `out`, `json` and `quiet` — three fields the function ignores — and five more it had no
  opinion about. `RenderDeckOptions` is the shape now: `slide`, `width`, `fontDirs`,
  `systemFonts` and `text`, all optional. `renderDeck(bytes)` is a whole call, and
  `renderDeck(bytes, { slide: 1, width: 640 })` is a thumbnail.

  `RenderOptions` is the verb's: it extends `RenderDeckOptions` with `out`, `json` and
  `quiet`, and `runRender` is what takes it. `RENDER_DEFAULTS` is gone, because the
  defaults are in `renderDeck` and `DEFAULT_WIDTH` is still exported. A width that is not
  a positive whole number now throws `CLI_WIDTH` instead of drawing an empty picture.

  Breaking: a call written against 0.1.0 passes three properties this shape does not have.
  Delete them, and delete any field you were only passing to satisfy the type.

- c1d7d7d: The face box is read from the table this machine's rasteriser reads.

  `pptx-studio render` took a face's ascent and descent from `OS/2.usWinAscent`/`usWinDescent`
  everywhere, or `sTypo` where `fsSelection` bit 7 is set. That is what Chromium answers through
  DirectWrite and it is not what Chromium answers through FreeType, which reads
  `hhea.ascender`/`descender`. Neither reading fits both: the first scores 12/12 against Windows and
  76/79 against Linux, the second 79/79 against Linux and 7/12 against Windows, and the misses are
  75.8 px and 100 px on a 1000 px em rather than roundings.

  On a Linux server — which is where this verb runs — the old reading put the IPA Gothic descent at
  195.8 px where the browser puts it at 120, 7.6% of the em on every line of Japanese text, moving
  every line break and every autofit decision.

  `indexFonts` now resolves a `FontBackend` from its platform, and `metricsOf` reads `hhea` under
  FreeType and `usWin` under DirectWrite, with bit 7 moving both onto `sTypo`. `FaceMetrics.source`
  gains `hhea`, which is also what a face carrying no `OS/2` now reports rather than claiming `usWin`
  while returning `hhea`'s numbers. macOS takes the DirectWrite reading and is unmeasured.

  **Breaking:** `facesIn` and `faceOf` take the backend as a third argument. There is no defensible
  default — a silent one would be a guess about the host — and `backendFor(platform)` is exported for
  callers that want the local answer.

  The rule is measured, not assumed: 79 real faces on a Linux runner against the 12 probe fonts T13
  built to disagree with themselves, both scored in
  `docs/adr/phase-3-text/0045-the-face-box-belongs-to-the-rasteriser.md`.

### Patch Changes

- 725ad6d: The ideographic baseline comes out of the font's `BASE` table.

  `pptx-studio render` placed every upright East Asian glyph on the face's box descent,
  which is right only for a face carrying no `BASE` table. Yu Gothic reports a box descent
  of 0.302 em and an ideographic baseline of 0.1201 em, so a vertical `eaVert` frame was
  drawn 0.18 em out of place on it — and the browser renderer, which asks Chromium, was
  never wrong about the same face.

  `sfnt.ts` now reads the `BASE` horizontal axis' `ideo` coordinate under the `DFLT`
  script, and `FaceMetrics` carries it as `ideographic`. Where there is none the box
  descent stands, and a coordinate outside the em answers zero, which is what
  `@pptx-studio/text`'s own probe does with an unusable one.

  The rule is measured, not assumed: four probe fonts whose `DFLT`, `latn` and `hani`
  coordinates all disagree settle it at 30 of 30 in `corpus/ground-truth/font-metrics.json`.
  Chromium reads `DFLT` and no other script — a `BASE` table naming only `latn`, `hani`
  and `kana` is ignored whole — and it reads the same coordinate for a Latin run as for a
  Han one, so the answer is a property of the face rather than of the text.

  A coordinate outside the em is handed over as written rather than clamped — Chromium
  reported 1.2 em for a font that says so — so both engines see the same unusable number
  and both answer zero.

- cec9ef2: Kerning is read through a type 9 GPOS Extension lookup.

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

- cec9ef2: `renderDeck` substitutes a typeface nothing stands in for rather than refusing to draw.

  A deck naming `Calibri Light` on a Linux server holding 53 perfectly good faces threw
  `CLI_NO_FACE` and rendered nothing. The chain ran exact family, the substitution table,
  Calibri, Carlito — and then gave up, which is the one place the font guard's design was
  not applied: `FaceUse` carries `asked`, `drawn` and `substituted` so the answer can be
  "Calibri Light was drawn in DejaVu Sans" rather than an exception.

  Two steps now follow the measured chain. The shorter forms of the asked-for name, so
  `Calibri Light` draws in Calibri and `Segoe UI Semilight` in Segoe UI where the machine
  has them — DrawingML has no weight axis, so a style word is part of the typeface name.
  Then, past every named candidate, the first family by name draws the run. Both are
  reported as substitutions under the name the deck asked for, and neither displaces the
  measured chain: ADR 0033 scored PowerPoint's own last resort at 22 of 22 and a guess must
  not override a measurement.

  The choice is by family name, not by the order `readdir` returned, and the font walk now
  sorts its entries — so two machines holding the same faces draw the same picture, where
  before the answer was the filesystem's own.

  `CLI_NO_FACE` still throws, and now means exactly one thing: no font was found at all,
  which is nothing to draw with rather than an unusual typeface.

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
  - @pptx-studio/census@0.1.0
  - @pptx-studio/model@0.1.0
  - @pptx-studio/opc@0.1.0
  - @pptx-studio/render-svg@0.1.0
  - @pptx-studio/text@0.1.0
  - @pptx-studio/validate@0.1.0
  - @pptx-studio/writer@0.1.0
  - @pptx-studio/xml@0.1.0
