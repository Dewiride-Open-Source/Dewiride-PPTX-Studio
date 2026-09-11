---
'@pptx-studio/cli': minor
---

The face box is read from the table this machine's rasteriser reads.

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
