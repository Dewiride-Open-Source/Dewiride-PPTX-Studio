# 0045 — The face box belongs to the rasteriser

**Sub-phase 3.10 · accepted · supersedes defect 3 and open questions 1 and 2 of
[0044](0044-what-the-runners-own-fonts-said.md), and re-derives its width tolerance**

[0044](0044-what-the-runners-own-fonts-said.md) fixed two of the three defects T14 found and left the
third — the face box 75.8 px out on the descent of the IPA Gothic faces — open, with
`FaceMetrics.candidates` added so the next run could score the rival readings instead of reporting an
unexplained miss. That run has happened, on `cec9ef2`. It answered, and the answer is that the
question was malformed: there is no single table the browser reads.

79 faces from `/usr/share/fonts` on `ubuntu24` image `20260907.300.1`, Chromium 151.0.7922.34, 2832
gated comparisons. One file skipped, `wqy-zenhei.ttc`, for being a collection.

## Neither reading fits both operating systems

Scored over the faces whose tables carry the pair each reading names, fitting within the 0.5 px one
rounding can cost:

| reading                                          | Windows, T13, 12 probes | Linux, T14, 79 real faces |
| ------------------------------------------------ | ----------------------- | ------------------------- |
| `hhea.ascender/descender`                        | 6/12                    | 71/79                     |
| `OS/2.usWinAscent/usWinDescent`                  | 11/12                   | 65/79                     |
| `OS/2.sTypoAscender/sTypoDescender`              | 7/12                    | 28/79                     |
| **`usWin`, or `sTypo` when `fsSelection` bit 7** | **12/12**, worst 0 px   | 76/79, worst 75.8 px      |
| **`hhea`, or `sTypo` when `fsSelection` bit 7**  | 7/12, worst 100 px      | **79/79**, worst 0.48 px  |

The Windows column is computed from `corpus/ground-truth/font-metrics.json` — the same committed
bytes T13 measured, re-scored for the reading T14 turned up. Both rows are perfect on one platform
and wrong on the other, and the misses are 100 px and 75.8 px on a 1000 px em, not roundings.

So **the shipped reader was not wrong; it was answering for one rasteriser while running on the
other.** DirectWrite reads `usWin`, FreeType reads `hhea`, and bit 7 moves both onto `sTypo`. The
same split already governs the advance quantum — whole pixels under FreeType, 1/65536 px under
DirectWrite, in 0044 — so this is the second face of one fact rather than a new one.

Three faces separate the two readings on Linux: `ipag.ttf`, `ipagp.ttf` and
`fonts-japanese-gothic.ttf` — two font files, since the first and third share SHA-256
`503af4a8…5a55e`. `ipag.ttf` is 1802/401 under `usWin` and 1802/246 under `hhea` on a 2048 em, and
the browser reports 880/120 px, which is `hhea` to within half a pixel. The eight faces where raw
`hhea` alone misses are the ones setting bit 7, which is why the composite fits 79 where the bare
table fits 71.

Two rivals are refuted outright. The browser does **not** cap the box at the em: 56 of the 79 faces
report a box larger than their em and the reader agrees with all 56, DejaVu Sans at 1164 px against
`usWin`'s 1164.1. And Linux does not refuse `usWin` generally — it agrees with it on the 68 faces
whose `hhea` and `usWin` say the same thing.

### What changed

`metricsOf` takes a `FontBackend`, and `indexFonts` resolves one from the platform whose fonts it is
reading. `FaceMetrics.source` gains `hhea`, which is also now what a face carrying no `OS/2` reports
rather than the `usWin` it used to claim while returning `hhea`'s numbers. `facesIn` and `faceOf`
take the backend explicitly: there is no defensible default, and a silent one would be a guess about
the machine.

macOS is **unmeasured**. It takes the DirectWrite reading, which is an assumption stated here rather
than a measurement, and open question 1 below.

## The width tolerance priced one rounding where the browser applies two

0044 derived the tolerance as `glyphs × (quantum/2 + reader step)`, from `W_browser = Σ round(aᵢ)`.
Six widths on three Lato faces stayed outside it. They are not a kerning defect: all 18 Lato weights
sit the same side of the browser by the same proportion, and the reader's kerning total is −7.53% to
−7.73% across every one of them. What separates the three is only that they cross the threshold.

The bias is systematic, and per glyph it exceeds what a single rounding can cost:

| sample             | glyphs | reader kerning | mean bias px/glyph | max px/glyph |
| ------------------ | -----: | -------------: | -----------------: | -----------: |
| `latin-short`      |     10 |         0.165% |             0.2611 |       0.4500 |
| `latin-spread`     |     40 |         0.282% |             0.1958 |       0.4250 |
| `cyrillic`         |     55 |         0.544% |             0.2687 |       0.3636 |
| `greek`            |     33 |         0.727% |             0.3249 |       0.4242 |
| `latin-kern-pairs` |     44 |         7.580% |             0.3996 |   **0.6477** |
| `latin-long`       |   1349 |         7.632% |             0.3982 |   **0.6553** |

Over the 18 faces with a 2000 unit em, at 1000 px. **A joint rounding of advance and adjustment is
refuted**: `round(aᵢ + kᵢ)` cannot cost more than half a pixel a glyph, and two samples cost 0.65.
The excess appears only where the kerning does — every low-kern sample stays under 0.45 — so the
browser rounds the advance and the kern adjustment separately, and a run carries one rounding per
glyph plus one per applied kern.

The em is what makes this visible at all, and the parity confirms the mechanism:

| units per em | faces | mean bias px/glyph | max    |
| ------------ | ----: | -----------------: | ------ |
| 64           |     4 |             0.0000 | 0.0000 |
| 1000         |    13 |             0.0000 | 0.0000 |
| 2048         |    40 |            −0.0299 | 0.1345 |
| 2000         |    18 |             0.3996 | 0.6477 |

At 1000 px a 1000 em scales by exactly 1 and there is nothing to round, so the reader is exact. A
2048 em spreads fractions across the interval and the roundings cancel. A 2000 em puts every value on
`.0` or `.5` exactly, where rounding half away from zero always pushes up — which is why Lato, and
only Lato, accumulates.

`widthTolerance(glyphs, kerns)` is now `(glyphs + kerns) × quantum/2 + glyphs × reader step`, and the
measurement records the kern count per row. **This is a correction to a derivation, not a loosened
gate**: a row with no kerns keeps exactly the old bound, a gated row that carries no kern count fails
as unmeasured rather than defaulting to the generous reading, and `unkerned` still misses by 7.6% —
some forty times the tolerance — so the kerning path stays separable.

## Verification

`pnpm check` green. 8 mutants tried on the two changes, 8 killed; one — a face with no `OS/2`
reporting `usWin` again — survived the first sweep as a missing assertion and was killed by adding
one, not by being called equivalent.

| mutant                                  | killed by                                                      |
| --------------------------------------- | -------------------------------------------------------------- |
| `backendFor` always answers DirectWrite | names FreeType on Linux and DirectWrite everywhere else        |
| FreeType reads `usWin`                  | reads hhea through FreeType and usWin through DirectWrite      |
| bit 7 ignored under FreeType            | follows fsSelection bit 7 onto sTypo through either rasteriser |
| a face with no `OS/2` reports `usWin`   | names hhea alone for a face carrying no OS/2 table             |
| the tolerance drops the kern term       | adds one browser rounding per kern the reader applied          |
| the tolerance doubles the kern term     | the same assertion, at equality                                |
| the shipped box ignores the platform    | A6, which decides the reading a rival is measured against      |
| a missing kern count is read as zero    | A2 fails a gated row that never counted its kerns              |

The reader's own numbers are unchanged on Windows, where T13 measured them:
`packages/cli/src/render/measure.test.ts` pins `indexFonts` to `platform: 'win32'` and the T13
assertions read faces with `directwrite` explicitly, so the fixture is scored against the rasteriser
it was recorded on whichever machine the suite runs on.

## Deviations

`facesIn` and `faceOf` are exported, so requiring the backend is a breaking change to the public API.
Pre-1.0 with the release not yet out, that is free, and the alternative — an optional parameter
defaulting to one rasteriser — would put a guess about the host in the one place this record exists
to say cannot be guessed.

## Open questions

1. **macOS is unmeasured.** CoreText is neither of the two backends measured here and takes the
   DirectWrite reading by assumption. The experiment is the one that already exists: run T14 on a
   macOS runner and read the table.
2. **Which rounding the face box takes** stays open, for the reason 0044 gave: a whole-pixel box is
   reproduced by `round` whatever the reader answers. Only DirectWrite can separate them and T13 did
   not ask.
3. **The reader could be exact on Linux, and is not.** The 1000-em faces reproduce the browser to the
   bit at 1000 px, which shows the browser scales linearly and rounds rather than hinting. A reader
   that quantised to its host's advance quantum would match FreeType exactly instead of being bounded
   by a tolerance. That is a new capability rather than a defect in what shipped, and it needs
   measuring on both backends before the reader takes it.
4. **A deck rendered on one platform and viewed on another** still has two answers. The CLI bakes
   positions into SVG on the machine it runs on, so a Linux server and a Windows viewer disagree by
   whatever their rasterisers disagree by. Nothing here fixes that; it names it.
