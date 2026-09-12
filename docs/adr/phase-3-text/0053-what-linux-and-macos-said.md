# 0053 — What Linux and macOS said

**Sub-phase 3.10 · accepted · closes open question 1 of
[0052](0052-the-box-rounds-half-up-and-the-markup-names-its-face.md) and open question 1 of
[0045](0045-the-face-box-belongs-to-the-rasteriser.md); revises 0052's rounding for one rasteriser**

[0052](0052-the-box-rounds-half-up-and-the-markup-names-its-face.md) put T13's probe fonts in front
of the two real-faces jobs and asked macOS by dispatch. This is what the runs said, and what changed
because of it: the probes were being autohinted, CoreText reads a table the reader had assumed it
did not, and it rounds the box in a way no other rasteriser here does.

## The probes were autohinted

Run 34688374783, the first with the probes indexed, settled the box question on Linux and failed
the width gate on 25 of 2916 rows, every one on a probe's `latin-short` sample at 16 or 32 px: 87 px
where the reader says 92.8 at 16 — nine advances of 9.6 and one of 6.4 truncated, the delta 5.8 to
the digit — and 193 against 185.6 at 32, every advance a pixel wide. FreeType autohints any SFNT
with no `fpgm` and a `prep` of seven bytes or fewer, whatever its glyphs carry (`FT_Load_Glyph`,
`ftobjs.c`), and the autohinter's normal mode recomputes the advance from the hinted phantom points.
Every real face on the runner carries a font program and takes the native path, where the advance is
the whole-pixel quantum T14 models.

The builder now writes a one-instruction `fpgm` (`PUSHB[0] 0; POP`) into every probe, so they are
hinted the way the real faces are. T13 re-measured on this machine is identical apart from the font
bytes, and the Linux job is green on run 34689472607.

## What Linux said about the box

The FreeType composite fits 93/93 with the fifteen probes in the set: `hhea` alone 84, the
DirectWrite composite 83, `usWin` 71, `sTypo` 35. `split` under FreeType reads `hhea`, which is
what [0044](0044-what-the-runners-own-fonts-said.md) had never measured. Half up and half to even
both fit 93/93, because no face on that runner puts a box on the half; floor and ceil 55, exact 51,
and Blink's borrow-from-the-ascent 76, refuted on the whole-pixel runner as it was on DirectWrite.
`split-2560` below puts `hhea` on the half, and the run this record ships with is the first to
score it there.

## What macOS said

Run 34689472437 indexed 247 faces of 796 in 382 files (127 collections skipped, four with no
measurable glyph, four the browser refused) and failed three gates, as predicted: `advance-quantum`,
`width` and `face-box-rival`. The last is the finding.

**CoreText reads `hhea`, and ignores `fsSelection` bit 7.** `hhea.ascender/descender` fits 247/247
within half a pixel; the FreeType composite 246/247; the DirectWrite composite the reader had
assumed, 210/247, differing on 37. The one face separating `hhea` from the composite is
`split-usetypo` — the probe built with bit 7 set and `hhea` 800/200 apart from `sTypo` 700/100 —
and CoreText reported 800/200. Not one of the 233 real faces there could have asked that question:
eleven carry bit 7 and on every one of them `hhea` and `sTypo` agree.

**CoreText holds the em ratio in single precision before it rounds.** Rescored against `hhea` for
equality, not tolerance: round half up 246/247, half to even 245, the borrow 206, ceil 184, floor
183, exact 159. Three faces sit on the half at 1000 px, and they disagree with each other:

| face                   | `hhea` at 1000 px | browser | as a 32-bit float        |
| ---------------------- | ----------------- | ------- | ------------------------ |
| Hoefler Text Ornaments | 806.5             | 807     | 1613/2000 rounds **up**  |
| Zapfino                | 1502.5            | 1503    | 601/400 rounds **up**    |
| Krungthep              | 262.5             | 262     | 672/2560 rounds **down** |

`floor(fround(units / unitsPerEm) × 1000 + 0.5)` fits all 247 exactly, and it is the only model
here that does. The same arithmetic on DirectWrite is refuted by T13 itself: `split-2000`'s 0.9285
is below the half as a float, and DirectWrite rounded it up, 929. So T13 gained `split-2560`, whose
`usWin` 2336/672 puts both sides on the half with the float going opposite ways: DirectWrite
answered **913/263**, plain half up, 30/30 against 28/30 for the single-precision model. Its `hhea`
2464/544 does the same for the two rasterisers that read `hhea`, predicting 962/213 in single
precision against 963/213 half up and 962/212 half to even.

**Widths on macOS.** The quantum is 1/65536 px on all 3540 rows, DirectWrite's, not the whole
pixel T14's Linux gate models — which is why `advance-quantum` is red there and stays so. The
shipped reading is exact on 2113 of 2928 gated rows and on 165 of 247 faces on every gated row. The
82 others fall four ways:

- **28 within half a pixel everywhere** — a last-bit difference per glyph or per kern, largest on
  ems that are not a power of two (STIX, Bodoni, Kefa, Luminari).
- **26 off only on `latin-long` at 100 or 1000 px**, by 0.5 to 5 px over 450 glyphs — about 2e-6
  of the width, the size of a sum kept in single precision (Georgia, Verdana, Courier New, Geneva).
- **17 where the browser kerns less than the reader** on the pair samples and agrees to the bit
  wherever nothing kerns: Apple's builds of Times New Roman, Trebuchet MS, Arial Narrow, Arial
  Black, Impact, Tahoma, Chalkduster, Herculanum, DIN Alternate. Which kerning source CoreText's
  shaping path read is unmeasured.
- **11 Apple system faces** (`.SF`, `.New York`, Skia) wider than their tables at small sizes and on
  them at 1000 px — `.New York` 13.5% at 8 px, 5.9% at 100, exact at 1000 — which is optical
  tracking (`trak`), a table the reader does not read.

None of the four is a defect in what the reader claims: it reads the tables DirectWrite reads, and
on macOS 165 faces agree with it to the 1/65536. Each is a rule a CoreText-exact reader would need,
and each is on the record with its count.

## What changed in the reader

`FontBackend` gains `coretext`, `backendFor('darwin')` names it, and `metricsOf` reads `hhea` through
it whatever bit 7 says. `boxPixels` rounds per backend: the exact ratio through DirectWrite and
FreeType, the ratio as a 32-bit float through CoreText, half up in both. `FaceMetrics` carries the
backend so the measurer needs no second lookup. T14's `shippedBoxFor('darwin')` is `hhea` alone,
`rasteriserOf` says CoreText, and `BOX_ROUNDINGS` takes units and the em, so the single-precision
model can be one of them; T13 scores 28 models over 30 boxes and the T14 variants table seven.

`packages/text`'s ideographic fallback is answered at `FACE_BOX_PX` and divided, because that is the
one size the engine probes at: the browser rounds the fallback per size (26 at 100 px and 263 at
1000 for 672/2560), and no single fraction of the em is both. The reader answers the 1000 px one.

## Verification

`pnpm check` green. T13 rescored over fifteen probes: face box 30/30, kerning 90/90, widths 540/540,
ideographic 38/38. CI: run 34688374783 on `main` at 55f3073 — `probe-fonts` green on its first
Windows run, `candidate` green on its first, `real-faces` red on the autohinter; run 34689472607 on
this branch green everywhere; run 34689472437 the macOS dispatch above.

| mutant                                        | killed by                                                          |
| --------------------------------------------- | ------------------------------------------------------------------ |
| `backendFor('darwin')` returns `directwrite`  | names FreeType on Linux, CoreText on macOS and DirectWrite else    |
| CoreText follows bit 7 onto `sTypo`           | reads hhea through CoreText whatever bit 7 says                    |
| `boxPixels` drops the `fround` for CoreText   | rounds the box as CoreText does when indexed for macOS             |
| `boxPixels` applies `fround` to every backend | keeps the em ratio in double precision through DirectWrite         |
| `split-2560` dropped from the probes          | the `probe-fonts` diff of the regenerated fixture, not a unit test |
| `shippedBoxFor('darwin')` stays DirectWrite   | scores macOS against hhea alone                                    |
| the single-precision model dropped from T14   | separates the single-precision ratio from half up                  |
| `fpgm` removed from the builder               | the Linux `real-faces` job, run 34688374783                        |

## Deviations

1. **A probe added after the plan.** `split-2560` exists because the macOS run produced a model
   with a prediction DirectWrite could test in one font, and not testing it would have left the
   rounding rule ambiguous across the two rasterisers that round differently.
2. **T14's gates are not made platform-aware.** On macOS `advance-quantum` fails on the quantum
   itself and `width` on the four classes above; the job is dispatch-only and gates nothing, and a
   green macOS run would need the four rules, not a looser gate.
3. **The ideographic test asks one size.** It had asserted a size-independent fraction against
   both baseline sizes, which `split-2560` shows cannot hold.

## Open questions

1. **FreeType on the half.** Half up and half to even tied 93/93 on the runner's faces; the run this
   record ships with scores `split-2560`'s `hhea` 962.5/212.5 under FreeType for the first time.
2. **The four macOS width classes** — the last-bit quantiser on odd ems, the single-precision sum,
   Apple's kerning source, and `trak` — each unmeasured beyond its count.
3. **Which of Skia's CoreText port and CoreText itself holds the ratio as a float** is not something
   a browser measurement can separate; the model is the product's concern and it fits 247/247.
