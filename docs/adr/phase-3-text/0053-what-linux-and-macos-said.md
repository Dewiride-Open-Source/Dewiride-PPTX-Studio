# 0053 — What Linux and macOS said

**Sub-phase 3.10 · accepted · closes open question 1 of
[0052](0052-the-box-rounds-half-up-and-the-markup-names-its-face.md) and open question 1 of
[0045](0045-the-face-box-belongs-to-the-rasteriser.md); revises 0052's rounding for one rasteriser**

[0052](0052-the-box-rounds-half-up-and-the-markup-names-its-face.md) put T13's probe fonts in front
of the two real-faces jobs and asked macOS by dispatch. This is what the runs said, and what changed
because of it: the probes were being autohinted, CoreText reads a table the reader had assumed it
did not, and it rounds the box from a number no other rasteriser here holds.

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
bytes, and the Linux job has been green since run 34689472607.

## What Linux said about the box

The FreeType composite fits 94/94 with the fifteen probes in the set: `hhea` alone 85, the
DirectWrite composite 83, `usWin` 71, `sTypo` 35. `split` under FreeType reads `hhea`, which is
what [0044](0044-what-the-runners-own-fonts-said.md) had never measured. On the runner's own faces
half up and half to even tie, because none puts a box on the half; `split-2560` puts `hhea` there
at 962.5/212.5, and run 34690905750 reported **963/213**: round half up 94/94, the single-precision
and 16.16 models 93, half to even 93, the borrow 77, ceil 56, floor 55, exact 51. FreeType rounds
as DirectWrite does.

## What macOS said

Run 34689472437 indexed 247 faces of 796 in 382 files (127 collections skipped, four with no
measurable glyph, four the browser refused) and failed three gates, as predicted: `advance-quantum`,
`width` and `face-box-rival`. The last is the finding; run 34690903205, with the reader below and
the fifteenth probe, agrees on every face at 248/248.

**CoreText reads `hhea`, and ignores `fsSelection` bit 7.** `hhea.ascender/descender` fits 248/248
within half a pixel; the FreeType composite 247/248; the DirectWrite composite the reader had
assumed, 210/248, differing on 38. The one face separating `hhea` from the composite is
`split-usetypo` — the probe built with bit 7 set and `hhea` 800/200 apart from `sTypo` 700/100 —
and CoreText reported 800/200. Not one of the 233 real faces there could have asked that question:
eleven carry bit 7 and on every one of them `hhea` and `sTypo` agree.

**CoreText holds the em ratio as 16.16 fixed point before it rounds.** Rescored against `hhea` for
equality rather than tolerance, five sides sit on the half at 1000 px and they do not all round the
same way:

| face                   | `hhea` at 1000 px | browser | 16.16, nearest        | 32-bit float   |
| ---------------------- | ----------------- | ------- | --------------------- | -------------- |
| Hoefler Text Ornaments | 806.5             | 807     | 52854.8 → 52855, up   | up             |
| Zapfino                | 1502.5            | 1503    | 98467.8 → 98468, up   | up             |
| Krungthep              | 262.5             | 262     | 17203.2 → 17203, down | down           |
| `split-2560` ascent    | 962.5             | 962     | 63078.4 → 63078, down | down           |
| `split-2560` descent   | 212.5             | 212     | 13926.4 → 13926, down | **up**, so 213 |

The first macOS run had four of those sides and fit a single-precision ratio, 247/247, with
`split-2560` built to test it. Its descent refuted it: 0.2125 rounds up as a float and the browser
said 212. `floor(round(units / unitsPerEm × 65536) / 65536 × 1000 + 0.5)` — the ratio as CoreText's
`Fixed`, nearest — fits 248/248 and is the only precision that does: 2<sup>17</sup> and a float
miss `split-2560`, 2<sup>20</sup> and 2<sup>24</sup> miss Zapfino, 2<sup>28</sup> misses Hoefler,
2<sup>15</sup> misses two, plain half up 246, half to even 246, exact 159. Truncation to 16.16
misses Hoefler and Zapfino. The same model on DirectWrite is refuted by T13: `split-2000`'s 0.9285
and `split-2560`'s `usWin` 0.2625 both sit below the half in 16.16, and DirectWrite rounded both up,
30/30 against 28/30; FreeType likewise, 94 against 93.

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
FreeType, the ratio as 16.16 fixed point through CoreText, half up in both. `FaceMetrics` carries the
backend so the measurer needs no second lookup. T14's `shippedBoxFor('darwin')` is `hhea` alone,
`rasteriserOf` says CoreText, and `BOX_ROUNDINGS` takes units and the em so that a precision can be
a model: the fixed-point rule ships and the single-precision one stays as the rival it is. T13
scores 32 models over 30 boxes and the T14 variants table eight.

`packages/text`'s ideographic fallback is answered at `FACE_BOX_PX` and divided, because that is the
one size the engine probes at: the browser rounds the fallback per size (26 at 100 px and 263 at
1000 for 672/2560), and no single fraction of the em is both. The reader answers the 1000 px one.

## Verification

`pnpm check` green. T13 rescored over fifteen probes: face box 30/30, kerning 90/90, widths 540/540,
ideographic 38/38. CI: run 34688374783 on `main` at 55f3073 — `probe-fonts` green on its first
Windows run, `candidate` green on its first, `real-faces` red on the autohinter; 34689472607 and
34690905750 on this branch green everywhere, the second with `split-2560`; 34689472437 and
34690903205 the two macOS dispatches, red on `advance-quantum` and `width` as predicted and green
on the box on the second.

| mutant                                         | killed by                                                          |
| ---------------------------------------------- | ------------------------------------------------------------------ |
| `backendFor('darwin')` returns `directwrite`   | names FreeType on Linux, CoreText on macOS and DirectWrite else    |
| CoreText follows bit 7 onto `sTypo`            | reads hhea through CoreText whatever bit 7 says                    |
| `boxPixels` drops the fixed point for CoreText | rounds the box as CoreText does when indexed for macOS             |
| `boxPixels` holds a float instead              | the same, on 544/2560                                              |
| `boxPixels` fixes every backend                | keeps the em ratio exact through DirectWrite                       |
| `split-2560` dropped from the probes           | the `probe-fonts` diff of the regenerated fixture, not a unit test |
| `shippedBoxFor('darwin')` stays DirectWrite    | scores macOS against hhea alone                                    |
| the fixed-point model dropped from T14         | separates 16.16 fixed point from single precision and from half up |
| `fpgm` removed from the builder                | the Linux `real-faces` job, run 34688374783                        |

The single-precision model was the shipped rule for one commit on this branch, 3622457, and the
macOS run that followed refuted it on the probe built for the purpose. That is the experiment
working as designed, and the reason `split-2560` puts `hhea` on the half as well as `usWin`.

## Deviations

1. **A probe added after the plan.** `split-2560` exists because the first macOS run produced a
   model with a prediction one font could test on all three rasterisers, and not testing it would
   have shipped a float where a `Fixed` was.
2. **T14's gates are not made platform-aware.** On macOS `advance-quantum` fails on the quantum
   itself and `width` on the four classes above; the job is dispatch-only and gates nothing, and a
   green macOS run would need the four rules, not a looser gate.
3. **The ideographic test asks one size.** It had asserted a size-independent fraction against
   both baseline sizes, which `split-2560` shows cannot hold.

## Open questions

1. **A tie in 16.16.** A ratio landing exactly on a half-unit of 1/65536 rounds nearest by
   `Math.round`, up; which way CoreText's `Fixed` goes there is unmeasured, and no face here lands
   on one.
2. **The four macOS width classes** — the last-bit quantiser on odd ems, the single-precision sum,
   Apple's kerning source, and `trak` — each unmeasured beyond its count.
3. **Half to even on macOS** is refuted by Hoefler and Zapfino among the runner's faces and by no
   probe: on a 2560 em every value on the half has an even floor or agrees with the fixed-point
   rule, so the record leans on two Apple faces for that one.
