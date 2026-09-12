# 0054 — A hundred slides, at any zoom

Date: 2026-09-13
Status: **accepted** — Gate 3 holds: 100 slides × 5 widths through the page, offline, 28/28 mutants
killed

**Sub-phase 3.11.** Gate 3 asks for "a 100-slide deck rendered faithfully at any zoom, entirely
client-side". None of its four claims was true or checkable when this began: the largest committed
deck had eleven slides, the page drew no pictures and no `w="0"` lines, its stage was a constant 900
pixels wide with no zoom control, and nothing measured what the page did or where it fetched from.
This is the record of making each claim a measurement, and of what the measurements found.

Code: `tools/ground-truth/render/zoom/`, `tools/gate3/`, `apps/studio/src/slides/`,
`tools/corpus/tiers/a-generated/decks/deck/a46-hundred-slides/`, `tools/fidelity/{oracle,record}.ts`,
`tools/fidelity/metric/grid.ts`, `packages/paint/src/lines/line.ts`,
`packages/render-svg/src/{paint,shape,image/media}.ts`, `packages/render-svg/src/text/`. Fixtures:
`corpus/ground-truth/zoom.json`, `corpus/decks/a46-hundred-slides.pptx` and its oracle at five
widths, `corpus/ground-truth/render/fidelity/zoom/expected.<env>.json`.

---

## The sentence, as four measurements

| claim                | what now checks it                                                                                                                                                                                                       |
| -------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| a 100-slide deck     | `a46-hundred-slides`, generated, 100 slides, in `corpus/decks/`; the oracle holds PowerPoint's export of every slide at 120, 240, 960, 1920 and 3840 pixels wide                                                         |
| rendered faithfully  | every slide scored through the **page's own DOM** against that export, at every width; `pnpm fidelity` scores it through the `<img>` path as well                                                                        |
| at any zoom          | experiment F2 asked PowerPoint what its export does at seven widths that a linear scale does not; the rule is in the renderer; the gate asserts the SVG at every zoom is the SVG at 100 % apart from what that rule owns |
| entirely client-side | the gate fetches the deck once, takes the browser context offline, and logs every request; a request it does not recognise, or any request after going offline, fails it                                                 |

`pnpm gate3` runs all of it in `pnpm check` and CI from committed fixtures, needs no Office, and
exits 0 only when every count is zero. The scores are reported and gate nothing, for the reason ADR
0035 gives.

## F2 — what zoom does to PowerPoint's own picture

Seven export widths (120, 240, 480, 960, 1200, 1920, 3840), 38 probes on seven slides, each slide
exported twice as BMP and compared with itself; 266 cases against PowerPoint 16.0 build 20326.
Every family had its candidate readings enumerated — including the ones anyone would write by hand —
and scored at 0.15 px; `analyse.ts` throws unless exactly one model per family fits every case, and
`assertSeparable()` refuses a pair of models no width can tell apart.

| family          | winner                                                                     | fits  | closest loser                                       |
| --------------- | -------------------------------------------------------------------------- | ----- | --------------------------------------------------- |
| hairline        | **H1** one device pixel at every width                                     | 42/42 | H4 half a point, never under a pixel — 36/42        |
| thin stroke     | **T4** rounded to whole pixels, never under one                            | 56/56 | T5 rounded up — 48/56; T2 true width ≥ 1 px — 46/56 |
| border          | **T4**                                                                     | 14/14 | T2, T3 — 10/14                                      |
| border edge     | **O4** wholly outside, the whole-pixel stroke snapped to the grid, half up | 8/8   | O1 wholly outside, unsnapped — 7/8                  |
| dashed hairline | **D1** the period of the drawn pixel width; a zero width has no dashes     | 28/28 | D4 solid at every width — 21/28                     |
| marker          | **M8** a ten-point head at or under two points, else five whole-pixel pens | 49/49 | M5, M7 — 44/49                                      |
| pattern         | **P1** the tile is six points at every width                               | 6/6   | P2 six points snapped to pixels — 5/6               |

What the losers would have looked like in the field: H4 is the reading the `w-quarter` clue in
`lines.json` suggested and it is wrong at 120 and 240 px, where half a point is under a pixel; T2 is
what an SVG renderer does by default and it is wrong for every width between one and two pixels;
D4 — "a dashed hairline is solid" — is right 21 times and wrong exactly where a dash is wider than
the pixel it is drawn with.

Three families needed no rule. **Text is linear**: every run's extent at every width is within 1 px +
2 % of a linear scale of the 960 export (worst 4 px on a 3840-pixel slide), and nothing is dropped
at 120 px — a 6-pt Arial run is still there at 0.75 px per point. **Gradients are identical** from
480 px up (worst 1 level of 255; 5 levels at 120, where there are only 49 distinct levels to
have). **The frame is stretched**, never letterboxed, at every width, and 3840 wide is not a
ceiling: `Slide.Export` produced every width asked for.

Two findings are recorded and not implemented, and `zoom.json` carries them:

- **O4's half-up snap.** A 1-pt picture border at 960 px is one pixel wholly outside the frame, and
  at 1200 px (1.25 px) PowerPoint draws one pixel with its inner edge snapped half a pixel out. The
  renderer draws the band wholly outside at the drawn width (ADR 0037) and does not snap; the
  difference is under half a device pixel and only at non-integer scales.
- **M8, the arrowhead pen.** A triangle head on a line at or under two points is ten points tall at
  every width, and above two points it is five pens of the whole-pixel stroke. `markerGeometry` still
  scales from the nominal width, which is what ADR 0023 measured at one width. Carried, below.

**Pattern coverage under 480 px** is PowerPoint's own artefact: `horz` at 480 px exported with no
ink at all, and at 120 and 240 px with a 3-px period where six points is 0.75 and 1.5 px. P1 is
scored from 960 up, where the period is six points to 0.02 px, and the fixture keeps the low widths
under `patternBelow` rather than pretending a rule fits them.

## The rule in the renderer

`deviceStrokeWidth(widthEmu, pxPerPt)` in `packages/paint`: `max(1, round(w · s))` device pixels, a
zero width one pixel. `render-svg` applies it when `RenderOptions.width` names a device — the page
always does, the CLI does, and a caller that gives no width gets the true width and a hairline as
`stroke-width="1" vector-effect="non-scaling-stroke"`. A dashed hairline is drawn solid (D1). The
device scale has a floor, `MIN_PX_PER_PT = 0.05`: below a twentieth of a pixel to the point
`deviceStrokeWidth` throws `LINE_DEVICE_SCALE`, the page's fit zoom stops there, and the widest one
device pixel a stroke is ever rounded up to is twenty points.

**Zoom is a mount at that width, never a resize.** This is the plan's "measured exception": the SVG
for 25 % is not the SVG for 400 % with another `width`, because `stroke-width` and
`stroke-dasharray` are rounded to the device pixels of the width the slide was mounted at. The
architecture claim "zoom is a transform, never a re-layout" survives in the half that matters: the
text is laid out once, at one point size, and the gate proves the line breaks and every other byte
of the markup are the same at every zoom. `MountedSlide.resize()` is gone.

The stroke rule moved 29 of the 31 corpus slides it touched toward PowerPoint (mean 9768 → 9770
bp over the 154 slides then scored) and made none worse.

## The deck, and the limits it found

`a46-hundred-slides`: Tier A, generated, LCG-seeded, 100 slides of ten kinds — cover 1, agenda 1,
section header 8, bullets 25, two-column 13, picture with caption 12, shape grid 13, quote 13, bar
diagram 13, closing 1 — with a `sldNum` field on every content slide (90), two procedural PNGs shared
by twelve slides' relationships, a master text-style override, `normAutofit` on every fourth bullets
slide, and a shape grid of twelve presets under solid, gradient, pattern fills, dashed lines, a
connector with a triangle head, an outer shadow and a glow, rotations, one `flipH` and one nested
group turned 20°. 164,143 bytes deflated, 222 entries; the census literal locks 803 shapes, 746
preset geometries, 182 gradient fills, 35 pattern fills, 39 connectors, 90 fields, 13 shadows and 13
glows. No date field: `a09-fields` already carries the calendar drift ADR 0038 Q3 records.

Its claim, in a43's terms, was the limits nothing had ever hit. Each one hit:

| limit                                        | what happened                                                                                                                                                           |
| -------------------------------------------- | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| the grid set's one-byte slide count          | `encodeGridSet` silently truncated at 256; it now throws above 255, and sets shard by bytes at 480 KiB (`shardGridSets`) — a46 at 960 px is three files                 |
| `CAPS.perFile` (512 KiB) on one deck's grids | the same sharding; `oracle.json` records name their file                                                                                                                |
| `CAPS.total` (12 MiB)                        | 8.0 MiB + a46 0.16 + 1.17 at 960 + 4.68 at four widths = 13.9 MiB; raised to 16 MiB, the user's decision                                                                |
| the harness's per-slide deck reload          | `drawSlide` loaded the package once per slide — quadratic on a46; it now caches the loaded deck per URL for the run                                                     |
| a Linux baseline that could not grow         | `--record` is refused in CI and `--bootstrap` once a baseline exists; the stale `expected.linux-x64.json` is deleted and the dispatch job records both baselines afresh |
| `sldNum` resolved per slide                  | 1..100, correct in every export and every render                                                                                                                        |
| `<defs>` ids across 101 mounted SVGs         | `idPrefix` per mount — `thumb<i>`, `slide<i>` — and the gate renames one to the other to compare them                                                                   |
| one measurer per mount                       | one `TextEngine` per deck, passed to every mount; the strip went from one `OffscreenCanvas` per slide to one per deck                                                   |

Two defects surfaced only because a hundred slides of ordinary content went past the renderer at
once, and neither was in the plan:

- **Every bordered picture was blank.** The clipped outline band's `clip-path` was on the same path
  as the fill, so the fill was clipped to the band. Fill and band are two paths now; a test draws a
  one-pixel red PNG behind a border and reads the pixel. +54 bp on the corpus mean.
- **No bullet was drawn.** The model resolved `a:buChar`, `a:buAutoNum` and `a:buBlip` and
  `packages/text` laid them out, and `render-svg` never emitted one. It does now, through
  `drawableBullet`, `numberParagraphs` and `bulletLayout`; 49 slides improved, none worsened, +7 bp.
  One slide is refused as a consequence: `a08-bullets-02` uses `arabic1Minus`, whose glyphs T5 could
  not read back (PowerPoint shapes them to glyph ids), so `TEXT_AUTONUMBER_UNMEASURED` is thrown
  rather than a guess drawn. Recorded as vanished from the baseline.

## The page

`apps/studio/src/slides/`: a strip of every slide at 120 px, mounted in frame-sized chunks after
the stage has painted; a stage at 25, 100, 200, 400 % or fit; one text engine per deck; a media
resolver over the package (`mediaFromStore`, shared with the CLI and the harness); a delegated click
listener instead of one per mount; and an automation hook — `openFromUrl`, `show`, `stageSvg`,
`thumbSvg`, `stageShapes`, `stageBox`, `thumbBox` — that reports timings and never asserts them.

The first drive of a46 took **18.6 s** to draw the strip while every line of JavaScript accounted
for 0.45 s. Chromium's trace put it in two `RasterTask`s of nine seconds each, and bisecting one
glowed hexagon in a 240-pixel SVG found the whole cost in one place:

| the shape, in the emitted filter      |  raster |
| ------------------------------------- | ------: |
| as emitted: rotated, gradient, dashed | 9486 ms |
| the same with no rotation             |   28 ms |
| solid, no stroke, rotated             | 8431 ms |
| solid, no stroke, flat                |   31 ms |

`feMorphology operator="dilate"` under a `rotate()` transform is pathological in Chromium **when the
filter's user space is EMU**: the cost rises with the radius in user units (radius 1: 0.5 s; 2857:
6 s; 28575: 9 s; 285750: 15 s) and vanishes when the same filter is written in a coarser unit
(÷10: 8.3 s; ÷100: 30 ms; ÷12700: 32 ms). So an effect filter is now written in **points**, inside a
`scale(12700)` / `scale(1/12700)` pair around the shape's paths — the second written as a string,
because `num()`'s three decimals round it to zero. Fifteen effect slides moved within the noise
floor (largest −6 bp, every `maxD` unchanged) and the strip fell to **761 ms**.

| a46 through the page, win32                  |                     measured |
| -------------------------------------------- | ---------------------------: |
| package parsed and model loaded              |                        95 ms |
| first slide painted                          |                       243 ms |
| 100 thumbnails mounted                       | 761 ms, 430 ms on the thread |
| JS heap after the strip                      |                        22 MB |
| one stage mount, any zoom                    |                      ~1.5 ms |
| one screenshot at 240 / 960 / 1920 / 3840 px |       41 / 65 / 127 / 360 ms |

## What the gate found

`pnpm gate3` on win32-x64, 2 m 11 s: 100 slides, 400 stage renders and 100 thumbnails, all drawn,
all invariant, two requests for the deck and none after going offline, 500 rasters recorded.

| zoom          | px wide | mean bp | worst slide                 |
| ------------- | ------: | ------: | --------------------------- |
| 100 %         |     960 |    9954 | a46-hundred-slides-40, 9819 |
| 25 %          |     240 |    9878 | a46-hundred-slides-24, 9655 |
| 200 %         |    1920 |    9960 | a46-hundred-slides-40, 9799 |
| 400 %         |    3840 |    9964 | a46-hundred-slides-40, 9799 |
| strip, 12.5 % |     120 |    9727 | a46-hundred-slides-88, 9184 |

The strip is the honest column: at 120 px a cell is one pixel, text is under a pixel tall and every
antialiasing decision shows. Slide 88 is a bar diagram whose labels PowerPoint drops into grey at
that size and ours keeps legible. The stage columns climb with width because the cell grows with
it, as ADR 0022 predicted for anything that scales.

Two facts the gate established about the page that the plan had assumed. First, an inline `<svg>`
root is **not** pixel-snapped by Blink: `getBoundingClientRect()` put the stage at y = 817.1875, and
a Playwright clip is viewport-relative unless `fullPage`, and then only as wide as the viewport. The
gate pins the stage and the strip at page (0, 0) with an injected stylesheet, hides the one it is
not shooting, and refuses a fractional box rather than rounding one. Second, only `stroke-width`
and `stroke-dasharray` differed between zooms across every slide measured — and the outline band's
containment rectangle, whose reach was two drawn widths. Its reach is now
`2 × (nominal width + one device pixel at the floor)`, zoom-invariant, after measuring that a reach
past ~600 pt changes how Skia antialiases the clip (−10 bp on every bordered picture; identical up to
600 pt).

`pnpm fidelity` over 254 slides: corpus mean **9841 bp**, a46 mean 9955, minimum 9819; noise floor 4
bp. The five lowest decks are the ones the plan has not reached — a23-smartart 7255, b04-table 7653,
b06-smartart 8571, a22-chartex 8800, a37-mce 8823.

### The real deck

One deck outside the corpus was measured through `pnpm gate3 --deck`, unscored — there is no oracle
for it and it is not a fixture — handed over by name and read once from its own directory:
100 slides, every one drawn at all four zooms and in the strip, zoom-invariant, no page error,
two requests for the deck and none after going offline; parsed in 160 ms, first slide at 300 ms,
100 thumbnails at 659 ms (314 ms on the thread), 28 MB of heap. It is the first deck not written by
this repository to go through the page end to end, and it needed nothing fixed.

## Verification

- **Tests from the fixture.** `line.test.ts` reproduces every F2 probe at every width within 0.15 px
  and asserts each losing model misses at least one; `render.test.ts` renders the hairline slide at
  240 and 3840 through the `<img>` path and profiles it; the text test asserts a run's extent at 4×
  is 4× its 1× extent within `textTolerance`; the effect test derives the emitted radius from
  `paint`'s own graph and reads the glow's pixels; the band test compares the clip at 960 and 240.
- **The way it will fail.** The page in Chromium against PowerPoint's own pixels at five widths,
  offline; and a page that resizes instead of re-mounting on zoom is caught at the first slide
  (`FID_SVG_INTRINSIC_SIZE`: the stage was still at the fit size).
- **Mutants, 28/28 killed.** F2's renderer rule 9/9 (among them width ≤ 0 → null, no
  `vector-effect`, a 0.75-pt minimum, the minimum on every width, `ceil` for `round`, the dash
  array from width 0, the marker from width 0, `w·s` unconditionally); the device floor 1/1
  (0.05 → 0.01); bordered picture 1/1; bullets 8/8; the effect filter 3/3 (radius left in EMU, the
  inverse scale through `num()`, no inverse group); the band reach 1/1 (the device pixel dropped
  from it); the harness 4/4 (a mask that blanks every `width`, a rename without its boundary, the
  hundred-slides clause dropped, cross-origin counted as static); the page 1/1 (resize on zoom).
- **Behaviour-preserving refactors** — the media resolver, `geometryOf`'s cell, the oracle loader,
  `record.ts`'s `Baseline` — left all 155 raster digests of the day unchanged.
- `pnpm check` green: structure, layering, legal, references, corpus, format, build, lint,
  typecheck, round trip 55/55, fidelity, Gate 2, Gate 3, package QA, tests.

## Deviations from the plan

- **Zoom is a mount, not `resize()`.** The plan named `MountedSlide.resize()` and said a
  device-pixel rule, if measured, would be the one exception. It was measured (T4, H1), so the
  exception is the design and `resize()` is deleted rather than kept beside it.
- **`--extend` was dropped.** An append-only record mode for a baseline that cannot otherwise grow
  in CI was designed and not built: deleting the stale Linux baseline and letting the dispatch job
  bootstrap both files is one mechanism instead of two, and `record-linux-baselines` records
  whichever is missing.
- **A record run exits 0.** `pnpm fidelity --record` wrote the baseline and then threw over the
  digests it had just recorded. It reports what it recorded now, and `pnpm gate3` was written the
  same way.
- **Five corpus count locks** in `tools/corpus/{lexical,suites}` were still at fifty-four decks when
  a46 landed; the commit that added the deck left them red. Fixed with the page.
- **The effect filter's unit, the band's reach, bullets and bordered pictures** were not in the
  plan; each is above.
- **PowerPoint, not LibreOffice**, as ADR 0038 decided, and the gate runs from committed grids.
- The CI job `record-fidelity-baseline` that ADR 0041 names is now `record-linux-baselines`.

## Open questions

1. **Device pixel ratio.** The page mounts at CSS pixels per point and the rule rounds to those; on a
   DPR-2 display the device pixel is half a CSS pixel, and what PowerPoint does there is what its
   1920 export shows at 100 %. The page does not pass `devicePixelRatio` into the mount. Measured
   nowhere yet.
2. **The strip's letterbox.** At 120 px a 16:9 slide is 67.5 px tall; PowerPoint rounds to 68 and
   stretches (F2's frame finding), the page rounds to 68 and the SVG's `preserveAspectRatio` centres
   it — a quarter-pixel disagreement that the 9727 column includes and nothing separates.
3. **Pattern fills under 480 px** in PowerPoint's own export (`patternBelow`): whether the thinning
   is the export's or the screen's is not known.
4. **Arrowheads at zoom (M8)** and **the border's half-up snap (O4)**: measured, recorded, carried.
5. **The clip-reach edge.** Skia antialiases the outline band's clip differently once its reach
   passes roughly 600 pt; a picture border wider than ~280 pt would cross it at any zoom. No corpus
   deck has one.
6. **Effect filters at DPR > 1** are written in points, which is one CSS pixel at 100 %; whether
   Chromium's rotated-filter cost returns at a scale of 2 is not measured.
7. **A second real deck.** One deck outside the corpus has been through the page. That is an
   anecdote, not a rule.
