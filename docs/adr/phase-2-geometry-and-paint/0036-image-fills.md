# 0036 — Image fills, and the two readings everybody gets wrong

Date: 2026-09-07
Status: **accepted** — 63 probes, 12 rules, 14 refuted readings, 8/8 mutants killed

**Sub-phase 2.12.** A plan amendment: `a:blipFill` was named in Gate 2's text and assigned to no
sub-phase, and was re-flagged in [0022](./0022-fills.md), [0024](./0024-model-parse-and-resolve.md),
[0025](./0025-renderers-geometry.md), [0026](./0026-debug-overlay-preset-gallery.md) and
[0027](../phase-3-text/0027-the-text-cascade.md) without ever being built. It is built here.

Code: `packages/paint/src/fills/blip.ts`, `packages/model/src/parse/paint.ts`,
`packages/render-svg/src/image/{blip.ts, header.ts}`, `tools/ground-truth/paint/blips/`.
Fixture: `corpus/ground-truth/blips.json`.

---

## What this sub-phase had to decide

An `a:blipFill` states an image, a crop, a mode, and a list of colour effects. The schema names all
four and defines the meaning of none of them, so every question here was measured.

### PowerPoint answered the unit questions in writing

`author.ps1` drives the PowerPoint object model and reads back the markup it chose. That settled six
questions in one run, before any probe deck existed:

| question                     | PowerPoint's answer                                                      |
| ---------------------------- | ------------------------------------------------------------------------ |
| `a:tile/@tx` units           | EMU — 9 pt is written `tx="114300"`, negatives as negatives              |
| `@sx`/`@sy` units            | hundred-thousandths of a natural size — 0.5 becomes `50000`              |
| `Fill.Transparency = 0.4`    | `<a:alphaModFix amt="60000"/>` — **`@amt` is opacity, not transparency** |
| brightness 0.7, contrast 0.8 | `<a:lum bright="40000" contrast="60000"/>`, i.e. `(ui − 0.5) × 2`        |
| "Watermark"                  | `<a:lum bright="70000" contrast="-70000"/>`                              |
| "Black and white"            | `<a:biLevel thresh="50000"/>`                                            |
| "Set transparent color"      | `<a:clrChange>` from a colour **to that same colour at `alpha="0"`**     |

The crop probe also fixed the resolution question by arithmetic: cropping 12.5 pt off a 32-pixel
image is written `t="52083"`, so the source is exactly 24 pt tall, which is 32 px at 96 dpi.

### The tile lattice, measured on 63 probes

Four probe decks, exported by PowerPoint as BMPs at 1920 and 3840 wide, and scored against candidate
readings that each predict an exact colour for every sampled pixel. The rules:

1. **A tile's natural size is the image's pixel count over its own declared resolution.** Not a fixed
   96 dpi: the same 32 px declaring 150 dpi tiles at 15.36 pt rather than 24 pt.
2. `@sx`/`@sy` scale that natural size, per axis.
3. **`@algn` places one whole tile flush against that edge or corner of the shape box**, and the
   lattice repeats from there. All nine values agree, to the pixel.
4. `@tx`/`@ty` move the lattice from that anchor, positive right and down.
5. **`@flip` mirrors _alternate_ tiles, not every tile.** The visible period doubles on the mirrored
   axis.
6. **`a:srcRect` changes what is drawn inside a tile, not how big the tile is.**

Rules 5 and 6 are the ones worth the experiment. Both are the reading anyone writes first, and both
are wrong:

| refuted reading                          | its best score | the measured reading, same probe |
| ---------------------------------------- | -------------: | -------------------------------: |
| `@flip` mirrors every tile               |           0.19 |                             1.00 |
| the lattice starts at the shape's origin |           0.05 |                             1.00 |
| `@tx` moves the lattice the other way    |           0.00 |                             0.93 |
| `@sx` is a fraction of the shape         |           0.20 |                             1.00 |
| the tile shrinks with `a:srcRect`        |           0.19 |                             0.86 |
| `a:srcRect` is coordinates, not insets   |           0.00 |                             1.00 |
| `a:fillRect` is ignored                  |           0.35 |                             0.90 |
| the tile ignores the image's own dpi     |           0.24 |                             0.92 |

### The colour effects are BT.709, in sRGB, per channel

A 256-column ramp over red, green and blue bands gives 256 inputs per effect rather than the four a
quadrant image would.

- **`a:grayscl` weighs BT.709** — red at full scale measures 54 of 255, where BT.601 predicts 76.
  Scored 0.999 against BT.601's 0.32. The sum is taken in sRGB: computing it in linear light puts
  mid-red at 60 where PowerPoint puts it at 27.
- **`a:biLevel`** thresholds that same luminance against `@thresh × 255`, confirmed at three
  thresholds.
- **`a:duotone`** interpolates linearly in sRGB between its two colours by that luminance.
- **`a:alphaModFix`** is an opacity: black at `amt="60000"` exports as 102, which is `0 × 0.6 + 255 ×
0.4`.
- **`a:lum` is per channel, not on luminance.** The slope is `1/(1 − contrast)` for positive contrast
  and `1 + contrast` for negative — one formula for both signs is refuted — about a mid-grey pivot.
  The brightness term is `0.5 × (1 + k) × bright`: it **grows with the contrast slope**, and treating
  it as independent scores 0.00 on every probe carrying both.
- **`a:clrChange` is an exact match.** A colour one level away from the target is left alone.

Effects apply in document order, confirmed by a grayscale followed by an alpha.

---

## What was built

`blipPlacement` in paint is the whole geometry, as a pure function of the fill, the shape box and the
image size. The renderer emits a `<pattern>` for both modes — a pattern cell clips its own contents,
which is what crops a stretched image without a second clip path — and the mirrored cases emit a
2×1, 1×2 or 2×2 cell carrying that many `<image>` elements, which is rule 5 expressed directly.

The effects become one SVG `<filter>`. `a:clrChange` has no primitive that tests colour equality, but
a 256-entry `discrete` transfer table tests one channel exactly, and three of them averaged reach 1
only where all three matched — so the mask is built from three tables and a linear transfer.
Refusing it was tried first and rejected: it cost the whole of `a03-fills-02`, which is worse than
the gap it was meant to be honest about.

**`r:embed` is carried, never resolved.** The parser records the part the fill was written in,
because rIds are scoped to one `.rels`: a fill inherited from a layout means the _layout's_ `rId2`,
and resolving it against the slide would silently paint a different picture. The renderer takes a
`media` resolver of `(embed, part)`.

---

## What it moved

| slide                                          |  before |       after |
| ---------------------------------------------- | ------: | ----------: |
| `a03-fills-02` — the eight `a:blipFill` probes | 8199 bp | **9378 bp** |
| `b09-picture-02`                               | 8918 bp | **9951 bp** |
| corpus mean                                    | 9719 bp | **9732 bp** |

`a03-fills-02` is [0035](../phase-3-text/0035-the-fidelity-harness.md)'s **open question 3** — "scores
8199 and nothing in the plan explains it". The answer was that the slide is the image-fill probe
slide and image fills were not drawn. That question is closed.

- 63 probes, and the measured reading is the top-scoring candidate on every one.
- 28 unit tests, every expected value derived from the fixture or from the rule's own arithmetic.
- **Mutation sweep: 8 mutants, 8 killed**, control green. The anchor table both ways, the flip
  parity, the tile size, the offset sign, both contrast branches, and the brightness coupling.

---

## Deviations from the plan

| plan item                        | disposition       | why                                                                                                                                                                                                                  |
| -------------------------------- | ----------------- | -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| a sub-phase for `a:blipFill`     | **added as 2.12** | The plan named it in Gate 2 and assigned it to nobody. Five ADRs flagged it; this one builds it.                                                                                                                     |
| image decoding in a core package | **header only**   | `imageSize` reads PNG, JPEG, GIF and BMP headers for pixel count and resolution. Nothing is decoded, so the renderer stays synchronous and there is no codec in the core. EMF, WMF and TIFF stay Phase 10 and throw. |
| `p:pic` picture shapes           | **not done**      | A picture is a shape kind, not a fill, and `b09-picture-01` is still 8642 bp because of it. It belongs with 2.10's `pic`, and is now the largest single gap the score can see.                                       |

---

## Open questions

1. **Where the arithmetic rounds is not determined.** Round-to-nearest fits the luminance-derived
   effects; truncation fits one `a:lum` probe. The two differ by at most one level per channel and
   the probes cannot separate them; the fixture records both scores rather than asserting one.
2. **A non-zero `a:blipFill/@dpi` is unexplained.** PowerPoint writes `dpi="0"` for every fill it
   authors, and no candidate fits the five probes that set it — a resample to that resolution is
   close but not exact. The implementation uses the image's own resolution, which is the measured
   rule for the `dpi="0"` every real deck carries.
3. **`rotWithShape` is parsed and unused.** No probe distinguishes it yet on a rotated shape.
4. **The `<pattern>` is anchored in the shape's own space**, so a `grpFill` child of an image-filled
   group has not been probed and may restart the tiling rather than continue it.
5. **`a:tile` on a slide background** has no shape box, and the extent is assumed to be the slide.
   Same gap as the gradient background in 0022.
