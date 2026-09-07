# 0037 — Picture shapes, and the border that is drawn outside

Date: 2026-09-07
Status: **accepted** — 54 samples, 24 readings scored, 23 refuted at a full channel, 10/10 mutants
killed

**Sub-phase 2.13.** A plan amendment, and the last of Phase 2's carried debt.
[0036](./0036-image-fills.md) built image _fills_ and closed by recording that a picture is a shape
kind rather than a fill, that `p:pic` was therefore still not drawn, and that `b09-picture-01` scored
8642 bp as the largest single gap the fidelity score could see. It is built here.

Code: `packages/model/src/parse/{paint.ts, sheet.ts}`, `packages/paint/src/fills/fill.ts`,
`packages/render-svg/src/{paint.ts, shape.ts, image/blip.ts}`,
`tools/ground-truth/render/pictures/`. Fixture: `corpus/ground-truth/pictures.json`.

---

## The defect, and why nothing caught it

`CT_Picture` is `nvPicPr, blipFill, spPr`. The image is a **sibling** of the shape properties and
sits in PresentationML, so a picture's markup reads

```xml
<p:pic>
  <p:nvPicPr>…</p:nvPicPr>
  <p:blipFill><a:blip r:embed="rId2"/><a:srcRect …/><a:stretch><a:fillRect/></a:stretch></p:blipFill>
  <p:spPr><a:xfrm …/><a:prstGeom prst="rect"/></p:spPr>
</p:pic>
```

`parseShape` read every shape's fill with `parseFill(spPr)`, whose element set is `a:noFill`,
`a:solidFill`, `a:gradFill`, `a:blipFill`, `a:pattFill`, `a:grpFill` — all six in DrawingML, none of
them a child of `p:spPr` on a picture. So every `p:pic` in every deck parsed with `fill: undefined`
and painted nothing, and had done since 2.9.

The unit suite could not see it: 3121 tests passed throughout. Neither could the round trip, which
re-emits a part nobody edited byte-for-byte and is indifferent to whether we understood it. Only the
fidelity harness — our render against PowerPoint's own, at PowerPoint's sample points — reports a
picture that is not there, and it did, at `maxD 254` on `b09-picture-01`: the largest per-cell
disagreement the metric can express.

## What the corpus could not answer

Wiring the fill in is four lines. Knowing whether those four lines are _right_ is not, and the corpus
had nothing to say about it. A census of all 26 `p:pic` elements across `corpus/{decks,authored,
written}`:

| on a picture         | in the corpus |
| -------------------- | ------------: |
| `prstGeom` ≠ `rect`  |             0 |
| a fill in `p:spPr`   |             0 |
| an `a:ln`            |             0 |
| `flipH` or `flipV`   |             0 |
| `a:xfrm/@rot`        |             0 |
| `a:srcRect`          |             1 |
| an SVG-only `a:blip` |             2 |

So four questions were open and unexercised, and answering them by reading the schema would have
produced exactly the kind of rule this repository does not accept. Experiment **C7** asks
PowerPoint.

## C7: nine probes, twenty-four readings, one survivor

Four saturated quadrants on a black slide, so that no sample has two readings and a flip cannot hide
in a symmetry. Nine pictures, six samples each, exported by PowerPoint 365 at 1920 px. Every probe
rectangle is a whole number of points, so each shape edge lands on a pixel boundary.

The four questions are independent, so the analysis scores the **whole cross product** — 2 clip × 2
precedence × 3 outline × 2 mirror — rather than one axis at a time, and refuses to emit a fixture
unless exactly one reading fits.

| question       | rivals                           | PowerPoint                           |
| -------------- | -------------------------------- | ------------------------------------ |
| **clip**       | `geometry` / `bbox`              | **clips to the shape's own outline** |
| **precedence** | `blip` / `spPr`                  | **the image wins**                   |
| **outline**    | `outside` / `centred` / `inside` | **wholly outside the box**           |
| **mirror**     | `image` / `outline`              | **the image is mirrored**            |

`geometry/blip/outside/image` fits at **worst channel error 0**. All twenty-three rivals are refuted
at **255** — the maximum a channel can disagree by — each on a named sample: `clip-ellipse/corner`
for a bounding-box reading, `precedence-solid/tl` for a `p:spPr` win, `outline-thick/corner` for a
centred border, `mirror-h/tl` for an unmirrored image.

### The border is drawn outside, and that is new

A 12 pt `a:ln` on a picture, read off an edge profile rather than a colour sample, because a band's
placement is a distance:

| a 12pt outline on | outside | inside |
| ----------------- | ------: | -----: |
| a **picture**     | 12.0 pt | 0.0 pt |
| a **shape**       |  6.0 pt | 6.0 pt |

The shape control is on the same slide of the same export deliberately: that a picture differs is
only a finding if both are read off one bitmap, and 2.8's centred default is re-measured here rather
than recalled. A picture's border is the full line width **outside** the box — PowerPoint never lets
a border eat into the image. SVG strokes are centred, so this is not something a renderer gets for
free: `strokeAttributes` now returns a `StrokeBand`, and `out` is drawn at double width and clipped
to the **complement** of the outline — one `<path>` carrying the inflated rectangle and the shape's
own subpaths under `clip-rule="evenodd"`. It is the mirror of the construction `algn="in"` has used
since 2.8.

## The SVG-only blip, and a throw that was wrong

Wiring the fill in immediately broke three slides that had been rendering: `a25-svg-blips-01/02/03`
threw `a:blipFill has no a:blip/@r:embed to resolve`. They hold this, which PowerPoint wrote:

```xml
<a:blip><a:extLst><a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}">
  <asvg:svgBlip r:embed="rId2"/></a:ext></a:extLst></a:blip>
```

An `a:blip` with **no `@r:embed` at all** — a picture whose only image is the vector original, with
no raster fallback. The throw was the defect. `render-svg`'s own `errors.ts` already states the
contract: _"an `a:blipFill` whose image has not been decoded: PowerPoint draws all four without
complaint, and so does this."_ A legal file that names an image we cannot yet decode is not a
failure, and killing the slide is a worse answer than the missing picture.

So `BlipFill.embed` is now `string | null`, `svgEmbed` carries `asvg:svgBlip/@r:embed` for 10.7, and
`parseBlipFill` throws only when the blip names **neither**. The extension is matched by its `@uri`,
never by the name of what it contains — a blip routinely carries `{28A0092B-…}` `useLocalDpi`
alongside — and an adversarial probe asserts a decoy `svgBlip` under the wrong `uri` is ignored.

## What it moved

Eight slides changed. Seven improved; one got worse, and it is the interesting one. Whole fidelity
harness, 146 slides, the "before" column measured by disabling the fill wiring and re-running rather
than recalled:

| slide               | before |    after |
| ------------------- | -----: | -------: |
| `b09-picture-01`    |   8642 | **9776** |
| `a19-decorative-02` |   8893 | **9979** |
| `a24-media-02`      |   9317 | **9984** |
| `a24-media-01`      |   9652 | **9985** |
| `a33-thumbnail-02`  |   9657 | **9975** |
| `a40-unicode-01`    |   9797 | **9983** |
| `a35-zip-shapes-01` |   9879 | **9984** |
| `a25-svg-blips-01`  |   9292 | **9175** |
| **corpus mean**     |   9732 | **9756** |

`a24-media` moved because a media object's poster frame is a `p:pic`; `a33-thumbnail` and
`a35-zip-shapes` because their content pictures were blank rectangles. **No Phase 2 slide remains in
the worst ten** — it is now SmartArt (4.5, 4.6), tables (4.1–4.4) and ChartEx (10.1–10.4), all of
which are later phases' work.

### Drawing the raster fallback is worse than drawing nothing

`a25-svg-blips-01` **regressed**, 9292 to 9175. Its second picture carries both a raster `r:embed`
and an `asvg:svgBlip`; PowerPoint draws the vector, and we now draw the raster. Against a blank
rectangle the metric was scoring background-against-background over most of the picture's area and
doing well out of it; against a raster that disagrees with PowerPoint's vector everywhere, it does
worse.

This is worth stating plainly rather than smoothing over, because it says something about the metric
as well as about the renderer: **a mean over cells rewards an empty slide.** It is the same reason
[0035](../phase-3-text/0035-the-fidelity-harness.md) leads its report with the worst slide rather
than the corpus mean. The regression is not an argument for reverting — a picture that is present and
imperfect is the honest render, and 10.7 replaces it with the vector — but it is an argument for not
reading 24 basis points of corpus mean as if it were the whole result.

## Verification

- 3162 → **3170 tests**, 89 files. The new ones read `corpus/ground-truth/pictures.json` and
  re-derive the rule from the sampled colours; none reads the implementation.
- **10 of 10 mutants killed.** The first sweep killed 9; the survivor — matching the SVG extension by
  child name instead of by `@uri` — was a missing assertion, not an equivalent mutant, and the decoy
  probe added for it kills it.
- `pnpm check` green: structure, layering, references, corpus (146 entries), format, lint, typecheck,
  build, round trip 52/52, fidelity, package QA, tests.

## Deviations from the plan

The plan gives `p:pic` no sub-phase of its own. 2.10 names `pic` in its file list and built the
transform stack; 2.12 built `a:blipFill` and recorded that a picture is neither. This is the third
plan amendment in Phase 2, and all three are the same shape: the plan treats images as a fill
property, and PowerPoint treats them as a shape kind with a fill-shaped payload.

The outline rule is a change to 2.8's code, not only an addition to it. `StrokePaint.inset: boolean`
became `StrokePaint.band: StrokeBand`, and every caller changed in the same commit — there is no
compatibility path, per the repository's own rule.

## Open questions

1. **Whether `rotWithShape` on a picture's blip means anything.** Every corpus picture writes the
   default and no probe varies it; C6 measured it on a shape's fill, where it has an effect.
2. **`a:srcRect` against a non-rect geometry.** The two are independent in our implementation — crop
   the source, then clip to the outline — and one probe would settle whether PowerPoint agrees.
3. **A picture's outline on a non-rect geometry.** The band is measured outside a rectangle only;
   whether an ellipse's border is offset along its normal or merely double-stroked-and-clipped is the
   same question 2.8 left open for `algn="in"`.
4. **SVG-only pictures draw nothing until 10.7,** and a picture that has a raster _and_ an SVG draws
   the wrong one of the two. Two corpus slides carry one of each. A recorded gap rather than an open
   question, but it is why `a25-svg-blips-01` sits at 9175 and why it moved the wrong way.
