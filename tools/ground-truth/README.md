# ground-truth

Experiments that ask the real Microsoft PowerPoint on this machine what it
actually does, rather than what a specification says it should.

Everything here is run **by hand**, needs PowerPoint installed, and is not part
of `pnpm check`. That is not laziness — a test that requires a licensed
application cannot gate a pull request. What CI can see is the _answers_, which
are committed as fixtures in [`corpus/ground-truth/`](../../corpus/ground-truth),
and the write-up in
[`docs/adr/phase-0-foundation/0007-ground-truth.md`](../../docs/adr/phase-0-foundation/0007-ground-truth.md).

## The experiments

### A — what does PowerPoint write into `ppt/fonts/*.fntdata`?

Save a deck twice, with `p:presentation/@saveSubsetFonts` flipped in between, and
read the EOT headers.

```bash
node tools/ground-truth/fonts/embedding/dump-fntdata.ts <deck.pptx> ... --json corpus/ground-truth/eot-headers.json
```

The one bit that matters is `TTEMBED_TTCOMPRESSED` (0x4) in the flags word at
offset 0x0C. It is set, always.

### B — will PowerPoint render an EOT we built?

```bash
node tools/ground-truth/fonts/embedding/build-deck.ts <dir> [<a-powerpoint-deck.pptx>]
node tools/ground-truth/fonts/embedding/verify-in-browser.ts <dir>     # is the FONT valid?
# open <dir>/probe-deck.pptx in PowerPoint, export slide1 as BMP, then:
node tools/ground-truth/fonts/embedding/analyse.ts <dir>
```

`graft-font.ts` puts one of those EOTs into a package PowerPoint itself wrote,
which is how the experiment separates "our EOT is wrong" from "our deck is
wrong":

```bash
node tools/ground-truth/fonts/embedding/graft-font.ts <pp-deck.pptx> <probe.eot> <typeface> <out.pptx>
```

**Read the controls before trusting a result.** The first run of this experiment
reported a pass that was not one: the deck opened cleanly, `Font.Embedded`
returned `msoTrue`, and the slide showed eight perfectly good letters — from a
substituted font. Only the row in a deliberately-absent typeface gave it away.

### C — which colour space does DrawingML compute in?

```bash
node tools/ground-truth/paint/colour/transforms/build-deck.ts <dir>
# open <dir>/swatch-deck.pptx in PowerPoint, read colours back and export BMPs
node tools/ground-truth/paint/colour/transforms/analyse.ts <dir> --fixture corpus/ground-truth/color-transforms.json
```

### C2 — the colour questions C did not ask _(added in 2.6)_

Every swatch in experiment C was built on `a:srgbClr` or `a:schemeClr`, on purpose, so a
disagreement was attributable to the transform rather than to the base. That left four of the six
bases, all three alpha operators, any non-identity `clrMap` and the whole percentage grammar
unmeasured. C2 measures them.

```bash
node tools/ground-truth/paint/colour/bases/build-deck.ts <dir>
powershell -File tools/ground-truth/paint/colour/bases/read.ps1 -Dir <dir>
node tools/ground-truth/paint/colour/bases/analyse.ts <dir> --fixture corpus/ground-truth/color-bases.json
```

Unlike C, the readback is scripted: `tools/ground-truth/paint/colour/bases/read.ps1` opens each deck with
`Open2007(OpenAndRepair:=msoFalse)`, reads `Fill.ForeColor.RGB` and `Fill.Transparency`, and exports
every slide as a BMP. Transparency is why the object model is read at all — a translucent swatch's
pixel is a composite with the swatch behind it.

**One question per package, and that is a finding rather than a preference.** The first run grouped
the probes into six topical decks; one came back _repaired_, with the only sentence PowerPoint ever
says and no indication which of sixty-seven swatches caused it. A repaired deck measures the repair.
Split into 21 single-question packages, the same run localised it immediately: an `a:hslClr/@hue`
outside `ST_PositiveFixedAngle`. Those swatches are flagged `fromRepairedDeck` in the fixture and
excluded from every assertion.

### C3 — gradients, and the 54 pattern tiles _(added in 2.7)_

Two halves, and the cheap one came first.

**Ask PowerPoint to author the fills** — `tools/ground-truth/paint/fills/author.ps1`, then unzip
`ppt/slides/slide1.xml` from what it saved. `Fill.Patterned(i)` over the whole `MsoPatternType` range,
`Fill.TwoColorGradient` over every style and variant, and the 24 preset gradients — then read what
PowerPoint wrote. One script, and it settled three things before a probe existed: exactly 54 pattern
values are accepted and PowerPoint spells each name into `a:pattFill/@prst` itself; a two-colour
gradient is written with **two** stops, not 33; and `a:gsLst` comes out **unsorted** in PowerPoint's
own from-centre fill. Nothing else in this directory is as high-yield per line.

**Then measure what it paints.**

```bash
node tools/ground-truth/paint/fills/build-deck.ts <dir>
powershell -File tools/ground-truth/paint/fills/read.ps1 -Dir <dir>
node tools/ground-truth/paint/fills/analyse.ts <dir> --fixture corpus/ground-truth/fills.json
npx prettier --write corpus/ground-truth/fills.json
node tools/ground-truth/paint/fills/write-tables.ts
```

The last step regenerates the two tables `@pptx-studio/paint` ships. Neither is hand-written and
`fill.test.ts` re-derives both from the fixture, so a measured constant cannot quietly become
somebody's memory of one.

Three things about the rig are worth copying rather than rediscovering:

- **A gradient has no single colour**, so unlike C and C2 the object model is nearly useless here —
  `Fill.ForeColor.RGB` returns one number for a whole ramp. The bitmap is the oracle. COM is kept
  for two things it does better: `Fill.GradientStops` proves the file order, and `Fill.Pattern`
  round-trips the enumeration.
- **Every rectangle is a whole number of points.** The slide is 960 × 540 pt and one point is 12700
  EMU, so at 960, 1920 or 3840 pixels wide every shape edge lands on a pixel boundary. Without that,
  an edge falls mid-pixel and the edge pixel is a blend with the slide behind it.
- **A pattern tile can only be read at 1280 × 720**, which over a 13.333-inch slide is exactly 96
  DPI. The tile turns out to be a fixed physical size — 8 px at 96 DPI — so at 1920 it is resampled
  to 12 px and every diagonal picks up a grey edge. The other three widths are exported anyway,
  because the _period_ at four resolutions is what proves the tile is physical rather than
  device-fixed or shape-relative.

Fifteen of the 144 probes are hostile and each is alone in its package. Seven were refused, and
because of that split each refusal names its own cause instead of taking seven other answers with it.

### C4 — the stroke, and the things drawn around it _(added in 2.8)_

Same two halves as C3, and the cheap one first again.

**Ask PowerPoint to author the strokes** — `tools/ground-truth/paint/lines/author.ps1`, then unzip
`ppt/slides/slide1.xml` from what it saved. `Line.DashStyle` over the whole
`MsoLineDashStyle` range, `Line.Style`, `Line.InsetPen`, the arrowhead triples, and
`Shadow.Type` over `msoShadow1`..`msoShadow43` plus glow, soft edge and reflection.
One script, and it settled six things before a probe existed — among them that two of the twelve
dash styles differ **only by `@cap`**, that PowerPoint writes `@cap` for nothing else,
and that twenty of the forty-three shadow presets save as `a:prstShdw`, an element no renderer
this project has read implements.

**Then measure what it paints.**

```bash
node tools/ground-truth/paint/lines/build-deck.ts <dir>
powershell -File tools/ground-truth/paint/lines/read.ps1 -Dir <dir>
node tools/ground-truth/paint/lines/analyse.ts <dir> --fixture corpus/ground-truth/lines.json
npx prettier --write corpus/ground-truth/lines.json
node tools/ground-truth/paint/lines/write-tables.ts
```

The last step regenerates the three tables `@pptx-studio/paint` ships — the dash arrays, the
compound divisions and the arrowheads. None is hand-written and `line.test.ts` re-derives all
three from the fixture.

Four things about this rig are worth copying rather than rediscovering:

- **A position is a better measurement than a pixel.** An antialiased edge crosses half coverage
  somewhere _inside_ one pixel, and interpolating between the two straddling samples locates it to
  about a tenth of one. So the fixture stores sub-pixel **crossings**, in points, and keeps a raw
  profile only where the answer really is a shape — a blur's falloff, an arrowhead's silhouette.
  That is what makes a four-width dash segment measurable to better than half a percent.
- **Placement is absolute, not a grid.** C3 let the builder place each probe. A stroke is drawn
  _around_ its geometry, so half of what is being measured is outside the shape's own rectangle: a
  square cap on a 24pt line paints 12pt past the endpoint and a 24pt blur reaches 24pt past the
  shape. Every probe states its rectangle **and its read window** in points on the slide.
- **Leave more room than looks necessary.** Two separate rounds of this experiment measured the
  neighbouring shape instead of the intended one — a 36pt compound stroke reaches 22pt outside its
  own rectangle, and a doubled shadow grew up into the row above's window. Both produced
  plausible-looking numbers rather than errors.
- **Fit the model; do not threshold it.** An ellipse's silhouette reaches the shaft's own height well
  inside its back edge, so thresholding one reports every oval about nine percent short. Fitting the
  outline family recovers the extent and the residual says whether the family was right.

Fifteen of the 213 probes are hostile and each is alone in its package. **Thirteen were refused**,
against seven of fifteen in C3: every enumeration in the stroke vocabulary is policed, and the only
two that get through are both `a:custDash` degeneracies.

### C5 — what a shape inherits, and from where _(added in 2.9)_

Two halves again, and this time the cheap half answered the sub-phase's central architectural
question outright.

**Ask PowerPoint to author the inheritance** — `tools/ground-truth/model/author.ps1`, then unzip the slide,
layout and master parts from what it saved. Three slides on one layout: one untouched, one typed
into, one nudged a single point. **Only the nudged one has an `a:xfrm`**, and it has the whole
resolved rectangle, inherited two hops and baked in at once. Typing does not create geometry. That
is architectural bet 3 confirmed by the format's own author, for the cost of one script — along with
the eleven stock layouts' placeholder sets, the forty-two theme shape styles, and a two-master deck
whose 22 layouts sit in one flat folder.

**Then measure what it resolved.**

```bash
node tools/ground-truth/model/build-deck.ts <dir>
powershell -File tools/ground-truth/model/read.ps1 -Dir <dir>
node tools/ground-truth/model/analyse.ts <dir> --fixture corpus/ground-truth/sheets.json
npx prettier --write corpus/ground-truth/sheets.json
```

**The measurement is a position, and PowerPoint reports it directly.** C3 and C4 sampled bitmaps
because a fill and a stroke are pictures; an inheritance is not. A placeholder with no `a:xfrm`
of its own still has a position, `Shape.Left` reports it in points exactly, and that is the
resolver's own answer read out of the resolver rather than reconstructed from what it painted. So
every candidate parent gets a rectangle no other candidate shares — five bands of seven boxes, 80
points apart in y and 130 in x — and **the box a probe lands in names the parent it matched**. One
number, no fitting, no error bars, and the whole run takes about a minute. Bitmaps are exported
anyway, for a human to look at; nothing depends on them.

Three things about this rig are worth copying rather than rediscovering:

- **A repaired deck is not a measurement.** Four clean-looking decks came back REPAIRED and their
  readings were of files PowerPoint had rewritten. Bisecting eight variants localised it to one
  attribute value: a slide master may carry only `title`, `body`, `dt`,
  `ftr`, `sldNum` and `hdr`. A second bisection found that every master must own
  its own theme part. Both rules are now hostile probes, so nobody has to find them twice.
- **A probe must not declare the thing it is asking about.** The two-hop _style_ probes first
  reported the layout's magenta position marker rather than the master's style, because the shared
  placeholder helper carried a fill. A plausible colour, and the wrong question answered.
- **Two identical overrides cannot tell you which one applied.** The colour-map precedence probe had
  to be rebuilt with a slide override that _disagrees_ with its layout's before it said anything.

Twenty-three of the 37 packages are hostile and each is alone in its own file; 16 were refused or
repaired.

### C6 — what a group does to the shapes inside it _(added in 2.10)_

Third sub-phase running that the cheapest thing in this directory was the highest-yield.

**Ask PowerPoint in writing, first** — `tools/ground-truth/render/transforms/author.ps1`, then unzip what it saved. `a:xfrm`
carries `@rot`, `@flipH` and `@flipV` and says nothing about the order they compose in, and the two
readings are not equivalent: for any reflection `F R(t) F = R(-t)`, so they differ by the **sign of
the angle**. Hand PowerPoint a shape already rotated 30 degrees, ask it to mirror the shape, and read
back what it writes. It wrote `rot="19800000" flipH="1"` — minus thirty — at 30, 45, 120 and 200
degrees and on both axes. Only a renderer that flips _before_ it rotates has to negate the angle to
reproduce a mirrored figure. One script, and the sub-phase's central question was settled before a
probe existed — along with what a group resize writes (the group's `ext`, and nothing else) and what
`Ungroup` bakes.

**Then measure what a group does.**

```bash
node tools/ground-truth/render/transforms/build-deck.ts <dir>
powershell -File tools/ground-truth/render/transforms/read.ps1 -Dir <dir>
node tools/ground-truth/render/transforms/analyse.ts <dir> --fixture corpus/ground-truth/transforms.json
npx prettier --write corpus/ground-truth/transforms.json
pnpm build && node tools/ground-truth/render/transforms/verify-render.ts <dir>
```

The measurement is C5's again and cheaper still: a leaf inside a group has no slide position of its
own, the group's transform gives it one, and `Shape.Left` reports it. Confirmed against PowerPoint's
own output to four decimal places — a child of a group turned 30 degrees and mirrored reported
`L=321.2435302734375`, and `x="4079793"` appears in the file once that group is ungrouped. So 46 of
the 65 probes need no bitmap at all.

Four things worth copying:

- **Read the shapes twice.** `GroupItems(i).Rotation` reports the _composed_ rotation while
  `HorizontalFlip` reports the child's own attribute. The reader therefore ungroups everything **in
  memory** and reads again; nothing is saved, and `Saved` is forced true before the close.
- **A colour is not an answer; a position is.** The `a:grpFill` probes read a colour and needed a
  _rectangle_ out of it. Every gradient slide carries a reference shape with the identical gradient
  over a known rectangle, drawn behind everything, so a sampled colour turns back into a fraction
  along the ramp — and two fractions solve for the span outright. Without it the first run compared a
  `grpFill` child against the group's own ground and found white, because a group's fill is not
  painted at all.
- **One angle is not a rule.** At 45 degrees a rotated child in a non-uniformly scaled group came
  back with its two extents swapped. Eighteen angles turned that into a quadrant snap that rounds
  halves up, and refuted the two `|sin|`-versus-`|cos|` forms anybody would write first.
- **Close the loop the other way.** `verify-render.ts` renders the same decks with `render-svg`,
  rasterises them in Chromium at PowerPoint's export size, and reads PowerPoint's own sample points.
  It found a bug nothing else did: `svgStops` returns a fraction and the renderer divided it by a
  hundred thousand again, collapsing every gradient to its last stop. Seventeen samples at once, and
  not one unit test.

### T5 — bullets, fields and script runs _(added in 3.5)_

The cheap half first again, and it settled three whole vocabularies before a probe existed.
`tools/ground-truth/text/bullets/author.ps1` walks `PpNumberedBulletStyle` from 0 to 47 and saves: 0 to 40 are accepted and
each writes its own `a:buAutoNum/@type`, so the 41 scheme names and their order are PowerPoint's own
rather than a transcription. The same script settles that `a:buChar/@char` keeps a private-use code
point verbatim, that "use the text's colour" is expressed by writing **nothing**, that a numbered
list carries **no `startAt` anywhere** — the numbers are computed at layout time and are in no file —
and that a field's cached `a:t` is discarded on open.

```bash
node tools/ground-truth/text/bullets/build-deck.ts <dir>
powershell -File tools/ground-truth/text/bullets/read.ps1 -Dir <dir>
node tools/ground-truth/text/bullets/analyse.ts <dir> --fixture corpus/ground-truth/bullets.json
npx prettier --write corpus/ground-truth/bullets.json
node tools/ground-truth/text/bullets/write-tables.ts
```

**The instrument is new, and it is the point.** `Slide.Export(path, "EMF")` records PowerPoint's own
GDI calls, so a bullet comes back as the characters `ExtTextOutW` was given, with the face, size and
colour it was given them in. Every earlier experiment here fitted a model to a measurement; this one
mostly reads the answer. `a:buAutoNum type="romanLcParenBoth"` at 4 is `"(iv)"` because that is the
string in the file PowerPoint wrote, not because `(iv)` scored better than `(iiii)`.

Five things about the rig worth copying rather than rediscovering:

- **PowerPoint refuses the SVG filter by name** and exports EMF and WMF happily. EMF is the one with
  32-bit coordinates and `ExtTextOutW` rather than the ANSI-only original.
- **GDI splits a run at digit and script boundaries.** The marker `#1925#` comes back as three
  records — `"#"`, `"1925"`, `"#"` — so a record-level marker test matches nothing. The stream has to
  be flattened to one character per entry and the markers found in the concatenation. The first
  version of the analysis lost every field probe silently, because a rendered date splits at every
  slash.
- **Some text is not text.** `ETO_GLYPH_INDEX` means the record holds glyph ids, and PowerPoint sets
  it for everything it shapes — the two Arabic autonumber schemes, and four of the sixteen field
  locales. Reading those as characters gives a confident wrong answer for two of forty-one schemes,
  so the flag is carried and those rows are scored differently or not at all.
- **Keep a second instrument for what the first cannot see.** A picture bullet draws no text at all,
  so `Paragraphs(i).BoundLeft` is its only measurement; a field in a complex script is only readable
  through `TextRange.Text`.
- **One instant cannot tell a component from a coincidence.** Turning a rendered date back into a
  format pattern round-trips perfectly and is still wrong: the Umm al-Qura year 1448 derived as
  `h448` because the hour was one o'clock. Three guards were each paid for by a wrong answer that
  looked right.

Nine of the 44 packages are hostile and each is alone in its file; all nine were repaired, which is
the finding — `startAt` outside 1..32767 and `buSzPct` outside 25000..400000 are values PowerPoint
rewrites rather than reads.

### T6 - the text frame _(added in 3.6)_

```bash
powershell -File tools/ground-truth/text/frames/author.ps1 -Dir <dir>
node tools/ground-truth/text/frames/build-deck.ts <dir>
powershell -File tools/ground-truth/text/frames/read.ps1 -Dir <dir>
node tools/ground-truth/text/frames/analyse.ts <dir> --fixture corpus/ground-truth/frames.json
```

Where a block of text sits inside a shape, how far in from each edge, and which way it runs.
601 probes over 50 packages, 21 questions, each fitting exactly one candidate.

**Every question is a difference between two probes that differ in one attribute.** `BoundWidth`
overstates an advance by about 5pt - a centred `Ag CJK` at Arial 18 sits where a 61.0pt advance
puts it while reporting 66.0 - so the control probe carries the unknown and the difference cancels
it. `BoundHeight` is exact and the `control` family keeps checking that.

**Two of the five anchoring values are unreachable from any user interface.** PowerPoint answers
`msoAnchorTopBaseline` with "The specified value has been deprecated", so `just` and `dist` had to
be read off the layout rather than authored. They are bottom anchoring, and they stretch nothing.

**`lfEscapement` is always zero.** The experiment was designed around it and PowerPoint rotates
with a world transform instead; the `@`-prefixed face name and the per-glyph record count carry
the vertical answer in its place.

## The layout

One directory per experiment, grouped by **the package the experiment answers for**. That is the
only organising principle that says anything useful here: an experiment exists to produce a fixture,
and the fixture exists for one package.

```
lib/                     the harness every experiment shares
  zip.ts                 a small ZIP reader/writer, deliberately not @pptx-studio/opc
  pptx.ts                writes a minimal PresentationML package by hand
  sheet-pptx.ts          a package builder with more than one master and layout
  bmp.ts                 sample a pixel out of PowerPoint own bitmap export
  emf.ts                 enough of EMF to read what PowerPoint drew, in text

fonts/                   -> packages/fonts (phase 8)
  format/                sfnt.ts, eot.ts - the containers, read and written
  embedding/             experiments A and B

paint/                   -> packages/paint
  colour/transforms/     C   - 214 swatches: tint, shade, lumMod, satMod
  colour/bases/          C2  - 259 more: the six bases, alpha, clrMap, percentages
  fills/                 C3  - 144 probes: gradients, the 54 tiles, 15 hostile
  lines/                 C4  - 213 probes: dashes, caps, joins, arrowheads

model/                   -> packages/model.   C5  - 114 probes: matching, inheritance
render/                  -> packages/render-*
  transforms/            C6  - 65 probes: group maps, turns, compositing
  text/                  T8  - 160 probes: the turn, the baseline, alignment, rules

text/                    -> packages/text
  cascade/               T1  - which of the nine claimed sources is read, in what order
  metrics/               T2  - 2091 probes: the line box, spacing, kerning, tracking
  line-breaks/           T3  - where a line breaks, and what kinsoku does to it
  autofit/               T4  - 3521 probes: the ladder, and what a last line measures
  bullets/               T5  - 6849 probes: 41 schemes, PUA bullets, fields, script runs
  frames/                T6  - anchors, insets, vertical text, a:br, endParaRPr

fonts/                   -> packages/text/src/fonts, packages/fonts, packages/cli
  format/                A   - what PowerPoint writes into ppt/fonts/*.fntdata
  embedding/             B   - whether PowerPoint renders an EOT we built
  substitution/          T7  - 159 probes: which face a run is actually drawn in
  metrics/               T13 - 514 probes: the box, kerning, widths, the baseline
  real-faces/            T14 - the same reader, on faces nobody built for it

fixtures.test.ts         the committed answers, as assertions CI can run
```

Every experiment has the same file names for the same roles, and only the ones it needs:

| file              | what it is                                                          |
| ----------------- | ------------------------------------------------------------------- |
| `probes.ts`       | the probe table: what is asked, and why each probe is there         |
| `build-deck.ts`   | writes the probe packages and the `*-inputs.json` describing them   |
| `author.ps1`      | makes PowerPoint author the same thing, so it names its own values  |
| `read.ps1`        | reads the decks back over COM, and exports what is needed           |
| `analyse.ts`      | scores the candidate models and emits the fixture                   |
| `write-tables.ts` | turns the fixture into a table a package ships                      |
| `score.ts`        | the verdict as a pure function, where the experiment is also a gate |

`analyse.ts` throws rather than emitting a fixture it cannot fit perfectly. That is the point of
the whole directory: a model that scores 46 of 46 is a finding, and one that scores 45 is a
question nobody has answered yet.

### T8 — where the glyphs go _(added in 3.8)_

3.1 to 3.7 settled where a _line_ goes. T8 asks the drawing questions: which way a mirrored shape's
text faces, where in the line box the baseline sits, what an underline is a fraction of.

```bash
node tools/ground-truth/render/text/build-deck.ts <dir>
powershell -File tools/ground-truth/render/text/read.ps1 -Dir <dir>
node tools/ground-truth/render/text/measure-in-browser.ts <dir>
node tools/ground-truth/render/text/analyse.ts <dir> --fixture corpus/ground-truth/text-rendering.json
node tools/ground-truth/render/text/write-tables.ts
pnpm build && node tools/ground-truth/render/text/verify-render.ts <dir>
```

The instrument is the EMF, and further into it than any experiment before: `EMR_EXTTEXTOUTW` under
`TA_BASELINE` gives the baseline as a number, the world transform in force gives the angle and its
determinant says whether the glyphs are mirrored, and `offDx` gives every character's advance.

Three things about the rig are worth copying rather than rediscovering:

- **Track `EMR_SAVEDC` and `EMR_RESTOREDC`.** Without them the world transform accumulates across
  drawing calls and every text record after the first reports an origin twice too far from the page
  corner. It looks like a finding.
- **A rotated frame must fit the page.** The first run put a 600 x 240pt frame near the top left;
  turned 90 degrees it hangs off the slide, and four probes exported no ink at all — which reads as
  "PowerPoint drew nothing" rather than "the export clipped it". The turn questions now use a
  400 x 160 frame centred on the slide, whose half-diagonal fits at every angle.
- **Off a quadrant there is no text record.** PowerPoint draws the glyphs as filled paths, so a 30
  degree probe's angle has to be fitted from where the ink's centre landed - and the ink box's
  _extents_ are the same at an angle and its negation, so it is the offset from the frame centre
  that separates 30 degrees from 330.

`verify-render.ts` is the other half, and the one that matters: it loads the same seven packages
through `opc` and `model` in a real Chromium, lays each probe out with `render-svg`, and compares
every line box against the one PowerPoint reported. 415 of 423 land within a tenth of a point; the
eight that do not are the left edge of a centred or right-aligned line, carrying T2's measured
browser disagreement whole.

### T13 — what a font-table reader must do to agree with the browser _(added in 3.10)_

The only experiment here that never opens PowerPoint. `pptx-studio render` measures text in Node,
where there is no canvas, so it reads the face's own tables instead — and every one of those reads
has more than one plausible source. **No real font can say which is right**, because a real font's
`hhea`, `usWin` and `sTypo` metrics are the same numbers. So the probes are eleven fonts built to
disagree with themselves on purpose, loaded into Chromium as data URIs under family names that exist
nowhere.

```bash
node tools/ground-truth/fonts/metrics/measure-in-browser.ts <dir>
pnpm build   # analyse scores the built reader, which is what an npm install gets
node tools/ground-truth/fonts/metrics/analyse.ts <dir> --fixture corpus/ground-truth/font-metrics.json
```

Four answers, each with every rival scored:

- The font bounding box is **`OS/2.usWinAscent`/`usWinDescent`**, and **`sTypoAscender`/`sTypoDescender`
  when `fsSelection` bit 7 is set** — 22/22, where reading `usWin` alone scores 20, `sTypo` alone 12,
  and `hhea`, which is the first thing anyone reaches for, scores 10.
- Pair kerning comes from **GPOS when the font has a `kern` feature, and the legacy `kern` table
  otherwise** — 66/66. The font that carries both, saying −200 in one and −100 in the other, is the
  only probe that can separate them: preferring the legacy table scores 60.
- A width is **each glyph advance truncated to 1/65536 px, plus each kern adjustment rounded to the
  same step** — 396/396, against 136 for exact float arithmetic and 388 for truncating both. The
  advance is linear in size to within 1.04e-4 px, which is what makes a table reader viable at all:
  the browser does not hint it.
- The ideographic baseline is **the `BASE` table's `ideo` coordinate for the `DFLT` script, and the
  face box descent where the font names none** — 30/30, against 22 for the descent alone. Chromium
  reads `DFLT` and no other script, does not fall through to one when `DFLT` is absent, and hands
  back a coordinate outside the em unchanged.

`analyse.ts` throws rather than emit a fixture where two readings tie, and it scores the reader in
`packages/cli/dist` rather than a copy of its logic, so the fixture is a statement about the shipped
code.

### T14 — the same reader, on faces nobody built for it _(added in 3.10)_

`metrics/` asks with fonts built to disagree with themselves; `real-faces/` asks with the faces a
machine has. That is the half T13 cannot reach: class-based GPOS kerning, GSUB ligatures, `hmtx`
tables with thousands of entries, and a `unitsPerEm` that is not a round 1000. Eleven samples across
seven scripts at six sizes, against every face in the directories named on the command line.

```bash
node tools/ground-truth/fonts/real-faces/measure-in-browser.ts <dir> --font-dir <fonts>
node tools/ground-truth/fonts/real-faces/analyse.ts <dir> --summary <dir>/summary.md
```

Nothing here is committed as a fixture. A digest over whatever fonts a runner happens to carry would
be red the week the image changes its font package, so the answer is a job artifact and the number
is dated in the ADR instead. `--font-dir` is required and there is no system fallback: the job says
which faces it read.

Latin, Cyrillic, Greek, Hebrew and CJK are gated at 1e-4 relative — a fifteenth of the 0.149% by
which the same browser disagrees with PowerPoint in T2. Arabic, Devanagari and the `liga`
substitutions are **recorded rather than gated**, because the reader sums unshaped advances and a
shaper does not; that ratio is what turns "no shaping" from a caveat into a number. No gated sample
contains an `f`, since `liga` is on by default in `measureText` and would otherwise score the reader
against a shaper it never claimed to be.

## Where an experiment reads the browser too

`fonts/substitution/measure-in-browser.ts` (T7), `text/metrics/measure-in-browser.ts` (T2) and
`render/text/measure-in-browser.ts` (T8) ask Chromium the same questions PowerPoint answered,
because for anything the product measures at run time the browser's answer is the one that ships.
T7's whole detector question can only be scored there: PowerPoint has no opinion about whether
`document.fonts.check` works. T8's decoration question is the same shape from the other side — it
reads Chromium's own `text-decoration` off a screenshot, and the answer is that it agrees with
PowerPoint on none of eight faces, which is why both rules are drawn as geometry.

T13 goes further and asks **only** the browser. Its subject is what a second measurement engine has
to do to agree with the first, and PowerPoint is not a party to that question.

## Driving PowerPoint

From PowerShell, via COM. Two things that cost time:

- `Presentation.SaveAs(path, format, EmbedFonts)` — the third argument is
  `MsoTriState`, and passing `msoFalse` there is how you save a deck with no
  embedded fonts while believing you asked for some.
- PowerPoint has **no** `SaveSubsetFonts`. That is Word's API. The setting lives
  in `p:presentation/@saveSubsetFonts` in the file, which is the right place to
  change it anyway — it does not touch the machine's settings.
