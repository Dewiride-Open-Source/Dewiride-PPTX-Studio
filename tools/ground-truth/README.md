# ground-truth

Experiments that ask the real Microsoft PowerPoint on this machine what it
actually does, rather than what a specification says it should.

Everything here is run **by hand**, needs PowerPoint installed, and is not part
of `pnpm check`. That is not laziness — a test that requires a licensed
application cannot gate a pull request. What CI can see is the _answers_, which
are committed as fixtures in [`corpus/ground-truth/`](../../corpus/ground-truth),
and the write-up in
[`docs/adr/0007-ground-truth.md`](../../docs/adr/0007-ground-truth.md).

## The experiments

### A — what does PowerPoint write into `ppt/fonts/*.fntdata`?

Save a deck twice, with `p:presentation/@saveSubsetFonts` flipped in between, and
read the EOT headers.

```bash
node tools/ground-truth/dump-fntdata.ts <deck.pptx> ... --json corpus/ground-truth/eot-headers.json
```

The one bit that matters is `TTEMBED_TTCOMPRESSED` (0x4) in the flags word at
offset 0x0C. It is set, always.

### B — will PowerPoint render an EOT we built?

```bash
node tools/ground-truth/build-font-deck.ts <dir> [<a-powerpoint-deck.pptx>]
node tools/ground-truth/verify-font-in-browser.ts <dir>     # is the FONT valid?
# open <dir>/probe-deck.pptx in PowerPoint, export slide1 as BMP, then:
node tools/ground-truth/analyse-font-deck.ts <dir>
```

`graft-font.ts` puts one of those EOTs into a package PowerPoint itself wrote,
which is how the experiment separates "our EOT is wrong" from "our deck is
wrong":

```bash
node tools/ground-truth/graft-font.ts <pp-deck.pptx> <probe.eot> <typeface> <out.pptx>
```

**Read the controls before trusting a result.** The first run of this experiment
reported a pass that was not one: the deck opened cleanly, `Font.Embedded`
returned `msoTrue`, and the slide showed eight perfectly good letters — from a
substituted font. Only the row in a deliberately-absent typeface gave it away.

### C — which colour space does DrawingML compute in?

```bash
node tools/ground-truth/build-swatch-deck.ts <dir>
# open <dir>/swatch-deck.pptx in PowerPoint, read colours back and export BMPs
node tools/ground-truth/analyse-swatches.ts <dir> --fixture corpus/ground-truth/color-transforms.json
```

### C2 — the colour questions C did not ask _(added in 2.6)_

Every swatch in experiment C was built on `a:srgbClr` or `a:schemeClr`, on purpose, so a
disagreement was attributable to the transform rather than to the base. That left four of the six
bases, all three alpha operators, any non-identity `clrMap` and the whole percentage grammar
unmeasured. C2 measures them.

```bash
node tools/ground-truth/build-swatch-deck2.ts <dir>
powershell -File tools/ground-truth/read-swatches2.ps1 -Dir <dir>
node tools/ground-truth/analyse-swatches2.ts <dir> --fixture corpus/ground-truth/color-bases.json
```

Unlike C, the readback is scripted: `read-swatches2.ps1` opens each deck with
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

**Ask PowerPoint to author the fills** — `author-fills.ps1`, then unzip
`ppt/slides/slide1.xml` from what it saved. `Fill.Patterned(i)` over the whole `MsoPatternType` range,
`Fill.TwoColorGradient` over every style and variant, and the 24 preset gradients — then read what
PowerPoint wrote. One script, and it settled three things before a probe existed: exactly 54 pattern
values are accepted and PowerPoint spells each name into `a:pattFill/@prst` itself; a two-colour
gradient is written with **two** stops, not 33; and `a:gsLst` comes out **unsorted** in PowerPoint's
own from-centre fill. Nothing else in this directory is as high-yield per line.

**Then measure what it paints.**

```bash
node tools/ground-truth/build-fill-deck.ts <dir>
powershell -File tools/ground-truth/read-fills.ps1 -Dir <dir>
node tools/ground-truth/analyse-fills.ts <dir> --fixture corpus/ground-truth/fills.json
npx prettier --write corpus/ground-truth/fills.json
node tools/ground-truth/write-paint-tables.ts
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

## The files

| file                    | what it is                                                     |
| ----------------------- | -------------------------------------------------------------- |
| `zip.ts`                | a small ZIP reader/writer, deliberately not `@pptx-studio/opc` |
| `sfnt.ts`               | enough of the SFNT container to cross-examine an EOT header    |
| `eot.ts`                | EOT read and write, versions 1 and 2.2, compressed or not      |
| `pptx.ts`               | writes a minimal PresentationML package by hand                |
| `swatches.ts`           | the 214 colour swatches and why each one is there              |
| `swatches2.ts`          | the 259 C2 adds, and the question each block settles           |
| `fills.ts`              | the 144 C3 probes: gradients, tiles, and fifteen hostile ones  |
| `write-paint-tables.ts` | turns the C3 fixture into the tables `paint` ships             |
| `author-fills.ps1`      | makes PowerPoint write the fills, so it names its own values   |
| `color-models.ts`       | the candidate models, side by side, so one can win             |
| `bmp.ts`                | sample a pixel out of PowerPoint's bitmap export               |
| `build-font.ts`         | synthesises the probe font from nothing                        |

## Driving PowerPoint

From PowerShell, via COM. Two things that cost time:

- `Presentation.SaveAs(path, format, EmbedFonts)` — the third argument is
  `MsoTriState`, and passing `msoFalse` there is how you save a deck with no
  embedded fonts while believing you asked for some.
- PowerPoint has **no** `SaveSubsetFonts`. That is Word's API. The setting lives
  in `p:presentation/@saveSubsetFonts` in the file, which is the right place to
  change it anyway — it does not touch the machine's settings.
