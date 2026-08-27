# ground-truth

Experiments that ask the real Microsoft PowerPoint on this machine what it
actually does, rather than what a specification says it should.

Everything here is run **by hand**, needs PowerPoint installed, and is not part
of `pnpm check`. That is not laziness — a test that requires a licensed
application cannot gate a pull request. What CI can see is the _answers_, which
are committed as fixtures in [`corpus/ground-truth/`](../../corpus/ground-truth),
and the write-up in
[`docs/adr/0007-ground-truth.md`](../../docs/adr/0007-ground-truth.md).

## The three experiments

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

## The files

| file              | what it is                                                     |
| ----------------- | -------------------------------------------------------------- |
| `zip.ts`          | a small ZIP reader/writer, deliberately not `@pptx-studio/opc` |
| `sfnt.ts`         | enough of the SFNT container to cross-examine an EOT header    |
| `eot.ts`          | EOT read and write, versions 1 and 2.2, compressed or not      |
| `pptx.ts`         | writes a minimal PresentationML package by hand                |
| `swatches.ts`     | the 214 colour swatches and why each one is there              |
| `color-models.ts` | the candidate models, side by side, so one can win             |
| `bmp.ts`          | sample a pixel out of PowerPoint's bitmap export               |
| `build-font.ts`   | synthesises the probe font from nothing                        |

## Driving PowerPoint

From PowerShell, via COM. Two things that cost time:

- `Presentation.SaveAs(path, format, EmbedFonts)` — the third argument is
  `MsoTriState`, and passing `msoFalse` there is how you save a deck with no
  embedded fonts while believing you asked for some.
- PowerPoint has **no** `SaveSubsetFonts`. That is Word's API. The setting lives
  in `p:presentation/@saveSubsetFonts` in the file, which is the right place to
  change it anyway — it does not touch the machine's settings.
