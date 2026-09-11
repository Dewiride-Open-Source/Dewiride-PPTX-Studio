# @pptx-studio/cli

**Command line tools for PPTX packages.** As of sub-phase 3.10 that is five
verbs: `render`, which draws slides as SVG with no browser and no LibreOffice;
`inspect`, which reads a deck and tells you what is in it; `validate`, which
checks it against the twenty-nine rules a `.pptx` must not break; `roundtrip`,
which reads a deck, writes it back, and proves nothing moved; and `bisect`,
which narrows a broken deck down to the change that breaks it.

Apache-2.0 · part of
[PPTX Studio](https://github.com/Dewiride-Open-Source/Dewiride-PPTX-Studio).

```bash
npx @pptx-studio/cli render deck.pptx --out slides/
npx @pptx-studio/cli render deck.pptx --slide 1 --width 640 --out thumb.svg
npx @pptx-studio/cli render deck.pptx --json --font-dir ./fonts

npx @pptx-studio/cli inspect deck.pptx
npx @pptx-studio/cli inspect deck.pptx --parts --namespaces
npx @pptx-studio/cli inspect deck.pptx --json --out census.json

npx @pptx-studio/cli validate deck.pptx
npx @pptx-studio/cli validate deck.pptx --explain

npx @pptx-studio/cli roundtrip deck.pptx
npx @pptx-studio/cli roundtrip deck.pptx --write out.pptx

npx @pptx-studio/cli bisect deck.pptx
npx @pptx-studio/cli bisect original.pptx broken.pptx --oracle powerpoint
```

```
PACKAGE
=======
  archive        188.5 MiB in 1616 entries
  parts          995 (192.0 MiB inflated, 1.0:1 overall)
  compression    1303 deflated, 313 stored
  largest part   /ppt/media/image1.png (639.5 KiB)

PRESENTATION
------------
  slide size     13.33 x 7.50 in (screen16x9)
  sheets         300 slides, 4 layouts, 1 masters
  notes          300 notes slides, 1 notes masters, 0 handout masters
  structure      6 sections, 2 custom shows
  fonts          embedTrueTypeFonts=true saveSubsetFonts=true
                 ProbeAlpha                  regular  charset 0

FEATURES
--------
     5194  Shapes                            phase 2.10
       30  Tables                            phase 4.1
       12  Charts                            phase 9.1
       60  Animation timelines               carried across, never rendered
```

Exit status is **1** when the census found a structural error and **0**
otherwise. Warnings and notes never fail the command, so `inspect` answers
"did this deck load, and is anything in it structurally broken" as a yes/no
question a script can branch on.

## `validate`

```
  /ppt/slides/slide1.xml
    V022  fatal
      /p:sld/p:cSld/p:spTree/p:sp/p:nvSpPr/p:nvPr/p:ph/@type
      type="hdr" is a whole-package refusal here, alone and with no other change. …

  V027 did not run: it compares against the package as it was opened, and none was supplied. …

validate: 26 rule(s), 1 finding(s), 1 blocking
```

Exit status is **1** when anything fatal was found and **0** otherwise; a
warning never fails the command.

Three of the twenty-nine rules compare a package against the package **as it was
opened**, and a file on the command line has no such history. They are skipped,
and the report names them rather than counting them as passes. That makes
`validate` a diagnostic; the export gate is `assertValid`, which the writer calls
with both packages in hand.

| flag        | what it does                                   |
| ----------- | ---------------------------------------------- |
| `--json`    | the report as JSON                             |
| `--explain` | append the rationale for every rule that fired |
| `--quiet`   | fatal findings only                            |
| `--out`     | write to a file instead of stdout              |

## `roundtrip`

```
roundtrip deck.pptx

  read      47048 bytes, 29 entries
  written   47048 bytes, 29 entries
  export    0 part(s) re-serialized, 28 streamed
  compared  28/28 parts identical (19 xml, 9 rels, 0 binary)

  no differences
```

Exit status is **1** when the two packages differ — and also when the export was
refused, or the file would not open at all, because CI wants one bit. The three
are distinguished on stderr, since each needs completely different work next.

**It is not a byte comparison, and it must not become one.** Entry order,
deflate level, timestamps and attribute order all differ legitimately between two
archives holding the same document. Nine of the corpus's PowerPoint-authored
decks come out exactly 1832 bytes smaller than they went in — every stored byte
of every entry identical — because Office writes a `0xA220` growth-hint extra
field on five of their entries and we do not. A byte-equality gate would have
been red on ten of fifty-one decks before a line of the renderer existed, over
padding.

What is compared instead: the canonical XML of every XML part, the relationship
graph with ids treated as opaque labels, and the SHA-256 of everything else. So a
deck PowerPoint has re-saved — which renumbers every `rId` — still compares
equal, while one where two `r:embed` values swapped what they point at does not.

| flag             | what it does                                   |
| ---------------- | ---------------------------------------------- |
| `--json`         | the comparison as JSON, with a digest per part |
| `--quiet`        | the differences only, without the tally        |
| `--write <file>` | also save the package that was written         |
| `--out <file>`   | write the report to a file instead of stdout   |

`--write` is there for the check nothing in this repository can make: open the
result in real PowerPoint. `bisect --oracle powerpoint` scripts it.

## `bisect`

```
bisect deck.pptx                       # against our own export of it
bisect original.pptx broken.pptx       # against a package from somewhere else
```

The debugger for this project. PowerPoint's whole diagnostic channel is one
sentence naming no part, no element and no reason, so the only way to find out
what it objects to is to ask again with less of the change present, and keep
asking. This is delta debugging — Zeller's `ddmin` over the set of changes
between the two packages, applied level by level down the tree.

```
  delta     135 change(s) in 7 entry(s)
  oracle    powerpoint, 5 run(s)

  1 change(s) in 1 entry(s), from a delta of 135, in 5 oracle run(s)

  removed  /[Content_Types].xml /Types/Default[3]
    -  <Default Extension="fntdata" ContentType="application/x-fontdata"/>
    +  (nothing)
```

| flag              | what it does                                       |
| ----------------- | -------------------------------------------------- |
| `--oracle <name>` | `validate` (default), `powerpoint`, or `command`   |
| `--command <cmd>` | for `--oracle command`; `{}` becomes the candidate |
| `--max-runs <n>`  | ceiling on oracle runs (default 2000)              |
| `--timeout <ms>`  | per run, for the oracles that spawn something      |
| `--progress`      | a line per oracle run; a bisection is not quick    |
| `--write <file>`  | save the smallest package that still fails         |
| `--json`          | the result as JSON                                 |
| `--out <file>`    | write the report to a file instead of stdout       |

**`--oracle powerpoint` opens each candidate with `OpenAndRepair` switched off,
and that is not a detail.** `Presentations.Open` has no repair parameter and
`Open2007` documents the default as on, while `DisplayAlerts` is documented to
answer a message box with its default — which for "PowerPoint found a problem
with content" is Repair. So the obvious script reports success on precisely the
files this project exists to avoid producing. Measured on one deck: with repair
off it fails with `0x80CB8002`; with repair on it opens, all three slides
intact.

The harness opens read-only with no window, writes nothing back, and disables
macros — `AutomationSecurity` defaults to _enabling_ them. If PowerPoint is
already running it attaches to that instance and never quits it.

## `inspect` options

| flag           | what it does                                                    |
| -------------- | --------------------------------------------------------------- |
| `--json`       | the census as JSON, for a script or for committing as a fixture |
| `--parts`      | the per-part table: bytes, elements, depth, relationships       |
| `--namespaces` | every namespace, with the prefixes it was spelled with          |
| `--top <n>`    | rows per histogram before truncating (default 15)               |
| `--out <file>` | write to a file instead of stdout                               |

## Everything it knows lives somewhere else

`@pptx-studio/census` does the reading, `@pptx-studio/validate` does the judging
and `@pptx-studio/writer` does the writing and the comparing, and not one of the
three has any Node in it at all — it runs in a browser tab and in a Web Worker. This package supplies the
two things a browser cannot: a path off the filesystem and a stream to write to.

That split is the point. The drop-a-deck explorer in `apps/studio` computes
exactly the same object in a Worker, so an answer here and an answer in a tab
cannot drift apart.

## Not built yet

The plan gives this package one more verb. Asking for it says which sub-phase brings
it rather than "unknown command":

| verb      | what it will do                          | sub-phase |
| --------- | ---------------------------------------- | --------- |
| `resolve` | show where a resolved property came from | 7.x       |

## `render`

Server-side thumbnailing with nothing installed but Node. The SVG stands alone:
pictures are embedded as `data:` URIs, gradients and patterns are `<defs>` in
the same document, and there is no external reference of any kind to resolve.

```
1 slide(s) at 1920x1080 -> thumb.svg
3 typeface(s) from 412 indexed face(s), 1 substituted
  Aptos -> Carlito
```

**Text is measured, not guessed.** The browser renderer measures with
`OffscreenCanvas.measureText`; there is no canvas in Node, so this reads the
face's own `cmap`, `hmtx`, `GPOS` and `OS/2` tables instead. That is a second
measurement engine, and the risk of a second engine is that it quietly disagrees
with the first. Experiment T13 settled the arithmetic rather than assuming it:
twelve fonts built so their tables disagree on purpose, 432 widths measured in
Chromium, and the reader reproduces **all 432 exactly**. It also settled where
an upright East Asian glyph sits: on the `BASE` table's `ideo` coordinate for
the `DFLT` script, and on the face box descent only where there is none. The
rules it found are
in [`corpus/ground-truth/font-metrics.json`](../../corpus/ground-truth/font-metrics.json)
and the reasoning is in
[ADR 0042](../../docs/adr/phase-3-text/0042-rendering-without-a-browser.md).

**The face box depends on the rasteriser**, so the reader takes the answer
belonging to the platform it runs on. Chromium reports a face's ascent and
descent from `OS/2.usWinAscent`/`usWinDescent` through DirectWrite and from
`hhea.ascender`/`descender` through FreeType, with `fsSelection` bit 7 moving
both onto `sTypo`. Neither reading fits both: T13's probes score 12/12 and 7/12
on Windows where 79 real faces score 76/79 and 79/79 on Linux, missing by up to
100 px on a 1000 px em. See
[ADR 0045](../../docs/adr/phase-3-text/0045-the-face-box-belongs-to-the-rasteriser.md).

Fonts are found in this platform's own directories, plus any `--font-dir` you
name, which are searched first so you can override a face without installing
one. `--no-system-fonts` limits it to what you named; `--no-text` draws geometry
only and asks no font questions at all.

A typeface the machine does not have is substituted through the same table
`@pptx-studio/text` uses in the browser, so the two renderers cannot fall back
differently, and the substitution is reported rather than hidden. The chain is
that table's entry for the name, then Calibri and Carlito — what PowerPoint
itself falls back to, measured 22 of 22 in
[ADR 0033](../../docs/adr/phase-3-text/0033-font-substitution-and-the-guard.md) —
then the family the asked-for name is a variation of, so `Calibri Light` draws
in Calibri where the machine has it.

Past all of those the first family by name draws the run. The browser's stack
ends in a generic and this has none, so something has to be chosen, and it is
chosen by name rather than by the order the files were read in: two machines
holding the same faces answer identically. It is still reported as a
substitution, under the name the deck asked for. **`CLI_NO_FACE` is thrown only
when no font was found at all** — an empty library is nothing to draw with,
which is a different thing from an unusual typeface. A code point no indexed
face can draw is reported too.

| flag                |                                                                     |
| ------------------- | ------------------------------------------------------------------- |
| `--slide <n>`       | one slide, 1-based; every slide by default                          |
| `--width <px>`      | the `width` attribute; the height follows the deck's aspect         |
| `--out <path>`      | a directory, or a file when rendering one slide; stdout when absent |
| `--font-dir <d>`    | look here first, repeatable                                         |
| `--no-system-fonts` | do not look in this platform's own font directories                 |
| `--no-text`         | geometry only                                                       |
| `--json`            | what was drawn, and which face drew each typeface                   |
| `--quiet`           | no summary after writing                                            |

### From a program

`renderDeck` is what the verb runs, without the file handling: it takes bytes and
returns the markup. Every option has a default, so the smallest call is the deck and
nothing else.

```ts
import { renderDeck } from '@pptx-studio/cli';

const deck = renderDeck(bytes); // every slide, 1920 wide
const thumb = renderDeck(bytes, { slide: 1, width: 640 }).slides[0]?.svg;
const shape = renderDeck(bytes, { text: false }); // geometry only, no fonts read
```

| option        | default     |                                                             |
| ------------- | ----------- | ----------------------------------------------------------- |
| `slide`       | every slide | 1-based; `null` and omitted both mean every slide           |
| `width`       | `1920`      | the `width` attribute; the height follows the deck's aspect |
| `fontDirs`    | none        | searched before the platform's own                          |
| `systemFonts` | `true`      | also look in this platform's font directories               |
| `text`        | `true`      | `false` draws geometry only and asks no font questions      |

`--out`, `--json` and `--quiet` are the command's, not the library's: they say where
the markup goes and what is printed about it. `runRender` takes those, reads the deck
off the disk and writes the files. `renderDeck` does neither, and a width that is not
a positive whole number throws `CLI_WIDTH` rather than drawing an empty picture.

Fonts are indexed per call, so a server rendering many decks pays for the scan every
time.

### What it does not do

No PNG. Rasterising would mean shipping a rasteriser, and the point of this verb
is that it needs nothing but Node — pipe the SVG to whatever you already have.

No shaping. Latin, Greek and Cyrillic measure correctly, and so does CJK, whose
advances do not depend on context. Arabic, Devanagari and the other scripts that
need a shaper will measure wide, because the reader sums unshaped advances. The
browser renderer has a real shaper and does not have this limit.
