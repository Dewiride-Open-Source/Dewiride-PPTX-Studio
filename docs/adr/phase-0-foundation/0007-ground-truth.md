# 0007 — Ground truth: embedded fonts and colour transforms

Date: 2026-08-27
Sub-phase: 0.7
Status: accepted, and it changes the plan
Amended by: [0021](../phase-2-geometry-and-paint/0021-colour.md), which measured the questions experiment C did not ask and
overturned two things recorded here — see "Also confirmed" and "Rounding" below

---

## Context

Three questions had been deferred on the grounds that arguing about them was
cheaper than measuring them, right up until it wasn't. Each one gates work
later in the plan, and each has a published answer that is either absent,
contested, or wrong.

- **A.** Is `ppt/fonts/*.fntdata` MicroType Express compressed? If it is, roughly
  four weeks of Phase 8 look different, because there is no MTX decoder in
  JavaScript or WebAssembly anywhere.
- **B.** Will PowerPoint render an EOT we built? Phase 8's custom-font export
  depends on it.
- **C.** In which colour space does DrawingML compute `tint` and `shade`?
  LibreOffice says a power law with γ=2.3, Apache POI says the sRGB piecewise
  curve, ECMA-376 says only "10% of the input colour combined with 90% white",
  and the answers differ on **every themed fill in every deck**.

Everything below was measured against Microsoft PowerPoint 16.0 build 20326 on
Windows 11, driven over COM. The tooling is in `tools/ground-truth/` and the
fixtures are in `corpus/ground-truth/`.

---

## A — PowerPoint always writes MicroType Express. The risk fired.

The plan said this "likely evaporates", on the strength of `TTEmbedFont`
documenting uncompressed `TTEMBED_RAW` as its default. It does not evaporate.

**76 of 76 `.fntdata` parts, across four decks, have `TTEMBED_TTCOMPRESSED`
(0x4) set.** Every one is EOT version `0x00020002`. The payload never begins
with an SFNT signature, so the flag is telling the truth rather than mislabelling
raw bytes — a possibility worth ruling out, since it would have been cheap to
believe and expensive to be wrong about.

PowerPoint's object model has no `SaveSubsetFonts` — that is Word's API — so
rather than change a setting on the machine we were working on, we flipped
`p:presentation/@saveSubsetFonts` inside the file and had PowerPoint re-embed:

| `@saveSubsetFonts` | `TTEMBED_SUBSET` | `TTEMBED_TTCOMPRESSED` | Cambria regular |
| ------------------ | ---------------- | ---------------------- | --------------- |
| `1`                | set              | **set**                | 131,519 bytes   |
| `0`                | clear            | **set**                | 775,316 bytes   |

So subsetting is a user setting and compression is not. There is no configuration
in which PowerPoint writes an EOT we can read without an MTX decoder.

One trap on the way: re-saving a deck that already has embedded fonts copies the
`.fntdata` parts through byte-for-byte rather than re-encoding them. The first
attempt at this comparison measured nothing at all for that reason. The
measurement above starts from a deck saved with embedding _off_, so every
embedding is fresh.

### What else the headers say

Committed as `corpus/ground-truth/eot-headers.json` — header fields only. The
payloads are Microsoft's fonts and are not in this repository.

- **`FamilyName` and `StyleName` carry a terminating NUL _inside_ the counted
  length. `VersionName` and `FullName` do not.** `FamilyNameSize` for Cambria is
  16 bytes for a seven-character name. A reader that does not strip it compares
  `"Cambria\0"` against `a:latin/@typeface="Cambria"` and finds no match.
- **`Italic` is `0xFF` for italic faces, not `0x01`.** `italic === 1` is wrong.
- `Charset` is `0`, not `DEFAULT_CHARSET` (1).
- `fsType` is copied honestly from `OS/2` — `0x0008`, Editable, for the Microsoft
  core fonts here.
- The version 2.2 tail is fixed: `RootStringSize` 0, then `RootStringCheckSum`
  as the four bytes `42 53 47 50` (`"BSGP"`, a constant, not a checksum of
  anything), `EUDCCodePage` 1252, and zeros for the rest — 20 bytes after the
  root string, which is how the arithmetic closes on `EOTSize - FontDataSize`.
- `<Default Extension="fntdata" ContentType="application/x-fontdata"/>`,
  relationship type `.../officeDocument/2006/relationships/font` hanging off
  `ppt/presentation.xml`, and `p:presentation/@embedTrueTypeFonts="1"`, exactly
  as the plan's Phase 8.7 predicted.
- `p:font/@panose` is 20 hex characters, `@charset="0"`, `@pitchFamily` 18, 34 or
  49 across the fonts seen. All four slots — `regular`, `bold`, `italic`,
  `boldItalic` — are written when the faces exist.

### Consequence

Phase 8.2 currently says a compressed font gets "a typed diagnostic and a
substitution badge". That would badge **every PowerPoint deck with embedded
fonts**, which is not a fallback, it is the feature not existing. The plan's own
mitigation is now the main line: compile `libeot` to WebAssembly behind a
dynamic import.

The research pass found something that makes this materially easier than the
plan assumed: **Monotype's MTX Licensing Statement grants a royalty-free,
irrevocable, sublicensable patent and code licence**, so the patent posture that
made hand-porting the W3C LZCOMP sample unattractive does not apply to using
libeot. libeot is MPL-2.0, which is file-level copyleft and therefore
distributable alongside Apache-2.0 code as a separate module — the same boundary
`fonts-metric-compat` already exists to draw.

---

## B — PowerPoint will not render an uncompressed EOT

The plan budgeted an hour to confirm the export half of the custom-font
requirement. It refuted it instead.

### The probe font

Wrapping a font from this machine would have answered a different question
badly: the font is installed, so PowerPoint could render it perfectly while
ignoring our bytes. So the probe font is synthesised — `tools/ground-truth/fonts/embedding/build-font.ts`,
1,272 bytes, ten glyphs, family name `ProbeAlpha`. Its eight letters are solid
bars of increasing height, so `ABCDEFGH` renders as a staircase that no
substitute font can imitate. It is CC0 by construction, which is the difference
between a fixture we can commit and one we cannot.

Judged twice, by parties with no stake in the outcome:

- **Chromium's OTS sanitiser** accepts it. Advance widths exactly 60.0px at
  100px/1000upem; bar heights exactly 10, 20, 30, 40, 50, 60, 70, 80.
- **Windows GDI+**, via an in-process `PrivateFontCollection` — nothing
  installed, nothing written to the registry — accepts it, reports
  `emHeight=1000 ascent=800 descent=200`, and renders bars of 6, 12, 18, 24, 30,
  36, 42, 48 at 60px.

The font is not the problem.

### What PowerPoint rejects

Six payload shapes, each with its own font and family name so that no row could
be satisfied by another row's data:

| Payload                                                             | Rendered? |
| ------------------------------------------------------------------- | --------- |
| EOT v1, charset 1, names not NUL-terminated                         | no        |
| EOT v1, charset 0, names NUL-terminated                             | no        |
| EOT v2.2, charset 0, NUL-terminated — PowerPoint's own header shape | no        |
| EOT v2.2, charset 0, not NUL-terminated                             | no        |
| EOT v1, charset 0, not NUL-terminated                               | no        |
| A bare SFNT, no EOT wrapper at all                                  | no        |

The third row is the important one: byte for byte the header PowerPoint itself
writes, differing only in that the payload is not compressed.

And the failure is not our package. Rows three and six were repeated by
**grafting into a deck PowerPoint wrote** — changing only the bytes of one
`.fntdata` part, the `@typeface` on its `p:embeddedFont` entry, and the
`a:latin/@typeface` of the run that used it. Every other byte of the package is
Microsoft's. Same result.

### Two things this cost us, and both are findings

**`Font.Embedded` returning `msoTrue` does not mean the font loaded.** It means a
`p:embeddedFont` entry matched the typeface. Every rejected variant above
reported `msoTrue`. A pass/fail read off the object model would have been green
and wrong; only the negative control — a row in a typeface that is neither
installed nor embedded, which must fall back — caught it. It is worth being
blunt about how close that was: the first pass of this experiment reported
success. The deck opened without a repair prompt, the object model said the font
was embedded, and the slide showed eight perfectly good letters.

**PowerPoint matches an embedded font by the name inside the font data**, not by
`@typeface` and not by the EOT header's `FamilyName`. Demonstrated twice: taking
PowerPoint's own Cambria `.fntdata`, unmodified, and renaming only `@typeface`
makes it fall back; patching the EOT header's `FamilyName` to agree does not
rescue it. So the three names — `a:latin/@typeface`, the EOT header, and `name`
ID 1 inside the SFNT — must all agree. Phase 8.7 already required the first two
to match byte-for-byte; the third is new.

### What is _not_ established

We could not construct a compressed EOT, so "compression is required" is a
conclusion by elimination, not a demonstration. The font is eliminated, the
package is eliminated, the naming is eliminated, and four header shapes are
eliminated. Compression is the one variable left, and it is the one PowerPoint's
own writer never omits.

An attempt to get a _reason_ rather than a symptom — calling `t2embed.dll`'s
`TTLoadEmbeddedFont` directly through P/Invoke — is recorded here as
inconclusive: it returned `0x0105` with `bytesRead=0` for **every** input
including PowerPoint's own Cambria `.fntdata`, which PowerPoint demonstrably
loads. The harness never invoked its read callback, so it was measuring itself.
The control is what caught it, and the harness is not committed.

### Consequence

Phase 8's promise — "custom fonts survive into exported PowerPoint on a clean
machine" — needs an MTX **encoder**, not just a decoder, and libeot only decodes.
Before Phase 8 opens, one of these has to be true:

1. an encoder exists or can be written (LZCOMP is symmetric enough that a
   decoder is most of the specification, but this is real work);
2. some writer we have not tested produces an uncompressed EOT that PowerPoint
   accepts — **LibreOffice exports `.pptx` with embedded fonts and is the
   obvious next measurement**, and it costs an hour;
3. the feature ships as "custom fonts render in PPTX Studio, and export as a
   substituted face with a visible warning", which is a smaller promise than the
   plan makes.

This is a question for Phase 8's planning, not a decision to take here.

---

## C — Colour: the sRGB piecewise curve, and a rule that covers everything

214 solid-fill swatches, one hand-built deck, read back two independent ways:
PowerPoint's `Shape.Fill.ForeColor.RGB`, and the centre pixel of PowerPoint's own
bitmap export. **They agree on all 214.** That is worth recording on its own,
because it means the object model is a usable oracle for future ground-truth
work; but where they could disagree, the pixels are what a user sees, and the
fixture records the pixels.

### tint and shade

| model                            | worst error | mean     | exact       |
| -------------------------------- | ----------- | -------- | ----------- |
| **sRGB piecewise linearisation** | **0/255**   | **0.00** | **102/102** |
| power law γ=2.3 (LibreOffice)    | 8/255       | 1.67     | 41/102      |
| arithmetic on the 0–255 values   | 73/255      | 31.00    | 0/102       |
| POI's asymmetric pair            | 73/255      | 12.71    | 51/102      |

The blend happens in linear light, and the transfer function is the sRGB
piecewise one — `c/12.92` below 0.04045 and `((c+0.055)/1.055)^2.4` above it.
Exact on every swatch.

This is not the plan's expected answer in two ways. The plan said the candidates
"differ by up to ~4/255 in midtones"; the real spread is 8/255 for γ=2.3 and
73/255 for the naive model. And Apache POI, which the plan credited with the
right curve, only linearises `shade` — its `tint` is computed straight on the
sRGB values, so it is exactly right on half the question and 73/255 wrong on the
other half.

### The rule that covers everything else

Every other transform was measured too, and they fall into three groups. The
split is clean enough to state as a rule:

- **The RGB-space transforms operate in linearised sRGB.** `tint`, `shade`,
  `inv`, `gamma`, `invGamma`, and `redMod`/`redOff` and their green and blue
  counterparts. All exact.
  - `a:inv` is a complement in _linear_ light: `#4472C4` inverts to `F8EBB3`,
    not to the channel complement `BB8D3B`. A renderer doing `255 - c` is wrong
    by up to 100 units.
  - `a:gamma` **is** the sRGB de-linearisation and `a:invGamma` its exact
    inverse. The standard exposes the curve directly, which is a strong
    independent confirmation of the tint/shade result.
- **The HSL-space transforms do not linearise at all.** `lumMod`, `lumOff`,
  `lum`, `satMod`, `satOff`, `sat`, `hueMod`, `hueOff`, `hue` run on the sRGB
  values as they stand. Linearising here is wrong by up to 115/255 — the largest
  single error measured anywhere in this sub-phase.
- **`a:gray` is the Rec.709 luma of the sRGB values**, no linearisation:
  `0.2126R + 0.7152G + 0.0722B`. `#FF0000` becomes `363636`.

### Saturation and lightness are not clamped in HSL space

`satMod val="200000"` and `satMod val="300000"` on `#4472C4` produce **different**
colours — `0460FF` and `004EFF`. A model that clamps saturation to 1 cannot
express that; it gives the same answer for both, and is wrong by 19/255 on the
second. Clamping happens once, per channel, on the way out. Fixing this took the
HSL group from 64/69 exact to 66/69.

### Transform order

Not a footnote. On `#4472C4`:

| chain                          | result   |
| ------------------------------ | -------- |
| `lumMod 60%` then `lumOff 40%` | `8FAADC` |
| `lumOff 40%` then `lumMod 60%` | `517CC8` |
| `tint 50%` then `shade 50%`    | `8D93A7` |
| `shade 50%` then `tint 50%`    | `BEC2D1` |

A renderer that applies transforms in a fixed order rather than document order
gets one of each pair badly wrong.

### Rounding

The three swatches a correct model still missed were all exactly `x.5`, and
PowerPoint takes the lower value. Round-half-down fits **171 of 171**; JavaScript's
`Math.round` fits 168. Three discriminating samples is thin evidence for a rule,
so this is recorded as what fits rather than as settled: the fixture holds
PowerPoint's answers either way, and a renderer that is off by one on a colour
that lands precisely on a half is not the problem worth solving first.

> **Amended by ADR 0021.** Four more tie cases were measured in sub-phase 2.6, and
> round-half-down fits only four of the resulting seven. No rounding mode fits all
> of them, and neither does single-precision representation error. The caution
> above was warranted; the rule was a coincidence of three samples.

### Also confirmed

Theme resolution and the `clrMap` on all 16 scheme names, including that
`bg1`/`tx1`/`bg2`/`tx2` route through the map while `dk1`/`lt1`/`dk2`/`lt2` do
not; `a:sysClr/@lastClr` preferred over the system colour name; and seven
identity swatches with no transform, which is the control that says the whole
pipeline — hand-built package, PowerPoint, bitmap export, sampling — is not
introducing an error of its own.

> **Amended by ADR 0021: the `@lastClr` claim is wrong.** Every `sysClr` in this
> experiment named a colour whose `lastClr` agreed with what this machine says, so
> both readings gave the same number and nothing here could tell them apart.
> Sub-phase 2.6 wrote `<a:sysClr val="windowText" lastClr="FF00FF"/>` and PowerPoint
> painted it black, from the machine. `@lastClr` is a cache for a reader that has no
> machine to ask — which is what `@pptx-studio/paint` is, so it prefers `@lastClr`
> deliberately rather than because PowerPoint does.

---

## The first `.pptx` this project ever wrote

Both experiments needed decks PowerPoint's object model cannot express — COM
cannot set `lumMod` on a fill and certainly cannot embed a font we built — so
`tools/ground-truth/lib/pptx.ts` writes the package by hand: content types,
relationships, a theme with all twelve `clrScheme` children and exactly three
entries in each `fmtScheme` list, a master with all twelve `clrMap` attributes,
a blank layout, and the slides.

**PowerPoint opened all of them with no repair prompt** — 214 shapes over three
slides for experiment C, nine rows with five embedded fonts for experiment B.
That is early evidence for Phase 1.3's writer, arriving a phase early and for
free.

One thing it cost, worth writing down because the diagnostic is so poor: an
unescaped `<` in an `a:t` — the arrows in my own row labels — makes PowerPoint
refuse the entire package over COM with _"The file or directory is corrupted and
unreadable"_. No part name, no line number, no hint that the problem is one
character in one text run. This is precisely the failure mode `cli bisect`
exists for in Phase 1.5.

---

## Verification

- 76 EOT headers read, 4 decks, `TTEMBED_TTCOMPRESSED` set on 76.
- Probe font accepted and rendered exactly by Chromium/OTS and Windows GDI+.
- 6 uncompressed payload shapes rejected by PowerPoint, 2 of them inside
  PowerPoint's own package.
- 214 swatches, 2 independent readbacks, 214/214 agreement.
- tint/shade: 102/102 exact under the sRGB piecewise model.
- All six derived transform rules hold exactly on every swatch that exercises
  them.
- 171/171 exact once the rounding mode is right.

## What is deferred

- **An MTX decoder**, and the question of an encoder. Phase 8.
- **Measuring LibreOffice's embedded-font export.** The cheapest thing that could
  overturn experiment B's conclusion, and it should be done before Phase 8 opens.
- **The `TTLoadEmbeddedFont` probe**, if a specific error code is ever worth
  having. The harness needs fixing first, and its own control says so.
- **Bold, italic and bold-italic slots**, and non-Latin charsets — the probe font
  is regular-weight Latin only.
- **Chart and diagram colour styles** (`colors1.xml`, `cs:variation`). Phase 9.1,
  and a different question from this one.
