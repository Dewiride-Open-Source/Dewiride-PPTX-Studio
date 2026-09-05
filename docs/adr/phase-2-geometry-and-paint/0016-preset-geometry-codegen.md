# 0016 — The 187 preset shapes, and how they get here

Date: 2026-09-02
Status: **accepted** — 187 transcoded, cross-checked against an independently written enumeration
Sub-phase: 2.1 — preset codegen

---

## Context

Phase 2 draws shapes. Before anything can be drawn, the definitions have to exist, and ECMA-376
ships them only as prose: several hundred pages describing 187 preset geometries, with no
machine-readable form anywhere in the standard's downloads.

Every open implementation that draws presets therefore carries its own transcription. The plan
picked Apache POI's `presetShapeDefinitions.xml` for two reasons that still hold: it is Apache-2.0,
which is the licence this project publishes under, and it is a single self-describing file rather
than a generator welded into a rendering stack.

The file is not in this repository and this sub-phase did not put it there.

---

## Decisions

### The same arrangement as the XSDs: fetch by hand, commit only the output

`tools/schema-codegen` already established the pattern in 0.6 — the input is a free download, the
generator takes it by path, only the generated file is committed, and that file records the SHA-256
of every input. `tools/geometry-codegen` is the same shape.

```
presetShapeDefinitions.xml   538 970 bytes
sha256 4a762444d8d85876881c02a5b1dedf6f73006fcd8acb7b4e393435615b37c780
root element <presetShapeDefinitons>
```

The root element is misspelled in the source — `Definitons`, no second `i`. Reproduced verbatim, and
the reader refuses a file whose root is spelled correctly, because that would be a different file
from the one this was written against.

### The verification is checked against something that is not the input

The sub-phase's stated exit criterion is "all 187 `ST_ShapeType` names present". Checked against the
file the names came from, that sentence is a **tautology**: it reports 187 for any input, including
one that lost a shape between POI releases.

So `tools/geometry-codegen/shape-types.ts` writes the enumeration out independently, from ECMA-376's
`ST_ShapeType`, and the generator diffs the two sets:

```
ST_ShapeType cross-check, against the list in shape-types.ts:
  expected  187
  found     187
  missing   none
  extra     none
```

That is two sources agreeing rather than one source agreeing with itself. The file says so in the
place someone would go to change it: a future disagreement is a finding, and the fix is to work out
which source is wrong, not to paste one list over the other.

### Encoded strings, not object literals

A faithful object literal of this data is **larger than the XML it came from**:
`{ kind: 'lnTo', to: { x: 'x2', y: 't' } }` spends 34 characters saying what `L x2 t` says in 6.
Written that way the buckets would be most of a megabyte of source for every consumer's bundler to
parse, and would dominate the download of a library whose job is drawing shapes.

`schema-order.gen.ts` made the same trade in 0.6. The difference is scale: 539 KB of XML becomes
**119 619 characters** of encoded text across six buckets, which is the difference between shipping
preset geometry by default and making it an opt-in. Decoding is lazy and per shape, so a slide
holding three rounded rectangles decodes three presets and not 187.

### Six buckets, and what a bucket is not

`basic` 43, `arrows` 31, `callouts` 16, `flowchart` 29, `misc` 47, `stars` 21. Five of the six have a
mechanical signature in the name; `basic` has none and is therefore a hand-written list, which is
the one bucket rule that is a judgement rather than a derivation.

Rule **order** decides one real case: PowerPoint files the six `*ArrowCallout` shapes under Block
Arrows, and matching on the name alone would say the opposite.

These are packaging boundaries and nothing about a shape depends on which file it lands in.
`getPreset` resolves any name from any bucket and never moves; a consumer importing one bucket
directly for the size saving should pin a version.

### Nothing is evaluated, and every operand is still a string

A formula's operands are guide names, built-in names and literals mixed together, and nothing can
tell them apart without a shape size and an `avLst` to resolve against. Parsing `200000` into a
number at build time would mean the transcoder deciding which operands are numbers — 2.2's job.

So this package answers _what `blockArc` consists of_ and not _where its points are_.

### The source is reported, not repaired

Two things in POI's file are outside ECMA's grammar. Both are carried through verbatim and both are
exported so a caller can see them:

- **Eight `+-` formulas carry four operands** where `+-` is `x + y − z`. All eight are in the
  circular-arrow family — `circularArrow`, `leftCircularArrow`, `leftRightCircularArrow` — and all
  end in a spare `0`, so the intended value is unambiguous. Truncating them here would be the
  transcoder holding an opinion about semantics, and would hide that the correction was ever needed.
- **Six formulas contain a double space** — `heptagon`, `pentagon`, `star5`, `star7` in `svc`, and
  `leftArrow`, `upArrow` in a coordinate guide. The split is on runs of whitespace, so operands come
  out right; a single-space split gives each an empty operand and an arity one too high, which the
  reader would then report as an arity error on a formula that is fine.

---

## What building it found

### `pie`'s text rectangle has two sides transposed

POI writes `<rect l="il" t="ir" r="it" b="ib"/>` for `pie`. The guides are named for the sides they
compute — `il`/`ir` horizontal, `it`/`ib` vertical — so `t` and `r` are swapped. Of the 29 shapes
that use this guide set, the other 28 read `l="il" t="it" r="ir" b="ib"`.

Reproduced rather than corrected, for the same reason as the two anomalies above, and pinned by an
assertion so that a later POI which fixes it makes the test fail and forces a deliberate decision.
It will matter when something first places text inside a `pie`, which is Phase 3, not this one.

### The count that had never been checked, checked

`rect` is both an element name and a shape name, so the naive count of `<rect>` elements (183) is one
shape definition plus 182 text rectangles. Five shapes have no text rectangle at all — `line`,
`lineInv`, `chartPlus`, `chartStar`, `chartX` — and 182 + 5 = 187 is the arithmetic that says the
reading is right. `rect` being a preset in its own right is easy to miss and would have left the
roster one short.

### An eslint limit that scales with the package count

`typescript-eslint` caps its default project at eight files, and `packages/geometry/tsdown.config.ts`
was the ninth. The cap exists because a wide default project is a genuine performance problem; this
project's is one twelve-line build config per package. Raised to 32 with the reasoning recorded,
rather than nudged up once per sub-phase for the next dozen packages.

---

## Consequences

`pnpm check` green.

- **187 presets**, 3922 guides (300 adjust, 3622 computed), 243 adjust handles, 860 connection
  sites, 182 text rectangles, 320 paths, 2915 path commands.
- **119 619 characters** encoded, from 538 970 bytes of XML; 133 KB of generated TypeScript
  including headers. Built and gzipped, all six buckets plus the decoder are **33 KB** — and a
  consumer that only draws flowcharts pays 1.8 KB plus the 3.2 KB decoder, because the entry points
  are genuinely separate: `dist/presets/flowchart.gen.js` contains no callout.
- **`pnpm test` is 59 files, 1495 tests**, up from 57 and 1451.
- `layering: 7 package(s) checked, no violations` — `geometry` was declared at layer 1 in 0.1,
  before it existed, which is what that file is for.

### The check that spans both halves

The transcoder cannot import the package — nothing links a workspace package into `tools/` — so
encoder and decoder are genuinely separate programs and a field dropped by one would not be noticed
by the other. `codegen.test.ts` decodes every committed bucket with the package's decoder, re-encodes
it with the tool's encoder, and compares character for character: `encode(decode(x)) === x` for all 187. The package's suite comes at it from the other side, checking the decode against totals the
generator counted from its **parse**, before anything was encoded.

### What is not done

- **Nothing is evaluated and nothing is drawn.** No formula has been computed, no arc unskewed, no
  path emitted. A `PresetShape` is inert data.
- **`custGeom` is untouched.** 2.4 unifies it with this, and the conformance test there — rewriting
  each `prstGeom` as the equivalent `custGeom` and asserting identical output — is what will
  exercise this data properly for the first time.
- **The 17 operators are validated by arity only.** The reader knows `cat2` takes three operands; it
  has no idea what `cat2` means. Neither does anything else yet.
- **No preset has been compared against PowerPoint.** That the definitions transcribe faithfully is
  established; that they _draw_ what PowerPoint draws is not, and cannot be until 2.4.
- **The 44 built-in guides do not exist here.** `w`, `h`, `hc`, `vc`, `wd2`, `ssd6`, `cd4`, `3cd4`
  and the rest appear as operand strings this package never resolves. 2.2 seeds them.
