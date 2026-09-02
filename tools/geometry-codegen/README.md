# `tools/geometry-codegen` — the 187 preset shapes, transcoded

Turns Apache POI's `presetShapeDefinitions.xml` into
[`packages/geometry/src/presets/*.gen.ts`](../../packages/geometry/src/presets).

```sh
curl -LO https://raw.githubusercontent.com/apache/poi/trunk/poi/src/main/resources/org/apache/poi/sl/draw/geom/presetShapeDefinitions.xml
node tools/geometry-codegen/generate.ts --presets presetShapeDefinitions.xml            # report only
node tools/geometry-codegen/generate.ts --presets presetShapeDefinitions.xml --write
```

No formatting step follows: `*.gen.ts` is in `.prettierignore` and in eslint's ignores, exactly as
`schema-order.gen.ts` is.

**The input is not in this repository and nothing here fetches it.** Same arrangement as
[`tools/schema-codegen`](../schema-codegen): the source is a free download, the codegen is run by
hand when it changes, and only the generated files are committed — each recording the SHA-256 of the
input it came from, so "which POI is this geometry from" has an answer that does not rely on
anyone's memory. The [`NOTICE`](../../NOTICE) entry is the other half of the obligation.

## Why POI and not the specification

ECMA-376 describes the preset geometries in prose and ships no machine-readable form of them. Every
open implementation that draws presets carries its own transcription; POI's is Apache-2.0 — the
licence this project publishes under — and is one self-describing file rather than a generator
welded into a rendering stack.

## The 187, checked against something that is not the input

The sub-phase's stated verification is "all 187 `ST_ShapeType` names present". Checked against the
file the names came from, that sentence is a tautology — it would report 187 for any file, including
one that had lost a shape between POI releases.

So [`shape-types.ts`](shape-types.ts) writes the enumeration out independently, from ECMA-376's
`ST_ShapeType`, and the generator compares the two sets. They agree exactly: 187 expected, 187
found, none missing, none extra. **If they ever disagree that is a finding**, and the fix is to work
out which source is wrong — not to paste one list over the other, which converts the check back into
the tautology it was written to avoid.

## What it refuses to do

The reader accepts a measured profile — 19 element names, 31 attribute names, 17 formula operators,
no comments, no CDATA, no DOCTYPE, no entities, no non-ASCII — and **fails with a byte offset on
anything else**, so a future POI release that introduces a construct stops the build instead of
being silently dropped from a shape. The encoder does the same for its five reserved separators.

It also does not repair its input. Two things in the file are outside ECMA's grammar, and both are
reported and carried through verbatim:

- **Eight `+-` formulas carry four operands** where `+-` takes three. All eight are in the
  circular-arrow family and all end in a spare `0`. Truncating them would be the transcoder having
  an opinion about semantics, which is 2.2's job.
- **Six formulas contain a double space.** The split is on runs of whitespace, so the operands come
  out right; a single-space split would give each an empty operand and an arity one too high.

## How the two halves are kept honest

The transcoder cannot import `@pptx-studio/geometry` — nothing links a workspace package into
`tools/`, and the type-only import it does use is erased before Node sees it. So encoder and decoder
are genuinely separate programs, and a field dropped by one would not be noticed by the other.

[`codegen.test.ts`](codegen.test.ts) closes that: it decodes every committed bucket with the
package's decoder and re-encodes it with this tool's encoder, then compares against the committed
text character for character. `encode(decode(x)) === x`, for all 187. The package's own suite
approaches it from the other side, checking the decode against totals the generator counted from its
parse before anything was encoded.
