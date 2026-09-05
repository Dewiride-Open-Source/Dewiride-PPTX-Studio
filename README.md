# PPTX Studio

[![CI](https://github.com/Dewiride-Open-Source/Dewiride-PPTX-Studio/actions/workflows/ci.yml/badge.svg)](https://github.com/Dewiride-Open-Source/Dewiride-PPTX-Studio/actions/workflows/ci.yml)
[![round trip](https://img.shields.io/endpoint?url=https%3A%2F%2Fraw.githubusercontent.com%2FDewiride-Open-Source%2FDewiride-PPTX-Studio%2Fmain%2F.github%2Fbadges%2Froundtrip.json)](#the-round-trip-badge)
[![License](https://img.shields.io/badge/license-Apache--2.0-blue)](./LICENSE)

**A PowerPoint renderer and layout-aware editor that runs entirely in the browser.**

Open a `.pptx`, see it rendered faithfully, edit every element on the slide, switch layouts and
themes and have them genuinely cascade, then export a file that still opens in real PowerPoint with
everything you did not touch left byte-for-byte intact.

> **Status: pre-alpha, under active construction.** Nothing is published to npm yet. Sub-phases
> 0.1 (repository skeleton), 0.2 (ZIP reader, decompression budgets, OPC part-name grammar), 0.3
> (part store, content types, relationships, ZIP32 writer), 0.4 (XML tokenizer and node model), 0.5
> (serializer and the byte-identical round-trip gate), 0.6 (schema order, Markup Compatibility,
> invertible edits), 0.7 (ground-truth experiments) and 0.8 (`inspect`, the Worker boundary, the
> benchmark) are complete, which closes **Gate 0**: drop a `.pptx` on a page and get a live
> explorer of its internals — every part, content type, relationship edge and a feature census —
> computed in a Web Worker, entirely in the tab. A 188 MiB, 300-slide deck reads in about a third
> of a second.
>
> Phase 1 is complete too — 1.1 (the corpus), 1.2 (the repair firewall), 1.3 (the writer), 1.4
> (the round-trip oracle), 1.5 (`bisect` and the PowerPoint oracle) and 1.6 (the CI gate and the
> badge above) — which closes **Gate 1**: drop a deck carrying charts, SmartArt, animations, OLE
> objects and macros onto the page, press Save, and the file your browser downloads opens in real
> PowerPoint 365 with no repair prompt and the same ten slides and the same shapes on each of them.
> Measured, not asserted: `pnpm gate1` drives the actual page in Chromium, captures the actual
> download, and asks PowerPoint through COM with `OpenAndRepair` **off**, which is the only setting
> under which a repair is an error rather than a silent success.
>
> The table below tracks what actually works, and it will not be marked green ahead of the code.

---

## Why another one

Every existing web PowerPoint tool sits at one of two extremes.

Server-side viewers render well and hand you a picture. You cannot edit a picture.

Browser editors that import `.pptx` — Canva, PPTist, PPTXjs — flatten each slide into a fixed list
of absolutely-positioned boxes the moment they parse it. That flattening is exactly why their
slides feel fixed: once a shape's missing `a:xfrm` has been defaulted to a concrete number, the
information that the shape _never had explicit geometry_ is destroyed, and with it any ability to
switch layouts, swap themes, or cascade anything at all. They also regenerate the file on export,
so charts, SmartArt, animations and OLE objects are lost on the way out.

PPTX Studio takes the opposite position on both:

- **The OPC package is the document.** Our model is a projection of it. Parts we do not edit are
  streamed back out byte-for-byte, so preservation is the default state rather than a feature that
  has to be implemented per content type.
- **Inheritance is never flattened.** A property that is absent stays absent. `undefined` means
  "ask the resolver", and the resolver walks slide → layout → master → theme every time. That one
  decision is what makes Change Layout, theme swaps and per-property Reset possible at all.

Read [`SCOPE.md`](./SCOPE.md) for what this deliberately is not.

## What works today

| Area                                                 | Status         |
| ---------------------------------------------------- | -------------- |
| Repository, toolchain, architecture guards           | ✅ 0.1         |
| OPC container: ZIP reader, budgets, part names       | ✅ 0.2         |
| OPC container: parts, content types, relationships   | ✅ 0.3         |
| Byte-preserving XML tokenizer and node model         | ✅ 0.4         |
| XML serializer and the byte-identical round trip     | ✅ 0.5         |
| Schema order, Markup Compatibility, invertible edits | ✅ 0.6         |
| Ground truth: embedded fonts, colour transforms      | ✅ 0.7         |
| Feature census, `cli inspect`, the Worker boundary   | ✅ 0.8         |
| The corpus: 52 licensed decks, three producers       | ✅ 1.1         |
| The repair firewall: 29 rules, `cli validate`        | ✅ 1.2         |
| Writer: dirty-part export, media GC, prepare hooks   | ✅ 1.3         |
| The round-trip oracle and `cli roundtrip`            | ✅ 1.4         |
| `cli bisect` and the PowerPoint oracle               | ✅ 1.5         |
| The CI round-trip gate and badge                     | ✅ 1.6         |
| Open, re-save and download in the browser            | ✅ Gate 1      |
| The 187 preset shape definitions                     | ✅ 2.1         |
| The formula evaluator: 17 operators, built-in guides | ✅ 2.2         |
| Arc math: the unskew, winding, the whole-turn split  | ✅ 2.3         |
| Path emit, and `custGeom` as the same thing          | ✅ 2.4         |
| Adjust handles: drag one, get the value to write     | ✅ 2.5         |
| Colour: six bases, 28 transforms, theme and clrMap   | ✅ 2.6         |
| Gradients and the 54 pattern tiles, both measured    | ✅ 2.7         |
| Dashes, joins, arrowheads and effects, all measured  | ✅ 2.8         |
| Painted: fills, strokes, gradients, effects          | ⬜ Phase 2     |
| Text engine, viewer, fidelity scoreboard             | ⬜ Phase 3     |
| Tables and SmartArt                                  | ⬜ Phase 4     |
| Select, move, resize, rotate                         | ⬜ Phase 5     |
| Text editing                                         | ⬜ Phase 6     |
| Layout switching, theme verbs, backgrounds           | ⬜ Phase 7     |
| Font embedding, custom font upload and export        | ⬜ Phase 8     |
| Charts                                               | ⬜ Phases 9–10 |
| Accessibility, virtualization, 1.0                   | ⬜ Phase 12    |

## Repository layout

```
packages/
  opc/     OPC container: zip, parts, content types, relationships
  xml/     byte-preserving XML tokenizer, XNode, serializer
  census/  what is inside a package: parts, relationship graph, feature census
  geometry/ preset and custom geometry: 187 shapes, the evaluator, arcs, paths, handles
  paint/   colour, fills, strokes and effects, every rule measured against PowerPoint
  model/   sheets, the inheritance chain, and the resolver everything else reads
  render-svg/ slides to SVG: group transforms, flip order, grpFill, as a string
  render-dom/ the live renderer, over the same node tree as render-svg
  validate/ the repair firewall: the 29 rules a .pptx must not break
  writer/  export: dirty-part-only serialization, media GC, prepare hooks
  cli/     the Node entry point — `pptx-studio inspect` and `validate`
apps/
  studio/  drop a .pptx on a page; the parse Worker boundary lives here
tools/
  layering/      the dependency-direction guard (see below)
  eslint-rules/  local ESLint rules with no upstream equivalent
  schema-codegen/ the ECMA-376 element-order table generator
  geometry-codegen/ the preset-shape transcoder, and the 187-name cross-check
  ground-truth/  experiments that ask real PowerPoint what it actually does
  bench/         synthetic decks, and the browser benchmark that reads them
  corpus/        the corpus generators, its five rules, and the roster
corpus/
  decks/         42 synthetic probe decks, one feature each but the last
  authored/      9 decks PowerPoint wrote, so the corpus has a second producer
  written/       1 deck our own writer round-tripped
  reject/        17 packages PowerPoint refuses, one measured finding each
  ground-truth/  what PowerPoint answered, as committed fixtures
  bench/         benchmark deck recipes, their hashes, and recorded timings
docs/adr/        architecture decision records
```

Further packages arrive with the phases that need them. Every one of them is already declared in
[`tools/repo/layering/layers.ts`](./tools/repo/layering/layers.ts) with its layer and runtime, because a
guard that arrives after the code it guards has already missed the first violation.

## The architecture is enforced, not documented

Three constraints hold this design together, and all three are checked by CI rather than by
convention:

1. **The core never touches Node.** `@pptx-studio/*` runs in a browser tab and in a Web Worker.
   `node:*` imports and Node globals are banned by ESLint; a `"node"` condition in an exports map is
   banned by the layering checker; and the core test suite runs in real Chromium rather than jsdom,
   because under jsdom `node:fs` resolves and the leak stays invisible until someone opens a tab.
2. **Dependencies run one way only.** `opc`/`xml` → `geometry`/`paint`/`text`/`fonts` → `model` →
   renderers → `react`. Enforced at the manifest level by `pnpm layering`, which also catches
   cycles, upward edges, and React appearing anywhere below the one package allowed to know it
   exists.
3. **No top-level await.** A single one makes a module asynchronous and therefore un-requireable
   from CommonJS, for every consumer, forever. A local ESLint rule rejects it.

Run all three with `pnpm check`.

## Open it, save it, open it in PowerPoint

That is Gate 1, and it is one command:

```sh
pnpm build
pnpm gate1                  # a real browser, a real download, and real PowerPoint
pnpm gate1 --no-powerpoint  # everything a machine without Office can do
```

It serves [`apps/studio`](./apps/studio), drives the page in Chromium, presses Save twice — once
with no edit and once with one — captures what the browser actually downloads, reads it back, and
then asks PowerPoint through COM whether it opens. Measured on 2026-09-02, PowerPoint 365 build
16.0.20326, on [`corpus/decks/a43-kitchen-sink.pptm`](./corpus/decks) — ten slides carrying four
charts, three SmartArt diagrams, three OLE objects, three animation timelines and a VBA project:

| save         | parts | rewritten | streamed | differs in           | PowerPoint        |
| ------------ | ----- | --------- | -------- | -------------------- | ----------------- |
| unchanged    | 55    | 0         | 55       | nothing              | opened, no repair |
| with a stamp | 55    | 1         | 54       | `/docProps/core.xml` | opened, no repair |

The first row is the gate. **Zero parts rewritten** is the architecture in one number: a part
nobody edited is copied out of the source archive still compressed, so there is no category of
content that can be silently dropped on the way out. PowerPoint reported the same ten slides with
the same shape counts — `[7, 2, 2, 3, 2, 2, 2, 3, 3, 3]` — for the downloaded file as for the
original.

The second row is why the demo has two buttons. Setting `cp:lastModifiedBy` is the smallest edit in
the package and changes nothing on any slide, but it is the only path in which a part is serialised
afresh — and a demo that can only ever emit the bytes it read has not exercised the writer.

`OpenAndRepair` is passed **off**, which matters more than it sounds. It defaults to on, and under
automation the repair is silent: a file PowerPoint would prompt about opens successfully, already
repaired, and reports success. With it off the repair becomes a catchable error, which is the only
way that answer is worth anything.

## The round-trip badge

The badge at the top says **52/52 decks**, and it means something narrow and checkable: every
licensed deck in [`corpus/`](./corpus) is read, written back out, and compared against the original
— 1474 parts in all, 875 as canonical XML, 553 as relationship graphs with the ids treated as
opaque labels, and 46 by SHA-256.

**It is not a byte comparison and must not become one.** Entry order, deflate level, timestamps and
attribute order all differ legitimately between two archives holding one document; nine of the
PowerPoint-authored decks come out exactly 1832 bytes smaller with every stored byte of every entry
identical, because Office writes a `0xA220` growth-hint extra field on five entries and we do not.

**It does not say the decks open in PowerPoint.** No hosted runner has Office on it — GitHub's
Windows images ship Visual Studio's Office _development_ workload, which cannot open a file. That
question is answered by `pptx-studio bisect --oracle powerpoint`, on a machine with PowerPoint, by
hand.

The number is committed to [`.github/badges/roundtrip.json`](./.github/badges/roundtrip.json)
rather than published from a workflow run, because it is a fact about the contents of this
repository and nothing about a run decides it. `pnpm roundtrip` recomputes it and fails if the two
disagree, so the badge cannot claim 52/52 unless the last run to touch it measured 52/52 — and no
workflow here needs write access to anything.

## Development

Requires Node ≥ 24.11 and pnpm ≥ 11.

```sh
pnpm install
pnpm browsers        # one-off: Chromium for the core test suite
pnpm check           # structure, layering, docs, corpus, format, lint, typecheck, build, test
```

Individual steps: `pnpm layering`, `pnpm corpus`, `pnpm lint`, `pnpm typecheck`, `pnpm build`,
`pnpm roundtrip`, `pnpm test`.

To look inside a deck:

```sh
node packages/cli/dist/cli.js inspect deck.pptx        # after pnpm build
node packages/cli/dist/cli.js validate deck.pptx       # the 29 must-not-break rules
node packages/cli/dist/cli.js roundtrip deck.pptx      # is it still the same deck
node packages/cli/dist/cli.js bisect a.pptx b.pptx     # which change broke it
node tools/bench/serve.ts --decks <dir>                # then drop one on the page
```

`tools/bench/` also generates the decks the benchmark reads — up to 300 slides and 200 MB, from a
recipe, deterministically. See [`tools/bench/README.md`](./tools/bench/README.md).

`pnpm test` and `pnpm typecheck` read package **source**, not `dist`, so neither needs a build
first. Publishing reads `dist`.

## Contributing

Contributions are welcome. See [`CONTRIBUTING.md`](./CONTRIBUTING.md) — this project uses the
Developer Certificate of Origin (a `Signed-off-by` line), not a CLA.

If you are looking for somewhere to start, the geometry presets in Phase 2 are unusually
well-suited to a first contribution: each one has an objective visual test and a number that moves.

## Licence

[Apache-2.0](./LICENSE). See [`NOTICE`](./NOTICE) for third-party attributions and
[`LEGAL.md`](./LEGAL.md) for the reimplementation basis and the clean-room record.

Not affiliated with or endorsed by Microsoft. PowerPoint is a trademark of Microsoft Corporation.
