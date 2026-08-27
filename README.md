# PPTX Studio

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
> of a second. The table below tracks what actually works, and it will not be marked green
> ahead of the code.

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
| Byte-perfect round trip across a 50-deck corpus      | ⬜ Phase 1     |
| Geometry, fills, strokes, effects                    | ⬜ Phase 2     |
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
  cli/     the Node entry point — `pptx-studio inspect`
apps/
  studio/  drop a .pptx on a page; the parse Worker boundary lives here
tools/
  layering/      the dependency-direction guard (see below)
  eslint-rules/  local ESLint rules with no upstream equivalent
  schema-codegen/ the ECMA-376 element-order table generator
  ground-truth/  experiments that ask real PowerPoint what it actually does
  bench/         synthetic decks, and the browser benchmark that reads them
corpus/
  ground-truth/  what PowerPoint answered, as committed fixtures
  bench/         benchmark deck recipes, their hashes, and recorded timings
docs/adr/        architecture decision records
```

Further packages arrive with the phases that need them. Every one of them is already declared in
[`tools/layering/layers.ts`](./tools/layering/layers.ts) with its layer and runtime, because a
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

## Development

Requires Node ≥ 24.11 and pnpm ≥ 11.

```sh
pnpm install
pnpm browsers        # one-off: Chromium for the core test suite
pnpm check           # layering, format, lint, typecheck, build, test
```

Individual steps: `pnpm layering`, `pnpm lint`, `pnpm typecheck`, `pnpm build`, `pnpm test`.

To look inside a deck:

```sh
node packages/cli/dist/cli.js inspect deck.pptx        # after pnpm build
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
