# PPTX Studio

**A PowerPoint renderer and layout-aware editor that runs entirely in the browser.**

Open a `.pptx`, see it rendered faithfully, edit every element on the slide, switch layouts and
themes and have them genuinely cascade, then export a file that still opens in real PowerPoint with
everything you did not touch left byte-for-byte intact.

> **Status: pre-alpha, under active construction.** Nothing is published to npm yet. Sub-phases 0.1
> (repository skeleton), 0.2 (ZIP reader, decompression budgets, OPC part-name grammar), 0.3 (part
> store, content types, relationships, ZIP32 writer) and 0.4 (XML tokenizer and node model) are
> complete: a real deck can be read, its relationship graph walked, every XML part parsed into a
> tree that remembers exactly where each node came from, and the whole thing written back
> byte-for-byte. The table below tracks what actually works, and it will not be marked green ahead
> of the code.

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

| Area                                               | Status         |
| -------------------------------------------------- | -------------- |
| Repository, toolchain, architecture guards         | ✅ 0.1         |
| OPC container: ZIP reader, budgets, part names     | ✅ 0.2         |
| OPC container: parts, content types, relationships | ✅ 0.3         |
| Byte-preserving XML tokenizer and node model       | ✅ 0.4         |
| XML serializer, schema order, invertible edits     | ⬜ 0.5–0.6     |
| Byte-perfect round trip across a 50-deck corpus    | ⬜ Phase 1     |
| Geometry, fills, strokes, effects                  | ⬜ Phase 2     |
| Text engine, viewer, fidelity scoreboard           | ⬜ Phase 3     |
| Tables and SmartArt                                | ⬜ Phase 4     |
| Select, move, resize, rotate                       | ⬜ Phase 5     |
| Text editing                                       | ⬜ Phase 6     |
| Layout switching, theme verbs, backgrounds         | ⬜ Phase 7     |
| Font embedding, custom font upload and export      | ⬜ Phase 8     |
| Charts                                             | ⬜ Phases 9–10 |
| Accessibility, virtualization, 1.0                 | ⬜ Phase 12    |

## Repository layout

```undefined
  xml/     byte-preserving XML tokenizer, XNode, serializer
tools/
  layering/     the dependency-direction guard (see below)
  eslint-rules/ local ESLint rules with no upstream equivalent
docs/adr/       architecture decision records
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
