# Where the build has got to

**37 of 109 sub-phases recorded, 3 of 13 gates closed.**

A sub-phase is done when it has an ADR, because that is when the working agreement says the
record gets written. ⚠️ means done with something still on the record — the note says what.

**This file is generated** from `docs/plan/phases.json` by `tools/repo/plan-status.ts`;
`pnpm references` fails if it is out of date. Update the JSON at the end of a sub-phase, in the
same commit as its ADR.

| phase                                                           | sub-phases | gate           |
| --------------------------------------------------------------- | ---------- | -------------- |
| [0 — Foundation and ground truth](#phase-0)                     | 8/8        | ✅ done        |
| [1 — Byte-perfect round trip](#phase-1)                         | 6/6        | ✅ done        |
| [2 — Geometry and paint](#phase-2)                              | 14/14      | ✅ done        |
| [3 — Text engine, viewer, fidelity scoreboard](#phase-3)        | 9/10       | ⬜ not started |
| [4 — Tables and SmartArt](#phase-4)                             | 0/6        | ⬜ not started |
| [5 — Select and move](#phase-5)                                 | 0/8        | ⬜ not started |
| [6 — Type](#phase-6)                                            | 0/5        | ⬜ not started |
| [7 — Layout-aware](#phase-7)                                    | 0/11       | ⬜ not started |
| [8 — Fonts, end to end](#phase-8)                               | 0/9        | ⬜ not started |
| [9 — Charts, core](#phase-9)                                    | 0/8        | ⬜ not started |
| [10 — ChartEx, 3-D, and the remaining content types](#phase-10) | 0/11       | ⬜ not started |
| [11 — Theme and template assets](#phase-11)                     | 0/6        | ⬜ not started |
| [12 — 1.0](#phase-12)                                           | 0/7        | ⬜ not started |

<a id="phase-0"></a>

## Phase 0 — Foundation and ground truth

| #          | sub-phase                                          | status  | record                                                                     |
| ---------- | -------------------------------------------------- | ------- | -------------------------------------------------------------------------- |
| 0.1        | Repo skeleton                                      | ✅ done | [0001](../adr/phase-0-foundation/0001-toolchain.md)                        |
| 0.2        | `opc`: ZIP reader and safety                       | ✅ done | [0002](../adr/phase-0-foundation/0002-zip-reader.md)                       |
| 0.3        | `opc`: parts, content types, relationships, writer | ✅ done | [0003](../adr/phase-0-foundation/0003-the-package.md)                      |
| 0.4        | `xml`: tokenizer and XNode                         | ✅ done | [0004](../adr/phase-0-foundation/0004-the-xml-layer.md)                    |
| 0.5        | `xml`: serializer and the round-trip gate          | ✅ done | [0005](../adr/phase-0-foundation/0005-the-serializer.md)                   |
| 0.6        | `xml`: order, MCE, invertible edits                | ✅ done | [0006](../adr/phase-0-foundation/0006-order-mce-and-edits.md)              |
| 0.7        | Ground-truth experiments A, B and C                | ✅ done | [0007](../adr/phase-0-foundation/0007-ground-truth.md)                     |
| 0.8        | `cli inspect`, Worker shell, benchmark             | ✅ done | [0008](../adr/phase-0-foundation/0008-inspect-worker-and-the-benchmark.md) |
| **Gate 0** | Drop a deck, get a live explorer of its internals  | ✅ done | [0008](../adr/phase-0-foundation/0008-inspect-worker-and-the-benchmark.md) |

<a id="phase-1"></a>

## Phase 1 — Byte-perfect round trip

| #          | sub-phase                                                         | status  | record                                                                                            |
| ---------- | ----------------------------------------------------------------- | ------- | ------------------------------------------------------------------------------------------------- |
| 1.1        | Corpus, with a CI-enforced licence field                          | ✅ done | [0009](../adr/phase-1-round-trip/0009-the-corpus.md)                                              |
| 1.2        | `validate`: the 29-rule repair firewall                           | ✅ done | [0010](../adr/phase-1-round-trip/0010-the-repair-firewall.md)                                     |
| 1.3        | `writer`: dirty-part export and media GC                          | ✅ done | [0011](../adr/phase-1-round-trip/0011-the-writer.md)                                              |
| 1.4        | `cli roundtrip`                                                   | ✅ done | [0012](../adr/phase-1-round-trip/0012-the-round-trip-oracle.md)                                   |
| 1.5        | `cli bisect` and the PowerPoint loop                              | ✅ done | [0013](../adr/phase-1-round-trip/0013-bisect-and-the-powerpoint-oracle.md)                        |
| 1.6        | CI gate and badge                                                 | ✅ done | [0014](../adr/phase-1-round-trip/0014-the-ci-gate-and-the-badge.md)                               |
| **Gate 1** | A deck out of a browser and into PowerPoint with no repair prompt | ✅ done | [0015](../adr/phase-1-round-trip/0015-gate-1-the-browser-round-trip.md) — 52/52 decks, 1474 parts |

<a id="phase-2"></a>

## Phase 2 — Geometry and paint

| #          | sub-phase                                                                   | status  | record                                                                                                                                                                                        |
| ---------- | --------------------------------------------------------------------------- | ------- | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| 2.1        | Preset codegen, all 187 shapes                                              | ✅ done | [0016](../adr/phase-2-geometry-and-paint/0016-preset-geometry-codegen.md)                                                                                                                     |
| 2.2        | Formula evaluator                                                           | ✅ done | [0017](../adr/phase-2-geometry-and-paint/0017-the-formula-evaluator.md)                                                                                                                       |
| 2.3        | Arc math                                                                    | ✅ done | [0018](../adr/phase-2-geometry-and-paint/0018-arc-math.md)                                                                                                                                    |
| 2.4        | Path emit and `custGeom` unification                                        | ✅ done | [0019](../adr/phase-2-geometry-and-paint/0019-path-emit-and-custgeom.md)                                                                                                                      |
| 2.5        | Adjust handles                                                              | ✅ done | [0020](../adr/phase-2-geometry-and-paint/0020-adjust-handles.md)                                                                                                                              |
| 2.6        | Colour                                                                      | ✅ done | [0021](../adr/phase-2-geometry-and-paint/0021-colour.md)                                                                                                                                      |
| 2.7        | Gradients and patterns                                                      | ✅ done | [0022](../adr/phase-2-geometry-and-paint/0022-fills.md) — the off-centre path focus is SVG's focal radial, 17 probes at rms 0.53; @flip is inert on a gradient fill                           |
| 2.8        | Lines and effects                                                           | ✅ done | [0023](../adr/phase-2-geometry-and-paint/0023-lines.md)                                                                                                                                       |
| 2.9        | Model: parse and resolve                                                    | ✅ done | [0024](../adr/phase-2-geometry-and-paint/0024-model-parse-and-resolve.md)                                                                                                                     |
| 2.10       | Renderers, geometry only                                                    | ✅ done | [0025](../adr/phase-2-geometry-and-paint/0025-renderers-geometry.md)                                                                                                                          |
| 2.11       | Debug overlay and 187-preset gallery                                        | ✅ done | [0026](../adr/phase-2-geometry-and-paint/0026-debug-overlay-preset-gallery.md)                                                                                                                |
| 2.12       | Image fills (`a:blipFill`)                                                  | ✅ done | [0036](../adr/phase-2-geometry-and-paint/0036-image-fills.md) — 63 probes, 14 refuted readings; a03-fills-02 8199 to 9378 bp                                                                  |
| 2.13       | Picture shapes (`p:pic`)                                                    | ✅ done | [0037](../adr/phase-2-geometry-and-paint/0037-picture-shapes.md) — 23 of 24 readings refuted at a full channel; a picture's border is drawn wholly outside it; b09-picture-01 8642 to 9776 bp |
| 2.14       | Gate 2's side-by-side, and the corpus gaps it found                         | ✅ done | [0038](../adr/phase-2-geometry-and-paint/0038-gate-2.md) — no committed slide carried a rotation, a flip or a non-default background until a44 and a45; a03-fills-02 9378 to 9638 bp          |
| **Gate 2** | A real deck rendered silent: fills, strokes, effects, nested rotated groups | ✅ done | [0038](../adr/phase-2-geometry-and-paint/0038-gate-2.md) — all nine named features carried by a committed slide and drawn beside PowerPoint's own rather than LibreOffice's; pnpm gate2       |

<a id="phase-3"></a>

## Phase 3 — Text engine, viewer, fidelity scoreboard

| #          | sub-phase                                                              | status         | record                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                    |
| ---------- | ---------------------------------------------------------------------- | -------------- | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| 3.1        | Text parse and the 10-source cascade                                   | ✅ done        | [0027](../adr/phase-3-text/0027-the-text-cascade.md)                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                      |
| 3.2        | Measurement and the line model                                         | ✅ done        | [0028](../adr/phase-3-text/0028-measurement-and-the-line-model.md)                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                        |
| 3.3        | Line breaking                                                          | ✅ done        | [0029](../adr/phase-3-text/0029-line-breaking.md)                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                         |
| 3.4        | Autofit, view and edit modes                                           | ✅ done        | [0030](../adr/phase-3-text/0030-autofit.md)                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                               |
| 3.5        | Bullets, fields, script runs                                           | ✅ done        | [0031](../adr/phase-3-text/0031-bullets-fields-and-script-runs.md)                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                        |
| 3.6        | Anchors, insets, vertical text                                         | ✅ done        | [0040](../adr/phase-3-text/0040-the-wordart-column.md) — every `ST_TextVerticalType` draws: the WordArt cell is seven sixths of the font box and square, exact on 7 of 9 faces, so the corpus scores 155 of 155 slides with none refused. Two bugs the fixtures had been catching unread: the insets did not turn with the frame, and `mongolianVert` stacked its columns in reverse                                                                                                                                                                                                                                                                                                                                                                                                                      |
| 3.7        | Font substitution and the guard                                        | ✅ done        | [0033](../adr/phase-3-text/0033-font-substitution-and-the-guard.md)                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                       |
| 3.8        | Text in both renderers                                                 | ✅ done        | [0034](../adr/phase-3-text/0034-text-in-both-renderers.md)                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                |
| 3.9        | Fidelity harness                                                       | ⚠️ done        | [0041](../adr/phase-3-text/0041-what-the-harness-was-not-gating.md) — the gate had been reporting PASS over 155 slides while gating 152, because a slide the baseline had never heard of counted as unchanged; it now refuses to score one it holds no digest for, and `--record` needs a written reason and the exact count it will move. Every differing cell is charged to the shape that painted it, which found a third bad shape on `a03-fills-02` that two passes over its heatmap had missed and put numbers on two of 2.7's recorded open questions. The clock a slide draws moves 31 cells at maxD 6, an order of magnitude under the rasteriser's own noise — Green as a developer gate; one workflow_dispatch run of record-fidelity-baseline records the Linux baseline CI will gate against |
| 3.10       | First npm release: `render-svg` and `cli render`                       | ⬜ not started |                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                           |
| **Gate 3** | A 100-slide deck rendered faithfully at any zoom, entirely client-side | ⬜ not started |                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                           |

<a id="phase-4"></a>

## Phase 4 — Tables and SmartArt

| #          | sub-phase                                                                       | status         | record |
| ---------- | ------------------------------------------------------------------------------- | -------------- | ------ |
| 4.1        | Table parse: grid, spans, merges                                                | ⬜ not started |        |
| 4.2        | The 74 built-in table styles, re-derived from PowerPoint                        | ⬜ not started |        |
| 4.3        | The 13-layer table cascade                                                      | ⬜ not started |        |
| 4.4        | Table render and cell editing                                                   | ⬜ not started |        |
| 4.5        | SmartArt fallback resolution                                                    | ⬜ not started |        |
| 4.6        | SmartArt render, read-only                                                      | ⬜ not started |        |
| **Gate 4** | Banded tables with merges edit cell by cell; SmartArt renders with real content | ⬜ not started |        |

<a id="phase-5"></a>

## Phase 5 — Select and move

| #          | sub-phase                                                              | status         | record |
| ---------- | ---------------------------------------------------------------------- | -------------- | ------ |
| 5.1        | Command bus and history                                                | ⬜ not started |        |
| 5.2        | Selection and hit-test                                                 | ⬜ not started |        |
| 5.3        | Resize and rotate                                                      | ⬜ not started |        |
| 5.4        | Snapping and the overlay canvas                                        | ⬜ not started |        |
| 5.5        | Preview/Commit gesture                                                 | ⬜ not started |        |
| 5.6        | Slide operations                                                       | ⬜ not started |        |
| 5.7        | React shell and inspector v1                                           | ⬜ not started |        |
| 5.8        | Export verification after every gesture class                          | ⬜ not started |        |
| **Gate 5** | Drag, resize, rotate, recolour, reorder - and export still opens clean | ⬜ not started |        |

<a id="phase-6"></a>

## Phase 6 — Type

| #          | sub-phase                                                       | status         | record |
| ---------- | --------------------------------------------------------------- | -------------- | ------ |
| 6.1        | ProseMirror schema and lossless converters                      | ⬜ not started |        |
| 6.2        | A single lazy EditorView                                        | ⬜ not started |        |
| 6.3        | IME                                                             | ⬜ not started |        |
| 6.4        | Live autofit                                                    | ⬜ not started |        |
| 6.5        | Inspector provenance                                            | ⬜ not started |        |
| **Gate 6** | Click a placeholder once and type; one undo for the whole burst | ⬜ not started |        |

<a id="phase-7"></a>

## Phase 7 — Layout-aware

| #          | sub-phase                                                          | status         | record |
| ---------- | ------------------------------------------------------------------ | -------------- | ------ |
| 7.1        | Matcher hardening, the 121-case matrix                             | ⬜ not started |        |
| 7.2        | Change-layout plan, pure                                           | ⬜ not started |        |
| 7.3        | Preview UI                                                         | ⬜ not started |        |
| 7.4        | Change-layout apply                                                | ⬜ not started |        |
| 7.5        | Reset                                                              | ⬜ not started |        |
| 7.6        | Theme verb: Colors                                                 | ⬜ not started |        |
| 7.7        | Theme verb: Fonts                                                  | ⬜ not started |        |
| 7.8        | Theme verb: Design, the master swap                                | ⬜ not started |        |
| 7.9        | Backgrounds                                                        | ⬜ not started |        |
| 7.10       | Theme-ification                                                    | ⬜ not started |        |
| 7.11       | Adapt layouts                                                      | ⬜ not started |        |
| **Gate 7** | Switch layout from a picker that said in advance what would happen | ⬜ not started |        |

<a id="phase-8"></a>

## Phase 8 — Fonts, end to end

| #          | sub-phase                                                                      | status         | record |
| ---------- | ------------------------------------------------------------------------------ | -------------- | ------ |
| 8.1        | SFNT reader                                                                    | ⬜ not started |        |
| 8.2        | EOT reader                                                                     | ⬜ not started |        |
| 8.3        | EOT writer                                                                     | ⬜ not started |        |
| 8.4        | Rights gate on `fsType`                                                        | ⬜ not started |        |
| 8.5        | WOFF and WOFF2 decode                                                          | ⬜ not started |        |
| 8.6        | FontFace registry                                                              | ⬜ not started |        |
| 8.7        | The six-artifact export                                                        | ⬜ not started |        |
| 8.8        | `fonts-metric-compat` and the font panel                                       | ⬜ not started |        |
| 8.9        | Clean-machine verification                                                     | ⬜ not started |        |
| **Gate 8** | Custom fonts survive into exported PowerPoint on a machine that never had them | ⬜ not started |        |

<a id="phase-9"></a>

## Phase 9 — Charts, core

| #          | sub-phase                                                                 | status         | record |
| ---------- | ------------------------------------------------------------------------- | -------------- | ------ |
| 9.1        | Chart parse and colour resolution                                         | ⬜ not started |        |
| 9.2        | Axes and scales                                                           | ⬜ not started |        |
| 9.3        | Bar and column                                                            | ⬜ not started |        |
| 9.4        | Line and area                                                             | ⬜ not started |        |
| 9.5        | Pie and doughnut                                                          | ⬜ not started |        |
| 9.6        | Scatter, bubble, radar                                                    | ⬜ not started |        |
| 9.7        | Legends, labels, titles                                                   | ⬜ not started |        |
| 9.8        | Combo charts                                                              | ⬜ not started |        |
| **Gate 9** | Charts render recognizably; export leaves every chart part byte-identical | ⬜ not started |        |

<a id="phase-10"></a>

## Phase 10 — ChartEx, 3-D, and the remaining content types

| #           | sub-phase                                    | status         | record |
| ----------- | -------------------------------------------- | -------------- | ------ |
| 10.1        | ChartEx data model                           | ⬜ not started |        |
| 10.2        | Waterfall, funnel, Pareto                    | ⬜ not started |        |
| 10.3        | Treemap, sunburst                            | ⬜ not started |        |
| 10.4        | Histogram, box-whisker, region map           | ⬜ not started |        |
| 10.5        | 3-D chart variants                           | ⬜ not started |        |
| 10.6        | EMF and WMF                                  | ⬜ not started |        |
| 10.7        | SVG, TIFF, animated GIF                      | ⬜ not started |        |
| 10.8        | Media                                        | ⬜ not started |        |
| 10.9        | OLE                                          | ⬜ not started |        |
| 10.10       | Ink and 3-D models                           | ⬜ not started |        |
| 10.11       | Kitchen-sink verification                    | ⬜ not started |        |
| **Gate 10** | Nothing on a real slide is a blank rectangle | ⬜ not started |        |

<a id="phase-11"></a>

## Phase 11 — Theme and template assets

| #           | sub-phase                                         | status         | record |
| ----------- | ------------------------------------------------- | -------------- | ------ |
| 11.1        | Palette generator in OKLCH                        | ⬜ not started |        |
| 11.2        | Contrast validation, WCAG and APCA                | ⬜ not started |        |
| 11.3        | Font pairing sets                                 | ⬜ not started |        |
| 11.4        | Master authoring pipeline                         | ⬜ not started |        |
| 11.5        | Three to five original CC0 masters                | ⬜ not started |        |
| 11.6        | Design gallery UI                                 | ⬜ not started |        |
| **Gate 11** | Pick a palette and watch a real deck restyle live | ⬜ not started |        |

<a id="phase-12"></a>

## Phase 12 — 1.0

| #           | sub-phase                                                                 | status         | record |
| ----------- | ------------------------------------------------------------------------- | -------------- | ------ |
| 12.1        | Slide virtualization                                                      | ⬜ not started |        |
| 12.2        | Image cache                                                               | ⬜ not started |        |
| 12.3        | Accessibility                                                             | ⬜ not started |        |
| 12.4        | Accessibility check panel                                                 | ⬜ not started |        |
| 12.5        | Fuzzing                                                                   | ⬜ not started |        |
| 12.6        | Docs site                                                                 | ⬜ not started |        |
| 12.7        | Release                                                                   | ⬜ not started |        |
| **Gate 12** | A public demo: 200 slides, first paint under 3s, published fidelity score | ⬜ not started |        |

## Carried debt

Open across sub-phases, and not owned by any of them.

### CI cannot run the fidelity gate

A baseline for a platform has to be recorded on that platform and this machine is Windows, so there is no `expected.linux-x64.json`. One `workflow_dispatch` run fixes it. Turborepo inputs for the task are deferred until then.

Raised in: ADR 0035, open questions 1 and 2.

### `packages/validate` is still at 1.2's 29 rules

Four text sub-phases each identified rules that belong in the repair firewall and none were added: `p:kinsoku/@lang`, three `bodyPr` rules, `@panose` shape, and `@u`/`@strike` simple types.

Raised in: ADR 0029, ADR 0032, ADR 0033, ADR 0034.

### The `@pptx-studio` npm scope and GitHub org are unclaimed

The plan names this the first action on approval. Every package is version 0.0.0 and publishable, so 3.10 is now blocked on it.

Raised in: the plan.
