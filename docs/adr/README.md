# Architecture decision records

One record per sub-phase: what was decided, what was measured to decide it, which alternatives
the measurement ruled out, and what is still open. Records are grouped by phase so the directory
stays navigable at ninety-six of them, and this index is the ordered log that grouping would
otherwise cost.

Numbers are permanent. A record is never renumbered and never edited to say something it did not
say — a decision that turns out to be wrong gets a new record that supersedes it, and the
measurement that refuted the old one goes in the new one.

**This file is generated** by `tools/repo/adr-index.ts`; `pnpm references` fails if it is out of
date.

## Phase 0 — foundation

| #                                                                   | decision                                                              |
| ------------------------------------------------------------------- | --------------------------------------------------------------------- |
| [0001](phase-0-foundation/0001-toolchain.md)                        | Toolchain for the PPTX Studio monorepo                                |
| [0002](phase-0-foundation/0002-zip-reader.md)                       | Reading a ZIP we did not write                                        |
| [0003](phase-0-foundation/0003-the-package.md)                      | The package on top of the archive                                     |
| [0004](phase-0-foundation/0004-the-xml-layer.md)                    | The XML layer                                                         |
| [0005](phase-0-foundation/0005-the-serializer.md)                   | The serializer and the round-trip gate                                |
| [0006](phase-0-foundation/0006-order-mce-and-edits.md)              | Schema order, Markup Compatibility, and invertible edits              |
| [0007](phase-0-foundation/0007-ground-truth.md)                     | Ground truth: embedded fonts and colour transforms                    |
| [0008](phase-0-foundation/0008-inspect-worker-and-the-benchmark.md) | `inspect`, the Worker boundary, and what a 200 MB deck actually costs |

## Phase 1 — round trip

| #                                                                   | decision                                                                  |
| ------------------------------------------------------------------- | ------------------------------------------------------------------------- |
| [0009](phase-1-round-trip/0009-the-corpus.md)                       | The corpus, and where a fixture is allowed to come from                   |
| [0010](phase-1-round-trip/0010-the-repair-firewall.md)              | The repair firewall, and the difference between invalid and broken-by-us  |
| [0011](phase-1-round-trip/0011-the-writer.md)                       | The writer, and what a no-op export has to prove                          |
| [0012](phase-1-round-trip/0012-the-round-trip-oracle.md)            | The round-trip oracle, and why it is not a byte comparison                |
| [0013](phase-1-round-trip/0013-bisect-and-the-powerpoint-oracle.md) | `cli bisect`, and asking PowerPoint a question it does not want to answer |
| [0014](phase-1-round-trip/0014-the-ci-gate-and-the-badge.md)        | The CI round-trip gate, and a badge that cannot lie                       |
| [0015](phase-1-round-trip/0015-gate-1-the-browser-round-trip.md)    | Gate 1: a deck out of a browser, and into PowerPoint                      |

## Phase 2 — geometry and paint

| #                                                                       | decision                                               |
| ----------------------------------------------------------------------- | ------------------------------------------------------ |
| [0016](phase-2-geometry-and-paint/0016-preset-geometry-codegen.md)      | The 187 preset shapes, and how they get here           |
| [0017](phase-2-geometry-and-paint/0017-the-formula-evaluator.md)        | The formula evaluator                                  |
| [0018](phase-2-geometry-and-paint/0018-arc-math.md)                     | Arc math                                               |
| [0019](phase-2-geometry-and-paint/0019-path-emit-and-custgeom.md)       | Path emit, and `custGeom` as the same thing            |
| [0020](phase-2-geometry-and-paint/0020-adjust-handles.md)               | Adjust handles, and the inverse of a shape             |
| [0021](phase-2-geometry-and-paint/0021-colour.md)                       | Colour, and the second time we asked PowerPoint        |
| [0022](phase-2-geometry-and-paint/0022-fills.md)                        | Gradients and pattern fills                            |
| [0023](phase-2-geometry-and-paint/0023-lines.md)                        | Strokes and effects                                    |
| [0024](phase-2-geometry-and-paint/0024-model-parse-and-resolve.md)      | Model: parse and resolve                               |
| [0025](phase-2-geometry-and-paint/0025-renderers-geometry.md)           | Renderers, geometry only                               |
| [0026](phase-2-geometry-and-paint/0026-debug-overlay-preset-gallery.md) | Debug overlay and preset gallery                       |
| [0036](phase-2-geometry-and-paint/0036-image-fills.md)                  | Image fills, and the two readings everybody gets wrong |
| [0037](phase-2-geometry-and-paint/0037-picture-shapes.md)               | Picture shapes, and the border that is drawn outside   |
| [0038](phase-2-geometry-and-paint/0038-gate-2.md)                       | Gate 2, and the two features the corpus never had      |

## Phase 3 — text

| #                                                                     | decision                                                             |
| --------------------------------------------------------------------- | -------------------------------------------------------------------- |
| [0027](phase-3-text/0027-the-text-cascade.md)                         | The text cascade                                                     |
| [0028](phase-3-text/0028-measurement-and-the-line-model.md)           | Measurement and the line model                                       |
| [0029](phase-3-text/0029-line-breaking.md)                            | Line breaking                                                        |
| [0030](phase-3-text/0030-autofit.md)                                  | Autofit                                                              |
| [0031](phase-3-text/0031-bullets-fields-and-script-runs.md)           | Bullets, fields and script runs                                      |
| [0032](phase-3-text/0032-anchors-insets-and-vertical-text.md)         | Anchors, insets and vertical text                                    |
| [0033](phase-3-text/0033-font-substitution-and-the-guard.md)          | Font substitution and the guard                                      |
| [0034](phase-3-text/0034-text-in-both-renderers.md)                   | Text in both renderers                                               |
| [0035](phase-3-text/0035-the-fidelity-harness.md)                     | The fidelity harness                                                 |
| [0039](phase-3-text/0039-upright-glyphs-in-vertical-text.md)          | Upright glyphs in vertical text, and the insets that never turned    |
| [0040](phase-3-text/0040-the-wordart-column.md)                       | The WordArt column, and the lines that stacked the wrong way         |
| [0041](phase-3-text/0041-what-the-harness-was-not-gating.md)          | What the harness was not gating, and where the difference is         |
| [0042](phase-3-text/0042-rendering-without-a-browser.md)              | Rendering without a browser, and the first release                   |
| [0043](phase-3-text/0043-what-3-10-left-open.md)                      | What 3.10 left open, and what installing it found                    |
| [0044](phase-3-text/0044-what-the-runners-own-fonts-said.md)          | What the runner's own fonts said                                     |
| [0045](phase-3-text/0045-the-face-box-belongs-to-the-rasteriser.md)   | The face box belongs to the rasteriser                               |
| [0046](phase-3-text/0046-nothing-gated-the-thing-being-published.md)  | Nothing gated the thing being published                              |
| [0047](phase-3-text/0047-ten-packages-were-never-attested.md)         | Ten packages were never attested                                     |
| [0048](phase-3-text/0048-every-package-proves-it-can-publish.md)      | Every package proves it can publish, before any of them do           |
| [0049](phase-3-text/0049-the-right-to-publish-cannot-be-monitored.md) | The right to publish can only be asked by the workflow that holds it |
