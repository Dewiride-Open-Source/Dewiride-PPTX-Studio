# Scope

Two lists. The first is what PPTX Studio is for. The second is what it will not become, and it is
published so that it survives contact with enthusiastic contributors — including us.

## What this is

A browser PowerPoint renderer and editor that:

1. **Renders any `.pptx` faithfully**, entirely client-side. Nothing is uploaded anywhere.
2. **Makes every element on a slide selectable and editable**, rather than rendering a picture.
3. **Stays bound to slide layouts and masters**, so layouts, themes, colours and backgrounds
   genuinely cascade. This is the difference between us and every other web editor that imports
   `.pptx`: they flatten inheritance at parse time, which is precisely why their slides feel fixed.
4. **Exports a file that still opens in real PowerPoint**, with everything we did not touch left
   byte-for-byte intact — charts, SmartArt, animations, OLE objects, macros, ink.
5. **Handles fonts end to end.** Normal fonts work; users can upload custom fonts; those fonts
   survive into the exported `.pptx` and render on a machine that has never had them installed.

## What this is not

**Not a PowerPoint clone.** The target is Canva's simplicity with PowerPoint's inheritance model,
not feature parity with a thirty-year-old desktop application. A request of the form "PowerPoint can
do X" is not by itself an argument for doing X.

**Not two hundred toolbar options.** The contextual toolbar is capped at seven controls. Adding an
eighth means removing one or moving it to the inspector. This is a hard constraint, not a
guideline.

The following are **omitted by design**. Each is a decision with a reason, not a gap waiting for a
volunteer:

| Omitted                                                                         | Why                                                                                                                                                                                                                                                                                               |
| ------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| **The SmartArt layout engine** (`dgm:layoutDef`, `alg`, `constrLst`, `forEach`) | It is a full declarative layout language. We render the `drawingN.xml` fallback instead, read-only. That is a _correctness_ requirement, not a shortcut: PowerPoint regenerates the fallback from data + layout + quickStyle + colors on any diagram interaction, so edits to it silently vanish. |
| **Animation authoring**                                                         | We preserve `p:timing` byte-for-byte and never rewrite it. Building an authoring UI for it means understanding it, and understanding it means being able to corrupt it.                                                                                                                           |
| **Editing chart data**                                                          | Charts are render-only. `chart1.xml` and its relationships are copied verbatim. Rewriting them is a leading cause of PowerPoint repair prompts, and the embedded workbook is a second format we would then own.                                                                                   |
| **Editing OLE payloads**                                                        | `ppt/embeddings/*.bin` is OLE2/CFB. Any rewrite is a guaranteed repair prompt. We render the preview and preserve the payload.                                                                                                                                                                    |
| **A server-side rendering service as a product**                                | `apps/render-service` exists only as a pinned reference renderer for CI fidelity comparisons. The product is client-side.                                                                                                                                                                         |
| **Real-time collaboration in v1**                                               | The CRDT taxes are paid up front — stable ids on every slide, shape, paragraph and run; addressing by id and never by array index; every mutation through the command layer — so that adopting Yjs later is an addition rather than a rewrite. Adopting it now is not.                            |
| **A `.ppt` (binary, pre-2007) reader**                                          | Different format, different decade, no shared code.                                                                                                                                                                                                                                               |
| **Feature parity with the desktop Format pane**                                 | See "not two hundred toolbar options".                                                                                                                                                                                                                                                            |

## The content contract

Every kind of content on a slide falls into one of four tiers, and the tier determines what the
editor lets you do with it. This is uniform on purpose: "every element is selectable and movable"
holds for all of it, at near-zero risk of producing a file PowerPoint refuses.

| Tier           | Content                                                                                 | You can                                                                     |
| -------------- | --------------------------------------------------------------------------------------- | --------------------------------------------------------------------------- |
| **Full edit**  | Shapes, text, pictures, groups, tables                                                  | Everything                                                                  |
| **Frame edit** | Charts, SmartArt, OLE, ink, 3-D models                                                  | Move, resize, rename, set alt text; the content renders but is not editable |
| **Preserve**   | Animations, macros, custom XML, extension lists, unknown `mc:AlternateContent` branches | Nothing — round-tripped byte-for-byte                                       |
| **Reject**     | Nothing yet                                                                             | —                                                                           |

The `Frame edit` row is why sub-phase 10.11's verification is "every element selectable and movable,
and every chart/diagram/OLE/ink part still byte-identical after export". Both halves, at once.

## Changing this document

Adding to "what this is not" is cheap. Removing from it requires a written case in an issue that
explains what the feature costs in maintenance and in export risk, and what gets removed to make
room in the UI. Silence is not agreement.
