# @pptx-studio/cli

**Command line tools for PPTX packages.** As of sub-phase 0.8 that is one verb:
`inspect`, which reads a deck and tells you what is in it.

Apache-2.0 · part of
[PPTX Studio](https://github.com/Dewiride-Open-Source/Dewiride-PPTX-Studio).

```bash
npx @pptx-studio/cli inspect deck.pptx
npx @pptx-studio/cli inspect deck.pptx --parts --namespaces
npx @pptx-studio/cli inspect deck.pptx --json --out census.json
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

## Options

| flag           | what it does                                                    |
| -------------- | --------------------------------------------------------------- |
| `--json`       | the census as JSON, for a script or for committing as a fixture |
| `--parts`      | the per-part table: bytes, elements, depth, relationships       |
| `--namespaces` | every namespace, with the prefixes it was spelled with          |
| `--top <n>`    | rows per histogram before truncating (default 15)               |
| `--out <file>` | write to a file instead of stdout                               |

## Everything it knows lives somewhere else

`@pptx-studio/census` does the reading, and that package has no Node in it at
all — it runs in a browser tab and in a Web Worker. This package supplies the
two things a browser cannot: a path off the filesystem and a stream to write to.

That split is the point. The drop-a-deck explorer in `apps/studio` computes
exactly the same object in a Worker, so an answer here and an answer in a tab
cannot drift apart.

## Not built yet

The plan gives this package six more verbs. Asking for one says which sub-phase
brings it rather than "unknown command":

| verb        | what it will do                                              | sub-phase |
| ----------- | ------------------------------------------------------------ | --------- |
| `validate`  | run the must-not-break rules over a package                  | 1.2       |
| `roundtrip` | read a deck and write it back, then prove nothing moved      | 1.4       |
| `bisect`    | find the smallest change that makes PowerPoint repair a deck | 1.5       |
| `fidelity`  | score a render against a reference                           | 3.9       |
| `render`    | render slides to SVG or PNG without a browser                | 3.10      |
| `resolve`   | show where a resolved property came from                     | 7.x       |
