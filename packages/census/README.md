# @pptx-studio/census

**What is inside a `.pptx`.** Open a package, walk its parts and its
relationship graph, scan every XML part, and get back a plain-JSON description:
the archive, the parts, the edges, the presentation's own structure, a feature
census, every namespace with the prefixes it was spelled with, and the timings.

Apache-2.0 · browser and Web Worker only, no Node · part of
[PPTX Studio](https://github.com/Dewiride-Open-Source/Dewiride-PPTX-Studio).

```ts
import { censusPackage, formatCensus } from '@pptx-studio/census';

const census = censusPackage(new Uint8Array(await file.arrayBuffer()));

console.log(census.presentation?.slides); // 300
console.log(census.features); // [{ key: 'table', label: 'Tables', count: 30, phase: '4.1' }, …]
console.log(formatCensus(census, { namespaces: true }));
```

## Three properties, all load-bearing

**The result is plain JSON.** No `Map`, no `Set`, no class instance, no
`undefined` — histograms are arrays of `{ name, count }` and absence is `null`.
A census is computed in a Web Worker and has to survive `postMessage`, and it is
written to disk by `pptx-studio inspect --json` and has to survive
`JSON.stringify`. Choosing the intersection of the two means the two consumers
cannot disagree about what a census is.

**No tree ever leaves.** Parts are scanned over tokens, never parsed into
`XNode`s, so the working set does not grow with the deck. A tree carries a
parent pointer on every node and the whole decoded source on the document;
returning one would not be a message, it would be a second copy of the deck.
`ppt/presentation.xml` is the single part parsed into a tree, because it is a few
kilobytes and the facts wanted from it are nested — and that tree does not escape
the function that built it.

**It does not refuse.** A part that will not tokenize becomes a problem entry
and the scan carries on, because a census is what you reach for when a file is
already broken. The only thing that ends it is an archive that will not open at
all, and that throws an `OpcError` with a code.

## Features are found by namespace, never by prefix

`mc:Ignorable` and `mc:Choice/@Requires` hold **prefixes**, and the same prefix
resolves to different namespaces in different parts of one corpus. Every lookup
here goes through the scope the element was actually written in, so `zz:tbl` is
a table when `zz` is bound to DrawingML and `a:tbl` is not one when `a` is bound
to something else.

Each detected feature carries the sub-phase of the plan that renders it, which
is what turns a census into a schedule rather than a disclaimer:

```
       30  Tables                            phase 4.1
       12  Charts                            phase 9.1
        7  SmartArt                          phase 4.5
       60  Animation timelines               carried across, never rendered
```

## What this is not

It is not the validator. Sub-phase 1.2 owns the must-not-break rules and refuses
to hand over bytes when one fires; those guard a file we are about to write.
The problems reported here are observations about a file somebody else wrote,
and several of them are things PowerPoint tolerates.

It answers _what is in this file_, never _how it is arranged_. Structural
questions — which shapes are inside which group — belong to the document model
in phase 2.
