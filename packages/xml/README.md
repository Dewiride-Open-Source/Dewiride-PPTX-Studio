# @pptx-studio/xml

A byte-preserving XML tokenizer, node model and serializer for OOXML.

> **Pre-alpha.** Sub-phase 0.1 has landed the package; the tokenizer arrives in 0.4, the serializer
> and its round-trip gate in 0.5, and schema ordering, MCE handling and the edit ops in 0.6. Today
> this exports only the namespace vocabulary those are written against.

Part of [PPTX Studio](https://github.com/Dewiride-Open-Source/Dewiride-PPTX-Studio). Runs in the
browser and in a Web Worker.

## Why not `DOMParser`

Because the DOM specification permits a serializer to rewrite namespace prefixes, and Markup
Compatibility attributes hold **prefixes, not URIs**. `mc:Ignorable="a14 p14"` and
`mc:Choice Requires="a14"` name prefixes declared elsewhere in the document. Rename one and
ignorable extension markup silently becomes a hard error in PowerPoint — with no diagnostic, because
PowerPoint does not emit one.

`@xmldom/xmldom` loses namespaces on `createElementNS` beneath a prefixed parent. `fast-xml-parser`
makes no guarantees about whitespace, self-closing form or quote style. And `DOMParser` is not
reliably available in a Web Worker, which is where our parsing runs.

So: a hand-rolled pull tokenizer over `Uint8Array` that records a byte span for every node. A node
nobody edited re-emits by slicing the original buffer. Editing one clears the span on that node and
its ancestors, and nothing else in the part moves — which matters because a slide part routinely
contains a `p:timing` animation tree and an `mc:AlternateContent` ink block that dragging a shape
must not disturb.

`DOCTYPE` is rejected outright, which removes the XXE and billion-laughs surface entirely rather
than defending against it.

## Licence

Apache-2.0
