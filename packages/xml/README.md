# @pptx-studio/xml

A byte-preserving XML tokenizer and node model for OOXML.

> **Pre-alpha.** Sub-phase 0.4 has landed the tokenizer and the tree. The serializer and its
> byte-identical round-trip gate are 0.5; schema-ordered insertion, the Markup Compatibility walker
> and the invertible edit operations are 0.6.

Part of [PPTX Studio](https://github.com/Dewiride-Open-Source/Dewiride-PPTX-Studio). Runs in the
browser and in a Web Worker; there is no Node build and there never will be.

## What it does today

```ts
import { parseXml, descendantElements, attributeValue, textContent } from '@pptx-studio/xml';

const doc = parseXml(store.read('/ppt/slides/slide1.xml'));

for (const el of descendantElements(doc.root)) {
  if (el.qname === 'a:t') console.log(textContent(el));
}

// Every node knows where it came from. While `dirty` is false, this *is* the node.
const shape = doc.root.children[0];
doc.source.slice(shape.start, shape.end); // byte-for-byte what was in the file
```

## Why not `DOMParser`

Because the DOM specification permits a serializer to rewrite namespace prefixes, and Markup
Compatibility attributes hold **prefixes, not URIs**. `mc:Ignorable="a14 p14"` and
`mc:Choice Requires="a14"` name prefixes declared elsewhere in the document. Rename one and ignorable
extension markup silently becomes a hard error in PowerPoint — with no diagnostic, because PowerPoint
does not emit one.

`@xmldom/xmldom` loses namespaces on `createElementNS` beneath a prefixed parent.
`fast-xml-parser` makes no guarantee about whitespace, self-closing form or quote style. And
`DOMParser` is not reliably available in a Web Worker, which is where our parsing runs.

## The bar is set by PowerPoint, and PowerPoint is lossy

Hand PowerPoint a slide part and ask it to save. Measured, by doing exactly that:

|                                   | PowerPoint's own resave      |
| --------------------------------- | ---------------------------- |
| comments, processing instructions | **discarded**                |
| `<![CDATA[x]]>`                   | **rewritten as plain text**  |
| `&#72;`                           | **resolved to `H`**          |
| `<p:spPr></p:spPr>`               | **collapsed to `<p:spPr/>`** |
| single-quoted attributes          | **rewritten as double**      |
| byte order mark                   | **stripped**                 |
| `xml:space="preserve"`            | **removed**                  |
| an unused `xmlns:zz`              | **dropped**                  |

Every row is something this package hands back untouched. That is the whole product: preservation is
the default state, not a feature.

## A node's source and a node's value are different things

XML 1.0 requires the parser to transform what it reads — line endings (§2.11), attribute values
(§3.3.3), references (§4.6) — and a tokenizer that stores only the transformed form can never
reproduce its input. So spans are the truth and values are derived.

The trap worth knowing, because the obvious implementation gets it backwards:

```ts
attr('a\tb'); // "a b"   — a literal tab is normalized to a space
attr('a&#9;b'); // "a\tb" — a reference to a tab is not
```

Expand references first and squash whitespace second and those two collapse into one value. The
mistake passes every test written against a corpus with no character references in it — such as ours,
which has zero across 2834 parts.

## Things that are easy to get wrong

- **`TextDecoder`'s `ignoreBOM` is named backwards.** The default, `false`, means "delete the BOM".
  **1037 of the 2834 XML parts** in our corpus carry one, so the default corrupts 37% of them — and
  invisibly, until an export is diffed against its input.
- **String offsets are not byte offsets.** Decode to a string, scan it, and report byte offsets and
  they are wrong for every part containing non-ASCII. We work in string space throughout and encode
  once at the edge; `encode(decode(bytes))` is byte-identical for all 2834 parts, and for any valid
  UTF-8, by construction.
- **A prefix does not have one meaning.** `p14` is bound to two different URIs in different parts of
  our corpus, 213 namespace declarations sit on non-root elements, and a template Microsoft ships
  contains `<p14:discardImageEditData xmlns="" xmlns:p14="…" val="0"/>` nested mid-document. So
  resolution walks up the tree; there is no table.
- **An unprefixed attribute is in no namespace**, not the default one (_Namespaces in XML_ §6.2).
- **`<x/>` and `<x></x>` are different**, and so is `<x />`. 70 822 of the 98 777 self-closing tags in
  our corpus have a space before the slash.
- **`xml:space` is not needed in PresentationML.** There is none in the corpus, 120 `<a:t>` elements
  carry edge whitespace without it, and PowerPoint round-trips them intact. This package reads the
  attribute and never writes one.

## The gate is coverage, not "no errors"

An off-by-one in a span raises no error. It parses every part of every deck happily and loses a
character the first time anything re-serializes. So `checkSpanCoverage` and `checkTreeCoverage` assert
that spans **tile the source exactly** — no gaps, no overlaps, and inside a start tag the name, each
attribute and the closing delimiter account for every character too. Every test that parses anything
runs through them.

The tag interior was added after a review: the first version tiled the nodes but never read
`trailingSpaceStart` or any attribute offset, so an off-by-one there passed the whole suite _and_
would have passed 0.5's byte-identical gate, surfacing for the first time in 0.6 on an edited deck.

Across the corpus: 2834 parts, 314 814 tokens, 194 148 elements, 170 019 attributes, **zero** gaps,
overlaps or byte mismatches.

## Hostile input

`DOCTYPE` is rejected outright, with its own error code. That removes XXE, parameter entities and
billion-laughs from the attack surface entirely rather than defending against them — and it costs
nothing, because **PowerPoint refuses a DOCTYPE too.**

The builder never recurses, because PowerPoint opens a part nested 5000 elements deep and a
recursive-descent parser would answer that with a `RangeError`. Depth, attribute count and node count
all have configurable ceilings. Every failure throws an `XmlError` with a machine-readable `code`.

## Known limitation

UTF-8 only. PowerPoint honours the `encoding` pseudo-attribute and will open a part written in
windows-1252 or UTF-16; we refuse it with `ERR_UNSUPPORTED_ENCODING`, naming the encoding. Zero of
2834 parts in our corpus are affected. A part _labelled_ windows-1252 whose content is ASCII parses
fine — the refusal is about bytes we cannot reproduce, not about labels.

## Licence

Apache-2.0
