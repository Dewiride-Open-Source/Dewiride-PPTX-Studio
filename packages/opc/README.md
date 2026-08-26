# @pptx-studio/opc

The Open Packaging Conventions container for `.pptx`: ZIP reading and writing, parts, content types
and relationships.

> **Pre-alpha.** Sub-phase 0.3 has landed the part store, content-type resolution, relationships and
> the ZIP32 writer. Part _content_ is still opaque bytes — the XML layer is 0.4.

Part of [PPTX Studio](https://github.com/Dewiride-Open-Source/Dewiride-PPTX-Studio). Runs in the
browser and in a Web Worker; there is no Node build and there never will be.

## What it does today

```ts
import { PartStore, REL_TYPE, CONTENT_TYPE } from '@pptx-studio/opc';

const store = PartStore.open(new Uint8Array(await file.arrayBuffer()));

// Nothing finds anything by filename. Walk the relationship graph.
const root = store.rootRelationships();
const presentation = root.resolve(root.firstOfType(REL_TYPE.officeDocument)!);

for (const rel of store.relationships(presentation).byType(REL_TYPE.slide)) {
  const slide = store.relationships(presentation).resolve(rel);
  const xml = store.read(slide); // inflated here, CRC-checked, budget charged once
}

// Add a slide: the part, its content type, and the relationship that reaches it.
store.addPart('/ppt/slides/slide9.xml', CONTENT_TYPE.slide, bytes);
store.relationships(presentation).addTo(REL_TYPE.slide, '/ppt/slides/slide9.xml');

const out = store.write(); // every untouched part passes through uninflated
```

## Preservation is the default state, not a feature

A part you did not edit is never decompressed. `write` copies its stored DEFLATE stream, CRC-32 and
both sizes straight across, recomputing only the local header offset. So charts, SmartArt,
animations, ink and OLE objects survive an export without this package understanding any of them —
which is the only way that survival is achievable at all.

Measured across 37 packages from two different producers: **29 come back byte-for-byte identical as
whole files**, and the other 8 differ only in ZIP header metadata that carries no OPC meaning
(timestamps, compression-level hint bits, version-made-by). Every entry in all 37 has a
compressed-size delta of exactly zero.

## Read leniently, write strictly

Anything PowerPoint opens, this opens. Anything PowerPoint refuses, this refuses to write.

`contentTypes.for(part)` returns `undefined` for an untyped part so an odd deck still loads;
`write` throws `ERR_MISSING_CONTENT_TYPE`. The write-time assertions are not configurable, and each
one corresponds to a package we broke deliberately and watched the installed PowerPoint reject:

| Broken how                                       | PowerPoint           |
| ------------------------------------------------ | -------------------- |
| `<Default>` for a media extension in use removed | refused `0x80CB8002` |
| `[Content_Types].xml` removed                    | refused `0x80CB8002` |
| `_rels/.rels` removed                            | refused `0x80070570` |
| `officeDocument` target dangling                 | refused `0x80070570` |
| `Id="1rId"` — not an `xsd:ID`                    | refused `0x80070570` |
| duplicate `Id` in one `.rels`                    | refused `0x80070570` |

Two consequences of that split are worth knowing, because both are the opposite of the obvious
choice and both were settled by measuring rather than reasoning:

**A broken relationship is refused only if we broke it.** A relationship pointing at a missing part
is _not_ automatically fatal — PowerPoint opens a deck containing one that nothing dereferences.
Refusing all of them would make a deck that arrived with a stale relationship impossible to export.
So every `Relationship` carries an `origin`, and `write` refuses one we added or one whose target we
deleted, while preserving one that arrived broken. `danglingRelationships()` reports the survivors,
so tolerating them is not the same as hiding them.

**Round-tripping repairs some packages.** Anything tolerated on read that we would not write
ourselves marks the part dirty, so it regenerates instead of passing through. A duplicate
`<Default>` — which PowerPoint refuses with `0x80CB8000` even when the two entries are identical —
goes in refused and comes out opening.

## Things that are easy to get wrong

- **A relationship target resolves against the _source part's_ folder**, not the `.rels` part's. For
  `/ppt/slides/_rels/slide1.xml.rels` the base is `/ppt/slides/`; for `/_rels/.rels` it is `/`.
- **`Id` is scoped to one `.rels` file.** Every part starts again at `rId1`. A package-global
  allocator is a bug.
- **Allocate from the highest id, not the count.** 99 of 1296 real relationship parts have gaps in
  their numbering, and 175 list their relationships out of numeric order.
- **`Default Extension="xml"` is not always `application/xml`.** In 29 packages we measured it is
  the presentation content type, and `ppt/presentation.xml` has no `Override` at all.
- **`/_rels/.rels` has extension `rels`**, not none. It is typed through the `Default`, and treating
  it as a dotfile leaves it with no content type — the one state PowerPoint refuses outright.
- **`[Content_Types].xml` first is a convention, not a requirement.** We checked: a deck rebuilt
  with it written last opens intact. Entry order is preserved rather than normalised; pass
  `normalizeEntryOrder` for Office's layout.

## The XML in here is not the XML layer

`flat-xml.ts` reads the two package-level grammars and nothing else. They are flat, attribute-only,
fixed by ECMA-376 Part 2, and regenerated wholesale when they change, so they need none of the
byte-offset `XNode` machinery that `@pptx-studio/xml` builds in 0.4 for slide parts.

It refuses a **DOCTYPE outright**, with its own error code. That single rule removes XXE and
billion-laughs from this package's attack surface, because there is no DTD subset left to expand.

## Reading an untrusted archive

Every number in a ZIP header is attacker-controlled, including the ones that say how big things
are, so no limit here is enforced by believing the archive. Reads inflate into a pre-allocated
buffer sized from the declared length, which bounds allocation by a number we choose; `fflate`
silently discards anything past that, so **every read ends in a CRC-32 check**. On top of that sit a
running inflated-bytes counter spanning the whole archive, an entry-count ceiling, a compression
ratio ceiling that only applies above an absolute size, and structural rejection of entries whose
byte ranges overlap — the shape of Fifield's quoted-overlap bomb, and something no legitimate
archive ever does.

Every failure throws an `OpcError` with a machine-readable `code`. Never a `RangeError`, never
whatever a dependency happened to throw.

## What arrives later

- **0.4–0.6** `@pptx-studio/xml`: the tokenizer, `XNode`, byte-identical re-serialization, and
  `XmlEdit` with exact inverses.
- **1.3** Dirty-part-only export and media mark-and-sweep. "Dirty" here currently means "replaced".
- **1.2** The 29 validation rules. This package asserts only the OPC layer.

## Licence

Apache-2.0
