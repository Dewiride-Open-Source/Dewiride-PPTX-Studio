# @pptx-studio/opc

The Open Packaging Conventions container for `.pptx`: ZIP reading and writing, parts, content types
and relationships.

> **Pre-alpha.** Sub-phase 0.2 has landed the ZIP reader, its decompression budgets and the OPC
> part-name grammar. The part store, content-type resolution, relationships and the writer arrive
> in 0.3.

Part of [PPTX Studio](https://github.com/Dewiride-Open-Source/Dewiride-PPTX-Studio). Runs in the
browser and in a Web Worker; there is no Node build and there never will be.

## What it does today

```ts
import { readZip, partNameFromZipEntry, isContentTypesStreamName } from '@pptx-studio/opc';

const archive = readZip(new Uint8Array(await file.arrayBuffer()));

for (const entry of archive.entries) {
  // Entries are an array, in central-directory order, because that order is
  // what the writer replays to keep an untouched export byte-identical.
  if (isContentTypesStreamName(entry.name) || entry.isDirectory) continue;

  const partName = partNameFromZipEntry(entry.name); // throws on anything hostile
  const bytes = archive.read(entry); // inflates, verifies CRC-32, charges the budget
}
```

Nothing is inflated until you ask for it. `archive.raw(entry)` hands back the stored bytes without
inflating or charging budget — that is what the writer will stream through for parts nobody edited.

## The threat model, in one paragraph

Every number in a ZIP header is attacker-controlled, including the ones that say how big things
are, so no limit here is enforced by believing the archive. Reads inflate into a pre-allocated
buffer sized from the declared length, which makes allocation bounded by a number we choose;
`fflate` silently discards anything the stream produces past that, so **every read ends in a
CRC-32 check** — the cap makes the reader safe, the checksum makes it honest. On top of that sit a
running inflated-bytes counter that spans the whole archive, an entry-count ceiling, a compression
ratio ceiling that only applies above an absolute size, and structural rejection of entries whose
byte ranges overlap — which is the shape of Fifield's quoted-overlap bomb, and something no
legitimate archive ever does.

Every failure throws an `OpcError` with a machine-readable `code`. Never a `RangeError`, never
whatever a dependency happened to throw.

## Part names

`validatePartName` implements the ECMA-376 Part 2 §9.1.1 grammar by its own rule identifiers
(`M1.1`–`M1.12`), plus four of ours for hostility the spec never contemplated: control characters
(`X1.1`), backslashes (`X1.2`), a leading drive letter (`X1.3`) and a length ceiling (`X1.5`).

Worth knowing: **the OPC grammar makes zip-slip structurally impossible.** `M1.9` forbids a
segment ending in a dot and `M1.10` requires a non-dot character in every segment, so a conformant
part name cannot contain `.` or `..` at all. Traversal is reported as `X1.4` only so the failure
reads as traversal rather than as a dot-placement nit.

Violations carry a severity. A media file named `my image.png` is out of spec (`M1.6`) and
completely harmless, so it is a warning; everything structural or security-relevant is fatal and
`toPartName` throws on it.

## What arrives in 0.3

- A `PartStore` that keeps every part's original bytes and parses lazily.
- `[Content_Types].xml` `Default`/`Override` resolution with case-insensitive extension matching.
- Per-part relationship handling with a **per-part** rId allocator — `rId` values are `xsd:ID`
  scoped to a single `.rels` file, so a global registry is a bug, not an optimisation.
- A ZIP32 writer that emits `[Content_Types].xml` first, `_rels/.rels` second, then the original
  entry order, with no directory entries and no ZIP64.

## Licence

Apache-2.0
