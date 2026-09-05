# ADR 0002 — Reading a ZIP we did not write

- **Status:** Accepted, with one decision corrected by [ADR 0003](./0003-the-package.md)
- **Date:** 2026-08-26
- **Sub-phase:** 0.2

## Context

`readZip` is the first line of code in this project that touches a file a stranger produced.
Everything above it — the XML tokenizer, the model, the renderer — is written on the assumption
that the bytes it receives are the bytes the archive actually contained, and that the reader
below it did not hang, exhaust memory, or hand back a part named `../../etc/passwd`.

Two threats shape the design, and they are not the ones people usually name:

- The classic recursive `42.zip` is **irrelevant** here. We never unpack a nested archive, so the
  recursion never happens. Citing it is a way of not thinking about the problem.
- The one that matters is David Fifield's non-recursive **quoted-overlap** bomb (WOOT '19, best
  paper). Each entry's compressed data contains the local file headers of the entries that follow
  it, so many central-directory entries reference overlapping byte ranges and the declared total
  grows quadratically against the file size: 10 MB in, 281 TB out, one round of decompression,
  ordinary DEFLATE, compatible with most parsers. A per-entry ratio check does not see it.

The second shaping fact is that **every number in a ZIP header is attacker-controlled**, including
the ones that say how big things are. No limit can be enforced by believing the archive.

## Decisions

### Parse the central directory ourselves; use `fflate` only to inflate

`fflate` is already a dependency and its `unzipSync` would have been the obvious choice. It is
the wrong tool here for four separate reasons, any one of which would be sufficient:

1. **It decompresses every entry eagerly.** Architectural bet #1 is that a part is parsed only
   when something reads it. Eager inflation is both the wrong shape and precisely the
   over-allocation we are defending against.
2. **It returns a plain object keyed by entry name.** JavaScript object key order puts
   integer-like keys first, so an archive containing an entry named `1` comes back reordered.
   Entry order is what the writer replays in 0.3 to keep an untouched export byte-identical, so
   losing it silently corrupts the one property Phase 1 exists to prove. Entries are an array.
3. **It discards everything except the bytes** — general-purpose flags, CRC-32, local header
   offsets, DOS timestamps, external attributes. Every one of those is needed for either safety
   or round-tripping.
4. **It verifies nothing.** No CRC, no cross-check between the local header and the central
   directory, no bound on total inflation.

The central-directory parser is about 250 lines of `DataView` reads. That is a smaller cost than
any one of the four.

### `inflateSync(data, { out })` is a hard allocation ceiling — measured, not assumed

Reading `fflate`'s source: `inflateSync` calls `inflt(data, { i: 2 }, opts.out)`, and the internal
`resize` flag is `noBuf || st.i != 2` — false whenever a buffer is supplied. So a caller-supplied
`out` is never grown. Verified on this machine rather than inferred:

| Input                                   | `out` size | Result                      |
| --------------------------------------- | ---------- | --------------------------- |
| 10 MiB of zeros (10 239 bytes deflated) | 10 MiB     | 10 485 760 bytes            |
| the same                                | 1 024      | **1 024 bytes, no error**   |
| the same, stored DEFLATE blocks         | 1 024      | **raw `RangeError` thrown** |
| truncated stream                        | 10 MiB     | `Error: unexpected EOF` (0) |

That first result is the defence: allocation is bounded by the size we choose, so a bomb cannot
make us allocate. The second is the catch — **the surplus is dropped without an error**, so the
cap alone is silently lossy. The third is a different catch, below.

### CRC-32 is mandatory, and is not an option

`fflate` exports no CRC function, so `crc32.ts` exists. It is not there for integrity in the
disk-corruption sense; it is there because it is the only thing that distinguishes "this part is
1 024 bytes" from "this part was capped at 1 024 bytes and the rest was thrown away". The
allocation cap makes the reader safe; the checksum makes it honest. There is deliberately no
option to skip it — a flag that disables the check is a flag that gets set in a hot path one day.

### Every failure throws `OpcError`, and third-party calls are wrapped

The stored-block row of that table throws a bare `RangeError: offset is out of bounds`, which
would violate the parser invariant Phase 12.5 asserts (_"never hangs, never over-allocates, never
throws a raw `RangeError`"_) — two phases before anyone went looking for it. Every call into
`fflate` goes through `guard()`, which re-throws anything that is not already an `OpcError` as
`ERR_INFLATE_FAILED` with the original attached as `cause`. A test sweeps single-byte corruptions
across a whole archive and asserts nothing else ever escapes.

### Five independent bomb defences, because each has a blind spot

| Defence                             | Catches                                      | Blind to                                      |
| ----------------------------------- | -------------------------------------------- | --------------------------------------------- |
| `maxEntryInflatedBytes`             | one entry claiming an absurd size            | many entries each claiming a plausible size   |
| `maxTotalInflatedBytes`, running    | the aggregate, across the archive's lifetime | nothing — but only rejects, does not diagnose |
| `maxCompressionRatio` above a floor | an honest 1000:1 expansion                   | overlap, where every entry looks reasonable   |
| structural overlap detection        | the quoted-overlap construction, by shape    | —                                             |
| `maxEntries`                        | the construction needs a great many entries  | —                                             |

Two details that are easy to get wrong:

- **The counter is charged per entry, not per read.** Re-reading a part must not consume the
  archive's allowance a second time, or opening the same deck twice in one session fails for a
  reason no user could act on.
- **The ratio test needs a floor.** A 40-byte DEFLATE stream expanding to a 12 KB
  `[Content_Types].xml` is a 300:1 ratio and completely ordinary. A ratio only means anything once
  the absolute size is large enough to hurt, so it is not applied below `ratioFloorBytes` (1 MiB).
  Without that floor the check either fires on real decks or is set so loose it fires on nothing.

Overlap detection is worth stating plainly: a legitimate archive **never** shares bytes between
entries, so refusing the shape costs nothing and answers the quoted-overlap bomb structurally
rather than merely surviving it.

### The OPC part-name grammar makes zip-slip structurally impossible

This was the surprise of the sub-phase. ECMA-376 Part 2 §9.1.1 rule `M1.9` forbids a segment
ending in a dot, and `M1.10` requires every segment to hold at least one non-dot character.
Together they mean a conformant part name **cannot contain `.` or `..` at all**. Path traversal
is not a special case bolted onto the reader; it is the grammar refusing to parse.

We still report it as its own rule (`X1.4`) rather than as "segment ends with a dot", because the
latter is not what someone reading a bug report needs to know. And `..` still has to be handled
in exactly one place — `resolveRelativeTarget`, since relationship targets in `.rels` parts are
ordinary relative URI references that legitimately use it. That resolution is clamped at the
package root.

Four rules are ours rather than the spec's, for hostility ECMA-376 never contemplated: control
characters and NUL (`X1.1`), backslashes (`X1.2`), a leading drive letter (`X1.3` — a colon is a
legal `pchar`, so `/C:/Windows/System32/evil.dll` satisfies the OPC grammar completely), and a
length ceiling (`X1.5`).

### Violations carry a severity, and only some are fatal

`M1.6` (non-`pchar`) and `M1.8` (over-encoded unreserved characters) are warnings; everything
else is fatal. A media part named `my image.png` is out of spec and completely harmless, and
failing a whole deck over it would be worse than tolerating it. Everything security-relevant or
structural is fatal.

> **Corrected in 0.3.** `my image.png` is _not_ harmless: PowerPoint refuses a package containing
> it outright (`0x808D1001`), as it does for `#` and any non-ASCII byte, while accepting `%20`.
> The rule and its severity are both kept — reading stays lenient — but the writer now refuses a
> `M1.6` violation. See ADR 0003.

The cost of that choice is that `[Content_Types].xml` — whose `[` and `]` are not `pchar` — would
pass as a part name on warnings alone. It is refused explicitly in `partNameFromZipEntry`
instead, which is the right place: the content-type stream is not a part, has no content type of
its own, and never appears in a relationship. The ZIP-to-OPC boundary is exactly where that fact
lives.

### The end-of-central-directory scan is strict

The EOCD signature can occur inside compressed data, and a scan that stops at the first match can
be steered onto an attacker-chosen central directory. The record is accepted only when its
declared comment length lands its end exactly at the end of the buffer. That also rejects
archives with data appended after the comment, which is the shape of a ZIP polyglot — a
restriction we take deliberately.

### ZIP64 is read but will never be written

Our writer emits ZIP32 only, and at a 200 MB ceiling nothing needs ZIP64. But other producers
emit a ZIP64 end-of-central-directory record even for small archives, and refusing to read one
would fail files that are otherwise perfectly ordinary. Roughly 45 lines buys that compatibility:
the locator, the record, and extra field `0x0001` — whose fields are in a fixed order but present
only when the corresponding 32-bit field held the `0xFFFFFFFF` sentinel, so the layout cannot be
parsed without knowing which sentinels fired.

### `DecompressionStream('deflate-raw')` considered, and deferred

Baseline since May 2023 and at 94.25% global support (Chrome/Edge 103, Firefox 113, Safari 16.4),
it would remove the `fflate` dependency from the read path entirely and is natively fast. It is
not used, for two reasons: it is asynchronous and stream-shaped, where the part store wants a
synchronous call inside a Worker; and it offers no output-size ceiling, so the allocation bound
would have to be rebuilt as a counting transform. The inflate call sits behind one function, so
revisiting this later — most plausibly for large media parts, off-thread — is a local change.

## What the first real decks taught us

The reader was run against 37 files written by Office itself (nine template `.potx` files shipped
with Office 16 and 28 sample decks): 2 961 entries, 10.0 MiB inflated, every CRC-32 verified, in
471 ms.

- **Every part name in every deck is fully conformant** — not one warning, let alone a fatal.
  That is reassuring about the strict/lenient split, and a caution: the lenient half is untested
  by Office output and only earns its place against files from other producers.
- **Office writes none of the ZIP features that are hardest to handle.** No data descriptors, no
  directory entries, no UTF-8-flagged names, no ZIP64. All four are covered by synthetic fixtures
  precisely because a corpus of Office output would never have exercised them.
- **A password-protected `.pptx` is not a damaged ZIP.** Saved through PowerPoint COM, it comes
  back as an OLE2 compound document (`d0cf11e0a1b11ae1`) with the presentation inside an
  `EncryptedPackage` stream. "Not a ZIP archive" is true and useless; `ERR_ENCRYPTED_PACKAGE`
  says the thing the user can act on. A pre-2007 binary `.ppt` is indistinguishable at this layer,
  and the message says so.

## Deferred

- **Streaming reads.** The whole archive is held in memory, which the 200 MB ceiling makes
  acceptable and which the writer in 0.3 needs anyway for byte-for-byte passthrough.
- **A CP437 round-trip test.** The high-byte table exists so an error message names a non-UTF-8
  entry correctly rather than as mojibake; no Office file exercises it.
- **`InflationBudget.reset()`.** Deliberately absent. A budget that can be reset is a budget an
  attacker can get reset.
