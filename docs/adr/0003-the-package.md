# ADR 0003 — The package on top of the archive

- **Status:** Accepted
- **Date:** 2026-08-26
- **Sub-phase:** 0.3

## Context

Sub-phase 0.2 got bytes out of an archive that might be hostile. This one turns those bytes into a
package: what each part _is_, how parts refer to each other, and how to put it all back without
damaging anything we did not touch.

The specification answers most of these questions and PowerPoint answers some of them differently.
Where the two disagree, PowerPoint decides, because acceptance by PowerPoint is the only test with
consequences. So the method here was not to reason from ECMA-376 about what a consumer _should_
reject — it was to take real decks, break them one way at a time, and hand each to the PowerPoint
installed on this machine.

| What we broke                                       | PowerPoint               |
| --------------------------------------------------- | ------------------------ |
| `<Default>` for a media extension actually in use   | **refused** `0x80CB8002` |
| `[Content_Types].xml` deleted                       | **refused** `0x80CB8002` |
| `_rels/.rels` deleted                               | **refused** `0x80070570` |
| `officeDocument` relationship target does not exist | **refused** `0x80070570` |
| `Id="1rId"` — not an `xsd:ID`                       | **refused** `0x80070570` |
| duplicate `Id` in one `.rels`                       | **refused** `0x80070570` |
| `[Content_Types].xml` written **last**              | opens, intact            |
| directory entries added                             | opens, intact            |
| UTF-8 flag set on every entry name                  | opens, intact            |
| `<Override>` for `slide1.xml` deleted               | opens, intact            |

Two things in that table are worth pausing on.

The **two distinct error codes** are a map of PowerPoint's own layering. `0x80CB8002` is the
content-type layer refusing the package before anything reads a slide; `0x80070570` is the
relationship graph. Knowing which one you have tells you where to look.

The **asymmetry in the last row** is the opposite of what you would guess. Deleting an `Override`
is survivable — the slide falls back to the `Default` for `xml`, gets typed `application/xml`,
and PowerPoint opens the deck anyway. Deleting a `Default` that some part depends on is fatal.
A part with _no content type at all_ is the thing that cannot be forgiven, and that, rather than
"every part has the right content type", is the assertion worth making.

## Decisions

### Read leniently, write strictly

Anything PowerPoint opens, we open. Anything PowerPoint refuses, we refuse to write.

Concretely: `contentTypes.for(part)` returns `undefined` for an untyped part rather than throwing,
so a deck with an odd corner still loads and still renders; `write` calls `require` on every part
and throws `ERR_MISSING_CONTENT_TYPE`.

The asymmetry is the point. Refusing to load is a bad experience for a user who has a file that
works everywhere else. Refusing to save is a good one, because the alternative is handing back a
file that fails silently on somebody else's machine hours later.

None of the write-time assertions is configurable, for the reason the CRC check in 0.2 is not: a
flag that turns off a correctness check is a flag somebody sets in a hot path.

Drawing the line in the right place turned out to be the hard part, and it took two goes. The next
two sections are both corrections: one where being strict broke preservation, and one where being
lenient quietly laundered a broken file straight back out.

### A broken edge is refused only if _we_ broke it

The first version of this code threw on any relationship whose internal target was missing, on the
reasoning that a relationship pointing at nothing is a mistake either way. That was wrong, and
measuring it settled it: we added an unreferenced image relationship pointing at a part that does
not exist to a real deck, and PowerPoint opened it with every slide and picture intact. The failure
is not "dangling" — it is "a hole in the traversal PowerPoint actually performs". A stale image
relationship nothing follows is survivable; a theme relationship a master follows is not.

We cannot tell those apart, because deciding it means reading `r:id` attributes out of part markup
and this layer cannot see inside a part until 0.4. But refusing all of them is the worse error: it
makes a deck that _arrived_ with a stale relationship impossible to export, which is exactly the
preservation this package exists to provide.

So the test is not "is this edge broken" but **"did we break it"**. Every `Relationship` carries an
`origin` of `archive` or `added`, and the store remembers the names of parts removed this session.
`write` refuses an edge we minted that points at nothing, and an edge whose target we deleted —
both are bugs in the code that did it, caught while the cause is still nameable. An edge that
arrived broken is preserved, and reported by `danglingRelationships()` so that tolerating it is not
the same as hiding it.

The bookkeeping is a set of names rather than severing inbound edges on delete: severing means
walking and parsing every `.rels` part on every removal, hundreds of parses per deleted slide, to
learn something one `Set.has` already knows.

### Anything we tolerate on read but would not write must set `dirty`

This one is a gap that only shows up where two reasonable rules meet. `ContentTypes.parse` accepts
a duplicate `<Default>` whose two content types agree, because failing a whole deck over a
redundant line would be worse than tolerating it. Separately, `write` streams the original
content-type bytes through untouched when nothing changed. Together, they meant we read a package
with a duplicate, said nothing, and wrote the duplicate straight back out.

And that duplicate is fatal. `M2.5` is unconditional, and PowerPoint enforces it literally: two
`<Default>`s for one extension are refused with `0x80CB8000` **even when the two elements are
byte-identical**. So lenient reading was quietly laundering corruption — accepting a file we could
have fixed and handing back one that still does not open.

The rule that closes it: **any anomaly we accept on read but would not write ourselves sets the
dirty bit**, so the part regenerates instead of passing through. Two lines. The effect is that
round-tripping such a package _repairs_ it — verified end to end: PowerPoint refuses the input with
`0x80CB8000` and opens our output.

### A non-`pchar` character in a part name is fatal — correcting ADR 0002

ADR 0002 rated `M1.6` a warning, reasoning that "a media part named `my image.png` is out of spec
and completely harmless, and failing a whole deck over it would be worse than tolerating it."

The first half is right and the second is false. `ppt/media/my image.png` — orphaned, fully
content-typed, referenced by nothing — makes PowerPoint refuse the entire package with
`0x808D1001`. So does `#`, and so does any non-ASCII byte. Meanwhile `%20` is accepted, along with
every `pchar` we already allow literally: `+ , ~ ! ( ) $ @ :`. In other words the rule is exactly
right and only its severity was wrong.

The fix follows the same lenient/strict split as everything else rather than reversing the ADR:
`M1.6` stays a warning in the grammar, so a package containing such a name still _loads_, and
`write` refuses it — pointing at the percent-encoded spelling, which is the same name written the
way OPC intends.

### The content-type stream is found case-insensitively and written canonically

ZIP entry names are case-sensitive and §7.3.7 fixes the spelling, so an exact-name lookup is the
defensible reading. It is also wrong in practice: PowerPoint opens a package whose stream is named
`[content_types].xml`, and refusing one would fail a file that works everywhere else for a reason
no user could act on.

Read either spelling, write the canonical one. The normalisation happens at exactly one place — the
entry loop in `open` — so that everything downstream compares against `CONTENT_TYPES_PART` and is
right without knowing this happened.

### A relationship may not target another relationship part

ECMA-376 Part 2 says implementers "shall treat any such relationship as invalid", and PowerPoint
agrees by refusing the package. Relationship parts are reached by naming convention — `_rels/` plus
`.rels` — and never by relationship, so a package claiming otherwise has a cycle in a graph that
does not admit one. `ERR_INVALID_RELATIONSHIP_TARGET`, checked at write time alongside the others.

### Entry order is preserved, not normalised — a deliberate change from the plan

The plan specified a writer "emitting `[Content_Types].xml` first, `_rels/.rels` second, then
original entry order", and verified by "read→write a deck, assert identical entry name set and
order". Those two clauses contradict each other for any package that did not already have that
layout, and a lot of them do not.

The evidence settled it:

- PowerPoint 365 does write `[Content_Types].xml` first and `_rels/.rels` second. So does every
  PowerPoint-authored file we have.
- A third-party producer of 29 packages in our corpus writes `ppt/presentation.xml` first and puts
  the content-type stream elsewhere entirely. Those files open in PowerPoint, and PowerPoint
  re-saves them without complaint.
- We rebuilt a real deck with the content-type stream written **last** and PowerPoint opened it
  with every slide, layout, shape and picture intact.

So the ordering is a convention that helps a consumer streaming an archive front to back, not a
requirement. Reordering buys nothing measurable and costs the one property this sub-phase exists to
establish. **We preserve the original order.** `normalizeEntryOrder` gives Office's layout on
request, and a package built with `PartStore.create` gets it anyway, having no original order to
preserve.

The gate is then stated exactly: identical entry name set **and** order, plus identical bytes for
every part.

### `flat-xml.ts` is not the XML layer, and does not become one in 0.4

`[Content_Types].xml` and every `.rels` part are flat: a root element and one level of empty
children, no text content, no mixed content, no `mc:AlternateContent`, no `extLst`. Their grammar
is fixed by ECMA-376 **Part 2**, not by the Part 1 schemas that 0.6 generates order tables from, so
it cannot drift. And they are never edited node by node — when the content-type map or a
relationship collection changes, we regenerate the whole part, so there is nothing for byte-level
preservation to preserve.

They get a reader sized to that job: about 250 lines, no `XNode`, no offsets. The real tokenizer in
0.4 exists for slide parts, where none of the above is true.

What this reader does share with the rest of the package is its posture. **A DOCTYPE is refused
outright**, with its own error code rather than a generic parse failure — that one rule removes XXE
and billion-laughs from this package's attack surface entirely, because there is no DTD subset left
to expand. An undeclared entity reference is likewise a hard error. The five predefined entities
and numeric character references are resolved, which is not exotic: every hyperlink target with a
query string carries `&amp;`.

### An untouched part is never decompressed

`ZipArchive.raw` hands back the stored DEFLATE stream; `passthroughEntry` moves it, its CRC-32 and
both sizes into the new archive together, recomputing only the local header offset. A part we did
not edit is therefore not merely re-encoded to the same bytes — it is never inflated at all.

That is what makes preservation the default state rather than a feature. We do not have to
understand a chart, a SmartArt diagram, an animation tree or an OLE object to carry it across, so
there is no category of content we silently drop.

Measured on the corpus: 29 of 37 packages come back **byte-for-byte identical as whole files**. The
other 8 differ only in ZIP header metadata that carries no OPC meaning — DOS timestamps we zero for
reproducibility, the deflate level-hint flag bits, version-made-by, and a handful of local-header
extra fields. Every one of those 8 has a **compressed-size delta of exactly zero on every entry**:
no part's data changes in any file.

### `rId` allocation counts from the highest id, not from the count

`Id` is an `xsd:ID`, so it is an XML `NCName` — it cannot begin with a digit, cannot contain a space
or a colon — and its scope is the single `.rels` part. Every part's relationships start again at
`rId1`. A package-global allocator is not an optimisation, it is a bug that surfaces as a duplicate
in the one file that mattered.

Allocating `rId{count + 1}` is the same bug wearing a different hat, and the corpus says so: of
1296 relationship parts measured, **99 have gaps in their numbering** and **175 list their
relationships out of numeric order**. PowerPoint 365's own `_rels/.rels` lists `rId3`, `rId2`,
`rId1` in that document order. So we take one past the highest number actually used, then step over
anything that would still collide.

Relationship order is preserved on write but means nothing. Slide order lives in `p:sldIdLst`.

### A relationship target resolves against the source part's folder

Not against the folder of the `.rels` part. For `/ppt/slides/_rels/slide1.xml.rels` the base is
`/ppt/slides/`, which is what makes `../slideLayouts/slideLayout1.xml` mean the right thing; for
`/_rels/.rels` the source part is the package root, so the base is `/`. Getting this wrong turns
every root relationship into `/_rels/ppt/...` and nothing resolves.

`sourcePartNameForRels` is the inverse of `relsPartNameFor` and exists so that base is computed in
exactly one place.

Both absolute and relative targets are read — roughly three quarters of the relationships in the
third-party corpus are absolute — and only relative ones are written, because that is what Office
writes.

### Part names keep their percent-encoding all the way through

A part name may contain percent-encoded characters, and the ZIP entry name is the part name minus
its leading slash, verbatim. So nothing decodes and nothing re-encodes, at any layer. That keeps
the two namespaces in exact correspondence and is why `M1.8` — over-encoded unreserved characters —
is a warning rather than a fatal: we are not normalising, so we have no reason to insist.

### A leading dot is still an extension — a bug in 0.2, found by 0.3

`partExtension('/_rels/.rels')` returned `''` in sub-phase 0.2, following the Unix convention that
a dotfile has no extension. OPC does not have that convention: the extension is whatever follows
the final dot of the final segment.

It matters because the root relationship part is typed by `<Default Extension="rels"/>` in every
package we have measured and carries an `Override` in none of them. With the old behaviour it
resolved to no content type at all, which is the one package state PowerPoint refuses outright. The
write-time assertion caught it on the first real deck.

### `CONTENT_TYPES_PART` is an entry name, `ROOT_RELS_PART` is a part name

Deliberate asymmetry, previously an accident. `[Content_Types].xml` has no part-name form because
it is not a part — nothing types it and nothing relates to it. `_rels/.rels` _is_ a part, so it is
named the way parts are named, with a leading slash. Conflating the two is how a lookup silently
misses, which it did, once, before the constant was changed.

### The writer's field values were measured, not chosen

Every value in the local and central headers was read out of a file PowerPoint 365 saved on this
machine: version made by 20 with host 0, version needed 10 for stored and 20 for deflated, general
purpose flags 0, timestamp 1980-01-01, empty extra fields, zeroed attributes, no archive comment.

Two of those are worth naming. PowerPoint **zeroes the timestamp** on every entry, which is what
makes our output reproducible for free rather than by policy. And it **stores** already-compressed
media at method 0 rather than inflating a PNG by a few bytes for nothing, which `deflatedEntry`
reproduces by falling back to stored whenever DEFLATE would not be smaller.

The UTF-8 flag is a choice rather than a compatibility matter: Office never sets bit 11, it does
not need to because the part-name grammar admits only `pchar`, and we confirmed that setting it
anyway does not bother PowerPoint. The honest flag for an all-ASCII name is zero, so that is what
we write. A non-ASCII entry name is refused rather than quietly flagged, because it means something
skipped percent-encoding.

`fflate.deflateSync` emits a **raw** DEFLATE stream — verified, not assumed: its output begins
`cb 48` where a zlib-wrapped stream would begin `78 9c`, and Node's `inflateRawSync` round-trips it.
A zlib header inside a ZIP entry is a corruption no reader reports as one.

### Reads are not cached; parses are

`PartStore.read` inflates on every call and checks the CRC on every call; the decompression budget
is charged once. Caching inflated bytes as well as the parsed form would double the memory of a
200 MB deck to keep a copy nobody reads twice. Relationship collections _are_ cached, because they
are the parsed form and because a caller that mutates one must see its own change.

### Removing a part does not collect garbage

`removePart` removes the part, its relationship part and its content-type `Override`. It does not
scan for relationships that pointed at it. Finding them means walking every `.rels` in the package,
and deciding what to do about each is a question about the document rather than the container.

`write` refuses to emit a dangling relationship, so the mistake surfaces at export with the part
named — rather than as a file PowerPoint declines to open. Media mark-and-sweep is 1.3, and it
needs the whole graph.

## What the corpus taught us, and a correction

The 0.2 report described the corpus as "37 real Office files". That is right for 8 of them and
wrong for 29. Only the `.potx` templates carry a `docProps/app.xml` naming PowerPoint 12 or 14; the
29 sample decks have no `docProps/app.xml` at all and were written by some other OOXML producer.

The distinction matters because the two producers disagree about nearly everything optional:

|                                | PowerPoint                                 | the other producer                 |
| ------------------------------ | ------------------------------------------ | ---------------------------------- |
| first entry                    | `[Content_Types].xml`                      | `ppt/presentation.xml`             |
| `Target` form                  | relative                                   | absolute (`/ppt/...`)              |
| `Relationship` attribute order | `Id, Type, Target`                         | `Type, Target, Id`                 |
| `Default Extension="xml"`      | `application/xml`                          | `…presentation.main+xml`           |
| byte-order mark                | none                                       | `ef bb bf`                         |
| XML declaration                | `UTF-8`, `standalone="yes"`, then **CRLF** | `utf-8`, no `standalone`, no break |
| empty-element tag              | `.../>`                                    | `... />`                           |
| `docProps/app.xml`             | present                                    | absent                             |

Two rows are traps worth stating on their own.

**You cannot assume `Default Extension="xml"` is `application/xml`.** In 29 of our packages it is
the presentation content type, so `ppt/presentation.xml` is typed by the `Default` and carries no
`Override` at all.

**PowerPoint writes a CRLF after the XML declaration** and nothing else anywhere — the rest of the
document is one line. It is invisible in every tool that shows a part as text, and we missed it
first time round because our census decoded through `TextDecoder`, which had already eaten the
other producer's BOM as well. Both were caught by comparing raw bytes instead. `XML_DECLARATION`
now carries the CRLF and a test pins it to `0d 0a`, so a regenerated package part is byte-comparable
against an Office-authored one and the useful diff is the empty one.

A note on citations: the rule identifiers this package uses — `M1.1`–`M1.12`, `M2.1`–`M2.10` — come
from ECMA-376 Part 2 **1st edition** and ISO/IEC 29500-2:2008. The 5th edition (Dec 2021) renumbered
every clause and dropped the identifiers entirely, and it also renamed the content-types stream to
"the Media Types stream". We keep the old identifiers because they are what POI, python-pptx and
every bug report in this ecosystem say, and a rule number nobody recognises helps nobody.

One place the editions genuinely disagree matters to us: the 5th edition exempts relationship parts
from needing a content-type entry, and the 1st does not. **PowerPoint sides with the 1st edition** —
delete the `rels` `Default` and it refuses the package — so our write-time assertion covers every
part, `.rels` included.

Having two disagreeing producers is better evidence than 37 files from one, and it is the reason
the reader is lenient in exactly the places it is. What neither producer supplies is a single
dangling relationship, duplicate id, missing content type, colliding part name, percent-encoded
target or malformed part name — 2961 parts, and not one. So every defence in this package is
exercised by synthetic fixtures, which is what `testing/build-package.ts` is for.

## Verification

- 242 unit tests in real Chromium, up from 120.
- Read→write across 37 packages from two producers: identical entry name set and order, identical
  bytes for every one of 2924 parts, 2884 relationships resolved.
- All 37 round-tripped files, **and all 37 with every part re-deflated and re-framed by our own
  writer**, opened in the installed PowerPoint with identical slide, layout, shape, picture and
  character counts. The rewritten set is 3.6% smaller.
- Each write-time assertion corresponds to a package we broke deliberately and watched PowerPoint
  refuse — and each tolerated case to one we broke and watched it open.
- A package with a duplicate `<Default>`, which PowerPoint refuses with `0x80CB8000`, opens after
  being round-tripped through this code.

## Deferred

- **Streaming or incremental writes.** The whole output is built in memory. At a 200 MB ceiling
  that is affordable, and the passthrough path means we hold compressed bytes rather than inflated
  ones.
- **Dirty-part-only export.** Nothing is parsed above the byte level yet, so "dirty" currently
  means "replaced". The real thing arrives with `XmlEdit` in 1.3.
- **Package-level validation beyond the OPC layer.** The 29 rules are 1.2. This sub-phase asserts
  only what we watched PowerPoint reject.
- **`.rels` parts written for a part that does not exist.** Legal, meaningless, and currently
  neither produced nor rejected.
- **Digital signature parts.** `_xmlsignatures/` round-trips as ordinary parts; nothing validates
  or re-signs, and editing a signed package silently invalidates its signature.
