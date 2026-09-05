# 0011 — The writer, and what a no-op export has to prove

Date: 2026-08-29
Status: **accepted** — 51/51 decks export with zero parts re-serialized, 1470 entries compared
Sub-phase: 1.3 — `writer`: dirty-part-only export, media GC, `prepare` hooks

---

## Context

The plan's 1.3 is three clauses and one verification: _dirty-part-only export;
media mark-and-sweep GC with layouts/masters excluded by default; `prepare`
hooks for later phases. Verify: no-op export of all 50 decks._

That last line is the whole of Phase 1 compressed into a sentence. The first
architectural bet — the OPC package **is** the document, and our model is a
projection of it — is falsifiable exactly here: if a deck can be opened and
written back with nothing re-serialized, then preservation is the default state
of the export path and every later phase inherits it for free. If it cannot, the
architecture is wrong, and the point of putting this before a single renderer
line is to find out now.

Most of the mechanism already existed. `PartStore.write` streams archive-backed
parts through `passthroughEntry` since 0.3, and `assertValid` has refused bad
packages since 1.2. What did not exist was the **path**: the order things happen
in, who is allowed to contribute to it, and — the part that turned out to
matter — whether anything checks the archive we actually emit rather than our
intention to emit it.

---

## Decisions

### The export is four steps and the order is load-bearing

`prepare` hooks → collect → write → check (preservation, then the 29 rules).

**Hooks before collection**, because a hook is the thing most likely to change
what is referenced. Sub-phase 8.7 both adds font parts and drops the ones no run
uses; a sweep that ran first would be answering a question about a different
package.

**Hooks before validation**, because a hook writes markup, and markup this
session synthesized is the riskiest kind there is. A hook that ran after
validation would be the one part of an export nothing checked.

**Both checks after the write**, which is the ordering that looks backwards and
is not. A check that runs before the write is checking an intention; what a user
opens is an archive. `V003` (no directory entries, no ZIP64), the duplicate half
of `V002`, and the preservation check below are all questions about the emitted
container that no amount of inspecting the store can answer. So bytes are
produced, then judged, then either returned or thrown away — never returned
unjudged.

### The preservation check reads the archive, because `V027` structurally cannot

`V027` is "a part nobody edited comes back out byte for byte", and it is a good
rule. It compares what the store would hand you for a part against what the
baseline store would hand you — and for an untouched part **both answers come
from the same archive**. It can therefore catch a part replaced without being
marked edited, and it cannot, even in principle, catch a writer that
re-serialized a clean part on the way out: by the time those bytes exist the rule
has already run against a store that knows nothing about them.

So the writer keeps its own check, and it compares the **stored** bytes — still
compressed — of the emitted archive against those of the source archive, along
with the method, the CRC and the uncompressed size. Comparing compressed bytes is
both cheaper (nothing is inflated on either side) and stricter: two different
DEFLATE encodings of identical content pass an inflate-and-compare and fail this.
Stricter is what is wanted, because any difference at all means something
re-compressed, and re-compressing is the failure.

It is not raw ZIP byte equality, which the plan rules out for 1.4 and which would
be wrong here for the same reason: `passthroughEntry` deliberately does not carry
an entry's DOS timestamp or external attributes, because those describe where a
file came from rather than what is in it.

Not run is reported as not run. `preservation.skipped` carries a sentence when
the source bytes were not supplied, because a check that reports zero either way
is worse than no check.

### Media GC: timid in three separate directions, on purpose

**The sweep set is `/ppt/media/`.** The plan says layouts and masters are
excluded by default and the reason generalises: PowerPoint keeps layouts no slide
uses, which is how the layout picker has anything to offer. Tidying one away
means "change layout" silently loses an option the user had a moment ago, to save
a few kilobytes, while the media is the megabytes.

**The walk is rooted at the package root, not at the slide list.** This is the
same decision made a second time and it is the half that is easy to miss. The
natural implementation asks "which images do the slides use?"; it is wrong
because a layout no slide currently uses still has its background picture, and
switching a slide onto that layout later has to still work. Rooting above the
masters gets it right with no special case.

**By default it collects only what this session orphaned**, by differencing
reachability against the baseline — the same question `origin` asks about
findings in ADR 0010, asked about parts. A deck can arrive with an orphaned
image; collecting it would mean that merely opening and saving a file changes it,
which breaks the property this sub-phase exists to establish, in the direction
where nobody notices for months.

Two properties fall out and are worth stating as properties rather than hopes:

- **It cannot create a dangling relationship.** A part is collected only when
  nothing in the package resolves to it, so there is no edge left to break. The
  case that kills naive implementations — two slides sharing one image, one
  deleted — cannot arise, because the surviving relationship still marks the
  image live.
- **It is therefore a safety net and not a delete.** It collects nothing until
  something else has removed the relationship, and removing it is part of
  deleting a picture, which is a document operation belonging to Phase 5. The
  corpus test asserts this from the other side: removing only the relationship is
  **refused** by `V006`, because eight bytes of markup still name it.

The walk refuses rather than guesses. If any relationship part will not parse,
nothing is collected — an unreadable edge is an invisible edge, and an invisible
edge makes its target look like garbage. `@pptx-studio/census` has a walk that
swallows the same error and carries on, which is correct for a report and would
be a data-loss bug here; its own comment says as much.

`planCollection` reports and never throws, so an export is never refused over an
unrelated malformed `.rels` that may well have arrived that way. `collectGarbage`
throws, for a caller who asked for a sweep and would misread silence as "nothing
to sweep".

### `prepare` hooks, so this package never learns what a font is

Written inline, save-time work becomes a list of special cases in the export
function that grows by one every phase, that the writer's own tests have to know
about, and that cannot be tested without standing up a font stack. As hooks, each
lands in the package that owns it and ships with the sub-phase that needs it.

Three are already known: 8.7's six font artifacts plus its collection pass, 3.4's
`@fontScale`/`@lnSpcReduction` write-back, and 10.8's `p14:media` /
`a:videoFile` duplicate. The order is the caller's and is never inferred — some
pairs genuinely depend on each other (autofit must settle before a font pass
decides which typefaces are used) and sorting by a notion of priority would be
guessing at a dependency the caller knows for certain.

---

## What building it found

### Relationship edits were invisible to everything that reads bytes

The one real defect, and it was latent in `PartStore` since 0.3 rather than new.

A relationship part has two representations: its bytes, and the parsed
`Relationships` the store hands out and caches. `relationships(...)` returns a
live object, so adding or removing an edge mutates the parsed form and leaves the
bytes alone until `write` reconciles them.

That was invisible while `write` was the only consumer, because it asked the
parsed form. It stopped being invisible the moment 1.2's validator read the same
package: `validate` reads relationship markup as **raw XML on purpose**, because
`Relationships.parse` refuses exactly the duplicate and malformed ids that three
of the rules exist to report. So it was checking the `.rels` as it arrived rather
than as it was about to be written — an edge removed still looked present, one
added still looked absent, and a `.rels` created this session was not in
`partNames` at all, so nothing checked it.

Found the way these things are found: a test exported a deck with one image
unlinked and got a dangling-relationship finding for the edge it had just
removed.

Fixed at the store, where the knowledge lives. `materializeRelationships()`
writes dirty collections back into the part table; `write` calls it, and so does
`exportPackage` before anything reads bytes. Idempotent, and it makes `read`,
`info`, `partNames`, the validator and a census all agree on the package about to
be emitted.

`PartStore.rewrittenParts()` was added alongside it for the same reason:
"dirty-part-only export" is the claim this package makes, and `PartInfo.fromArchive`
alone cannot express it — a collection edited in place leaves `fromArchive` true
while `write` re-emits the part.

### The reachability walk could not borrow the census's optimisation

The census restricts its walk to parts that own a `.rels`, which on a package
being read is free. Here it is a bug: a collection created this session has no
`.rels` in `partNames` yet, so keying on existing relationship parts makes exactly
the edges this session added invisible — and those are the edges a writer is
about. One failing test, and the optimisation it replaced was buying a `Map`
lookup per part.

### The corpus writes no `xml:space="preserve"` at all

1.2 measured 120 `<a:t>` elements across the corpus carrying edge whitespace
_without_ the attribute, which is why `V029` treats it as preserved rather than
required. The other half of that measurement turned up here: not one deck writes
it either. So nothing in the corpus can exercise that clause of the rule, and the
negative test has to seed a deck with the attribute before it can take it away.
Recorded as its own assertion rather than worked around silently.

---

## Consequences

`pnpm check` green. 47 test files, 1269 tests.

The verification the plan asked for, with the numbers it produced:

- **51/51 decks export.** `rewritten` is empty for every one — not a single part
  of a single deck was re-serialized.
- **1470 entries compared and identical**, stored bytes, method, CRC and size.
  That is every entry of every archive, including each deck's content-type
  stream.
- **All 29 rules ran on each**, against a baseline. `validate.test.ts` runs 26
  because a file on the command line has no history; this is the first time the
  three preservation rules have seen real markup.
- **Entry names and order preserved** across all 51.

Negative tests, because "the rules found nothing" is only worth something if they
can find something: an edit that drops an `xml:space="preserve"` is refused with
`V029`; removing an image relationship without removing the markup that uses it
is refused with `V006`; and deleting the `<p:pic>` as well collects the image,
leaves the other picture's image alone, and re-serializes exactly two parts.

### What is not done

- **`cli roundtrip` is 1.4 and `cli bisect` is 1.5.** This package refuses bad
  exports; it does not diagnose them, and there is no CLI verb for it yet.
- **No export has been opened in PowerPoint.** The corpus decks were opened in
  PowerPoint when they were built, and the preservation check proves an exported
  archive holds the same stored bytes — but "PowerPoint opens the file we wrote"
  is a separate claim, and 1.5's scripted COM loop is where it gets made.
- **The GC has never collected anything on a real user deck**, because nothing
  orphans media yet. Its corpus test does the orphaning by hand. Phase 5.6 is the
  first caller.
- **`prepare` has no hooks.** The protocol is tested; the three known
  implementations belong to 3.4, 8.7 and 10.8.
- **Two `PartStore` methods are new public API** — `materializeRelationships` and
  `rewrittenParts` — added rather than reaching into the store from the writer.
  Both are consequences of the store's lazy relationship model, and if that model
  ever changes, both go.
