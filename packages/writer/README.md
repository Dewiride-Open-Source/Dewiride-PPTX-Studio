# @pptx-studio/writer

**Handing the file back.** A part nobody edited is never serialized — its stored
DEFLATE stream is copied from the source archive to the output archive without
ever being inflated. Charts, SmartArt, animations, OLE objects and macros
survive because nothing had to understand them.

Apache-2.0 · browser and Web Worker only, no Node · part of
[PPTX Studio](https://github.com/Dewiride-Open-Source/Dewiride-PPTX-Studio).

```ts
import { exportPackage, openPackage } from '@pptx-studio/writer';

const pkg = openPackage(bytes); // { store, baseline, baselineBytes }
pkg.store.replacePart('/ppt/slides/slide1.xml', edited);

const { bytes: out, rewritten, streamed } = exportPackage(pkg);
// rewritten: ['/ppt/slides/slide1.xml']  -- everything else streamed
```

`exportPackage` throws rather than returning bytes that would make PowerPoint
show a repair prompt.

## Why the writer is not allowed to understand the document

A writer that emits a `.pptx` from a model can only emit the features that model
has. Every feature it has not been taught is a feature the file loses on save —
which is why a deck that goes through most browser editors comes out without its
charts, its SmartArt, its animations and its macros. Not because anyone decided
to drop them, but because the export path had no way to write them and no way to
notice.

Preservation cannot be a feature with a coverage percentage. The percentage is
never a hundred, and the missing part is always somebody's.

So it is the default state instead: parts are opaque bytes, and only the ones
something explicitly replaced are re-encoded.

## What happens on an export, and in what order

| Step               | Why it is there                                                  |
| ------------------ | ---------------------------------------------------------------- |
| `prepare` hooks    | Save-time work owned by other packages — font embedding, autofit |
| Media collection   | Runs after the hooks, because a hook is what orphans things      |
| `PartStore.write`  | Untouched parts stream through still compressed                  |
| Preservation check | The emitted archive against the source archive, entry by entry   |
| `assertValid`      | The 29 rules; refuses to return the bytes if we broke one        |

The two checks come **after** the write, which is the one ordering that looks
backwards. It is right because a check that runs first is checking an intention,
and what a user opens is an archive. Bytes are produced, then judged, then either
returned or thrown away — never returned unjudged.

## Media garbage collection

Off in effect unless something was actually orphaned, and deliberately timid
about what it will touch.

**The sweep set is `/ppt/media/`.** Not layouts, not masters, not
`ppt/embeddings/*.bin`, whatever the relationship graph says about them.
PowerPoint keeps slide layouts no slide uses — that is how the layout picker has
anything to offer — and a package where an unused layout has been tidied away is
one where "change layout" has quietly lost options. The layout is kilobytes; the
media is the megabytes.

**The walk is rooted at the package, not at the slide list.** The obvious
implementation asks "which images do the slides use?" and it is wrong in a way
that shows up weeks later: a layout no slide currently uses still has its
background picture, and switching a slide onto it has to still work.

**By default it only collects what this session orphaned**, by differencing
reachability against the baseline. A deck can arrive with an orphaned image, and
collecting it would mean that merely opening and saving a file changes it.

**It cannot break a relationship.** A part is collected only when nothing in the
package resolves to it, so there is no edge left to dangle — and two slides
sharing one image, one of them deleted, cannot lose it, because the surviving
relationship still marks it live. The corollary is that the sweep is a safety
net and not a delete: removing the picture's `r:embed` is part of deleting the
picture, and that is the command layer's job.

```ts
// Everything unreferenced, however it got that way. Refuses if the graph
// could not be walked completely — an unreadable edge is an invisible edge.
collectGarbage({ store, policy: 'every-orphan' });
```

## Prepare hooks

Several later phases owe the file something at save time, and none of it is this
package's subject: embedded fonts are six artifacts that must all be present or
PowerPoint reports a problem; autofit has to write back `@fontScale`; the
`p14:media` / `a:videoFile` relationship pair is a duplicate on the way out.

```ts
exportPackage({ ...pkg, prepare: [embedFonts, commitAutofit] });
```

Hooks run in the order given, before collection and before validation. Before
collection because a hook is exactly the thing that changes what is referenced;
before validation because a hook writes markup, and newly synthesized markup is
the riskiest kind there is. One that throws stops the export — five of the six
font artifacts written is a package PowerPoint refuses with no way to find out
which.

## The baseline is not optional in practice

`openPackage` hands back the package and an untouched twin. Leaving the twin out
costs three things: the three preservation rules are skipped rather than passed,
the media sweep collects nothing because there is no way to tell what this
session orphaned, and every fatal finding counts as one we introduced — so the
firewall gets _stricter_, not quieter.

## Is it still the same deck?

`comparePackages` answers it, and `roundTripPackage` does the whole thing —
open, export, compare — in one call. `pptx-studio roundtrip` is a thin front end
over it.

```ts
const { ok, comparison } = roundTripPackage(bytes);
// comparison.differences: [] · counts: { xml, binary, relationships, same }
```

**Not by comparing bytes.** Entry order, deflate level, timestamps and attribute
order differ legitimately between two archives holding one document — nine of
the corpus's PowerPoint-authored decks come out exactly 1832 bytes smaller with
every stored byte of every entry identical, because Office writes a `0xA220`
growth-hint extra field on five entries and we do not.

`assertPreserved` above and this are not competing. That one asks whether the
entries we promised not to touch came out identical, which is the right question
for a no-op export and the only question it can answer. This one asks whether a
package we _edited_, or one somebody else wrote, is the same document.

Three comparisons, one per kind of content:

- **XML parts** by canonical form — `canonicalXml` in `@pptx-studio/xml`, which
  sorts attributes and normalises escaping but keeps prefixes, unused namespace
  declarations and every space, because in OOXML those are meaning.
- **Relationship parts** as a graph with the ids treated as opaque labels,
  matched by what they resolve to. The mapping is then applied to the referring
  markup, so a deck PowerPoint re-saved — renumbering every `rId` — compares
  equal, while two `r:embed` values that swapped targets do not.
- **Everything else** by SHA-256. Not `crc32`: a 32-bit code collides within
  about 2^16 inputs, which is fewer parts than the corpus has.

`[Content_Types].xml` is compared as neither. It is a map with several spellings,
so what is compared is the resolved content type of every part.

## Which of those differences is the one that matters?

`bisectPackages` answers it, and `pptx-studio bisect` is the front end.
`comparePackages` above says _that_ two packages differ and where; this narrows a
delta of hundreds down to the changes that are actually load-bearing.

```ts
const { minimal, bytes, runs } = bisectPackages(original, broken, { oracle });
// minimal: [{ entry: '[Content_Types].xml', kind: 'children', where: '/Types/Default[3]', … }]
// bytes:   the smallest package that still fails — ready to open, or to keep
```

It is Zeller's `ddmin` over the **set of changes** between the two packages,
applied level by level down the tree (Misherghi and Su's HDD). Isolation rather
than simplification, and the difference is not academic: because a configuration
is a subset of changes applied to the original, and the splice unit is a whole
node replaced by a whole node, **every candidate is well-formed XML by
construction** — no oracle run is ever spent on markup no editor could produce.

The delta is taken over **ZIP entries, not parts**, because `[Content_Types].xml`
is not a part, and a missing `<Default Extension="fntdata"/>` in it is the
canonical repair prompt. Anything that parses as XML is descended into; anything
else is one atom.

The oracle is yours: `(bytes) => 'fails' | 'passes' | 'unresolved'`. The third
value is Zeller's too and it earns its place — a timed-out PowerPoint read as
"passes" would let the reducer discard the guilty change and blame an innocent
one.

## What is not here

The CI round-trip gate and the badge are 1.6. And no bisection has yet been
driven by a failure nobody planted — which is the good news about 1.3 and 1.4,
and also the reason this is untested against the case it was built for.
