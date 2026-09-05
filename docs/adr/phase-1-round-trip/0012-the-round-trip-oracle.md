# 0012 — The round-trip oracle, and why it is not a byte comparison

Date: 2026-08-29
Status: **accepted** — 51/51 decks round-trip; 1419 parts compared as documents rather than as bytes
Sub-phase: 1.4 — `cli roundtrip`

---

## Context

The plan's 1.4 is one sentence of design and one of prohibition:

> Canonical-XML per part, relationship-graph isomorphism with rIds treated as
> opaque labels, SHA-256 on media. **Never raw ZIP byte equality** — entry
> order, deflate level, DOS timestamps and attribute order all legitimately
> differ, so a byte-equality test is red on day one and disabled on day two.

Sub-phase 1.3 already compares bytes, and does it strictly: `assertPreserved`
takes the _stored_ — still compressed — bytes of every entry we did not edit and
requires them to be identical, and all 1470 entries of all 51 decks pass. That
is a good check and it answers a narrow question: on a no-op export, was
anything re-serialised behind our backs?

It cannot answer the question this project actually runs on. The moment an edit
lands, the deck legitimately changes, and "is it still the same deck" stops
being a question about archives. It is also the wrong question to ask of a file
somebody else wrote: PowerPoint renumbers relationship ids on save, so a
comparison that cannot see past that is useless the first time it is pointed at
the output of the program we are trying to be compatible with.

### The prohibition is not theoretical — it fails on day one, measurably

Running the no-op export over the corpus and comparing archive sizes:

**Ten of the fifty-one decks come out a different size**, 17 200 bytes smaller
in total, while every entry's stored bytes are identical. Nine of them are the
PowerPoint-authored `b01`–`b09` and each loses **exactly 1832 bytes**, whatever
its entry count — 37 entries for `b01-blank`, 57 for `b02-layouts`, the same
1832 either way.

The 1832 is not a coincidence and it is not ours. Each of those decks carries a
ZIP local-header extra field with header id **`0xA220`** — Microsoft's Open
Packaging Growth Hint — on exactly five entries: 520 bytes on
`[Content_Types].xml` and `_rels/.rels`, 264 on `docProps/core.xml`,
`docProps/app.xml` and `ppt/_rels/presentation.xml.rels`. It is padding
reserved so that those parts can grow in place on an incremental save. It says
nothing about the document, our writer does not emit it, and dropping it costs
a resave optimisation and nothing else.

So a byte-equality gate would have been red on ten of fifty-one decks before a
single line of the renderer existed, over padding. That is precisely the test
that gets disabled on day two.

---

## Decisions

### Three comparisons, because a package has three kinds of content

**XML parts compare by canonical form.** A new module, `canonicalXml` in
`@pptx-studio/xml`, which the plan's layout already anticipated.

**Relationship parts compare as a graph, with the ids as opaque labels.** Never
as XML: two `.rels` that disagree only about whether the theme is `rId1` or
`rId7` describe the same package.

**Everything else compares by SHA-256.** The bytes are in hand on both sides, so
the digest is not needed to decide equality — it is needed to _report_, and to
give `cli bisect` a way to name a part in 1.5. `crc32`, which
`@pptx-studio/opc` already has for the archive, is a 32-bit error-detecting
code: a birthday collision is expected within about 2^16 inputs, which is fewer
parts than the corpus has, and it is not evidence that two files are the same.

`[Content_Types].xml` is compared as _neither_. It is a map, and the same map has
several spellings — a `Default` for an extension and an `Override` on every part
with that extension mean the identical thing. What is compared instead is the
resolved content type of every part, which is the only thing the map is for.

### `canonicalXml` is not W3C C14N, and two of its rules are actively wrong here

The escaping rules are C14N's and the attribute ordering is C14N's. Three things
depart, and two of them would be data loss:

**Namespace declarations are kept, including unused ones.** C14N prunes a
declaration no name in scope refers to. PowerPoint writes `xmlns:a14="…"` on a
slide root alongside `mc:Ignorable="a14"` with no `a14:` element anywhere in the
part — the declaration exists so that the `mc:Ignorable` token resolves. Prune
it and a well-formed part becomes one naming an undeclared prefix.

**Prefixes are kept as written, never renamed.** The same reason one layer up,
and the reason `@pptx-studio/xml` exists at all rather than a call to
`XMLSerializer`. A canonical form free to rename `a14` to `ns3` would report two
documents as identical when one of them had quietly become invalid — undoing the
project's second architectural bet at the far end of the pipeline.

**The XML declaration is kept**, normalised to version and `standalone`. C14N
drops it; Office writes one on every part, and losing one is a real change. The
encoding is dropped, because this is a comparison of characters: a part that
arrived UTF-16 and left UTF-8 is the same part.

Whitespace is not normalised anywhere. Collapsing the whitespace _between_
elements is tempting and would hide the exact failure this exists to catch — a
writer that reformats a part on the way out. Inside `a:t` it is content.

### Ids are matched by what they point at, and the mapping is applied to the markup

Within one `.rels`, relationships are grouped by `(type, target-mode, resolved
target)`; groups of the same size pair up in document order, and that pairing is
the label mapping. It is then applied to the _referring_ part's markup before
canonicalising, so `r:embed="rId3"` on one side and `r:embed="rId7"` on the other
are the same reference exactly when rId3 and rId7 resolve to the same part.

The relabelling is keyed on the **namespace** of the attribute rather than on a
list of names, and the corpus settles that choice rather than taste. Measured:
582 attributes across the fifty-one decks are in the relationships namespace,
under eight distinct local names, and **every one of their values matches
`rId\d+` or is empty** — so the namespace rule is exact here, not approximate.

The plan also names eight — `r:id`, `r:embed`, `r:link`, `r:pict`, `r:dm`,
`r:lo`, `r:qs`, `r:cs` — and they are not the same eight. The corpus has
`r:blip` on 46 attributes and no `r:pict` at all. A list of names would already
have been one entry short, on a corpus we built ourselves, before meeting a
single deck from outside. That measurement is now an assertion in
`tools/corpus/roundtrip/roundtrip.test.ts`.

The mapping is applied to the original side **only**. Applying it to both maps
the ids twice and cancels out, which is a bug that passes every test where the
two packages agree and fails silently on the one case the machinery exists for:
two ids that swapped what they point at while both documents still say `rId1`.
That is one deck showing image A and the other showing image B. It was caught by
the test written for exactly that case, and it is the reason that test exists
rather than a simpler one.

### The exit code says one thing; the message says three

`pptx-studio roundtrip` exits 1 on a difference _and_ on a refusal _and_ on a
file that will not open — CI wants one bit. All three need completely different
work next, so stderr distinguishes them by hand.

---

## What building it found

### Every package PowerPoint refuses round-trips cleanly, and that is correct

All seventeen packages in `corpus/reject` — the ones PowerPoint will not open,
where `validate` exits 1 on every one — exit **0** under `roundtrip`.

It looks like a bug for about ten seconds. Round-tripping a broken package is
supposed to hand back the same broken package, defect intact; a `roundtrip` that
failed here would be claiming it had _changed_ the file, which is the one thing
it must not do. Asserted, so that nobody later "fixes" it.

### A lower-cased lookup that could never match

Consolidating three private copies of "is this content type XML" into one export
on `@pptx-studio/opc` turned a latent inconsistency into a real bug and then
into a caught one. The helper lower-cases its input, as RFC 2045 requires; the
set it consults had been copied verbatim from Office's spelling,
`…officedocument.vmlDrawing`, capital D and all. Nothing matched.

The failure was invisible in the pass/fail sense — a VML part compared as opaque
bytes still compares correctly — and visible in the count, which is why the
corpus test pins the split three ways rather than the total. VML is where an OLE
object's on-slide preview lives and where `@spid` resolves; comparing it as a
blob would have meant a difference in it could never be located, only detected.

The narrower private copy in `@pptx-studio/validate` was **not** changed. It
would widen the set of parts the firewall parses for its order rules, which is a
change to 1.2's behaviour and does not belong in a sub-phase about comparison.
Recorded here rather than done quietly.

### The firewall stopped this sub-phase's own test

The "re-compress every part" test replaced each part with its own bytes to give
`deflateLevel` something to act on. `V027` refused the export: `ppt/embeddings`
holds OLE2 compound files and rewriting one is a guaranteed repair prompt,
whatever the replacement is — including a byte-identical one. The test now skips
that directory. The rule was right and the test was wrong.

### A no-op export cannot demonstrate the thing this sub-phase is about

The first version of the corpus test set `normalizeEntryOrder` and
`deflateLevel` and asserted the archive changed. It did not: with nothing
re-serialised there is nothing to deflate, and the entry order was already
canonical. Worth stating rather than working around — the knobs only bite once
something is rewritten, so the test rewrites everything.

---

## Consequences

`pnpm check` green. 52 test files, 1366 tests (up from 47 and 1269).

- **51/51 decks round-trip**, with `differences` empty on every one.
- **1419 parts compared**: 840 as canonical XML, 536 as relationship graphs, 43
  as SHA-256. Pinned three ways rather than as a total, because a deck that
  silently lost half its parts would still report zero differences.
- **No relationship id is renamed** by anything this writer produces. The
  mechanism exists for files PowerPoint saved; asserting it stays unused is what
  makes the deliberate-renumbering test mean something.
- **A deck whose every relationship id is renumbered by +1000**, in the `.rels`
  and in the markup, compares equal — and one whose two `r:embed` values are
  swapped does not.
- **Every part of every deck re-compressed at a different level, in a normalised
  entry order**, compares equal.
- **Every `<a:off x y/>` in every slide, layout and master written with its
  attributes swapped** compares equal — 40+ parts across the corpus.
- **582 attributes in the relationships namespace, eight local names, every
  value an `rId`** — the premise the relabelling rests on, pinned rather than
  assumed.

Negatives, because a comparator that never disagrees is not a comparator: one
character of a shape name, one byte of an image, one relationship removed, a
content type changed, a part added, a part gone, and a part that will not parse.

### What is not done

- **No exported file has been opened in PowerPoint.** `roundtrip --write` exists
  so it can be done by hand; scripting it is 1.5's COM loop. Everything above is
  a claim about our own reading of the file.
- **The comparator has never been pointed at a package PowerPoint wrote from
  one of ours**, which is the case the id relabelling was built for. The
  renumbering in the corpus test is done by hand, in the shape PowerPoint is
  documented to use.
- **`cli bisect` is 1.5 and the CI gate is 1.6.** This says _that_ two packages
  differ and where; narrowing a difference to the smallest change that causes a
  repair prompt is the next sub-phase.
- **`@pptx-studio/opc` gained a SHA-256** and is now the home of two digests.
  It is there because it is the only layer-0 package that deals in bytes, and
  because `crypto.subtle` is both asynchronous and gated on a secure context —
  absent on a plain `http:` origin, which a client-side library does not get to
  choose.
