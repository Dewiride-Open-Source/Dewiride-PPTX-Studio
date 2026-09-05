# ADR 0006 — Schema order, Markup Compatibility, and invertible edits

- **Status:** Accepted
- **Date:** 2026-08-27
- **Sub-phase:** 0.6

## Context

Sub-phase 0.5 could turn a tree back into bytes. This one is about changing it first, and the plan
gives the sub-phase four parts and one exit criterion:

> `schema-order.gen.ts` codegen from the ECMA-376 Transitional XSDs; `insertInOrder()` as the _only_
> sanctioned way to add an OOXML child; an MCE walker resolving `mc:Choice/@Requires` prefixes to
> URIs; `extLst` as an ordered opaque list keyed by `@uri`, never rebuilt; `XmlEdit` ops each
> constructed with its exact inverse. _Verify:_ inserting a child at every position in a synthetic
> sequence lands schema-correct; **every edit's inverse restores byte identity**.

## Where the table comes from, and what is not in it

The schemas are not in this repository and are not vendored into it. They are a free download from
Ecma International — ECMA-376 Part 4, 5th edition, December 2016, which carries
`OfficeOpenXML-XMLSchema-Transitional.zip` — and `tools/schema-codegen/generate.ts` reads them from
a directory named on the command line. Only the generated table is committed, together with the
SHA-256 of each input, so "which schemas is this from" has an answer that does not depend on
anyone's memory.

Four of the twenty-six schemas are read: `pml.xsd`, `dml-main.xsd`, `dml-picture.xsd`,
`dml-lockedCanvas.xsd`. **Not `dml-chart`, not `dml-diagram`, not `sml`, not `wml`** — and that is a
design rule rather than an omission. We never rewrite a chart, a diagram or an embedding; those
parts are copied byte-for-byte. An order table for them would be code shipped to every browser to
serve an operation the plan forbids.

The consequence is visible in the corpus and worth stating plainly: 6072 elements that have element
children have no entry in the table. They are, in order of frequency, the `.rels` root (1296), then
SmartArt `dgm:`, chart `c:` and chart-style `cs:` elements. Inserting into any of them raises
`ERR_SCHEMA_ORDER`, which is the correct answer to a question we have promised not to ask.

The reader in `tools/schema-codegen/read-xsd.ts` is a hundred lines and does not use
`@pptx-studio/xml`, because that would make the codegen depend on the build output of the package it
generates source for. It is safe to hand-roll only because it **refuses everything outside a
measured profile**: across all twenty-six schemas the constructs used are `xsd:sequence`,
`xsd:choice`, `xsd:group`, `xsd:element` and `xsd:any`, with no `xsd:all`, no `complexContent`, no
`simpleContent`, no `substitutionGroup`, no `abstract`, and no type derivation of any kind.

## A rank, not a list — and the reason is z-order

The obvious model for an `xsd:sequence` is an ordered list of names, and it is wrong.

`p:spTree` admits `sp`, `grpSp`, `graphicFrame`, `cxnSp`, `pic` and `contentPart` under a single
`maxOccurs="unbounded"` choice, and their freedom to interleave **is** the z-order of the slide. A
strict list would make inserting a picture reorder every shape on it.

So every child name gets an integer rank, children appear in non-decreasing rank order, and names
sharing a rank may appear in any order. Three rules produce it:

1. a `sequence` advances the rank once per particle;
2. a `choice` starts every branch at the same rank and takes the highest any branch reached, because
   at most one branch is instantiated;
3. **a compositor that repeats freezes everything beneath it to one rank**, because a second
   repetition can put a first-particle element after a last-particle one.

Rule 3 is the one that earns its keep. `a:p` is `pPr?`, then `EG_TextRun` with
`maxOccurs="unbounded"`, then `endParaRPr?` — so `a:r`, `a:br` and `a:fld` share a rank and
interleave, `a:pPr` stays before them, and `a:endParaRPr` stays last. That is the plan's appendix
rule, arrived at mechanically rather than transcribed.

**Zero rank collisions** across all 382 complex types: no walk ever assigned one element name two
different ranks, so the three rules are consistent with the whole of PresentationML and DrawingML.

`xsd:any` is not a rank. Three types have it — `p:CT_Extension`, `a:CT_OfficeArtExtension`,
`a:CT_GraphicalObjectData` — and all three are places the plan forbids rebuilding: inside an `extLst`
entry, and inside a `graphicFrame`'s payload. They get no table, so insertion is refused rather than
guessed at.

## Keys are namespaces; prefixes are only spelling

`p:` and `a:` are conventions. A producer may bind PresentationML to any prefix, and 0.4 already
found one prefix bound to two different URIs in one corpus. So the table is keyed `{namespace}local`
and the prefix is resolved through the tree at lookup time.

## Three names the parent alone cannot place

The table is keyed by parent element name, because that is what a caller holding an `XElement` has.
The schema is keyed by complex type, and the two are not the same map. Almost always the difference
is invisible — `a:ext` is `CT_PositiveSize2D` under `a:xfrm` and `CT_OfficeArtExtension` under
`a:extLst`, and only the second has element children at all. Measured, exactly three names remain:

| name     | under        | children                  |
| -------- | ------------ | ------------------------- |
| `a:xfrm` | `p:spPr`     | `a:off`, `a:ext`          |
| `a:xfrm` | `p:grpSpPr`  | plus `a:chOff`, `a:chExt` |
| `a:path` | `a:pathLst`  | `a:moveTo`, `a:lnTo`, …   |
| `a:path` | `a:gradFill` | `a:fillToRect`            |
| `p:to`   | `p:animClr`  | a colour                  |
| `p:to`   | `p:set`      | an animation variant      |

Dropping ambiguous names would have been the cheap answer and would have cost us `a:xfrm`, which is
where every drag writes. So those eleven pairs get one more level of key, and **the codegen asserts
that one more level is enough** — it reports every pair the grandparent still fails to separate, and
there are none.

## Byte-identical undo needs the dirty flags back, and that is sound

This is the substance of the exit criterion, and it is sharper than it first reads.

`dirty === false` is not a hint. It is an invariant: _`source.slice(start, end)` **is** this node's
serialization._ An inverse that restores the value but leaves the flag set produces the same
_document_ and not the same _bytes_: the subtree is rebuilt rather than sliced, so
`<a:off x="0" y="0" />` may come back without its space, `&#62;` may come back as `>`, and a CRLF
between two elements comes back as LF — which 0.5 measured on **656 of 2834 real parts**. Undo would
silently rewrite a quarter of the corpus.

So `markDirty` takes an optional journal of the nodes it actually changed, every edit records that
journal, and its inverse clears exactly those flags — never a flag it did not set. Clearing is sound
precisely when the invariant has been restored, and an exact inverse is what "restored" means.

Three consequences shape the operations:

- **Removal keeps the object.** `removeChild`'s inverse holds the `XNode` itself, not a description
  of one, because a structural copy would carry no source span and would have to be rebuilt.
- **Removing an attribute does not dirty the attribute**, only its element, so restoring it restores
  its original spacing and quote character.
- **Removal is addressed by node; insertion by index.** An index is a position at the moment it is
  applied, so a command planning two edits against one parent would have the second silently shifted
  by the first. Removal has a stable address available and uses it; insertion does not, because the
  position is the whole content of the edit.

## What the corpus says

37 decks, 2834 parts, 194 148 elements — the same corpus as 0.4 and 0.5.

|                                            |                                              |
| ------------------------------------------ | -------------------------------------------- |
| **children in schema order**               | **194 148 / 194 148 — 0 out of order**       |
| children the table could rank              | 174 242 / 174 244                            |
| …the two it could not                      | one `mc:AlternateContent`, one `p14:gallery` |
| parents with a content model               | 130 737                                      |
| parents with element children and no model | 6072 — `.rels`, `dgm:`, `c:`, `cs:`          |
| **edits applied, then undone**             | **194 244**                                  |
| …of which schema-ordered insertions        | 31 330 (2117 more refused)                   |
| **undo → byte-identical**                  | **2834 / 2834 — 100%**                       |
| edited document → reads back as itself     | 2834 / 2834                                  |
| undone document → reads back as itself     | 2834 / 2834                                  |
| dirty flags left behind after undo         | 0                                            |
| markup-compatibility problems              | 0                                            |
| `mc:AlternateContent` elements             | 4, all selecting `mc:Fallback`               |

The first row is the one that could have falsified this work rather than merely exercising it. The
table was generated from a standard; 194 148 elements written by PowerPoint are 194 148 chances for
the codegen to have mis-read an `xsd:group` ref, collapsed a choice it should not have, or got the
repeating-compositor rule backwards. None of them disagreed. Coverage is reported beside it because
a table that ranked nothing would also have scored zero.

The edit battery is mixed on purpose and applied as one batch, because a single edit is easy to
invert and the interesting failures are in the interaction — one edit's `markDirty` stopping at an
ancestor another edit already dirtied, and the undo having to unpick exactly its own share of that.
By kind: 105 651 attribute values set or added, 19 537 attributes removed, 31 330 elements inserted
in schema order, 34 917 children removed, 2809 text nodes rewritten. All 194 244 applied and undone
in 743 ms.

## What the instruments found, which is the part worth reading

Three instruments, and two of them found real defects in work that was already passing its tests.

**Mutation testing — 31 deliberate defects, 31 killed.** It also found that the byte comparison was
not enough on its own. An inverse that restores the flags but _not_ the values serializes correctly,
because the restored flags send every node back to its source slice and the right bytes hide the
wrong value behind them — leaving a tree that reads back as something the file does not say. The
mutant survived until `checkRoundTrip` was added _after_ the undo as well as after the edit.

**Fuzzing — the 40 000-case sweep found two defects the corpus could not.**

The first is the more serious. **The mixed-content guard could refuse an inverse.** Removing an
element from a parent that already held text — which a real file never does and a mangled one can —
and then putting it back looked like "creating mixed content", so undo was refused and the document
could not be returned to what it was. A guard that can refuse an inverse is worse than no guard,
because the property it breaks is the one everything above this layer depends on. The distinction
that fixes it is worth keeping: every other precondition in `edit.ts` is _structural_ — an index in
range, a node not already attached, no attribute of that name — and an exact inverse satisfies all of
them by construction. Only this one is _semantic_, about the shape of the surrounding document, and
semantic checks apply to forward edits only.

The second: **`applyEdits` was not atomic.** Some refusals are only knowable at apply time, because
they depend on the state earlier edits in the same batch left behind. A caller whose fifth edit was
refused held a half-applied document and no way back, since the inverses of the first four are only
returned on success. `applyEdits` now rolls back with those inverses and rethrows, so a refused batch
leaves no trace — dirty flags included. 26 of the sweep's cases exercise that path.

The sweep otherwise: 2855 parses and 37 145 typed errors, **disjoint**, so nothing throws after a
successful parse; zero untyped throws, coverage violations or fidelity failures; 2829 documents
genuinely changed by the edit battery and 48 322 edits undone exactly. Slowest single case 197 ms,
which is the 2000-fold repeated open tag being edited rather than merely parsed.

## Markup Compatibility

`Requires` is the one MCE attribute written **unprefixed**, and therefore, by _Namespaces in XML_
§6.2, in no namespace at all. `mc:Ignorable`, `mc:MustUnderstand` and `mc:ProcessContent` sit on
arbitrary elements and carry the prefix; `Requires` is an attribute of `mc:Choice` itself and does
not. Looking for `mc:Requires` finds nothing in any real file, and a walker that then treats the
Choice as requiring nothing selects the **first** branch of every `mc:AlternateContent` in the
document. There is a test for it and a mutant that dies on it.

The selection checks the whole element before returning, rather than stopping at the winning branch.
Short-circuiting would be faster on an element with three children, and it would mean a document
whose _second_ `mc:Choice` names an unbound prefix reads perfectly on a machine where the first
Choice is supported and is an error on one where it is not.

An unresolvable prefix fails loudly rather than counting as unsupported. Treating it as unsupported
would silently select a different branch of the document, which is the corruption this whole package
exists to prevent, arrived at from the other direction.

`effectiveChildren` **keeps** an element in a namespace nobody declared ignorable. MCE says the
producer should have annotated it; a consumer that deletes what a producer forgot loses content, and
for an editor whose thesis is preservation that is the worse error.

Nothing in `mce.ts` mutates. The plan's rule is _never rewrite an `mc:AlternateContent` branch you
don't understand_, and the way to keep it is to have no code there that could.

## Adjacent text nodes, which ADR 0005 deferred to here

Two text children in a row serialize to one run and reparse as one node. A parse never produces that
shape; **removing an element from between two whitespace runs does**, and that removal is correct.
`checkRoundTrip` now compares adjacent text runs as one entry, which is not a concession: XML 1.0 has
character data, not a list of text nodes, and the split between two of them is an artefact of how the
tree was built. CDATA is deliberately _not_ coalesced into a neighbouring run, so a CDATA section
that came back as ordinary text is still a difference.

## Decisions taken quietly, recorded so they can be argued with

- **No whitespace is synthesized.** An element inserted into a pretty-printed part lands without
  matching indentation. Measured: 11 of 372 slide parts in the corpus have any inter-element
  whitespace at all, so there is nothing to match in 97% of the parts we edit, and inventing text
  nodes we never read is the wrong default for this package.
- **An emptied `extLst` is kept.** `<p:extLst/>` is schema-legal and real files contain them;
  removing it would be a rewrite nobody asked for, in the part of the tree where unrequested rewrites
  are most expensive.
- **Extensions are appended, never sorted.** GUIDs have no meaningful order.
- **Values are not validated at edit time.** The serializer already refuses unpaired surrogates,
  characters outside `Char`, and the terminators of constructs that have no escape. Two answers to
  one question in one tree is worse than one answer in the wrong place.

## Verification

`pnpm check` green: layering, format, lint, typecheck, build, publint/attw, and **500 tests in real
Chromium** across 23 files, up from 417 across 20.

1. **The corpus gate**, above. Also the plan's literal wording, taken literally: every one of the 32
   subsets of `p:sp`'s five children, crossed with every child missing from it — 80 insertions, each
   of which must leave the result a subsequence of the schema's order.
2. **Fuzzing**, above.
3. **Mutation testing.** 31 defects introduced one at a time: each half of the dirty journal, the
   ordering of the restore against the mutation, the object-preserving removal, the `<=` that puts a
   new shape on top of the z-order rather than under it, the pinning of unrankable siblings, the
   grandparent lookup, namespace-keying, the rank separators, `Requires`, `every` versus `some`,
   Choice versus Fallback, the ancestor walk for `mc:Ignorable`, `ProcessContent` as qnames, the
   rollback, and the forward-only guard. **31 of 31 killed.**

## What is deferred, and to where

- **The codegen is not runnable in CI**, because the schemas are not in the tree. A contributor who
  wants to regenerate the table downloads the Ecma zip; the SHA-256s in the generated file say
  whether they got the same one. This is the same shape of gap as the corpus gate and it is smaller:
  the input is a fixed published artefact from a standard whose last edition is 2016.
- **The corpus gate is still not automated in the repo.** Third ADR in a row to say so. Sub-phase
  1.1's checked-in, licence-audited corpus is where it gets fixed.
- **Declaration quoting**, unchanged from 0.5: a rebuilt `<?xml … ?>` is written with double quotes.
  Nothing edits a declaration.
- **`insertInOrder` decides its position at plan time.** Two insertions planned against one parent
  before either is applied can therefore both compute the same index. Removal was made
  node-addressed for exactly this reason; insertion cannot be. Sub-phase 5.1's command bus is where
  a plan-then-apply discipline belongs, and it will need to know this.
- **The `.rels` schema is not read**, so `Relationships` has no content model. It does not need one:
  it admits a single child name, so any position is in order.
