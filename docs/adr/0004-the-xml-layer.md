# ADR 0004 — The XML layer

- **Status:** Accepted
- **Date:** 2026-08-26
- **Sub-phase:** 0.4

## Context

Sub-phase 0.3 turned an archive into a package of parts whose content was opaque bytes. This one
opens the bytes.

The constraint is unusual enough to be worth restating, because it is what disqualifies every
general-purpose XML library: we must be able to hand a part back **exactly as we received it**,
except for the one node someone edited. Not equivalent XML. The same characters.

Two sources of ground truth, and they answered different questions.

**The corpus says what is normal.** All 2834 XML parts across the 37 packages, scanned at the byte
level:

|                                         |                                                              |
| --------------------------------------- | ------------------------------------------------------------ |
| XML parts                               | 2834, from two producers                                     |
| Distinct XML declarations               | **3** — `utf-8`, `UTF-8 standalone`, `utf-8 standalone`      |
| What follows the declaration            | **3 variants** — nothing (2161), `\r\n` (656), `\n` (17)     |
| Byte order marks                        | **1037 of 2834** carry one                                   |
| Attribute quoting                       | 170 019 double, 0 single                                     |
| Self-closing tags                       | 98 777, of which **70 822 are written `<x />`** with a space |
| Long-form empty elements (`<x></x>`)    | 33, all in `docProps`                                        |
| Entity references, whole corpus         | **39** — 26 `&quot;`, 5 `&lt;`, 5 `&gt;`, 3 `&amp;`          |
| Numeric character references            | **0**                                                        |
| CDATA sections, comments, PIs, DOCTYPEs | **0**                                                        |
| `xml:space`                             | **0**                                                        |
| Mixed content                           | **0**                                                        |
| Deepest nesting                         | 23, in a `p:timing` tree                                     |
| Most attributes on one element          | 15, on an `a:defRPr`                                         |
| Largest single part                     | 69 KB                                                        |

Two things stand out. First, the declaration, the BOM and the trailing newline **vary within a single
package** — `Pitchbook.potx` has parts with no BOM and `\r\n`, and parts with a BOM and `\n`. A
serializer that emits a canonical declaration loses byte fidelity on most of the corpus. Second, the
corpus is a weak test: it contains none of the constructs a tokenizer is most likely to get wrong.

**PowerPoint says what is legal.** So we built 62 variants of a real slide part, one lexical
construct each, and opened every one with COM automation against the installed PowerPoint 365.
**47 of the first 50 opened.** The failures are the interesting half:

| Construct                                                             | PowerPoint                     |
| --------------------------------------------------------------------- | ------------------------------ |
| `<!DOCTYPE p:sld>`                                                    | **refused** `0x80070570`       |
| `<a:off x="0" y="0" x="0"/>` — duplicate attribute                    | **refused** `0x80070570`       |
| literal U+000B in an `<a:t>`                                          | **refused** `0x80070570`       |
| `&nbsp;` — undeclared entity                                          | **refused** `0x80070570`       |
| a lone `0xE9` byte under `encoding="UTF-8"`                           | **refused** `0x80070570`       |
| truncated UTF-8, overlong UTF-8                                       | **refused** `0x80070570`       |
| UTF-16 bytes declaring `encoding="UTF-8"`                             | **refused** `0x80070570`       |
| UTF-8 bytes declaring `encoding="UTF-16"`                             | **refused** `0x80070570`       |
| comments, before / inside / after the root                            | opens                          |
| processing instructions                                               | opens                          |
| `<![CDATA[…]]>` inside `<a:t>`                                        | opens, read as text            |
| `&#72;`, `&#x48;`, `&#x1F600;`                                        | opens — the last renders as 😀 |
| single-quoted attributes, single-quoted declaration                   | opens                          |
| **no XML declaration at all**                                         | opens                          |
| `<p:spPr></p:spPr>` long form                                         | opens                          |
| `<a:off x="0" />`, newlines inside a tag, spaces around `=`           | opens                          |
| a fully pretty-printed slide part                                     | opens                          |
| `xml:space="preserve"`, and trailing whitespace **without** it        | opens, whitespace kept         |
| unescaped `>` in text and in an attribute value                       | opens                          |
| a BOM added to a part that had none                                   | opens                          |
| literal CRLF and TAB inside an attribute value                        | opens                          |
| `xmlns=""` mid-part; one prefix bound to two URIs in sibling subtrees | opens                          |
| `encoding="windows-1252"` with a real `é`                             | opens, reads `café`            |
| genuine UTF-16LE with a matching declaration                          | opens                          |
| nesting 5000 elements deep                                            | opens                          |

### What PowerPoint does to its own files

The most useful measurement of the lot. We took the variants PowerPoint accepted, opened each one,
saved it through PowerPoint's own writer, and diffed the slide part:

| Construct                        | Survives PowerPoint's resave?                          |
| -------------------------------- | ------------------------------------------------------ |
| comments                         | **discarded**                                          |
| processing instructions          | **discarded**                                          |
| `<![CDATA[x]]>`                  | **rewritten as plain text**                            |
| `&#72;`                          | **resolved to `H`**                                    |
| single-quoted attributes         | **rewritten as double**                                |
| `<p:spPr></p:spPr>`              | **collapsed to `<p:spPr/>`**                           |
| byte order mark                  | **stripped**                                           |
| `xml:space="preserve"`           | **removed** — though the whitespace it guarded is kept |
| an unused `xmlns:zz` declaration | **dropped**                                            |

PowerPoint is a lossy XML round-tripper. Every row above is something this package hands back
untouched, which means sub-phase 0.5's byte-identical gate is **strictly stronger than what
PowerPoint does to its own files**. It also sets the standard: if we ever canonicalize helpfully, we
have merely reinvented the thing we are trying to beat.

## The tokenizer works in string space, not byte space

**This deviates from the plan**, which specifies "pull tokenizer over `Uint8Array` emitting byte
offsets". The guarantee the plan wanted is preserved exactly; the representation is different, and
it is different because of two measurements.

**It is lossless.** For all 2834 XML parts, `encode(decode(bytes))` is byte-for-byte `bytes`. That is
not luck — valid UTF-8 and sequences of Unicode scalar values are in bijection, since UTF-8 admits no
overlong forms and no encoded surrogates and `TextEncoder` emits only the shortest form — but the
measurement confirms there is no part in our hands where the reasoning fails to apply.

**It is much faster.** Scanning an 8 MB part four ways:

|                                                 |           |
| ----------------------------------------------- | --------- |
| raw bytes, `TextDecoder` per token              | 289 ms    |
| raw bytes, hand-rolled ASCII fast path          | 162 ms    |
| **decode once, scan the string**                | **65 ms** |
| decode once, scan, intern names through a `Map` | 139 ms    |
| `TextDecoder` on its own                        | 7 ms      |

Decoding is essentially free, and hand-decoding UTF-8 per token costs 4.4× the thing it was meant to
avoid. The interning row is the surprise: with only 67 distinct names in 262 429 tags, interning
_doubled_ the cost. V8 makes `String.prototype.slice` an O(1) `SlicedString`; hashing that slice for
a `Map` lookup forces exactly the copy the slice avoided. A standard trick, measured, and rejected.

Byte offsets are not merely slower, they are a trap. Decode a part to a string, scan the string, and
report the offsets as byte offsets and they are silently wrong for every part containing non-ASCII —
which is 255 text nodes and 687 attribute values in our corpus, and every deck with CJK, curly
quotes, or a Wingdings bullet. String space removes the failure mode rather than defending against
it.

The cost is that `decodeXmlSource` is a wall: anything that is not UTF-8 is refused, by name. See
"Where we differ from PowerPoint" below.

> `TextDecoder`'s option is named backwards and it matters. `ignoreBOM: false` — the **default** —
> means "consume the BOM and do not emit it", i.e. delete the first three bytes. That would corrupt
> **1037 of our 2834 parts**, invisibly, until an export was diffed against its input. We pass
> `ignoreBOM: true`, which means "treat it as an ordinary U+FEFF".

## A node's source and a node's value are different things

XML 1.0 mandates transformations of the input, and each one is a place where storing the transformed
form destroys byte fidelity:

- **§2.11 end-of-line.** `\r\n` and lone `\r` must be reported as `\n`. Observed in PowerPoint:
  writing `line1\r\nline2` into an `<a:t>` and reading the text back yields `line1\nline2`.
- **§3.3.3 attribute-value normalization.** A literal tab, line feed or carriage return becomes a
  space.
- **§4.6 reference expansion.**

So every node carries a span, and its `value` is derived. The subtlety worth writing down, because
the obvious implementation gets it backwards:

> **A literal whitespace character in an attribute value is normalized to a space. A character
> reference to the same character is not.** `name="a&#9;b"` and `name="a<TAB>b"` are two different
> values. An implementation that expands references first and then squashes whitespace collapses them
> into one — and it will pass every test written against a corpus that contains no character
> references at all, which ours does not.

The same asymmetry applies one tier up: `text()` normalizes literal line endings, but `&#xD;` stays a
carriage return.

## Coverage, not "parses without error", is the gate

The plan's stated verification for 0.4 is "tokenize every part of the corpus without error". That is
too weak to catch the failure mode that matters. **An off-by-one in a span raises no error at all.**
It tokenizes every part of every deck happily and then silently drops or duplicates a character the
first time 0.5 re-serializes.

So the gate is that the spans **tile the source exactly** — no gaps, no overlaps, starting at the BOM
boundary and ending at `source.length`. `checkSpanCoverage` asserts it over a token stream and
`checkTreeCoverage` over a tree, where an element's children must additionally tile the range between
its two tags. It is falsifiable and cheap enough to run on every part of the corpus and inside every test that
parses anything.

It was not, at first, _total_. See "What the review changed" below: the version that shipped tiled
the nodes but never looked inside a tag, which left the offsets that exist solely for 0.6 as the
only ones nothing verified.

## Namespaces are resolved by walking the tree, never from a table

The tokenizer emits `(prefix, local)` and resolves nothing. Both plausible shortcuts are wrong on
files we hold:

- **A package-wide prefix table is wrong.** `p14` is bound to `.../powerpoint/2010/main` in 46 parts
  of the corpus and to `.../powerpoint/2007/7/12/main` in 8 others.
- **A per-part table read off the root element is wrong.** 213 namespace declarations in the corpus
  sit on a non-root element. `ClassicPhotoAlbum.potx` — a template Microsoft ships — contains, nested
  inside `<p:extLst>`:

  ```xml
  <p14:discardImageEditData xmlns="" xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2007/7/12/main" val="0"/>
  ```

  A mid-document prefix binding beside a default-namespace _undeclaration_.

We also built a part binding one prefix to two different URIs in sibling subtrees and PowerPoint
opened it. So `resolvePrefix` walks up from the element. Two consequences fall out for free:
`xmlns=""` resolves to `''` ("in no namespace"), which is distinct from `undefined` ("nothing binds
this"); and `attributeNamespaceOf` returns `''` for an unprefixed attribute, because _Namespaces in
XML_ §6.2 says an unprefixed attribute does not pick up the default namespace — get that wrong and
every unprefixed attribute in a slide part appears to be in the PresentationML namespace.

`prefixMap()` returns `Map<prefix, Set<uri>>` rather than `Map<prefix, uri>` for the same reason, and
it is what 0.5 will compare across a round trip.

## DOCTYPE is refused, and PowerPoint agrees

Rejecting the construct outright removes XXE, parameter entities and billion-laughs entirely, because
there is no internal subset left to expand. The worry with a rule like that is always that it refuses
a real file. It does not: **PowerPoint refuses a DOCTYPE too**, and there are none in the corpus.

The rule also closes the entity question. With no DTD there is nowhere for an entity to have been
declared, so the vocabulary is the five predefined ones — and PowerPoint enforces exactly that, since
`&nbsp;` in an `<a:t>` is refused.

## Flat offsets, not nested span objects

The plan sketches `XNode { …, raw?: {start, end}, dirty }`. An 8 MB part holds 262 429 elements;
giving each a `raw`, an `openTag` and a `closeTag` object is three quarters of a million allocations
to express six integers. Elements carry `start`, `end`, `openTagEnd`, `closeTagStart`, `nameEnd` and
`trailingSpaceStart` directly, and `dirty` carries the meaning `raw === undefined` would have. The
information is identical.

Two related changes, both measured on the same part: the tokenizer's attribute objects are **adopted**
by the tree rather than copied into a parallel type, and duplicate-attribute detection is a linear
scan over the attributes collected so far rather than a `Set` — because a `Set` must be allocated once
per element to deduplicate a list whose longest instance in the whole corpus is 15 entries. Together
these took full-tree parsing from 12.3 MB/s to 24.5 MB/s across the corpus.

`trailingSpaceStart` deserves its own line: without it a serializer emits `<a:off x="0"/>` where the
input said `<a:off x="0" />`, and that is 70 822 of the 98 777 self-closing tags in our corpus.

## The builder never recurses

PowerPoint opens a slide part nested **5000 elements deep**. A recursive-descent builder would meet a
well-formed file that PowerPoint accepts and fail on it — announcing the failure with a `RangeError`,
the one thing this package has promised never to throw. So the builder uses an explicit stack, and
`maxDepth` is a memory bound rather than a stack guard. It defaults to 1024, against a real-world
maximum of 23.

## Where we differ from PowerPoint, deliberately

Running all 62 variants through both, **58 verdicts agree exactly**. The four that do not are all on
the two axes we chose:

| Case                                      | PowerPoint | Us                         | Why                        |
| ----------------------------------------- | ---------- | -------------------------- | -------------------------- |
| `encoding="windows-1252"` with a real `é` | opens      | `ERR_UNSUPPORTED_ENCODING` | we read UTF-8 only         |
| genuine UTF-16LE                          | opens      | `ERR_UNSUPPORTED_ENCODING` | same                       |
| nesting 5000 deep                         | opens      | `ERR_LIMIT_EXCEEDED`       | `maxDepth` is configurable |
| UTF-8 bytes declaring `UTF-16`            | refused    | refused                    | closed after measuring     |

The last row was a genuine gap: we accepted it until this run showed PowerPoint does not. The rule
that resolves it is that a document contradicting itself is refused, while a document whose label and
bytes merely _could_ disagree is read — which is why `encoding="windows-1252"` over ASCII-only content
still parses and round-trips perfectly, and only fails once a byte actually needs the other codec.

The encoding gap is documented, typed and named rather than silent. Zero parts in the corpus are
affected; extending `decodeXmlSource` is additive when one turns up.

## A correction to the plan

The plan's must-not-break appendix lists `xml:space="preserve"` on any `a:t` with leading or trailing
whitespace. **That rule does not hold for PresentationML.** There is no `xml:space` anywhere in the
37-package corpus; 120 `<a:t>` elements carry edge whitespace without one; a slide part built with
trailing spaces and no `xml:space` round-trips through PowerPoint with the spaces intact; and
PowerPoint's own writer _removes_ an `xml:space` we add while keeping the whitespace. The rule is a
WordprocessingML convention that does not transfer. `xmlSpace()` therefore reads the attribute and
never synthesizes one.

## Verification

- **145 unit tests** in the package, run in real Chromium. Every test that tokenizes or parses goes
  through a helper that asserts span coverage, so a span bug anywhere fails everywhere. (The
  character-production, reference, encoding and runtime suites do neither, and do not.)
- **All 2834 XML parts of the 37-package corpus**: 314 814 tokens, 194 148 elements, 170 019
  attributes. Zero token gaps, zero tree gaps, zero slice mismatches, zero byte mismatches, zero
  undeclared prefixes.
- **62 hand-built lexical variants**, every one of which was also handed to PowerPoint. 58 verdicts
  identical; the four differences are the table above.
- **40 000 structured mutations** of real slide parts - truncation, injection, deletion,
  attribute floods, tag-repetition, ampersand storms. 6653 parsed, 33 347 threw, **zero untyped
  throws and zero coverage violations**; slowest single parse 16.8 ms. The plan defers this
  invariant to Phase 12's fuzzing harness, but waiting means building five sub-phases on an
  unverified promise, so a deterministic 4000-case slice runs in the test suite.
- 24.5 MB/s full-tree parsing across the corpus; 8 MB single part in 714 ms.

## What the review changed

0.4 was reviewed adversarially after it was built: four lenses over the source, every finding then
handed to a verifier whose job was to **refute** it by running the code. 29 findings raised, 6
survived, 2 refuted. I reproduced all six myself before changing anything, because the equivalent
review at 0.3 was wrong twice in five findings by working from stale files.

**The one that mattered was found by mutation testing, not by reading.** `checkTreeCoverage` tiled
document children and element children, and `checkSpanCoverage` tiled token spans — but neither ever
read `nameEnd`, `trailingSpaceStart`, or any of an attribute's six offsets. Those exist _solely_ for
0.6's rebuild path, so they were the only offsets in the package that nothing verified.

Introducing an off-by-one into `trailingSpaceStart` for tags with two or more attributes passed all
123 tests, both coverage checks, and the fuzz slice. It would also have passed **0.5's byte-identical
corpus gate** — because in a round trip nothing is dirty and every element re-emits as
`slice(start, end)`. The first failure would have been in 0.6, on a user's edited deck, silently
rewriting `<p:cNvPr id="2" name="Title 1" />` as `<p:cNvPr id="2" name="Title 1"/>`. That is the
whole lesson of the sub-phase restated: a gate that does not read a field cannot defend it.
`checkTagInterior` now tiles the tag as well, and the zero-attribute arm is not redundant with the
closing-delimiter check — `<a:bodyPr  />` with a corrupted offset still leaves a slice of `" />"`.

The other five, each verified by running it:

|                                                                 | Was                                                                                                                    | Now                                                                                     |
| --------------------------------------------------------------- | ---------------------------------------------------------------------------------------------------------------------- | --------------------------------------------------------------------------------------- |
| `decodeReference` ignored the `end` its callers passed          | `<r a="&amp" b="x;"/>` → `no entity named "amp" b="x"`                                                                 | reads no further than the run it was given                                              |
| the declaration was scraped with three unanchored regexes       | `<?xml version="1.0"?>` parsed; `<?XML` was a declaration; `version` was read out of the _inside_ of an encoding value | `XMLDecl` parsed as the fixed grammar it is, and character-checked like everything else |
| `&#X41;` decoded to `A`                                         | the `x` in `'&#x'` was treated as case-insensitive                                                                     | it is a terminal in a case-sensitive grammar; refused                                   |
| no ceiling was denominated in what governs allocation           | 20 000 elements × 256 attributes = 41.9 MiB source, 20 001 nodes (0.5% of `maxNodes`), **817 MB of heap**              | `maxSourceLength`, 16 MiB                                                               |
| `resolvePrefix` rescanned every ancestor's whole attribute list | a legal 1.7 MB document took **37 s** in `undeclaredPrefixes`                                                          | elements carry `hasNamespaceDeclarations`, computed free in the attribute loop          |

And four smaller ones: qualified names are now validated on attributes and end tags rather than
elements only (`<r a="1" :a="2"/>` produced two attributes that both resolved to `('', 'a')` and slid
past the duplicate-attribute check); CDATA outside the root is refused, as text already was; an
unpaired surrogate is refused (the single value for which decode and encode are not inverses —
reachable only through `parseXmlString`, never through `parseXml`); and `<!DocType r>` gets
`ERR_DOCTYPE_FORBIDDEN` rather than degrading to generic malformed XML.

Two findings were **refuted** and are recorded here so they are not raised again: that `prefixMap()`
is blind to the rewrite it exists to catch (it catches every rewrite a serializer bug can produce;
only an exact URI _permutation_ between sibling subtrees escapes), and the tag-interior gap as
originally framed via hand-assigned struct fields (not an input — the mutation-testing form above is
the real one).

`ERR_UNDECLARED_PREFIX` was removed from `XML_ERROR_CODES`: it shipped in the public type and no
path ever threw it.

### Two things the review found that are still open

- **`&#X41;` has not been checked against PowerPoint.** It is refused on the strength of the grammar
  — `CharRef ::= '&#x' [0-9a-fA-F]+ ';'` is case-sensitive in the `x` — and none of the 62 variants
  used that spelling, so the "58 verdicts agree exactly" figure below is unaffected. But this project
  settles disagreements by asking PowerPoint, and that has not been done here.
- **The 2834-part corpus gate is not automated in this repository.** It lives in a throwaway script.
  CI runs the unit tests and the fuzz slice; a contributor cannot re-run the measurement this ADR
  leans on. Sub-phase 1.1's licensed, checked-in corpus is where that gets fixed.

## What is deferred, and to where

**0.5** — the serializer, dirty propagation to ancestors only, the prefix-map assertion, and the
byte-identical round-trip gate across the whole corpus. 0.4 stops at reading; nothing here mutates a
tree, and `dirty` is written but never set.

**0.6** — `insertInOrder` against the generated schema tables, the Markup Compatibility walker that
resolves `mc:Choice/@Requires` prefixes through the scope this package exposes, `extLst` as an ordered
opaque list, and `XmlEdit` with exact inverses. The offsets 0.4 records — `trailingSpaceStart`,
per-attribute leading whitespace, `nameEnd` — exist so that 0.6 can rewrite one attribute of a tag and
leave the spacing of every other alone.

**Out of scope entirely** — DTD processing of any kind, XML 1.1, schema validation (that is
`@pptx-studio/validate` in 1.2), and encodings other than UTF-8 until a real file needs one.
