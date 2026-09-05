# ADR 0005 — The serializer and the round-trip gate

- **Status:** Accepted
- **Date:** 2026-08-26
- **Sub-phase:** 0.5

## Context

Sub-phase 0.4 built a tree over a source string in which every node remembers the span it came from.
This one turns the tree back into bytes.

The plan calls 0.5 "the load-bearing gate of the entire architecture" and sets its exit criterion as
**parse → serialize → byte-identical for 100% of parts across the whole corpus**, with the
instruction that it must not be softened.

It is worth saying plainly what that criterion does and does not establish, because the gap is the
whole substance of this sub-phase.

## The stated gate is passable by `return document.source`

In a round trip nothing is dirty. Every node re-emits as `source.slice(start, end)`, the spans tile
the source exactly — which 0.4 already proved with `checkSpanCoverage` and `checkTreeCoverage` — so
the concatenation is the source by construction.

A byte comparison against the corpus therefore tests the **tokenizer's spans**, which were already
gated, and says nothing at all about the serializer's ability to write anything. The one line above
scores 2834/2834.

That is not a criticism of the plan; the gate is a proxy, and a good one, for the property that
actually matters. But building only what the proxy measures would ship a serializer whose entire
rebuild path — the half sub-phase 0.6 runs on — has never been executed.

So the gate here is the property, and the byte comparison is kept alongside it as the special case
it is:

> **Serialize the document, parse the result, and the two trees must agree on everything that
> carries meaning** — node order and kind, qualified names, attribute names, values, order and
> quoting, text, and the namespace bindings of every prefix.

This holds whether or not anything was edited. Which means it can be run with **every node and every
attribute in a real part marked dirty**, forcing the whole part through the rebuild, and that is how
the corpus is used below.

## Rebuilding is a right inverse, and four characters are lost without one

The tokenizer's central asymmetry — a node's source and a node's value are different things — runs
in the opposite direction here. Writing a value out as we find it is wrong, because XML transforms
what it reads on the way back in:

| a value containing | written literally, reparses as | by     |
| ------------------ | ------------------------------ | ------ |
| text `\r`          | `\n`                           | §2.11  |
| attribute `\r`     | a space                        | §3.3.3 |
| attribute `\n`     | a space                        | §3.3.3 |
| attribute `\t`     | a space                        | §3.3.3 |

So each becomes a character reference. This is the same trap that makes `name="a&#9;b"` and
`name="a<TAB>b"` different values on the way in — except that on the way out, getting it wrong
deletes data rather than merging two values.

`&`, `<` and the delimiting quote are escaped for the ordinary reasons.

## The corpus overruled my escaping rule for `>`

Nothing requires `>` to be escaped except inside `]]>`, which §2.4 forbids in content. So the narrow
rule is to escape it only there and write a bare `>` everywhere else, and that is what this did
first — reasoning that a rebuilt node should keep its producer's spelling, so that the rebuild path
could be compared byte-for-byte and not merely value-for-value.

Measured, across all 2834 parts, 22 461 text nodes and 170 019 attribute values:

| spelling of `>` inside a text node or attribute value | occurrences |
| ----------------------------------------------------- | ----------- |
| a literal `>`                                         | **0**       |
| `&gt;`                                                | **5**       |

The producers never write a bare `>` inside content. The narrow rule therefore re-spelled all five
references and preserved nothing; escaping `>` unconditionally re-spells nothing. It also deletes
the `]]>` special case rather than handling it — if `>` is never written literally, the sequence
cannot occur.

The reverse case is not a judgement call and went the other way: **only the delimiting quote is
escaped.** All 170 019 attributes are double-quoted and 26 carry a `&quot;`, so escaping the
apostrophe as well would re-spell real values for nothing.

## What a forced rebuild changes on 2834 real parts, and it is one thing

|                                                               |                        |
| ------------------------------------------------------------- | ---------------------- |
| parts                                                         | 2834, from 37 packages |
| **(1) clean → byte-identical**                                | **2834 / 2834 — 100%** |
| **(2) forced rebuild → value-identical**                      | **2834 / 2834 — 100%** |
| (3) forced rebuild → byte-identical                           | 2178 / 2834 — 76.85%   |
| parts where (3) differs                                       | 656                    |
| …of those, differing **only** by CRLF → LF in a text node     | **656**                |
| …differing for any other reason                               | **0**                  |
| entity references re-spelled (`&amp;` `&lt;` `&gt;` `&quot;`) | **0 of 39**            |
| coverage gaps, dirty-invariant violations, failures           | 0                      |

Row (3) is measured, not asserted, and the single cause is worth stating precisely: CRLF and LF are
the same value after §2.11, so a text node being rebuilt from its value cannot know which it came
from. 656 parts have a CRLF between the declaration and the root element.

**This costs nothing in real operation, and the reason is the dirty rule.** `dirty` travels up, never
down — so a text node is rebuilt only when someone edited _that node_, and the whitespace between
siblings in an untouched part is never dirty. Marking everything dirty is a test instrument for
reaching the rebuild path, not a thing 0.6 will ever do. `serialize.test.ts` pins both halves:
under a forced rebuild the CRLF becomes LF, and under `markDirty(root)` alone the source comes back
byte-for-byte.

Throughput: 28.3 MB/s parse-and-serialize, 44.9 MB/s for a full rebuild.

## The walk carries its own stack

For the same reason the parser's does. PowerPoint opens a part nested 5000 elements deep; a
recursive serializer answers that with a `RangeError`, which is the one thing this package has
promised never to throw. End tags ride the work stack as plain strings. Pinned with a 5000-deep
document.

## The serializer validates what it is about to write

Only on the rebuild path, so a clean document pays nothing. Sub-phase 0.6 will hand it strings, and
a part that leaves here looking well-formed and is refused by PowerPoint with no diagnostic is the
most expensive failure in this project to debug. So: names must still be names, values must hold
only characters XML permits and no unpaired surrogates, and the three constructs with no escape
mechanism — comments, CDATA sections, processing instructions — refuse a value that would contain
their own terminator.

A carriage return in a comment or a CDATA section is refused outright rather than silently altered.
Neither construct admits a character reference, so §2.11 would turn it into a line feed. The parser
never produces such a value; only a caller that synthesized one can.

## Lexical choices the rebuild keeps

Each of these is preservation the plan's byte gate would not have noticed the loss of, because a
clean document never reaches the code:

- the **quote character** each attribute was written with;
- the **whitespace before `>` or `/>`** — 70 822 of the corpus's 98 777 self-closing tags are written
  `<a:off x="0" y="0" />`;
- the **indentation between attributes**, taken from the source rather than guessed, because
  `attribute.start` is the start of its leading whitespace and not of its name;
- `<x></x>` versus `<x/>`, in both directions — and the empty form is **derived** from
  `selfClosing && children.length === 0` rather than read from the flag, so no edit operation in 0.6
  has to remember to flip it when it adds a child.

## `markDirty`, and the invariant that makes it safe

`dirty` travels to ancestors and stops at the first one already dirty, which keeps a 60-frame drag
from being quadratic in depth. An attribute is not a node and has no parent, so `markAttributeDirty`
exists; setting `attribute.dirty` by hand and forgetting the element is a **silent no-op** — the
element stays clean, re-emits its whole start tag as a slice, and the edit vanishes with no error
anywhere.

`checkDirtyInvariant` reports exactly that, and `checkRoundTrip` catches it independently.

## Verification

`pnpm check` green: layering, format, lint, typecheck, build, publint/attw, **417 tests in real
Chromium** across 20 files, up from 384.

Three instruments, deliberately different in kind:

1. **The corpus gate**, above. 2834 parts, all three rows.
2. **Fuzzing.** The 4000-case slice of the 40 000-case sweep now also asserts that every mutant which
   parses serializes back byte-identically, and still reads back as the same document once forced
   onto the rebuild path. 40 000 mutations: 2855 parses, 37 145 typed errors, **zero** untyped
   throws, coverage violations or fidelity failures; slowest single case 17.9 ms.

   One finding from building it, which is why it is listed as its own instrument: the escaping rules
   were **unreachable** from the original fuzz corpus. A tab, line feed or carriage return written
   literally has already been normalized away by the time the serializer sees the value, so only a
   character reference can put one there — and the corpus had none. Three deliberate breaks to the
   escapers survived all 4000 cases. A source carrying `&#9;`, `&#xA;`, `&#xD;`, `&#62;` and `]]&gt;`
   fixed it, and now kills them.

3. **Mutation testing.** 17 deliberate defects introduced one at a time — each of the four lost
   characters, both `>` rules, the quote rules, the self-closing derivation, the trailing space, the
   attribute indentation, the quote character, an off-by-one in the clean slice, both halves of the
   dirty check, the end tag, and the byte order mark. **17 of 17 killed.** This is the instrument
   that says the tests are load-bearing rather than merely green, and it is the one that found 0.4's
   worst defect.

## What is deferred, and to where

- **Adjacent text nodes.** Two text children in a row serialize to one run and reparse as one node.
  A parse never produces that shape; an edit could. It belongs to 0.6, which owns node insertion,
  and `checkRoundTrip` will report it as a structural difference if 0.6 gets it wrong.
- **Declaration quoting.** `XDeclaration` records the three values and not their quote characters, so
  a _rebuilt_ declaration is written with double quotes. Nothing edits a declaration, and a clean one
  is a slice. If 0.6 ever needs to, the field is one line.
- **The corpus gate is still not automated in the repo.** It lives in a scratchpad script pointed at
  decks outside the working tree, so CI runs the unit tests and the fuzz slice only, and a
  contributor cannot re-run the measurement this ADR rests on. Sub-phase 1.1's checked-in,
  licence-audited corpus is where that gets fixed — it is now blocking two ADRs rather than one.
