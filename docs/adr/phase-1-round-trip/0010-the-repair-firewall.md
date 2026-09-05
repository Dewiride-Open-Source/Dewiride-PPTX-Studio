# 0010 — The repair firewall, and the difference between invalid and broken-by-us

Date: 2026-08-28
Status: **accepted** — 29 rules, 17 refusal fixtures, quiet on all 51 corpus decks
Sub-phase: 1.2 — `validate`: the repair firewall

---

## Context

The plan's 1.2 is one paragraph and a 29-line appendix: _all 29 rules; runs on
every export in dev and prod and refuses to hand over bytes on failure; verify
by hand-corrupting a deck 29 ways._ The appendix is explicitly abbreviated —
"the full 29 land in 1.2" — so choosing them is part of the sub-phase rather
than transcription.

The reason the package exists is stated in the plan and is worth restating,
because everything below follows from it: **PowerPoint emits no diagnostic log
whatsoever.** When it refuses a file the message names no part, no element and
no line. When it _repairs_ one it says less than that. A user who gets a repair
prompt has no way at all to find out why, and neither do we — except by building
a package with one change in it and watching what happens, which is what
sub-phases 0.7 and 1.1 spent months doing and what `cli bisect` will automate in
1.5.

So this is not a schema validator that produces a tidy report. It is the only
feedback loop that exists.

---

## Decisions

### The twenty-nine are seven groups, not six

The appendix names six: package, relationships, element order, required, IDs,
preservation. Those account for twenty-four rules and every bullet in the
appendix maps onto one of them.

The seventh is **`refused`** — five rules for schema-legal markup PowerPoint
declines. `p:ph type="hdr"` on a slide. A geometry guide referenced but never
defined. A `c:strLit` inside a series `c:tx`. A `p:control`, in any of eight
forms. A `cs:chartStyle` with thirty of its thirty-one entries. None of those
can be derived from ECMA-376, because ECMA-376 permits all five.

Giving them their own category is not filing. It changes what the table means: a
rule in `package` or `required` is a citation, and a rule in `refused` is a
_measurement_, falsifiable in a way a citation is not. If a later PowerPoint
build opens one of these, the rule is wrong and should be deleted — so each
carries a `why` recording what was tried and what opened, and the entries are
tagged `evidence: 'measured'` rather than `'schema'`, with a test that holds the
distinction.

The count landed on twenty-nine because the plan says twenty-nine, and it was
not padded to get there. Two appendix bullets — "`a:endParaRPr` last inside
`a:p`; `a:rPr` before `a:t`" — turned out to be already covered by the generated
ordering table, which ranks both. Rather than write a duplicate rule for the
message's sake, the freed slots went to the measured refusals, which had no rule
at all.

### `origin`: it answers "did we break it", not "is this valid"

This is the decision the rest of the package is shaped around.

A fatal finding in a file the user imported ten seconds ago is information. The
same finding in a file they just edited is a bug in us. Refusing both is the
strict-looking choice and it is wrong: a deck with one pre-existing defect — a
dangling image relationship, which PowerPoint tolerates and which is common in
files that have been through three other tools — could then be opened in this
editor and never saved again. The editor would be refusing to give somebody back
their own file over a problem it did not cause and cannot fix.

`PartStore.write` already made exactly this call for dangling relationships in
sub-phase 0.3. This generalises it to all twenty-nine.

**And it is computed, not judged per rule.** Every finding carries an `origin`,
worked out by running the rules a second time against the package _as it was
opened_ and differencing the two reports: a finding in both is `inherited`, one
only in the new report is `introduced`, and only `introduced` fatals refuse an
export. That is exact, it needs no rule to reason about history, and it cannot
drift from what the rules actually do — which a per-rule judgement would, on the
first rule somebody added without thinking about it.

The second pass runs only when the first found something fatal, and only the
rules that fired, so a clean export — nearly all of them — pays nothing.

### Findings carry an XPath, in the document's own prefixes

The plan asks for "part URI and XPath". Both are on every finding, plus a byte
offset for `cli bisect`.

The offset alone would have been cheaper and it is useless to a person: it names
a position in a file that is usually one line long and often two megabytes wide,
and it stops being true the moment anything above it changes. `ppt/slides/slide3.xml`
at `/p:sld/p:cSld/p:spTree/p:sp[2]/p:nvSpPr/p:cNvPr/@id` says what is wrong in
one line; byte 41 207 does not.

The prefixes are the document's own, never canonicalised. A path that rewrote
`pp:` to `p:` would be a more correct XPath expression and a worse _location_ —
the reader would not find it in the file in front of them. It is also the same
rule the whole project runs on: `mc:Ignorable` and `mc:Choice/@Requires` hold
prefixes rather than URIs, so a prefix is load-bearing data in OOXML and
inventing one is never neutral.

Positional predicates are emitted only where they disambiguate. `/p:sld[1]/p:cSld[1]`
is noise; `/p:spTree/p:sp[2]` is the finding.

### Nothing is skipped silently

Three rules compare against the package as it was opened. One reads the archive
rather than the part store. Given neither, they do not run — and the report
lists them under `skipped` **with the reason**, because a preservation rule that
reports nothing because it had nothing to compare against looks exactly like one
that found nothing, and those are opposite answers. Parts that would not parse
land in `problems` for the same reason, and `V002` records a `problem` when it
could check the parsed content-type map but not the markup.

That is three separate channels for "this report means less than it looks like",
and all three are printed.

### The corpus is the check on false positives, and it earned its keep

A false negative costs one missed defect. A false positive blocks an export the
user wanted, and a tool that blocks exports the user wanted gets switched off,
after which it catches nothing at all. Those are not symmetric and the second is
much harder to notice, because a rule that is too strict looks exactly like a
rule that is working.

So the second half of the verification is `tools/corpus/validate.test.ts`: all
fifty-one committed corpus decks through the validator, zero fatal findings.
Every deck in that corpus opens in PowerPoint — that is what the corpus _is_,
established one bisection at a time across 0.7 and 1.1 — so a rule that fires on
one is a bug and there is no third possibility to argue about.

The first run produced **134 findings across eleven decks**, and every one was
worth having:

- **82 × `V012`**, mostly `mc:AlternateContent`, plus `a14:m`, `p14:honeycomb`
  and `a37-mce`'s deliberate `zz:` markup. The cause is a real ambiguity in
  `childRanks`, which returns `undefined` both for "the schema forbids this
  child here" and for "we have never read the schema this child comes from".
  That collapse is _correct_ for a caller trying to insert a child — neither can
  be placed — and catastrophic for one trying to judge an existing one. Fixed by
  exporting `SCHEMA_NAMESPACES` and `isSchemaNamespace` from `@pptx-studio/xml`
  and checking only children from the four vocabularies the table was generated
  from.
- **52 × `V020`**, in two kinds. Shape ids reused across the `mc:Choice` and
  `mc:Fallback` of one `mc:AlternateContent` — which are alternatives, never
  both present, and PowerPoint's own writer does exactly this in `a22-chartex`.
  And `dsp:cNvPr id="0"` on every shape of a SmartArt drawing part, which
  PowerPoint generates and regenerates. Fixed by scoping the uniqueness check to
  Markup-Compatibility branches that can coexist, and to PresentationML.
- **13 × `V012` that were right.** `a20-tables` wrote `<a:tblBg><a:solidFill>`
  and `<a:tcStyle><a:solidFill>` where the schema has a `fill|fillRef` choice, so
  the fill belongs one level deeper — while `a:tblPr`, two elements away, takes
  the fill directly. A real defect in a corpus deck, found by a validator that
  had been running for an hour. The deck was fixed and regenerated; `C-REGEN`
  and `C-CENSUS` both still pass.
- **2 × `V021` warnings**, both correct and both left standing. `a02-placeholders`
  carries a placeholder its generator comments as _"Tier 5: orphan. Nothing to
  inherit"_; the rule found it and described it in the deck's own words without
  being told it was there. `a19-decorative` has the same shape incidentally. Both
  are pinned in the test by XPath, so a third warning is somebody's problem.

Two rules changed, one shared helper was added to another package, and one
corpus deck was wrong. None of that would have been visible from the twenty-nine
unit tests, all of which passed throughout.

### `corpus/reject/`: seventeen packages that must fail

The mirror image, and the thing `ROSTER.md` promised sub-phase 1.2 would build.
Four rules — `V022`, `V024`, `V025`, `V026` — have **no instance anywhere in the
good corpus**, because every file containing one is a file PowerPoint refuses.
Without fixtures they would be enforced entirely on trust.

Three decisions inside it:

**Minimal packages, not chassis decks.** `tools/corpus/gen/package.ts` builds a
realistic deck — master, layout, theme with the two gradients `fillRef/@idx`
needs, `docProps`, twenty-two entries. That is right for a probe. It is wrong
for a refusal, where the point is _the smallest markup that reproduces it_,
because that is what `cli bisect` reduces towards and what a person reads when
deciding whether a rule is still true. So `tools/corpus/reject/deck.ts` builds
ten parts and no theme, and each fixture changes one thing.

**`kind: "fixture"`, not `kind: "deck"`.** They are `.pptx` files, so the obvious
choice would corrupt two other rules: `C-COV` counts a deck's non-zero `features`
as coverage, and a package PowerPoint refuses is not evidence a feature works;
`C-LEX` counts distinct XML producers, and these share ours.

**The limitation is on every entry, not in a footnote.** The _markup_ in each
fixture is markup that was measured and refused. The _package_ around it is new
and much smaller, and nobody has opened these seventeen files in PowerPoint. That
is stated in `fixtures.ts` and in each entry's `refusal` field, which records
what PowerPoint said and where the finding is written down. `C-REJECT` asserts
only what is checkable here — that our validator refuses each one, by the rule
its manifest names. Whether PowerPoint agrees with our _reason_ is a separate
claim that `tools/corpus/gen/open-in-powerpoint.ps1` settles, on a machine with
PowerPoint, by somebody who chooses to run it.

### `M1.8` is fatal here, and a warning in `opc`

The carried follow-up from 1.1, discharged. `@pptx-studio/opc` rates a
percent-escape of an unreserved character a warning, and that is right _there_:
the package layer has to open what it is given, and a file that arrives with
`%5F` in a part name is a file the user still wants to see. Here the package is
on its way out, and PowerPoint refuses it with `0x808D1005` — measured nine for
nine on one-name packages.

`V004` reports every `validatePartName` violation at the rule's own severity,
warning or not, for the same reason. Read leniently, write strictly; and a name
a deck _arrived_ with comes back `inherited` and does not block anything.

---

## Consequences

`pnpm check` gains 112 tests, in seven files. `pptx-studio validate <deck>` exists and exits 1
on anything fatal — a diagnostic, not the gate, because a file on the command
line has no history and the three preservation rules cannot run on it. The gate
is `assertValid`, which sub-phase 1.3's writer calls with both packages in hand.

The corpus gains a sixth collection and its fifth named rule. `corpus/` is now
2.3 MiB across 78 entries.

`@pptx-studio/xml` gained two exports, `SCHEMA_NAMESPACES` and
`isSchemaNamespace`, which is a small widening of a layer-0 package's surface in
service of a layer-3 one. It is the right place for it: the four namespaces the
ordering table was generated from are a fact about that table, and the
alternative was hard-coding them in `validate` where they would drift the first
time the codegen read a fifth schema.

### What is not done

**`V027`, `V028` and `V029` have never run on a real edit**, because nothing
edits a package yet. They are tested against synthetic ones — a store opened
from modified bytes with no `replacePart` (which is what a writer re-serialising
a clean part looks like), and a store with `replacePart` called (which is what
an edit looks like). Sub-phase 1.3 is the first caller that will exercise them
for real, and 1.3's own gate — a no-op export of all 51 decks — is exactly the
shape that would catch a mistake in them.

**`V021` uses a five-tier matcher that sub-phase 7.1 owns.** It is implemented
here because the rule needs it and 7.1 is a long way off, and it is a _warning_
partly for that reason. While writing it: the plan says `<p:ph/>` means
`type="obj"` in sub-phase 2.9 and `type ??= 'body'` in 7.1, and those cannot both
be right. ECMA-376 gives `CT_Placeholder/@type` a default of `body`, so that is
what this uses; 7.1 settles it against the 121-case matrix.

**The built-in geometry guides are a stand-in.** `V023` needs the seeded guide
table that sub-phase 2.2 owns and `@pptx-studio/geometry` will hold. Until then
there is a deliberately generous list here plus two shape tests, because the
rule is fatal: a name we wrongly think is undefined refuses a file that works,
while a name we wrongly accept only misses one instance of a defect whose
real-world form is a typo or a deleted `a:gd` — and neither of those looks
anything like `wd12`.

**Passing all twenty-nine remains necessary and not sufficient.** PowerPoint
rejects some schema-legal markup for reasons nobody has enumerated. The
strongest defence is architectural — never synthesize markup we did not read —
and the second strongest is 1.5's `bisect`.
