# corpus

Sub-phase 1.1's gate: **CI fails on any corpus entry lacking a licence.**

```bash
pnpm corpus
```

Runs second in `pnpm check`, after `layering` and before `format:check`. It needs no build, no
workspace package and no browser, so it is the cheapest thing in the gate and a contributor learns
about a licensing problem in a second rather than four minutes into a Chromium install. Ordering it
after formatting would mean a run that stops on whitespace never reports that a fixture has no
licence, which is backwards.

## What it reads

Every `corpus/**/manifest.json`. There is deliberately no single root manifest: the six
collections have genuinely different envelopes (`bench` records a generator, `ground-truth` records
what it was measured on, `reject` records what PowerPoint said) and flattening them would force the
union of all six onto every entry,
while a roll-call of includes would let deleting one line silently uncover a whole directory.
Discovery plus `C008` is stronger than either, with less code.

Entry ids are unique across **all** manifests, which is what lets one collection reference another's
entry rather than duplicate it.

## The three storages

| `storage`   | carries                 | means                                                          |
| ----------- | ----------------------- | -------------------------------------------------------------- |
| `committed` | `path`                  | the bytes are in the repository and hash to `sha256`           |
| `generated` | `outputName` + `recipe` | too large to commit; the recipe reproduces exactly these bytes |
| `pinned`    | neither                 | a hash of a file the repository cannot locate, and must not    |

`generated` exists because `corpus/bench/`'s largest deck is 188 MiB. `pinned` exists for
measurements of files that stay on the machine that made them — the same posture as
`eot-headers.json`, which records the header of 76 Microsoft-authored font parts and not one byte
of the fonts.

## The rules

| id                    | asserts                                                                 |
| --------------------- | ----------------------------------------------------------------------- |
| `C001-no-licence`     | a `license` from the closed set; `NONE`/`NOASSERTION` rejected by name  |
| `C002-no-source`      | a `source` from the closed set                                          |
| `C003-source-licence` | the licence is one that source may claim                                |
| `C004-attribution`    | third-party work carries a complete CC BY §3(a)(1) attribution          |
| `C005-unknown-key`    | no unrecognised key in an envelope or an entry                          |
| `C006-id`             | kebab-case, unique across every manifest                                |
| `C007-storage`        | the storage's required fields, and no path escapes `corpus/`            |
| `C008-orphan`         | every file under `corpus/` is claimed by exactly one manifest           |
| `C009-double-claim`   | no file is claimed twice                                                |
| `C010-sha256`         | committed bytes hash to what the manifest says                          |
| `C011-bytes`          | committed size matches                                                  |
| `C012-size-cap`       | 512 KiB per file, 12 MiB total                                          |
| `C013-feature-key`    | every `features` key is a real census key; decks have one, fixtures tag |
| `C014-recipe-tool`    | `recipe.tool` names a file that exists under `tools/`                   |
| `C015-derivation`     | a redacted derivative names its original's hash, and it differs         |
| `C016-gitattributes`  | every committed binary extension has a `binary` line                    |
| `C017-sorted`         | entries are sorted by id                                                |
| `C018-serializers`    | a collection with committed decks names what wrote its XML and its ZIP  |
| `C019-coverage`       | every census key is exercised by a deck, or declared in `uncovered`     |

### Three that are worth explaining

**`C003` is the rule that does the legal work.** `C001` — the one sub-phase 1.1's plan actually
names — is satisfied by typing six characters into a field. `C003` is what says a deck somebody else
wrote and released under CC BY is not ours to relicense as Apache-2.0, whatever the field says.

**`C008` is the one that fires in real life.** Nobody adds a manifest entry and forgets the licence;
what happens is that a deck gets dropped into `corpus/` and no manifest is touched at all. `C008`
enumerates with `readdirSync` rather than `git ls-files` for three reasons: it is roughly sixty times
faster, it works in a source tarball with no `.git`, and — decisively — it sees **untracked** files,
so a local `pnpm check` catches the deck before it is ever committed.

The corollary is a discipline: nothing may write scratch output into `corpus/`. Every generator in
this repository takes an output path and defaults nowhere.

**`C019` is `C-COV`, and it is the only rule here that is about what the corpus does _not_ have.**
`C013` checks that every feature key a deck claims is a real one; `C019` is the converse and the
harder direction — every key the census can report is exercised by some deck, or a manifest carries
an `uncovered` entry saying why not and what would close it. Both halves fire: a key nothing covers
and nobody declared is how a coverage badge becomes a lie, and a key still declared after a deck
started covering it is how the declaration rots into an apology for a hole that was filled.

Two definitions inside it are load-bearing. A `features` count of **zero** is not coverage — it is
`a01-minimal` saying the census looked and found none, and reading it as coverage would let one deck
declaring every key at zero score the corpus at 100%. And a `pinned` deck is not coverage either,
because its bytes are a hash of a file that never left the machine that made it.

## Not in turbo, on purpose

`pnpm layering` is the precedent. If anyone wires this into a turbo task the hazard is precise:
turbo's default inputs cover a package's git-tracked files, and the moment somebody writes an
`inputs` array to tweak one thing they opt out of that default — `corpus/**` silently stops being
hashed and a deck added without a manifest entry replays a cached PASS. Turbo also hashes tracked
files only, so an untracked deck is invisible to it however inputs are declared. That is the third
reason `C008` walks the filesystem itself.

## The files

| file                        | what it is                                                               |
| --------------------------- | ------------------------------------------------------------------------ |
| `ROSTER.md`                 | the fifty-one decks, deck by deck, and what each is the probe for        |
| `schema.ts`                 | the enums — the legal posture in executable form                         |
| `check.ts`                  | `checkCorpus(input)`, pure: no filesystem, no process                    |
| `check-corpus.ts`           | discovery, hashing, formatting, exit code                                |
| `check.test.ts`             | one `it` per rule, synthetic input, no temp directories                  |
| `conventions.test.ts`       | every producer here writes the bytes PowerPoint writes                   |
| `lexical.ts`                | `C-LEX`: the header reader and the byte-level lexer, no declared policy  |
| `lexical-forms.ts`          | the 60 declared forms, their evidence in the wild, and every gap         |
| `lexical.test.ts`           | `C-LEX` itself: the table against a fresh scan, in both directions       |
| `coverage.test.ts`          | `C-COV` as a number: what the corpus covers, and the one gap it declares |
| `census-keys.gen.ts`        | generated; the 47 census feature keys                                    |
| `census-drift.test.ts`      | asserts the generated list still matches the census                      |
| `write-census-keys.ts`      | regenerates it (`pnpm build` first)                                      |
| `gen/`                      | the Tier A generator: chassis, shape and text helpers, one file per deck |
| `gen/*.ts` (bytes)          | `png` `jpeg` `emf` `xlsx` `media` `cfb` `vba` — every binary we author   |
| `gen/cfb.test.ts`           | the two layers `a32-macros` needs that no other test can see             |
| `authored/`                 | the PowerPoint COM scripts, so Tier B's provenance is auditable          |
| `authored/build-tier-b.ps1` | the nine Tier B decks, one block per deck                                |
| `authored/decks.ts`         | their reviewed censuses and descriptions                                 |
| `authored/decks.test.ts`    | `C-CENSUS` for Tier B, and the privacy and font constraints              |
| `written/`                  | the Tier C deck: our own writer, one deck round-tripped                  |
| `written/build-written.ts`  | builds it and its manifest; `--check` is `C-REGEN` (`pnpm build` first)  |
| `written/decks.ts`          | its reviewed census, and the source deck it pins                         |
| `written/decks.test.ts`     | `C-CENSUS` and `C-REGEN` for Tier C, and the three headers we drop       |

### `gen/` — the Tier A decks

```bash
node tools/corpus/gen/build-probes.ts --out corpus/decks --manifest
npx prettier --write corpus/decks/manifest.json
node tools/corpus/gen/build-probes.ts --check --out corpus/decks
powershell -File tools/corpus/gen/open-in-powerpoint.ps1 -Dir <dir>
```

The Prettier line is not optional. `--manifest` writes `JSON.stringify(…, null, 2)`, which puts
every `recipe.args` element on its own line, and Prettier collapses short arrays — so
`pnpm format:check` fails on a freshly generated manifest. Nothing downstream cares which form it
is in: `C010` and `C011` hash the `.pptx` files, not the manifest, so reformatting is safe.

`--out` has no default, which is the corollary of `C008-orphan`: a generator that wrote into
`corpus/` by accident would break the gate for whoever ran it next.

`--check` is `C-REGEN` — it rebuilds every deck and compares against what is on disk without
writing. Committed bytes and a recipe that reproduces them are two different claims and hashing the
file only checks the first.

`open-in-powerpoint.ps1` is the gate no amount of schema conformance substitutes for. It reports the
shape count of every slide, because the failure that matters is not the refusal — which is loud —
but the **silent repair**, which opens successfully with content quietly dropped.

A shape count only sees slides, though, and from `a12-masters` onward most of what a deck probes is
not on one. A deck carrying masters, notes pages, comments, sections, custom shows, transitions or
timings needs a second pass that reads them back through PowerPoint's object model — `Designs.Count`,
`SectionProperties.Count`, `NamedSlideShows.Count`, `Slide.NotesPage.Shapes.Count`,
`Slide.Comments.Count`, `CustomLayouts(i).SlideShowTransition.EntryEffect`,
`Slide.TimeLine.MainSequence.Count`, `PageSetup.SlideWidth`. Every one of those can come back zero
on a package PowerPoint opened without complaint, which is exactly the failure the shape count was
introduced to catch, one level in.

And a shape count sees only what PowerPoint chooses to enumerate, which from `a20-tables` onward is
not everything it kept. So the hard-content decks add a **third** pass: copy each deck into the
scratchpad, open it, `SaveAs` a copy, and diff the copy's parts and elements against the original.
That is the only pass that separates "not surfaced as a shape" from "silently dropped", and the
two look identical from a count — `a27-ink` reports two shapes a slide where three were written,
and only the re-save shows that every `p:contentPart` and its `mc:Fallback` picture are gone.

It cuts both ways, which is the point. The same pass showed PowerPoint _adding_ content: a VML
group arriving as a native `p:grpSp`, a missing SmartArt drawing being laid out and written, a
transitional `p:oleObj` upgraded to the switch form. None of that is a defect in a deck; all of it
is a list of things **our** writer must not do, and it is why sub-phase 1.4 compares canonical XML
and a relationship graph rather than bytes. Write what it finds into `ROSTER.md` beside the
refusals.

It is also the only pass that can tell an accepted value from a **kept** one. `a39-large-ids`
carries a `p:cNvPr id="4294967295"`; PowerPoint opens the package, reports the shape, and renumbers
the id to `4` on save — so the value is tolerated rather than valid, and nothing short of re-saving
distinguishes those. The same pass is how `a34-extlst` learned that an `a:extLst` survives with its
order, its duplicate URIs and its GUID case intact while only the **last** child of any one `a:ext`
does.

Every deck gains a `p:extLst` on two slides and three `p:ext` children on this pass. That is
`p14:creationId` and its neighbours, it happens to all forty-one, and it is the baseline to
subtract before anything else in the diff means something.

It is also the only thing that finds what the census cannot. `a05-geometry` passed every test in
this repository while being markup PowerPoint would not open, and `a02-placeholders` did the same:
two of the sixteen `ST_PlaceholderType` values are refused outright, which no schema states. When a
deck refuses, bisect it — package, then slide, then sheet, then one element per package — and write
what you found into `ROSTER.md`'s **What PowerPoint refuses**, because its message names no part,
no element and no line, and the next person gets nothing but that section.

`conventions.test.ts` is why there are three hand-written PresentationML builders in this repository
and no shared markup module. Sharing code makes them agree without making any of them right;
asserting all three against a measurement of PowerPoint's own output checks the property that
matters, and holds for a builder written next year.

### `authored/` — the Tier B decks

```bash
node tools/corpus/authored/make-assets.ts <dir>            # b09's two images, from our own encoders
powershell -File tools/corpus/authored/build-tier-b.ps1 -Out <dir>
node tools/corpus/authored/write-manifest.ts               # describes what is committed
npx prettier --write corpus/authored/manifest.json
```

**These four commands are not a build.** Running them replaces the corpus rather than refreshing
it: every `sha256` changes even when nothing about the decks did, because PowerPoint stamps
`dcterms:created` on every save. That is why Tier B entries carry no `recipe` and why
`write-manifest.ts` reads what is already committed instead of building anything. `C-REGEN` has no
Tier B counterpart and cannot have one.

What is auditable is the authoring rather than the output, which is the whole reason the script is
committed. It reads nothing outside its `-Out` directory: every presentation is created in memory,
`b07`'s workbook is a new empty one made in place, and `b09`'s images come from
`tools/corpus/gen/png.ts` and `jpeg.ts` — there is no stock imagery in this corpus and no
photograph from the authoring machine.

`decks.test.ts` asserts the two constraints that would otherwise be promises: no `ppt/fonts/` part,
and no author, company or machine name anywhere in the archive. Both are worth having as tests
rather than as care, because both have already failed once — the scrub through
`BuiltInDocumentProperties` that `ROSTER.md` originally prescribed cannot be reached from
PowerShell at all, so it ran silently and changed nothing, and the check that should have caught
that was itself reading a field name that does not exist and passing on an empty string. A test
that cannot fail is worse than no test, so this one was verified by pointing it at a string the
decks really contain and watching all nine fail.

### `written/` — the Tier C deck

```bash
pnpm build                                                  # this one needs it; see below
node tools/corpus/written/build-written.ts --out corpus/written --manifest
node tools/corpus/written/build-written.ts --check --out corpus/written
npx prettier --write corpus/written/manifest.json
```

One deck: `b01-blank` read through `PartStore.open` and written straight back out with nothing
touched in between. Its census is `b01`'s key for key and contributes nothing to `C-COV`; what it
contributes is thirty-seven ZIP headers written by the writer this project ships.

`--check` **is** `C-REGEN`, which Tier C has and Tier B cannot. A deflating writer normally cannot
promise reproducible bytes — deflated output depends on which zlib produced it, which is why Tier A
stores its entries rather than compressing them — but a no-op write touches no part, so
`passthroughEntry` moves every already-compressed stream across without inflating it and `fflate`
never runs.

`build-written.ts` imports `packages/opc/dist` and so needs `pnpm build`, for the reason
`write-census-keys.ts` gives: nothing links workspace packages into `tools/`, and the resolution
that lets a Vitest file write `import { PartStore } from '@pptx-studio/opc'` comes from Vitest's
alias table rather than from Node. Here that is a preference as well as a constraint — the deck
committed under `corpus/written` should be the output of the writer we _publish_, not of its
sources. The gate is `decks.test.ts`, which imports `src` and needs no build; the script is the
convenience.

On Windows that dynamic import needs `pathToFileURL`. Without it Node refuses the drive letter as
an unknown URL scheme, which is what `write-census-keys.ts` had been doing since it was written —
its own documented regeneration command could not run on this machine. Fixed in passing, and the
regenerated file came back identical.

`decks.test.ts` asserts the diff against the source rather than the deck's contents, because the
diff is the whole tier: every part byte-identical after inflation and in the same order, the deck a
fixpoint under a second write, and the three header fields our writer drops — version-made-by 45,
general-purpose flags `0x0006`, and five `0xA220` growth hints totalling the exact 1832-byte size
difference — asserted **in both directions**, present in the source and absent here. All three are
legal to drop and none is a bug, which is why they need a test: a reader ignores every one of them,
so nothing else here would notice if our writer started emitting them or if PowerPoint stopped.

`check.ts` is pure for the same reason `tools/layering/check.ts` is: it takes the parsed manifests
**and** a pre-computed file table, so `C008` and `C010` are testable against a corpus shape this
repository does not have yet.

## Exit status

`0` clean, `1` violations, `2` the check could not run. The third is separate on purpose — a
manifest that is not valid JSON is a different incident from a manifest that says something wrong.

## The other five rules

All five are closed. Three of them live outside this checker, because each needs to build a deck,
read a census or unzip an archive and this checker deliberately does none of those:

- **`C-CENSUS`** ✔ — a fresh `censusPackage()` of each committed deck reproduces its `features`
  map exactly, in both directions. Tier A in `gen/probes.test.ts`, Tier B in
  `authored/decks.test.ts`, Tier C in `written/decks.test.ts`.
- **`C-REGEN`** ✔ — every deck rebuilds to the bytes committed under `corpus/decks` and
  `corpus/written`. Tiers A and C; never Tier B, for the reason `authored/` gives.
- **`C-LEX`** ✔ — in `lexical.test.ts`, against the inventory in `lexical-forms.ts`. What it
  found is below.
- **privacy and font rules** ✔ — in `authored/decks.test.ts`, for the only tier that could
  violate them.

- **`C-COV`** ✔ — `C019` above, which unlike the other three needs nothing but the manifests and
  the committed key list, so it runs here on a bare clone. What it answers is below.

- **`C-REJECT`** ✔ — in `reject/reject.test.ts`, and it is the only one of the five that is
  about files that are **wrong**. Seventeen minimal packages, each reproducing one refusal
  `ROSTER.md` measured, each asserted to be caught by the validator rule its manifest names. It
  exists because four of the validator's twenty-nine rules have no instance anywhere in the fifty-one
  good decks — every file containing a `p:control`, a `c:strLit` in a series title, a `hdr`
  placeholder or a short `cs:chartStyle` is a file PowerPoint refuses — so without these they
  would be enforced entirely on trust. Arrived with sub-phase 1.2; see
  `docs/adr/0010-the-repair-firewall.md`.

## What `C-COV` says

**Forty-six of the census's forty-seven feature keys are exercised by a deck.** The one that is not
is `model3d`, and it is declared in `corpus/decks/manifest.json` rather than quietly absent:

> The `a:ext` GUID that carries `am3d:model3d` appears in no public specification. Microsoft's own
> Open XML SDK documents the element and its namespace but not the extension URI that hosts it, nor
> the relationship type and content type of the model part beside it. Guessing the four is worse
> than the hole, and `a34-extlst` is the deck that says why: an unknown `a:ext/@uri` is carried
> through untouched, so a fabricated GUID would open, render its raster and survive every resave
> with PowerPoint treating the extension as somebody else's. Nothing here would ever contradict it,
> and this rule would read green.

Closing it is experiment **E4** and the Tier A slot `a28-model3d`, and E4 is blocked on a download
nobody has approved. Whoever runs it deletes the `uncovered` entry in the same commit, because
`C019` fails if they do not.

**Fourteen of the forty-six rest on a single deck each** — `ink` and `contentPart` on `a27-ink`,
`audio`, `video` and `media` on `a24-media`, and so on. That is what a probe corpus looks like when
it is built one feature at a time, and it is not a violation of anything; `coverage.test.ts` pins
the list so that "46 of 47" cannot read as more breadth than it is.

## What `C-LEX` says

The rule: **every lexical convention is exercised by at least two decks from at least two distinct
serializers, or carries a written gap saying why not.** Sixty forms declared, and the answer splits
cleanly down the middle of the package:

| layer                    | forms | covered by two producers |
| ------------------------ | ----- | ------------------------ |
| the ZIP container        | 24    | **17**                   |
| the XML inside the parts | 36    | **12**                   |

Three container writers and only two XML serializers, and that ceiling is the whole story. It is
enforced rather than assumed: `C018` makes each collection name what wrote each of its two layers,
and `corpus/written` names PowerPoint for its XML because `c01-opc-writer`'s parts _are_
`b01-blank`'s parts. Two collections that share a serializer collide to one producer, so a rule
counting distinct **files** cannot mistake a copy for evidence.

**Every entity reference in this corpus was written by `tools/corpus/gen`.** The nine
PowerPoint-authored decks contain no ampersand at all across their 387 XML parts, so `&amp;`,
`&lt;`, `&gt;` and `&quot;` are each checked only against the escaper that wrote them. E8 measured
PowerPoint writing all four; the decks it measured were never committed. This is the cheapest gap in
the table to close and the only large one that a deck can close at all.

**Most of the rest cannot be closed by authoring anything.** ADR 0004 opened 62 lexical variants in
the installed PowerPoint and re-saved the ones it accepted: comments and processing instructions are
discarded, a CDATA section is rewritten as plain text, `&#72;` resolves to `H`, a single-quoted
attribute is rewritten double, a BOM is stripped. For those forms our second producer is not silent
but incapable, and the rule says so in each row rather than implying a deck is owed.

The scanner reads bytes and never calls `@pptx-studio/xml`. Asking the tokenizer which forms a file
contains would agree with the tokenizer by construction, and `C-LEX` exists to be an independent
reading.

Two things it found that nothing else had:

- **PowerPoint does not always write an XML declaration.** Four of the corpus's 1422 XML parts have
  none — `ppt/charts/style1.xml`, `style2.xml`, `colors1.xml` and `colors2.xml` in `b05-chart`
  begin at their root element. ADR 0004 measured 2834 real parts and E8 measured 38, and no part in
  either lacked one.
- **The majority real-world form of three conventions is absent here entirely.** 1037 of ADR 0004's
  2834 parts carry a BOM and none of ours do; 2161 of them put nothing between the declaration and
  the root element and all 1418 of ours put a CRLF there; two of the three declarations it counted
  spell the encoding `utf-8` and neither of our producers does.

### What `C-LEX` does not reach

**Entry order is not a convention it can check.** The two E8 decks and `b01` disagree about whether
`slide1.xml` precedes its own `.rels`, and `b05` interleaves layouts with their rels in no pattern
at all. Order is unstable _within_ one producer, so there is nothing to assert.

**No committed deck carries a part this project serialized.** `packages/opc` has two XML
serializers of its own — `ContentTypes.serialize()` and `Relationships.serialize()` — and a no-op
write invokes neither, because passthrough means nothing is dirty. Their lexical form is pinned by
`flat-xml.test.ts` (Office's exact declaration, `standalone="yes"`, CRLF, no BOM, tight `/>`,
`&quot;` for a quote), so it is measured but unrepresented. Closing it needs a Tier C deck that
edits something — sub-phase 1.3's export, not 1.1's corpus.
