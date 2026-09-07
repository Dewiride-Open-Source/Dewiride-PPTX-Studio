# 0009 — The corpus, and where a fixture is allowed to come from

Date: 2026-08-27
Status: **accepted** — provenance, the gate, all three producers and all four named rules are
settled. 51 decks built; the fifty-second, `a28-model3d`, is cut and its census key declared
Sub-phase: 1.1 — Corpus

---

## Context

Phase 1's whole argument is that a preserving package architecture round-trips a
real deck byte-for-byte, and every one of its deliverables reads the same input:
1.2's validator, 1.3's writer, 1.4's `cli roundtrip`, 1.5's `bisect`, and 1.6's
badge. Sub-phase 0.5's already-passed gate is stated against it too — "parse →
serialize → byte-identical for 100% of parts across the whole corpus".

So the corpus is the oracle for the load-bearing bet of the project, and the
question of what is in it is not a cataloguing exercise. Two problems have to be
answered at once:

- **Licensing.** Real decks carry third-party copyright and, routinely, licensed
  embedded fonts. `LEGAL.md` already restricts sources; nothing enforced it.
- **Circularity.** If every fixture is written by our own generator, then 1.4
  tests our reader against our own writer. That proves the two agree. It does
  not prove either is right, and it systematically cannot catch a lexical
  convention we do not happen to emit.

---

## Decisions

### The corpus is tiered by _producer_, not by feature

Feature coverage was never the binding constraint — it is easy to synthesize a
deck containing any element we like. Lexical coverage is the constraint, and it
is a property of who serialized the file:

| Tier  | Producer                              | Committed             | Why it exists                                                         |
| ----- | ------------------------------------- | --------------------- | --------------------------------------------------------------------- |
| **A** | `tools/corpus/tiers/a-generated`      | bytes + recipe + hash | Breadth. Every feature, on demand, deterministically.                 |
| **B** | Microsoft PowerPoint 365 (16.0.20326) | bytes only            | The byte conventions. This is the tier that makes 1.4 mean something. |
| **C** | `packages/opc`'s own writer           | bytes                 | Gives 1.4 a case against the writer this project ships.               |
| **D** | (`corpus/bench/`, sub-phase 0.8)      | recipe + hash         | Too large to commit. Not counted toward the 50.                       |
| **E** | (`corpus/local/`)                     | hash only             | Files that never enter the repository. See below.                     |

Tier B carries **no recipe**: PowerPoint's output is not byte-stable across
monthly builds, so a recipe that claimed to reproduce it would be a lie with a
shelf life.

### Committed bytes, not recipes, for anything under the size cap

A recipe's hash is the hash of a _build_. `tools/bench/zip-stream.ts` deflates
through `node:zlib` and `packages/opc` through `fflate`, whose catalog pin is a
caret range. A Node upgrade or a routine `pnpm update` would invalidate every
pinned hash with no source change — and that is the one failure a legal audit
cannot absorb, because `license: "CC-BY-4.0"` would then describe bytes nobody
reviewed. Committing the bytes makes the hash a hash of the reviewed object.

The recipes are kept as well, as a regeneration test. Bytes-only stops testing
the generator; recipe-only lets a generator bug move the fixture silently.

Caps: **512 KiB per committed file, 12 MiB total**, both enforced. Anything
larger is `storage: "generated"` — the recipe and the hash, not the bytes.

### Three storages, so that "not committed" stops being prose

`corpus/bench/manifest.json` used to carry `"path": "(not committed - 188.5 MiB)"`,
which is a sentence in a field that is supposed to be machine-checkable. Now:

- `committed` — carries `path`; the bytes are here and hash to `sha256`.
- `generated` — carries `outputName` and `recipe`, and **may not** carry `path`.
- `pinned` — carries neither, and nothing in the repository can locate the file.

### Three manifests discovered, not one root file

`LEGAL.md` said "every entry in `corpus/manifest.json`", and no such file
existed — there were two, with genuinely different envelopes (`bench` records a
generator; `ground-truth` records the PowerPoint build it was measured on).
Flattening them would force the union of three envelopes onto every entry and
would require editing two accepted ADRs that cite the paths by name. A roll-call
of `includes` is worse still: deleting one line silently uncovers a directory.

Discovery plus `C008-orphan` is strictly stronger with less code, so the checker
reads every `corpus/**/manifest.json` and `LEGAL.md` was corrected. Entry ids are
unique **across all** manifests, which is what lets one collection reference
another's entry rather than duplicate it.

### `features` is a map of census keys to counts, not an array of labels

Both existing manifests used labels that were not census keys — `smartart`,
`custom-shows`, `embedded-fonts` against `smartArt`, `customShow`,
`embeddedFont` — so 1.6's badge could not have been computed from the manifest
as it stood. Worse, `smartart` conflated two keys the census reports separately
(`smartArt` from `dgm:relIds`, `smartArtDrawing` from `dsp:drawing`).

A map, because **`0` is meaningful**: it is the fixture stating that a feature
was considered and left out, and an array cannot say that. Both producers
already emit maps.

Non-package fixtures — `color-transforms.json`, the probe fonts — can never
carry census keys, so they carry free-form `tags` instead, and an entry may have
one or the other and never both.

### Sub-phase 1.1's stated verification, and the rule that actually matters

The plan's verification is "CI fails on any entry lacking a license". That is
`C001`, and on its own it is satisfied by typing six characters into a field.
The rule that does the legal work is **`C003-source-licence`**: the licence must
be one that the declared _source_ may claim, so a CC-BY deck somebody else wrote
is not relicensable as Apache-2.0 whatever the field says.

The rule that fires in real life is **`C008-orphan`**: every file under
`corpus/` must be claimed by exactly one manifest. Nobody adds a manifest entry
and forgets the licence; what happens is a deck is dropped into the tree and no
manifest is touched at all. It enumerates with `readdirSync` rather than
`git ls-files` for three reasons — it is far faster, it works in a source
tarball with no `.git`, and decisively it sees **untracked** files, so a local
`pnpm check` catches the deck before it is ever committed.

The full table is `tools/corpus/README.md`; the enums are
`tools/corpus/manifest/schema.ts`.

### The gate runs second, before formatting

`pnpm layering && pnpm corpus && pnpm format:check && …`. It needs no build, no
workspace package and no browser. Ordering a legal gate behind a whitespace gate
would mean a run that stops on formatting never reports that a fixture has no
licence.

It is deliberately **not** a turbo task. If it ever became one the hazard is
precise: turbo's default inputs cover a package's git-tracked files, and the
moment anyone writes an `inputs` array to tweak one thing they opt out of that
default — `corpus/**` stops being hashed and a deck added without a manifest
entry replays a cached PASS. Turbo also hashes tracked files only, so an
untracked deck is invisible to it however inputs are declared.

---

## The four provenance questions, and the answers

These were put to the owner before any deck was built, because three of them are
licensing or permission judgements and not engineering ones.

### Tier B: PowerPoint-authored decks may be committed

**Approved.** PowerPoint 365 is scripted over COM, reading and writing only
inside the working tree and the session scratchpad, and the resulting decks are
committed. Every authoring script lives in `tools/corpus/authored/` so the
provenance is auditable line by line.

A Blank Presentation contains Microsoft-generated `ppt/theme/theme1.xml` and
eleven default layouts. `NOTICE` discloses this. Four constraints are enforced
rather than trusted:

- Blank Presentation only — **never** `File > New` from a shipped design
  template (Ion, Berlin, Facet…), whose theme is the part that is genuinely
  Microsoft's creative work.
- No `ppt/fonts/` part in any committed deck except the one CC0 probe font.
- No Designer suggestion and no stock imagery. Unlike the scaffolding, those
  would be actual infringement.
- No author name, company or machine name in `docProps`, scrubbed at author
  time via `BuiltInDocumentProperties` — never by editing the saved file, which
  would destroy the exact Microsoft serialization the deck exists to capture.

### Tier E: the 37 local decks are hash-pinned, never copied

**Approved.** Sub-phases 0.4, 0.5 and 0.6 were measured against 37 decks holding
2,834 parts, and all three ADRs close by noting the gate "is not automated in
the repo… it lives in a scratchpad script pointed at decks outside the working
tree". `corpus/local/manifest.json` records `sha256`, byte count and per-deck
element counts — `storage: "pinned"`, so no `path` and no `outputName`, and
nothing in the repository can locate the files.

`pnpm corpus:verify --dir <path>` takes the directory **from the operator**. The
repository never enumerates it. This is the `eot-headers.json` precedent: a
measurement about files is not a copy of them.

### A macro-enabled deck is committed, on its own path

**Approved.** A `.pptm` with an empty VBA project whose only body is
`Sub Noop()`, at `corpus/decks/*.pptm` so that anyone who must exclude it can do
so with a single glob. There is no other way to cover the `macros` census key: a
macro-enabled deck is a `vbaProject.bin` part and nothing in any markup says so.

### Tier C is our own writer, and nothing is downloaded

**Approved as the minimum.** One deck written by `packages/opc`'s own writer,
tagged as such. LibreOffice, Google Slides and third-party CC0/government decks
were all offered and declined, so **nothing is fetched from the network** and the
corpus is entirely produced on this machine by PowerPoint, our generator, or our
writer.

The cost is recorded honestly: with only two producers of foreign-to-us bytes —
PowerPoint and, marginally, our own `opc` writer — the lexical coverage rule
(`C-LEX`) rests almost entirely on Tier B. If Tier B is ever removed, the
round-trip gate degrades to a proof of idempotence.

**Built, and "marginally" turned out to be the precise word.** `c01-opc-writer`
is a third producer of ZIP headers and not of XML at all — see _Tier C, and what
a third producer is and is not_ below.

---

## Fixed while doing this

Four latent bugs in `tools/bench/deck.ts`, none reachable from the three
benchmark recipes and all reachable from the corpus's small and degenerate ones:

- `expectedFeatures` read `section` and `customShow` counts back off the recipe
  while `writeDeck` clamped both to the slide count, so a one-slide deck asking
  for six sections declared six and wrote one. Every count now comes from where
  the markup was written, and the function no longer takes the recipe at all —
  the input whose use caused the bug.
- A recipe with no free shapes allocated a hyperlink relationship and attached
  it to nothing. `danglingRelationships()` cannot see it either, because an
  External relationship has no target part to be missing.
- Comment cadence was hard-coded at every fifteenth slide, so the smallest deck
  that could carry a comment was fifteen slides long.
- `bench.test.ts`'s "finds every feature the recipe put in, **and no others**"
  only ever checked one direction. Ten keys the generator emits — `shape`,
  `placeholder`, `presetGeom` among them — were declared by nobody and checked
  by nobody. The declaration now names all 47 census keys, zeros included, and a
  second test fails if the census grows a rule the generator has no opinion
  about.

And one in the repository itself: `.gitattributes` marked `ttf`, `otf`, `woff`
and `fntdata` as binary but not `eot`, so sub-phase 0.7's two committed EOT
fixtures were subject to `* text=auto eol=lf`. `C016-gitattributes` found it on
the checker's first run.

---

## Verification

- `pnpm check` green end to end: layering, corpus, format, lint, typecheck,
  build, package QA, and **1055 tests across 33 files**.
- `pnpm corpus` reports `4 manifest(s), 60 entrie(s), 57 file(s), 2.1 MiB, no
violations`, and exits 1 with a legible message when a licence is removed, a
  source and licence disagree, an attribution block is missing, or an unclaimed
  file appears under `corpus/`.
- All three benchmark decks still rebuild to the SHA-256 pinned in
  `corpus/bench/manifest.json` — `small`, `xml-heavy` and the 188.5 MiB
  `media-200mb`, byte for byte. `tools/ground-truth/lib/zip.ts` grew a `store`
  option for the corpus and the benchmark generator uses that writer for its
  embedded workbooks, so this is a check on that change and not a formality.
- 49 tests, one per rule, against synthetic manifests: `checkCorpus` takes the
  file table rather than reading the disk, so `C008` and `C010` are tested
  against a corpus shape this repository does not have yet.
- All forty-one Tier A decks open in PowerPoint 16.0.20326 with **no repair
  prompt** — `a01` `[1]`, `a02` `[2,9,5]`, `a03` `[18,10,55]`,
  `a04` `[11,5,2]`, `a05` `[7,3,25]`, `a06` `[16,20,9]`, `a07` `[2,4,7]`,
  `a08` `[2,4,3]`, `a09` `[3,3,5]`, `a10` `[2,5,7]`, `a11` `[13,7,5]`,
  `a12` `[9,9,9]`, `a13` `[2,2,2,2]`, `a14` `[1,1,1]`, `a15` `[2,2,2]`,
  `a16` `[2,2,2]`, `a17` `[4,4,4]`, `a18` `[5,3,3]`, `a19` `[7,6,7]`,
  `a20` `[2,2,3]`, `a21` `[2,2,3]`, `a22` `[3,3,4]`, `a23` `[2,2,2]`,
  `a24` `[3,4,2]`, `a25` `[4,4,3]`, `a26` `[3,3,3]`, `a27` `[2,2,2]`,
  `a29` `[4,4,2]`, `a30` `[3,4,2]`, `a31` `[5,3,4]`, `a32` `[5,5,5]`,
  `a33` `[5,5,5]`, `a34` `[5,5,5]`, `a35` `[7,5,5]`, `a36` `[5,5,5]`,
  `a37` `[5,5,5]`, `a38` `[8,6,0]`, `a39` `[5,5,5]`, `a40` `[9,5,5]`,
  `a41` `[5,4,3]`, `a42` `[5,3,3]`. The count matters more
  than the open: a refusal is loud, a silent repair is not.
- `a38`'s third slide reports **zero**, which is correct: it is a deliberately
  empty `p:spTree` on a Blank layout. It is the only deck in the corpus with a
  slide that has nothing on it, and the only way to find out whether a shape
  count of zero is a probe or a repair is to have written one on purpose.
- **Two of those counts are not what the generator wrote, and both are
  explained rather than waved through.** `a27-ink` wrote three shapes a slide
  and PowerPoint reports two: the `p:contentPart` is not enumerated. `a30-vml`
  wrote three on its second slide and PowerPoint reports four: the extra is
  `probe_group`, an `msoGroup`, which is the `v:group` from the VML part
  materialised as a native shape. One is content PowerPoint does not surface and
  the other is content it invented, and a shape count alone cannot tell either
  from a repair — which is what the next bullet is for.
- **A shape count cannot see the Structure group**, so the eight decks that
  carry masters, notes, comments, sections, transitions and timings were read
  back through PowerPoint's object model as well. `a12` reports
  `Designs.Count = 3` with 2, 1 and 1 layouts and the third design named
  separately; `a13` reports `SectionProperties.Count = 3` and
  `NamedSlideShows.Count = 3`; `a14`'s three notes pages report 3, 6 and 2
  shapes, matching the three, six and two placeholders written; `a15` reports
  `Comments.Count` of 2, 2, 0 — the second slide's two being one classic and
  one 2018-format comment, so both formats are read; `a16`'s eighteen layouts
  each report a distinct non-zero `EntryEffect` and the Title Only layout
  reports the master's, which is inheritance working; `a17` reports a main
  sequence on every slide; `a18` reports 720 x 540 points, which is 4:3.
- **A third pass, because a count sees only what PowerPoint enumerates.** Each
  of the eleven hard-content decks was copied into the scratchpad, opened,
  `SaveAs`-ed, and the copy diffed against the original part by part and
  element by element. `a20`, `a24`, `a29` and `a31` come back with every
  counted element intact — tables, the media double relationship, `m:oMath`,
  `p:embeddedFontLst`. The other seven do not, and every difference is now a
  recorded finding rather than an unknown: `p:contentPart` 3 → 0 and its
  `mc:Fallback` picture with it; `v:shape` 4 → 0 with a `p:grpSp` in its
  place; `p:oleObj` 4 → 6 as the transitional form is upgraded to the switch;
  `dsp:drawing` 2 → 3 as PowerPoint lays out the diagram that had none; and
  `colors1.xml`, `style1.xml`, a PNG fallback and a second SVG dropped as
  parts PowerPoint did not write itself. None of it is a defect in a deck. All
  of it is a list of what **our** writer must not do.
- `C-CENSUS`: each deck's `features` map is a reviewed literal in its own
  module, and the census reproduces it exactly in both directions.
- `C-REGEN`: `node tools/corpus/tiers/a-generated/build-probes.ts --check --out corpus/decks`
  reports all forty-one rebuilding byte-for-byte, and a test asserts the same
  against the manifest's `sha256`, `bytes` and `features`. The five decks that
  predate the layout and text work reproduce their **original** committed
  hashes, so the chassis grew six options without moving a single byte of what
  was already reviewed; the eleven that predate the Structure work reproduce
  theirs through a second round of chassis changes; and all nineteen reproduce
  theirs through a **third**, which added `a:hlinkClick` to `cNvPrXml`, a font
  scheme to the theme, attributes to `p:presentation`, and exported `relsXml`
  and `REL` for the decks that bring parts of their own. All thirty then
  reproduce theirs through a **fourth**, which added `mc:Ignorable` on a part
  root, explicit `p:sldId` and sheet ids, a macro-enabled content type,
  package-level relationships and per-entry ZIP shape — and, because
  `a35-zip-shapes` needed it, changed the archive writer itself.
- **`a35-zip-shapes` is the one deck this claim is weaker for.** It deflates,
  so its committed hash is the hash of a build and a `node:zlib` change can
  move it with no source change. That is the cost of having any DEFLATE in the
  corpus at all, it is confined to the one deck that is _about_ compression,
  and the recovery is to check that nothing but the deflate streams moved and
  re-pin. A test asserts the deck really does deflate rather than quietly
  storing, so the exception cannot become vacuous.
- **The `.pptm` reaches further than the byte comparison.** `outputName()`
  feeds the writer, the `--check` reader, the manifest's `path`, the manifest's
  `format` and `C008-orphan`, and `format` is validated against nothing — so a
  deck whose `format` and `path` disagreed would pass every rule silently.
  Both now derive from one `ProbeDeck.extension`, which is the only reason they
  cannot. `.gitattributes` already carried `*.pptm binary`, so
  `C016-gitattributes` was satisfied before the deck existed.
- **`cfb.test.ts`, twelve assertions, for the two layers nothing else sees.**
  `a32-macros`'s `ppt/vbaProject.bin` is a compound file holding compressed
  streams; the census classifies it as binary and does not read it, `C-REGEN`
  compares one hash, and PowerPoint does not load a VBA project when it opens a
  file. So a broken sector chain or a compressor that cannot decompress its own
  output would pass every other gate. `readCfb` follows the FAT, the MiniFAT
  and the directory's sibling tree rather than running the writer backwards;
  `decompressOvba` was written from the format rather than from
  `compressOvba`; and the `dir` record walk has to land exactly on the end of
  the stream, which is what catches the records whose four-byte field is a
  reserved constant rather than a size.
- `tools/corpus/suites/conventions.test.ts`: 21 assertions across all three
  PresentationML builders in the repository, every one of them read out of
  `powerpoint-conventions.json` rather than chosen.

---

## What the first two probes measured

Recorded as `corpus/ground-truth/powerpoint-conventions.json`, produced by
`tools/corpus/tiers/b-authored/probe-conventions.ps1` and its analyser. PowerPoint
16.0.20326.20100, two decks, neither committed.

### E6 — OLE, and it is not what either report predicted

Build 2205 stopped writing the VML fallback for OLE objects. It did **not**
stop writing the wrapper. What 20326 writes for an embedded workbook is:

```xml
<a:graphicData uri="…/presentationml/2006/ole">
  <mc:AlternateContent xmlns:mc="…">
    <mc:Choice xmlns:v="urn:schemas-microsoft-com:vml" Requires="v">
      <p:oleObj name="Worksheet" r:id="rId2" imgW="5186562" imgH="3805404"
                progId="Excel.Sheet.12"><p:embed/></p:oleObj>
    </mc:Choice>
    <mc:Fallback>
      <p:oleObj …><p:embed/><p:pic>… <a:blip r:embed="rId3"/> …</p:pic></p:oleObj>
    </mc:Fallback>
  </mc:AlternateContent>
</a:graphicData>
```

So, against the research:

- The `mc:Choice` requires **`v`**, not `p14`, and declares `xmlns:v` **on
  itself** — both as one report argued and the brief's context had wrong.
- There is **no `vmlDrawing` part and no `@spid`**, even though the Choice
  still requires the VML prefix. A reader that resolves `Requires="v"` and then
  looks for a VML part finds nothing, and that is the file being correct.
- The payload is **`ppt/embeddings/Microsoft_Excel_Worksheet.xlsx`** through a
  `…/relationships/package` relationship, with a `Default Extension="xlsx"`.
  Route (b) is what PowerPoint itself does; hand-writing a CFB was never
  necessary.
- The preview is an **EMF**, not a PNG, and lives only in the `mc:Fallback`.
- PowerPoint writes a `docProps/thumbnail.jpeg`.

`a26-ole` therefore emits this shape for `oleObject`, and a _separate_,
deliberately transitional frame with a `vmlDrawing` part for the `vml` key.
Because the census counts raw, pre-MCE markup, one logical OLE object here
reports `oleObject` **2** — once in the Choice and once in the Fallback. That
is the right policy for a census and it has to be what `expectedFeatures`
counts too.

### E8 — the byte conventions, and one that changes a rule

Across all 38 XML parts of a PowerPoint-saved deck: **no BOM**; one declaration
form, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`, followed by
**CRLF**; no part broken across lines; **1305 self-closing tags, every one
written `<x/>`, none spaced**; no single-quoted attribute; and exactly three
entity references — `&amp;`, `&lt;`, `&gt;`. It escapes `>` in character data
although XML does not require it, and does not escape `"` or `'` there.

Our generator already matches all of it.

**The rule that changes.** The appendix says "`xml:space="preserve"` on any
`a:t` with leading/trailing whitespace". PowerPoint does **not** write it: a run
authored as `"  leading and trailing spaces  "` is stored with the spaces and
no attribute. That is WordprocessingML's rule, where a `w:t` does carry one; in
DrawingML the whitespace is simply significant. Our writer must not add it —
adding it would be synthesizing markup we did not read, which is the one thing
the architecture forbids.

**The ZIP conventions** are also worth having: general-purpose flag `0x0006`
(deflate level hints, **not** bit 11, so no UTF-8 flag), DOS timestamp
`0x0000`/`0x0021` — 1980-01-01, the same zeroed value our generator writes — no
directory entries, and a **520-byte `0xa220` extra field** on the first local
header. That is the Open Packaging _growth hint_, which reserves padding so a
part can grow without the archive being rewritten. No synthetic deck of ours
produces one, and `a35-zip-shapes` now has to.

### The finding that changes the roster

ADR 0005 measured the real-world corpus: **70 822 of 98 777 self-closing tags
are written `<a:off … />`**, spaced. This PowerPoint build writes **zero**
spaced out of 1305.

So the dominant real-world form is produced by neither PowerPoint 20326 nor by
us, and Tier B will never exercise it. Since Tier C is now a single deck from
our own writer and nothing is downloaded, **no producer available to this
project emits the spaced form at all.** The corpus must therefore contain a
deck whose self-closing tags are deliberately spaced, generated by an explicit
option rather than acquired — and `C-LEX` has to check that lexical form
against a list of forms known to exist in the wild, not merely against "two
distinct producers", because two producers that agree prove nothing.

This is exactly the gap the lexical rule was introduced to find, and it was
found on the first measurement rather than in Phase 1's round-trip failures.

---

---

## The generator, and the first five decks

### A third builder, and a test instead of shared code

Forty feature probes needed a home and neither existing builder was one.
`tools/bench/deck.ts` is a recipe interpreter streaming to a file descriptor,
built for three hundred slides and two hundred megabytes; its vocabulary is
cadences, and expressing forty probes in it would have meant roughly sixty more
`DeckRecipe` fields and a `writeDeck` nobody can read — after which the probes
could still only say what the loop already knows how to say.
`tools/ground-truth/lib/pptx.ts` is closer in shape but is a sub-phase 0.7 artifact
whose output two committed measurements were taken from, so changing what it
emits would make `color-transforms.json` and `eot-headers.json` describe a deck
that no longer exists.

So `tools/corpus/tiers/a-generated/markup/chassis.ts` is written fresh and is canonical from here on.
The obvious hazard in three builders is three divergent notions of correct
markup, and **sharing code is the weak answer to it** — it makes them agree
without making any of them right, and does nothing about a fourth written next
year. `tools/corpus/suites/conventions.test.ts` is the strong answer: it reads
`powerpoint-conventions.json` and asserts every producer in the repository
matches what PowerPoint measurably wrote — no BOM, one declaration form followed
by CRLF, no part broken across lines, no single-quoted attribute, only the three
entity references, no `xml:space` on an `a:t`, `[Content_Types].xml` first, no
directory entries. Twenty-one assertions across three producers, and it holds
for the fourth automatically. That converts E8 from a document into a contract.

### Probe decks are stored, not deflated

Committing the bytes fixed the legal half of the "a recipe's hash is the hash of
a build" problem — the hash is of the reviewed object — but not `C-REGEN`, which
rebuilds each deck and compares. Deflate would put zlib's version into that
comparison forever, so a Node upgrade would redden the gate with no source
change.

Storing removes it: these bytes are a pure function of our own XML. The decks
are 18–83 KiB against a 512 KiB per-file cap, so nothing is spent.

The cost is real and is not left to be an accident: no Tier A deck then
exercises DEFLATE. `a35-zip-shapes` is the deck that does, deliberately — mixed
methods, the general-purpose flag bits, the 520-byte `0xa220` growth hint — and
is the only corpus deck whose bytes depend on zlib. Putting the exception on the
deck that is _about_ compression keeps the dependency somewhere it can be
reasoned about.

### `features` is a literal, not a tally

`tools/bench/deck.ts` counts what it emitted and calls that the expectation,
which is right for a deck whose contents a recipe decides. For a probe it would
be circular: the deck exists to state what markup is in it, and a map derived
from the markup restates the markup instead of checking it. So each deck module
carries a written map, reviewed, and `probes.test.ts` demands the census — an
independent token scanner — reproduce it **exactly**. Every deck's markup change
fails that test until a person looks at what moved.

### Fixed while doing this

`collection` was one of `corpus | bench | ground-truth | local`, where the last
three are directory names and the first was a placeholder nothing used. It is
now `decks`, and the checker asserts a collection is named after the directory
it lives in — without which two manifests could both call themselves `bench` and
every message about either would name the wrong directory.

---

## Layouts, text, and two corrections to the plan

The next six decks — `a02`, and `a07` through `a11` — are the text and
inheritance batch, and building them grew the chassis in six places rather than
one: a package may now declare its own `layouts`, replace the master's two fixed
placeholders, append shapes to the master, carry `p:hf` on the master or on a
layout, replace `p:txStyles`, and add `a:objectDefaults` to the theme. Every one
of those is an option with a default, and the five decks that predate them
rebuild to their original hashes unchanged, which is the only evidence worth
having that a refactor of a fixture generator was safe.

A seventh file, `tools/corpus/tiers/a-generated/markup/text.ts`, holds the paragraph and run builders.
It exists for the same reason `shapes.ts` does: six decks needed every child of
`a:pPr` and `a:rPr`, the order of those children is what PowerPoint refuses
over, and writing it inline six times is six chances to get it wrong. The orders
are copied out of `packages/xml/src/edit/schema-order.gen.ts`, so the fixtures and
the runtime's `insertInOrder()` read the same generated table.

### Two places the plan is wrong, and the decks say so

**`a:fld` has fifteen reserved types, not fourteen.** ECMA-376 §21.1.2.2.4 lists
`slidenum`, `datetime`, and `datetime1` through `datetime13`. `a09-fields`
carries all fifteen, each with a literal `ST_Guid` and cached text for one fixed
instant — the one ECMA's own examples use — so the table is checkable against the
specification line by line rather than against taste. It also carries
`datetimeFigureOut`, which is in no specification and which PowerPoint writes
whenever a date placeholder updates automatically. `@type` is `xsd:string` with
no enumeration behind it, so a renderer that switches on a closed set of fifteen
shows nothing where PowerPoint shows a date; the general rule is to fall back to
the cached `a:t`, and a real unreserved value in the corpus is what makes that
rule testable instead of aspirational.

**`ST_PlaceholderType` has sixteen values and PowerPoint accepts fourteen.** See
the refusal section below.

### What `a11` is deliberately not authoritative about

It carries twelve stored `normAutofit` scales from 100% down to 25%, and it is
**not** the source of truth for PowerPoint's autofit ladder. Sub-phase 3.4
measures those steps from PowerPoint's own output for given content, and a
fixture that guessed them would turn that measurement into a tautology — the
same failure mode as a `features` map tallied at the emit site. What `a11` is
authoritative about is the property a viewer must have: stored values are
applied verbatim, so a PowerPoint-authored file pixel-matches, and a viewer that
recomputes instead is visibly wrong in eleven of the twelve boxes.

---

## Structure, and four more corrections

The eight decks `a12` through `a19` are the parts of a deck that are not a
slide: masters, notes, comments, sections, transitions, timings, page size and
the accessibility surface. Building them cost the chassis five more options —
`masters`, `notesMaster`, `handoutMaster`, `notesWidth`/`notesHeight` with
`slideSizeType`, and a `transition` on both master and layout — plus
`p:cNvPr`'s `@descr`, `@title`, `@hidden` and `a:extLst`, routed through one
`cNvPrXml` so every shape kind grew them at once.

It also cost a **second round of ground truth**. Five of the eight need markup
that no specification writes down, so on 2026-08-27 two throwaway decks were
authored through PowerPoint COM in the session scratchpad and read back. Nothing
was downloaded and nothing outside the scratchpad was touched; the decks are not
committed and no byte of them is.

What that measurement settled:

| question                                         | answer                                                   |
| ------------------------------------------------ | -------------------------------------------------------- |
| the `a:ext/@uri` carrying `adec:decorative`      | `{C183D7F6-B498-43B3-948B-1728B52AA6E4}`                 |
| does PowerPoint write `&quot;`                   | **yes**, in an attribute, and nowhere else               |
| the notes master's six placeholders and geometry | `hdr`, `dt` 1, `sldImg` 2, `body` 3, `ftr` 4, `sldNum` 5 |
| the handout master's four                        | `hdr`, `dt` 1, `ftr` 2, `sldNum` 3                       |
| what a comment looks like in build 16.0.20326    | `p188:cmLst`, not `p:cmLst`                              |
| what a transition looks like                     | an `mc:AlternateContent` pair, not a bare `p:transition` |

Four of those change what this repository says or does.

**`escapeAttribute` escapes rather than throws.** The text group left it
throwing on a double quote on the ground that `powerpoint-conventions.json`
recorded three entities and the measured deck's silence about a fourth was not
evidence. It still is not evidence — but a shape renamed to `a "quoted" name`
and saved produced `name="a &quot;quoted&quot; name"`, two occurrences and no
other entity in the package, which is. The fixture's `entitiesUsed` gains
`&quot;` with a comment naming the second measurement, and `a19-decorative`
puts a quote in an `@descr` — the one attribute in a deck that holds text a
user typed.

**The census undercounts comments.** Its `comment` rule matches the content
type ending `presentationml.comments+xml`; PowerPoint 365 writes
`application/vnd.ms-powerpoint.comments+xml` with a `p188:cmLst` root, an
`ppt/authors.xml` beside it and a `p188:commentRel` in the slide's own
`p:extLst`. So a deck commented in PowerPoint today censuses as having no
comments. `a15-comments` carries **both** formats — its `features` map says
`comment: 2` and those two are the classic parts — so the gap is a thing the
corpus states. Whether `packages/census` grows a rule for the 2018 form is a
decision for its own sitting, not a silent widening of this one.

**`p:transition` has twenty-one effects, not the plan's twenty-two**, and
`a16` needs twenty-two sheets to carry them because a transition is legal on a
master and a layout as well as a slide. That turned out to be the deck's most
useful property rather than a workaround: a renderer that reads transitions off
slides alone finds four of the twenty-three in it.

**One package holds one `p:sldSz`**, so the roster's "4:3, 16:9, A4 and a
custom size" is four decks. `a18` is 4:3 — the size sub-phase 7.11's reflow
work is actually about — and `a41-a4` and `a42-custom-size` join the roster,
taking Tier A's target from forty to forty-two — and back to forty-one when `a28` was cut, for
the reason the last section gives.

---

## Hard content, and the loop earning its keep

Eleven decks — `a20` through `a31`, less the cut `a28` — and the group is
worth an ADR section for a reason the first three groups were not: **five of the
eleven were refused by PowerPoint the first time they were built**, and every
one of the five turned into a rule that is now written down.

That is the sub-phase 1.5 loop doing exactly what it was specified to do, three
phases before the sub-phase exists. The procedure each time: build one package
per slide, then one package per feature, then one package per attribute, until
the smallest thing that reproduces the refusal is in hand. Twenty-nine
one-variable packages across five rounds.

### Four of the eleven were measured, not read

`a21`, `a22`, `a23` and `a24` were written from decks PowerPoint 16.0.20326
authored on 2026-08-27 through `Shapes.AddChart2`, `Shapes.AddSmartArt`,
`Shapes.AddPicture` on an SVG the corpus generator wrote, and
`Shapes.AddMediaObject2` on a WAV it wrote. Nothing was downloaded, everything
happened inside the session scratchpad, and no measured deck is committed.

Two of those measurements changed a decision rather than confirming one. An SVG
inserted through COM produced an `a:blip` with **no `@r:embed` at all** — the
`asvg:svgBlip` in its extension list is the only reference to any image, and no
PNG part was written — which is the reverse of the shape sub-phase 10.7 was
specified against, so `a25` carries both. And `chartEx1.xml` wrote **12 of its
12 self-closing tags spaced**, `<cx:title … />`, while the other 3,046
self-closing tags in the same sixty parts wrote none: PowerPoint has more than
one XML serializer, which makes ADR 0005's conclusion that no producer available
to this project emits the spaced form wrong. `ROSTER.md`'s `a36-spaced-tags`
entry now says so.

### The six refusals

Every one is schema-legal markup rejected for a reason no schema states, and
every one is now in `ROSTER.md`'s **What PowerPoint refuses** with the smallest
reproduction, ready to become a `corpus/reject/` fixture in sub-phase 1.2.

| markup                                                            | verdict |
| ----------------------------------------------------------------- | ------- |
| `c:strLit` inside `c:tx`                                          | refused |
| `cs:chartStyle` with 4 of its 31 entries                          | refused |
| `cs:chartStyle` with all 31                                       | opens   |
| `cx:chartSpace` with no `.rels` — including PowerPoint's own part | refused |
| the same part with `chartStyle` and `chartColorStyle` related     | opens   |
| one `.wav` as both a media object and a transition's `p:snd`      | refused |
| two copies of the same `.wav`, one for each                       | opens   |
| `p:contentPart` reached by `…/2006/relationships/customXml`       | refused |
| the same, reached by `…/office/2010/relationships/customXml`      | opens   |
| `p:control`, in any of eight forms                                | refused |
| `<p:controls/>` with no children                                  | opens   |

Three of them deserve a sentence more.

**The ChartEx one was proved with PowerPoint's own bytes.** Seven ChartEx frames
refused identically, which pointed away from their content; `chartEx1.xml` was
then lifted byte for byte out of the deck PowerPoint had just saved, dropped
into a package with nothing else changed, and refused too. Adding back its
`chartStyle` and `chartColorStyle` relationships made that same package open.
So the rule is not about our markup at all, and a classic `c:chartSpace` — which
needs neither part — proves it is ChartEx's alone.

**The ink one shows PowerPoint validating one relationship type rather than
types in general.** With only the `@Type` changed and nothing else, the ECMA
`customXml` type refuses while `…/image` and `…/slide`, which are nonsense for
an InkML part, both open. The ECMA type means the document's custom XML data
store; handed an ink part instead, PowerPoint refuses the package rather than
the part.

**The `p:control` one is bounded but unexplained.** Eight variants, one change
each — bare, name-only, with and without `r:id`, with a `p:pic` preview, with
an ActiveX part using `persistStreamInit` and its `.bin`, and the whole package
retyped as macro-enabled — and all eight refused, while an empty `p:controls`
opens. So it is the child element and not its attributes, its target, or the
absence of a VBA project. Real decks with ActiveX controls exist and open;
what they do differently is not something this machine can find out, because
PowerPoint has no COM entry point that inserts a control. `a30-vml` reaches its
VML through `p:oleObj/@spid` instead.

### And five things it does without saying so

A deck that opens is half the gate. Re-saving each deck through PowerPoint and
diffing the result is the other half, and it found things a shape count cannot:
`p:contentPart` discarded together with its fallback picture; a `v:group`
converted into a native `p:grpSp`; a transitional `p:oleObj` upgraded to the
switch form and its embedding renamed; a missing SmartArt drawing generated;
and every part PowerPoint did not itself write — chart colours and styles, a
raster fallback, a second SVG — dropped.

That last group is the reason sub-phase 1.4 compares canonical XML and a
relationship-graph isomorphism rather than bytes, and it is now demonstrated
rather than argued: PowerPoint renames media, renumbers relationships, rewrites
its own extension parts and adds content on every save. A round-trip test that
expected byte equality would be red on the first file.

### What the group cost the chassis

Five options, and all thirty decks still rebuild byte-for-byte: `a:hlinkClick`
on `cNvPrXml`, a font scheme on the theme, extra attributes on
`p:presentation`, and `relsXml` and `REL` exported so a deck can bring parts
with relationships of their own. Four new modules author bytes nothing else in
the repository had: `emf.ts` (a five-record metafile that genuinely draws),
`xlsx.ts` (a five-part workbook, which experiment E6 made possible by
establishing that a modern OLE payload is a package and not a compound file),
`media.ts` (a real 440 Hz WAV, and an MP4 container with no tracks), and
`chart-style.ts` (the thirty-one-entry chart style two decks now need for two
different reasons).

`a31-embedded-fonts` needed no new bytes at all: sub-phase 0.7 already writes an
SFNT from scratch and wraps it in EOT, so the deck embeds a font this repository
authored, with `fsType` 0 because that is the honest answer.

---

## The awkward ones, and three findings the schema cannot express

The last eleven decks — `a32` through `a42` — complete Tier A but for `a28`.
They are the ones with nowhere else to go: the archive rather than the markup,
the lexical forms nothing here emits, the empty cases, the id edges, the two
remaining slide sizes, and the two parts that are not XML.

Three were refused on the first attempt, and all three refusals turned out to be
**arithmetic and grammar rather than content** — which is what makes them worth
the ADR rather than only the roster.

- **`p:cNvPr/@id` is read as a signed 32-bit integer.** Sweeping one id per
  package: `0 … 2147483647` open, `2147483648 … 4294967294` are whole-package
  refusals, and `4294967295` opens. `ST_DrawingElementId` is `xsd:unsignedInt`,
  so PowerPoint accepts 2,147,483,649 of its 4,294,967,296 values and rejects
  the rest — and re-saving shows why the survivor at the top survives, because
  `0xFFFFFFFF` is renumbered away while `2147483647` is kept. It is a sentinel
  it tolerates rather than an id it accepts. No schema can say that, and the
  allocator in sub-phase 5.6 has to know it.
- **A part name may not percent-escape an unreserved character.** Nine
  one-name packages: `%2D`, `%41`, `%5F` and `%7E` refused with `0x808D1005`;
  `%20`, `%23`, `%24`, `%2C` and `%3A` all open. That is RFC 3986 §6.2.2.2
  normalisation enforced as a hard error, and it is the sharpest of the three
  because **this project already has the rule and rates it a warning** —
  `packages/opc`'s `M1.8`, which `toPartName` accepts. The measurement did not
  find a missing rule; it found a wrong severity, which is a thing only a
  measurement finds.
- **A slide has exactly one `slideLayout` relationship and it resolves to a
  layout part.** Zero, two, one pointing at a master and one pointing at a
  missing part are each a refusal. The binding lives only in the rels part, so
  `CT_Slide` cannot state this and does not.

Two chassis options and two new byte-authoring modules came out of the group.
`cfb.ts` writes an OLE2 compound file — allocation tables, a mini stream, and a
directory of red-black trees whose sibling order is by name **length** first —
and `vba.ts` writes an MS-OVBA project into it. `jpeg.ts` writes a baseline
greyscale JPEG whose Huffman tables were read out of the `DHT` segments of a
thumbnail PowerPoint wrote rather than recalled from the standard, and whose
all-ones quantization table makes the encoder exact: with every divisor one, a
block of flat grey has a DC coefficient of exactly `8 × (level − 128)` and no AC
coefficients at all, so the file round-trips through a decoder with no error and
needs no DCT. GDI+ decodes it back to the levels it was given.

The re-save pass earned its keep again, and this time on the corpus's own
contract rather than on PowerPoint's behaviour. **`a34-extlst` proves the
extension rule the must-not-break list asserts**: thirteen `a:ext` children come
back in the original order, with a URI used twice still used twice and two GUIDs
differing only in case still differing only in case. It also found the limit —
**only the last child element of any one `a:ext` survives**, so a writer that
appends a second child to an existing extension loses the first. And it found
that a comment inside an extension is dropped while a CDATA section is not.

## What PowerPoint refuses, and what it accepts

Every deck is opened in PowerPoint 16.0.20326 over COM, and the shape count of
every slide is compared against what the generator wrote — because the failure
that matters is not the refusal, which is loud, but the **silent repair**, which
opens successfully with content quietly dropped. All eleven now open with every
count matching exactly. Two of them did not, the first time, and both refusals
became findings.

_This section is the first eleven decks. The Structure group added one refusal —
a `p:sldMasterId` colliding with a `p:sldLayoutId` — the Hard content group added
six more, and the awkward ones three, all of which are above. That is thirteen,
and `tools/corpus/ROSTER.md` keeps the whole list in one place, because that is
the file the next person reads when PowerPoint says nothing._

`a05-geometry` did not, the first time: "PowerPoint could not open the file", no
part name, no element name, nothing else. Bisecting it by slide and then by
shape put it on slide 2, and rebuilding that slide's arrow by hand made it open
— which located the fault in the helper rather than the markup. Two findings
came out of it, and both are sub-phase 1.2 rules that no schema states:

- **A guide name referenced but never defined is a whole-package refusal.** Not
  a repair, not a dropped shape. It applies to an `a:gdLst` formula naming an
  adjust value no `a:avLst` defines, and to an `a:pt` coordinate naming a guide
  no `a:gdLst` defines. `ST_GeomGuideName` is an unconstrained token, so this is
  schema-legal markup rejected for an unstated reason.
- **`a:ahXY` or `a:cxn` as a direct child of `a:custGeom`**, missing its
  `a:ahLst`/`a:cxnLst` wrapper, is the same refusal with the same message. That
  was the actual bug: the helper took the wrapper for four children and the
  contents for the other two, and the callers passed contents for all six.

`a02-placeholders` refused too, and its bisection was cheaper because the deck
was built expecting it: `hdr` and `sldImg` had been given their own layout and
their own slide precisely so that one step would name them. It took four —
package, then slide, then master-versus-layouts, then one placeholder type per
package — and the answer is unambiguous:

- **`p:ph type="hdr"` and `p:ph type="sldImg"` are a whole-package refusal**,
  on a slide layout and on a slide alike, either one alone with no other change.
  The other seven content types were built as one-type packages in the same
  sweep and every one opens: `obj`, `chart`, `tbl`, `clipArt`, `dgm`, `media`,
  `pic`. Nothing in the schema says this. `CT_Placeholder` is one complex type
  shared by masters, layouts, slides, notes slides and handout masters, and
  `ST_PlaceholderType` is one enumeration holding all sixteen values — the
  restriction that two of them belong to the notes and handout families only is
  real, unwritten, and enforced by refusing the file.

That finding is why `a02` covers fourteen of the sixteen types and `a14-notes`
inherits the pair, and it is a sub-phase 1.2 rule: a validator that lets an
export carry `type="hdr"` on a slide hands the user a file that will not open.

`a12-masters` refused as well, and unlike the other two the fault was entirely
ours — which is the more useful kind of finding, because it names a rule the
validator has to enforce on markup this repository will generate again:

- **A `p:sldMasterId/@id` that collides with a `p:sldLayoutId/@id` is a
  whole-package refusal**, and those two are one number space. `ST_SlideMasterId`
  and `ST_SlideLayoutId` are both "2147483648 and up"; nothing in the schema says
  they may not overlap, and PowerPoint allocates from a single running counter —
  master, its layouts, next master, its layouts — and refuses a package that does
  not. The chassis had numbered masters and layouts from two counters, which is
  indistinguishable from correct with one master and collides on the second, so
  **every** package with two or more masters was refused.

  Seven one-variable packages located it in one round: a single master through
  the same code path opened, as did an inverted `p:clrMap`, a second
  `a:clrScheme`, and `a:overrideClrMapping` on a slide. Two masters with any
  content at all did not. `slideMasterXml` now takes its layout ids from the
  caller and the caller allocates from one counter, with the reason written
  where the next person will change it.

There was also a bug of our own that the census caught before PowerPoint could:
`escapeXml` is for character data and leaves `"` alone, and `a10` and `a11` were
first written with shape names like `vert="vert270"`. The result is
`name="vert="vert270""`, malformed from that byte on. `escapeAttribute` now
guards every attribute value — and, since the Structure group's measurement,
escapes rather than throws. It threw at first, deliberately, because
`powerpoint-conventions.json` recorded three entities and the measured deck
contained no attribute with a quote in it, which is not evidence that PowerPoint
never writes a fourth. It turned out to write one: see **Structure, and four
more corrections** above. The rule the throw enforced was right and the answer
it was waiting for arrived.

And one that closes an experiment: **`a:effectDag` is accepted.** PowerPoint
never authors one — it writes `a:effectLst` for everything a user can build — but
it reads a nested `a:cont` graph without complaint and the shape survives with
its effects. So the other half of `a:spPr`'s effect choice is not a
write-only branch of the specification, and the corpus can carry it.

Accepted without complaint, all spec features PowerPoint does not itself write:
`a:hslClr`, `a:scrgbClr`, `a:prstClr`, path gradients with `a:fillToRect`,
`a:tileRect`, `a:tile/@flip`, `a:blip` colour effects, `a:grpFill`, `a:ahPolar`,
and all 54 `ST_PresetPatternVal` values.

---

## Tier B, and the difference a second producer makes

Nine decks written by PowerPoint 16.0.20326 through
`tools/corpus/tiers/b-authored/build-tier-b.ps1`, in a new `corpus/authored/` collection.
They exist because of a sentence written into this ADR before any of them
existed: with only one producer of foreign-to-us bytes, `C-LEX` rests on nothing
and the round-trip gate degrades to a proof of idempotence. A generator checked
only against itself is checked against nothing.

**They live in their own collection for a mechanical reason.**
`build-probes.ts --manifest` rewrites `corpus/decks/manifest.json` in full from
`PROBE_DECKS`, so a Tier B entry parked there would survive exactly until the
next Tier A change. `C005` requires a manifest's `collection` to equal its own
directory name, so a new directory means a new enum value; `authored` is that
value. The two tiers also make incompatible claims — a Tier A entry carries a
`recipe` that reproduces its bytes, and a Tier B entry cannot.

**There is no `C-REGEN` for Tier B and there cannot be.** PowerPoint stamps
`dcterms:created` into `docProps/core.xml` on every save, so two runs one second
apart differ. What is auditable is the authoring rather than the output, which is
why the script is committed and reads nothing outside its output directory. The
manifest generator therefore _describes what is committed_ instead of building
anything — running the authoring script again replaces the corpus rather than
refreshing it, and that asymmetry is stated at the top of both files so nobody
discovers it from a diff of nine hashes.

### What the two floors, diffed, are worth

`b01-blank` and `a01-minimal` are the same deck by two producers, and three
divergences show up there and nowhere smaller. The first is the one that
matters: **PowerPoint writes `<p:sldSz cx="12192000" cy="6858000"/>` with no
`@type` at all**, where our chassis writes `type="screen16x9"` on the identical
extent. Both are legal. The consequence is that a renderer keyed on `@type`
rather than on `cx`/`cy` passes all 41 Tier A decks and fails on every file
PowerPoint ever wrote — which is precisely the class of bug a single-producer
corpus cannot surface. `PageSetup.SlideSize` reports `ppSlideSizeCustom` for
that deck, confirming the attribute is absent rather than defaulted, and
`a42-custom-size` omits `@type` only at a portrait custom extent, so it never
covered this.

The second is `_rels/.rels` in `rId3, rId2, rId1, rId4` order, which closes a
declared gap. The third is that the thumbnail is the only **stored** entry in an
archive that deflates everything else; `b05` and `b07` extend it — the embedded
`.xlsx` is stored too, so the rule appears to be "store what is already
compressed". A fourth finding is a constraint on `C-LEX` rather than an input to
it: **entry order is not stable even within one producer.** `b01` and the two E8
decks disagree about whether `slide1.xml` precedes its own `.rels`, and `b05`
interleaves layouts with their rels in no pattern. `C-LEX` may assert that a
convention is _exercised_, not that entry order is _fixed_.

### The privacy scrub did not work the way the roster said

`ROSTER.md` prescribed scrubbing `docProps` through `BuiltInDocumentProperties`.
That collection is unreachable from PowerShell on this machine: it is non-null
and enumerates to 34 items, but `.GetType()`, `.Item('Author')`, `.Name` and
`.Value` all throw `NullReferenceException`, and `[System.__ComObject].InvokeMember`
reports `Method 'System.__ComObject.Item' not found` because the Office type
library is not bound in this host. The failure is silent — the code runs, reports
success and changes nothing — and the first build of `b01` was saved carrying the
author's name twice.

`RemoveDocumentInformation(ppRDIRemovePersonalInformation)` works, in one call,
before the save and without editing the saved file. It **empties** `dc:creator`
and `cp:lastModifiedBy` rather than removing them, so the test asserts
`<dc:creator></dc:creator>` and not absence, and it writes
`removePersonalInfo="1"` onto `p:presentation` — a cost recorded rather than
hidden.

**The check that should have caught this was itself passing vacuously.** It read
`part.data` where `ZipEntry`'s field is `part.bytes`, so it decoded `undefined`
to an empty string and found no names in nine decks. A sibling assertion failing
for the same reason is what exposed it. The test now asserts the _result_ rather
than the method, and was verified by pointing it at a string the decks really
contain and confirming all nine fail. A test that cannot fail is worse than no
test, because it is also a claim.

### Three findings from driving the object model

`PpEntryEffect` is one enumeration serving both shape entry animations and slide
transitions, and **neither half accepts all of it** — `ppEffectFlashOnceFast`
(3841) is rejected for a transition by name. `Presentations.Add` returns a
presentation with **zero** slides, so deleting "the auto-created first slide"
costs a real one, which is how `b02` first came out with ten layouts instead of
eleven. And a slide bound to a layout inherits only that layout's _content_
placeholders: every layout carries `dt`, `ftr` and `sldNum`, and none reaches a
slide unless `p:hf` turns it on.

### The fixture that is better than the one we wrote

`b08-transitions` carries the best MCE case in the corpus, and Microsoft wrote
it. PowerPoint wraps every `p:transition` in `mc:AlternateContent`; for a
post-2010 effect the `mc:Choice` holds `p14:honeycomb` while the `mc:Fallback`
holds a plain `p:fade`. **The two branches are not the same transition.** A
consumer that resolves `mc:Choice/@Requires` wrongly shows something _different_
rather than something degraded — the exact failure the MCE walker exists to
prevent, and one `a37-mce` could only simulate with synthetic switches. It is
also why `transition` counts 10 for five slides: the census scans both branches,
by design.

`b05-chart`'s second chart is the companion case — one `c:plotArea` holding a
`c:barChart` **and** a `c:lineChart`, so code that reaches for
`firstElementChild` silently drops the line overlay. Sub-phase 9.8 names that
failure; this is it, in a file we did not author.

---

## Tier C, and what a third producer is and is not

`c01-opc-writer` is `b01-blank` read through `PartStore.open` and written
straight back out with nothing touched in between. That is the whole deck, and
the smallness is the point: it is the committed evidence for architectural bet 1
at package scale. Thirty-seven entries in, thirty-seven out, in the same order,
every part byte-identical after inflation, and **no entry recompressed** —
because an untouched part is passed through as the DEFLATE stream it already
was, never inflated and re-deflated.

That last property is also what makes the tier reproducible, and it inverts the
relationship between the tiers in a way worth stating plainly. Tier A **stores**
its entries rather than compressing them, deliberately, because deflated bytes
depend on which zlib produced them and a fixture with a pinned SHA-256 would
need re-pinning after a routine Node upgrade. Tier C deflates — its entries are
PowerPoint's DEFLATE streams — and is still bit-stable, because `fflate` never
runs. So Tier C has a `C-REGEN` and Tier B cannot have one, which reads
backwards until you see that Tier B's obstacle was never compression: it is
`dcterms:created`.

### A third producer of headers, and not of markup

The reason for Tier C in the first place was to stop `C-LEX` resting on a single
foreign producer. It does that only partly, and the rule has to say so rather
than count files.

Every byte inside every part of `c01` is still Microsoft's. Self-closing
spacing, quote character, entity spelling, BOM, declaration form, attribute
order, element order — on all of them `c01` is `b01` again, and a rule that
counted distinct **files** rather than distinct **serializers** would score a
copy as independent evidence. What is genuinely ours is the container: 37 ZIP
headers, where our writer and PowerPoint disagree in exactly three places.

| field           | PowerPoint 16.0.20326                | `packages/opc` |
| --------------- | ------------------------------------ | -------------- |
| version made by | 45 (host 0, ZIP 4.5)                 | 20             |
| general-purpose | `0x0006` deflated, `0x0000` stored   | `0x0000`       |
| extra field     | `0xA220` growth hint × 5, 1832 bytes | none           |

Each is legal to drop, none is a bug, and every one is invisible to a reader —
which is precisely why they are asserted in both directions in
`tools/corpus/tiers/c-written/decks.test.ts` rather than described in prose and checked
nowhere. `0x0006` is bits 1 and 2, a compression _level hint_ meaning "super
fast" that no decompressor reads; we deflate at level 6, whose encoding is
`0b00`. The growth hint is 512 bytes of padding so an editor can rewrite a part
slightly larger _in place_; we rewrite the whole archive on every save, and the
1832 bytes PowerPoint spends on it is the entire size difference between the two
files.

### The documentation was wrong where the code was not

All three sat in `zip-writer.ts`'s header table under the sentence "Every field
value below was measured from a file PowerPoint 365 wrote on this machine …
Where the two differ, PowerPoint wins". Two of them were listed as "as
PowerPoint writes". They are not, and
`corpus/ground-truth/powerpoint-conventions.json` had recorded `0x0006` and the
520-byte extra field since E8 — the contradiction was sitting in the repository,
between two files nobody had read side by side.

Nothing in the code was wrong. Every one of our three values is defensible on
its own terms, and `c01` opening in PowerPoint 16.0.20326 with no repair prompt
is the evidence that they are also harmless. What was wrong was a divergence
recorded as an agreement, which is the failure mode that matters most in a file
whose entire job is to say why each byte is the byte it is. The table now gives
both columns and a reason per row.

This is the second time in this sub-phase that a claim in this repository was
checked and found to be a claim rather than a measurement — `ROSTER.md`'s
`BuiltInDocumentProperties` scrub was the first. Both were found by building the
tier that depended on them, which is an argument for building the small
verifying deck rather than reasoning about whether it would tell you anything.

### What Tier C still does not cover

`packages/opc` has two XML serializers of its own — `ContentTypes.serialize()`
and `Relationships.serialize()` — and a no-op write invokes **neither**, because
passthrough means nothing is dirty. So no committed deck in this corpus carries
a part that this project serialized. Their lexical form is measured —
`flat-xml.test.ts` pins the declaration to Office's exact
`standalone="yes"` + CRLF + no BOM, and the tight `/>` — but it is
unrepresented in the corpus. Closing it needs a Tier C deck that edits
something, which is sub-phase 1.3's export rather than 1.1's corpus, and it is
recorded as a gap rather than folded into the claim.

---

## `C-LEX`, and the answer it gives

The rule sub-phase 1.1 stated and could not enforce until there was a second
producer: **every lexical convention is exercised by at least two decks from at
least two distinct serializers, or carries a written gap saying why not.** It is
`tools/corpus/lexical/lexical.test.ts`, against the sixty declared forms in
`lexical-forms.ts`.

It matters more than feature coverage ever did. Sub-phase 0.5's gate is "parse,
serialize, byte-identical for 100% of parts". Run against decks a single
generator wrote, that gate proves the generator and the serializer agree with
each other, which they would even if both were wrong. Only a second producer
turns it into evidence about the format rather than about us.

### The unit is a serializer, and making that true needed a schema change

Counting files would have scored `c01-opc-writer` as independent evidence about
XML lexical form when its parts are `b01-blank`'s parts, byte for byte. Counting
tiers makes the same mistake with a directory name attached. So the unit is the
code that decided the bytes, and a package has **two** of them: what serialized
the parts, and what wrote the ZIP headers.

The manifest is the only durable place that fact can live, and it could not say
it. `producer` was already carrying two different shapes — an object for Tier B,
a bare string for Tiers A and C — and nothing validated either. So `C018` was
added: a collection with committed decks declares `serializers.xml` and
`serializers.container`, per collection rather than per entry, because within a
collection it is true by construction and fifty-one repetitions of one pair is
fifty chances for one of them to be wrong.

`producer` was left exactly as it was. Widening it would have meant rewriting
fifty-one entries to answer a question the envelope answers once, and the two
fields say different things: `producer` is provenance — who made this file, and
on what build — while `serializers` is about which code decided which bytes. The
cost is that `producer` still holds two shapes and `C018` does not check it; see
**Still open**.

`corpus/written` names `Microsoft PowerPoint 16.0.20326` for its XML, spelled
exactly as `corpus/authored` spells it. The collision is the mechanism: two
collections sharing a serializer count once, so the caveat is enforced rather
than written in a paragraph and trusted.

`corpus/decks` names `tools/ground-truth/lib/zip.ts` as its container rather than
`tools/corpus/tiers/a-generated`, because the chassis hands its entries to `writeZip` and
decides no header field itself.

### It reads bytes, and does not ask `packages/xml`

Deliberately. `packages/xml` records the quote character, the self-closing form
and the whitespace before `>` because ADR 0005 decided it must; a rule that
asked the tokenizer which forms a file contains would agree with the tokenizer
by construction and could never catch it being wrong. The scanner is a
hand-written lexer over the text — a lexer and not a regex sweep, because `/>`
occurs inside attribute values and `&gt;` occurs in text.

### The answer, and it splits down the middle of the package

| layer                    | serializers | forms | covered by two |
| ------------------------ | ----------- | ----- | -------------- |
| the ZIP container        | **3**       | 24    | **17**         |
| the XML inside the parts | **2**       | 36    | **12**         |

Three container writers and two XML serializers is the ceiling, and it is the
whole story. The container is well evidenced. The markup is not, and the largest
single reason is one nobody had noticed:

**Every entity reference in the corpus was written by `tools/corpus/tiers/a-generated`.** The
nine PowerPoint-authored decks contain no ampersand at all across their 387 XML
parts, so `&amp;`, `&lt;`, `&gt;` and `&quot;` are each checked only against the
escaper that wrote them. That is exactly the failure the rule exists to name,
and it is not a limit of the producer: E8 measured PowerPoint writing all four,
and a follow-up on 2026-08-27 caught `name="a &quot;quoted&quot; name"` on a
renamed shape. The decks both measurements used were never committed. One more
Tier B deck closes four rows at once.

### A row has to cite the world, because two producers that agree prove nothing

Every form carries evidence that it occurs outside this repository, or the
counter-evidence. The counter-evidence turns out to be as useful: `a36`'s
single-quoted attribute is **0 of the 170 019 attributes** ADR 0004 measured, and
numeric character references are **0 of 2834 parts**. Those fixtures exist
because ADR 0005 chose to record the quote character per attribute rather than
per part, not because a producer was seen to write one — and the table now says
so rather than implying a second producer is owed.

That distinction runs through the whole inventory. ADR 0004 opened 62 lexical
variants in the installed PowerPoint and re-saved the ones it accepted: comments
and processing instructions are **discarded**, a CDATA section is **rewritten as
plain text**, `&#72;` **resolves to `H`**, a single-quoted attribute is
**rewritten double**, a byte order mark is **stripped**. For those forms our
second producer is not silent but incapable, and no deck could ever close them.
Saying "awaiting a second producer" of them would have been false.

### Two things it found that nothing else had

**PowerPoint does not always write an XML declaration.** Four of the corpus's
1422 XML parts have none: `ppt/charts/style1.xml`, `style2.xml`, `colors1.xml`
and `colors2.xml` in `b05-chart` begin at `<cs:chartStyle` and `<cs:colorStyle`.
ADR 0004 scanned 2834 real parts and E8 scanned 38, and no part in either lacked
one, so nothing written down here would have predicted it. It is the ChartEx
spacing finding seen from another angle — there is more than one XML serializer
inside PowerPoint and the chart ones are the odd members. What matters is not a
second producer for the form but that nothing here adds a declaration back,
which `packages/opc` gets for free by never re-serializing a part it did not
edit.

**The commonest real-world form of three conventions is absent here entirely.**
1037 of ADR 0004's 2834 parts carry a BOM and none of ours do. 2161 of them put
nothing between the declaration and the root element; all 1418 declared parts
here put a CRLF there. Two of its three declarations spell the encoding `utf-8`;
neither of our producers does. Each is handled — the tokenizer keeps a BOM, and
the declaration is never rebuilt — so these are unrepresented rather than
untested. But a corpus that cannot exhibit the majority form of three
conventions is a corpus whose round-trip badge means less than it looks, and the
rule now says that in a place a reader will hit.

### What it may not assert

**Entry order.** The two E8 decks and `b01` disagree about whether `slide1.xml`
precedes its own `.rels`, and `b05` interleaves layouts with their rels in no
pattern at all. The order is unstable _within_ one producer, so there is no
convention to check. That is a rule about the rule rather than a gap in it, and
it is why "entry order" appears in 1.1's plan for `C-LEX` and not in the
inventory.

## `C-COV`, and the one thing it could not cover

The last of the four rules, and the shortest to state: **every census feature key
is exercised by at least one deck, or a manifest declares it with why nothing
covers it and what would close it.** The answer is forty-six of forty-seven.

It is `C019` in `tools/corpus/manifest/check.ts` rather than a test, which is a
deliberate departure from where the other three live. `C-CENSUS`, `C-REGEN` and
`C-LEX` each have to build a deck, read a census or unzip an archive, so none of
them can run in `pnpm corpus`. `C-COV` needs nothing but the manifests and the
committed key list, so it runs on a bare clone before anything is built — which
matters, because the thing it catches is somebody adding a census feature and no
deck to exercise it, and that is a change a contributor makes in the census
package while thinking about something else entirely.

**Both directions fire, and the second is the one that keeps it honest.** A key
nothing covers and nobody declared is how a coverage badge becomes a lie. A key
still declared after some deck started covering it is how the declaration rots
into an apology for a hole that was filled — so `C019` fails on that too, and
whoever builds the deck deletes the entry in the same commit.

Two definitions inside the rule are load-bearing and neither was obvious:

- **A `features` count of zero is not coverage.** `a01-minimal` writes zeroes to
  say the census looked and found none, which is a statement worth recording and
  the exact opposite of exercising the feature. Reading a zero as coverage would
  let one deck declaring every key at zero score the corpus at 100%.
- **A `pinned` deck is not coverage.** Its bytes are the hash of a file that
  never left the machine that made it, so a coverage claim resting on one rests
  on something no contributor can reproduce. There are none in the corpus today;
  `coverage.test.ts` is what notices if one arrives.

### The gap it declares

`model3d`, and `a28-model3d` is cut rather than written. Experiment E4 needs the
`a:ext` GUID that carries `am3d:model3d`, and a search of the public record
found the element but not its host: Microsoft's Open XML SDK documents
`DocumentFormat.OpenXml.Office2019.Drawing.Model3D` and its serialized name —
which does at least corroborate the census rule's lower-case `model3d` spelling
against a first-party source — but publishes neither the extension URI, nor the
relationship type and content type of the model part beside it. ECMA-376 predates
the feature entirely.

Guessing the four strings was considered and rejected, and the reason is the
interesting part — it is the opposite of the one that first suggested itself.
The obvious objection is that a fabricated `@uri` would be thrown away, so the
probe would quietly evaporate. It would not. `a34-extlst` measured exactly this
question and found that **an unknown `a:ext/@uri` is carried through untouched**,
which is the whole contract `extLst` exists to provide and is elsewhere a
property this project depends on.

Here it cuts the wrong way. A deck with a made-up GUID would open without
complaint, render the baked raster, and preserve the extension through every
resave — so `C-CENSUS` green, `C-REGEN` green, `C-COV` green, permanently, on a
census key covered by markup no Office feature has ever read. There is no later
event that contradicts it. A hole that is written down is worth more than a
probe that cannot fail, and declaring the gap is the honest form of the same
information.

**Sixty forms in `C-LEX`, forty-seven keys in `C-COV`, and between them one
declared hole.** That is the state sub-phase 1.2 inherits.

## Still open

The roster is complete. Everything the corpus does not have is written down
deck by deck in `tools/corpus/ROSTER.md`, together with the experiments that
gate individual decks and the table of what PowerPoint refuses.

**All four named rules are closed.** `C-CENSUS` for every deck in all three
tiers; `C-REGEN` for Tiers A and C, and it does not apply to Tier B; `C-LEX` and
`C-COV` in the two sections above. The privacy and font rules are closed for the
only tier that could violate them.

**Tier B being complete does not close the declared gaps that say "Tier B", and
`C-LEX` added two more.** Of the original five only `_rels/.rels` was covered by
one of the planned nine. A shape-anchored comment, a VBA project that compiles,
a video with coded samples and ink each need a tenth deck the roster never
planned, and two of them additionally need something outside this project's
control — a Trust Center setting, and a machine with a pen.

The two new ones need none of that, only a decision:

- **a deck whose text and shape names need escaping**, which closes the four
  entity rows at once and is the largest single thing the corpus is missing;
- **a deck with a ChartEx chart**, which is the only way to get a second
  producer for `xml.selfClosing=spaced` — the dominant real-world form, 70 822
  of 98 777, currently witnessed here only by a deck we wrote.

Widening Tier B is still a decision rather than work that was skipped, but these
two are cheap and they buy more than the other four.

**`producer` holds two shapes and nothing validates either.** Tier B writes an
object — application, version, build, platform — and Tiers A and C write a bare
string naming a directory. Both are honest for what they describe and no rule
compares them against anything, so a third shape would pass too. `C018` covers
the question `C-LEX` actually depends on and deliberately left this one alone. A
rule of the same size as `C018` would close it whenever `producer` is next
touched.

**`M1.8` is a wrong severity, not a missing rule.** `packages/opc` reports an
over-encoded unreserved character in a part name as a warning; PowerPoint
refuses the package. Sub-phase 1.2 should raise it, and `a40-unicode` is the
fixture that says why.

**A `format` that disagrees with its `path` is unguarded.** Nothing in
`tools/corpus/manifest/check.ts` compares the manifest's `format` field against the
extension of its `path`, so the two could claim different things about one file
and every rule would pass. They cannot today because both derive from
`ProbeDeck.extension`, which is a property of the generator rather than of the
checker — a hand-written manifest entry has nothing stopping it. It is a
one-line rule whenever the manifest stops being generated.

**Whether the VBA project in `a32-macros` would compile is not known.** The
compound file and the compression are checked structurally against
independently written code, and PowerPoint does not load a VBA project when it
opens a file, so the deck opening is evidence about the package and not about
the project. Six `Reserved` constants inside `dir`, the `_VBA_PROJECT` version
word and the acceptability of a project with no `PROJECTREFERENCES` are
asserted from the specification. Closing it needs "Trust access to the VBA
project object model", a Trust Center setting on the user's machine.

**E4 remains blocked, and `a28-model3d` is cut on the strength of it.** The
reasoning is in the `C-COV` section above. The route back is unchanged and
cheap to describe: insert a 3-D model in PowerPoint, read the extension out of
the saved file, write the deck, delete the `uncovered` entry. What blocks it is
that PowerPoint's gallery fetches a Microsoft-licensed model over the network,
which is a download nobody has approved. Whether PowerPoint will accept a model
authored here instead — a glTF binary this repository could write, as it writes
its own PNG, JPEG, EMF and compound files — is untested and would be the first
thing E4 tried.

**`corpus/reject/` and `C-REJECT` are sub-phase 1.2's, not 1.1's.** Thirteen
refusal findings are recorded in `ROSTER.md` waiting to seed it, and two
`C-LEX` rows — `xml.markup=doctype` and `xml.entity=hex-upper-ill-formed` —
already point at it. The collection is for files this project must **refuse**,
which is the validator's subject rather than the corpus's.
