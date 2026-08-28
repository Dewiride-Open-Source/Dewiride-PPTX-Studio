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

Every `corpus/**/manifest.json`. There is deliberately no single root manifest: the three
collections have genuinely different envelopes (`bench` records a generator, `ground-truth` records
what it was measured on) and flattening them would force the union of all three onto every entry,
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

### Two that are worth explaining

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

## Not in turbo, on purpose

`pnpm layering` is the precedent. If anyone wires this into a turbo task the hazard is precise:
turbo's default inputs cover a package's git-tracked files, and the moment somebody writes an
`inputs` array to tweak one thing they opt out of that default — `corpus/**` silently stops being
hashed and a deck added without a manifest entry replays a cached PASS. Turbo also hashes tracked
files only, so an untracked deck is invisible to it however inputs are declared. That is the third
reason `C008` walks the filesystem itself.

## The files

| file                   | what it is                                                               |
| ---------------------- | ------------------------------------------------------------------------ |
| `ROSTER.md`            | the fifty-two decks, deck by deck, and what each is the probe for        |
| `schema.ts`            | the enums — the legal posture in executable form                         |
| `check.ts`             | `checkCorpus(input)`, pure: no filesystem, no process                    |
| `check-corpus.ts`      | discovery, hashing, formatting, exit code                                |
| `check.test.ts`        | one `it` per rule, synthetic input, no temp directories                  |
| `conventions.test.ts`  | every producer here writes the bytes PowerPoint writes                   |
| `census-keys.gen.ts`   | generated; the 47 census feature keys                                    |
| `census-drift.test.ts` | asserts the generated list still matches the census                      |
| `write-census-keys.ts` | regenerates it (`pnpm build` first)                                      |
| `gen/`                 | the Tier A generator: chassis, shape and text helpers, one file per deck |
| `gen/*.ts` (bytes)     | `png` `jpeg` `emf` `xlsx` `media` `cfb` `vba` — every binary we author   |
| `gen/cfb.test.ts`      | the two layers `a32-macros` needs that no other test can see             |
| `authored/`            | the PowerPoint COM scripts, so Tier B's provenance is auditable          |

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

`check.ts` is pure for the same reason `tools/layering/check.ts` is: it takes the parsed manifests
**and** a pre-computed file table, so `C008` and `C010` are testable against a corpus shape this
repository does not have yet.

## Exit status

`0` clean, `1` violations, `2` the check could not run. The third is separate on purpose — a
manifest that is not valid JSON is a different incident from a manifest that says something wrong.

## Still to come

Two of the four are now closed, in `gen/probes.test.ts` rather than here — both need to build a
deck and read a census, and this checker deliberately does neither:

- **`C-CENSUS`** ✔ — a fresh `censusPackage()` of each committed deck reproduces its `features`
  map exactly, in both directions.
- **`C-REGEN`** ✔ — every deck rebuilds to the bytes committed under `corpus/decks`.

The two that need the roster complete before they can mean anything:

- **`C-COV`** — every census key is covered by at least one deck, or is named in a top-level
  `uncovered` array. A declared gap is a signed statement; an undeclared one is how a 50/50 badge
  becomes a lie.
- **`C-LEX`** — every lexical convention (self-closing spacing, quote character, entity spelling,
  BOM, declaration form, ZIP general-purpose flags, DOS timestamps, entry order) is exercised by at
  least two decks from at least two distinct producers. Feature coverage was never the binding
  constraint on whether the round-trip gate means anything; this is.
- **privacy and font rules** for any deck authored by a real application — no author name, company
  or machine name in `docProps`, and no `ppt/fonts/` part outside the one CC0 exception.
