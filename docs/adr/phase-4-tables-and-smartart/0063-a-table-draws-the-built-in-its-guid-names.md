# 0063 — A table draws the built-in its GUID names

Date: 2026-09-29
Status: **accepted** — PowerPoint's own gallery gave up 74 styles and 74 GUIDs, each serialised one
way across 81 sources in two sessions and three themes (377/377). Across 84 packages one rule fits
every question: the built-in the GUID names, in any case, whatever the package says (230/230, 7/7),
and otherwise a 1-pt black grid (3/3 against the theme's `dk1` and `tx1`). 23/23 mutants killed,
one of them only after the assertion it exposed was added.

**Sub-phase 4.2.** The plan's row says "the 74 built-in table styles, re-derived from PowerPoint".
4.1 had parsed `a:tbl` and left `a:tableStyleId` a string: PowerPoint keeps a built-in id the package
does not define and discards one it does not know (ADR 0056), and nothing in this repository knew
which ids were built in. `LEGAL.md` binds this sub-phase to an independent derivation, because the
obvious source, LibreOffice's `predefined-table-styles.cxx`, is MPL-2.0. This is the record of
asking PowerPoint for the list, for each definition, and for which style a table is actually drawn
with.

Code: `packages/model/src/{table,parse/table,resolve/table,builtin/table-styles}.ts`,
`V032` and `V033` in `packages/validate/src/rules/{rules,id,required,table-style-ids}.ts`,
`tools/ground-truth/model/tables/styles/`, `tools/corpus/suites/tables.test.ts`. Fixture:
`corpus/ground-truth/table-styles.json`. Changeset:
`.changeset/a-table-draws-the-built-in-its-guid-names.md`.

---

## Where the 74 come from

PowerPoint's object model applies a style by GUID (`Table.ApplyStyle`) and reads one back
(`Table.Style.Id`, `.Name`), and cannot list them. No Microsoft document lists the GUIDs either;
Office.js's `PowerPoint.TableStyle` enum names the styles and nothing more. So `author.ps1` drives
PowerPoint's own Table Styles gallery through Windows UI Automation (the owner chose this over an
Office.js add-in and over clicking by hand): a visible window, a selected table, and each item of
the in-ribbon gallery invoked through its `InvokePattern`, with a sentinel style applied over COM
first so that an invoke that did nothing shows. A `-Discover` pass dumped the gallery's UI tree
before anything was invoked; the script refuses to start while any deck is open and walks only
PowerPoint's own window.

The gallery lists **74 items, 74 GUIDs**, in the order of the Office.js enum, each item's UI
Automation name equal to the `Style.Name` COM reads back and to the `@styleName` PowerPoint then
writes. The expanded gallery groups them as Best Match for Document (14), Light (21), Medium (28)
and Dark (11). A new table gets Medium Style 2 - Accent 1, `{5C22544A-…}`, and so does `@def`.
LibreOffice's list was not opened: every GUID here came from `Table.Style.Id` after a gallery
click, and every definition from PowerPoint's writer.

## What a definition is

With the GUIDs in hand, `author.ps1` applies each one to its own table and saves: `pp-styles-off`
and `pp-styles-on` hold all 74 with every flag off or on, and `pp-one-01` … `pp-one-74` one style per
deck. `read.ps1`, in a second PowerPoint process, applies every style again in two sweeps built in
our own themes and saves those too. Four readings of what PowerPoint writes were scored over 377
(source, style) rows from 81 sources:

| reading                                   | fits    |
| ----------------------------------------- | ------- |
| **one string per GUID**                   | 377/377 |
| baked into the theme it was applied under | 303/377 |
| different in another session              | 229/377 |
| pruned to the parts the table's flags use | 0/377   |

So a definition is a fixed string: theme colours by slot, tints and widths, never the theme's
values. All 74 come to 164,549 bytes. Every accent family is one template with the accent
substituted (5/5 each, Dark Style 2's paired accents 2/2), which also checks the GUID-to-name
binding a second way; the strings are shipped as written rather than rebuilt from templates.

**Legal standing.** These are PowerPoint's bytes, committed and shipped as the Tier B decks are
(ADR 0009): PowerPoint driven by our scripts, writing its own output. The owner approved committing
them verbatim in the fixture and in `@pptx-studio/model` before the experiment ran. `NOTICE` and
`LEGAL.md` record it, and `LEGAL.md`'s LibreOffice row now says neither thing the plan named was
taken from LibreOffice.

## C8 — which style a table draws

84 packages, one table each (or one per slide in five sweeps of all 74), in three themes of our
own: the stock-like palette, a second palette with other fonts and a clrMap that maps `bg1` and
`tx1` to the dark slots, and a third whose `dk1` is a brown RGB rather than a system colour. Each
slide is exported at one pixel per point and the table's region compared byte for byte with a
control: a built-in's own definition in the same theme, or a table with direct borders. The
instruments were checked before anything was scored: 79 slides exported twice were identical, the
controls read again after the probes were unchanged (no style leaked between packages), and
`Cell.Shape.Fill` matched the exported pixel on 325/325 cells.

| question                                      | winner                     | fits    | the losers                                                  |
| --------------------------------------------- | -------------------------- | ------- | ----------------------------------------------------------- |
| a built-in id the package does not define     | the built-in               | 230/230 | the first built-in 5/230; bare, `@def`, the grid 3/230 each |
| a built-in id the package defines differently | the built-in               | 7/7     | the package's definition 0/7; the grid 0/7                  |
| a custom id the package defines               | the grid                   | 6/6     | the package's definition, bare, `@def`, the first 0/6       |
| an id looked up by `@styleName`               | the grid                   | 1/1     | by id 0/1; by name 0/1                                      |
| an id nothing defines                         | the grid                   | 7/7     | bare, `@def`, the first built-in 0/7                        |
| no id                                         | the grid                   | 4/4     | bare, `@def`, the first built-in 0/4                        |
| the grid, in the second theme                 | 1-pt black                 | 3/3     | No Style, Table Grid 0/3; bare 0/3                          |
| the grid's colour, in the third theme         | black                      | 3/3     | the theme's `dk1` 0/3; `tx1` 0/3                            |
| a GUID in lower or mixed case                 | matched                    | 3/3     | every unmatched reading 0/3                                 |
| a GUID without braces, or padded              | repaired                   | -       | -                                                           |
| one GUID defined twice                        | ignored                    | 2/2     | the first 0/2; the last 0/2                                 |
| an inline `a:tableStyle`                      | resolved by its `@styleId` | 4/4     | its content 0/4; the part 1/4; ignored 2/4                  |

**What the losers would have looked like in the field.** A reader that honours the package's
definitions — the one anybody writes first, because the definitions are right there — draws every
redefined built-in as the package says and every custom style from other producers as its author
meant; PowerPoint draws neither. A reader that falls back to `@def` draws Medium Style 2 - Accent 1
under every table that names nothing. A reader that falls back to No Style, Table Grid is right in
the default theme and wrong in any theme that maps `tx1` elsewhere: that is exactly the reading
this experiment held until the second theme refuted it.

### What I had wrong on the way

The probe set changed twice, and both changes are measurements:

- I built my "package whole" and "merge" controls as custom ids with the candidate definitions. Every
  one of them drew the grid, which is how the rule for custom ids was found; those packages became
  custom-id probes, and the modified-built-in question is scored on the only prediction the pixels
  can refute — the package's own fill (0/7) against PowerPoint's built-in, exactly, to the byte (7/7).
  Any merge that took a single byte of the package's definition is refuted by that equality.
- In the first theme the default grid is pixel-identical to No Style, Table Grid, and the first
  analysis called it that. A second theme, where `tx1` maps to a light slot, showed a black grid
  where the named No Style, Table Grid draws light lines; a third, with an RGB `dk1`, separated
  black from `dk1`. The fallback is a fixed 1-pt black grid, not a built-in, and the model returns
  `null` for it.
- I marked the padded-GUID probes as schema-valid, reading `ST_Guid` as a whitespace-collapsing
  token. PowerPoint repaired both. They were alone in their packages, so the repair is the finding
  and they are recorded as hostile.
- The first serialisation comparison counted definitions PowerPoint had merely written back as they
  were read. `read.ps1` now re-applies every style in the sweeps, and only those count as a second
  session.

## What PowerPoint writes back

Every opened package was saved as a copy, and the copy saved again:

- PowerPoint writes the definition of every built-in applied in the session — including a restyled
  table's first style and a deleted table's — and the `@def` style always. It never adds the
  definition of a built-in a package names without defining: the sweeps come back with an empty list.
- A redefined built-in, a custom style and a duplicate are kept, normalised (an empty `a:tcBdr` added).
- An id nothing defines is dropped; a lower-case one is written upper case.
- An inline `a:tableStyle` is moved into the part and replaced by its `a:tableStyleId`.
- The part is found **by the relationship**: a part at another path keeps its custom id, an orphaned
  `ppt/tableStyles.xml` does not.
- A repaired id becomes `{00000000-0000-0000-0000-000000000000}`; a repaired part comes back empty.

## The model

`TableStyleRef`'s inline form now carries its `id`, because that is all PowerPoint reads of it.
`parseTableStyle` reads `CT_TableStyle` into `TableStyle`: the background, the thirteen parts, each
with its text style (`b` and `i` as `on`/`off`/`def`, a font collection or a `fontRef`, a colour) and
its cell style (eight edges as a `Line` or an `lnRef`, a fill or a `fillRef`, `cell3D` kept as a
node). Nothing absent is defaulted. The forms PowerPoint repairs throw: a GUID without braces or with
padding, a missing `@styleId`, `b="yes"` (`MODEL_TABLE_ATTR`), an edge with neither `a:ln` nor
`a:lnRef` and an empty `a:fill` (the new `MODEL_TABLE_STYLE`). `MODEL_ERROR_CODES` had drifted six
codes behind its union; it is now the union's source, checked at compile time.

`BUILTIN_TABLE_STYLES` is generated from the fixture by `write-tables.ts`: 74 entries of GUID, name
and PowerPoint's XML, 173 KB of source and 7.5 KB gzipped, reached only through the two functions
below, so a bundle that never resolves a table style carries none of it. `builtinTableStyle(id)`
looks a GUID up in any case and parses it afresh each time — an `XElement` is mutable, and a shared
one would let an edit in one document change another. `tableStyleOf(table)` is the rule above and
nothing else: it does not take the package's table-style part, because nothing in that part changes
the answer. The package still parses and preserves the part as it always has; `Document` carries no
new field.

## The firewall

`V032`, a warning, evidence `measured`: a table whose `a:tableStyleId` or inline `@styleId` names
none of the 74. PowerPoint opens it without a word and draws the grid, and an id nothing defines is
gone after the next save: the file says one table and PowerPoint draws another, the V030 pattern.
This closes ADR 0056's second open question. `validate` carries the 74 GUIDs in its own generated
list rather than depending on `model`.

`V033`, fatal, evidence `both`: a table style id that is not a GUID in braces, an `a:tblPr` naming
its style twice, a table-style part whose root is not `a:tblStyleLst`, a list without `@def`, a style
without `@styleId`, `b` or `i` outside `on`/`off`/`def`, an edge with neither `a:ln` nor `a:lnRef`,
and an `a:fill` with no fill. Every one is a repair C8 measured, and a test holds the whole firewall
to refusing all eleven packages PowerPoint repaired.

The corpus has one table that names no built-in — `a20-tables`' inline style — and `V032` now warns
on it, pinned beside the two placeholder warnings.

## Verification

- **Tests from the fixture.** `packages/model/src/table.test.ts` holds the catalogue to the fixture
  byte for byte in gallery order; parses all 74 and checks every part each one's XML carries; reads
  Medium Style 2 - Accent 1 and Themed Style 1 - Accent 1 value by value and a crafted style's eight
  edges to their own keys; and — the honest check for a resolver — predicts, for every one of the 59
  single-table probes PowerPoint opened as written, the control whose pixels the table must match,
  and for every one of the 74 built-ins in three sweeps the sweep slide it must match. Each rival
  reading's recorded score is held below its row count. The parser is held to refusing exactly the
  forms PowerPoint repaired.
- `packages/validate/src/validate.test.ts` holds `V032` to firing on exactly the clean probes that
  name an id and drew the grid, `V033` to the eight repair forms and to silence on every clean probe,
  and the firewall to refusing every repaired package. `rules.test.ts` re-derives the GUID list.
- `tools/corpus/suites/tables.test.ts` resolves every table in the 55 committed decks and finds
  `b04-table`'s Medium Style 2 - Accent 1 — written by PowerPoint in 4.1's session through
  `AddTable` — byte-identical to C8's, from the gallery weeks later.
- **Mutants, 23/23 killed.** The model 15/15: the lookup case-sensitive, the built-in lookup
  dropped, No Style, Table Grid as the fallback, the inline style ignored, `b` defaulted to `def`,
  `i` ignored, `insideH`/`insideV` and the diagonals swapped, the `fontRef` colour dropped, a padded
  GUID accepted, the opening brace made optional, an empty edge and an empty fill read as none, a
  missing `@idx` read as 0, and one byte of the catalogue. The firewall 8/8: `V032` ignoring an
  inline style, comparing case-sensitively, firing on built-ins; `V033` letting a brace go, checking
  no root, letting a style be named twice, letting `b` be anything; one GUID of validate's list.
  `V033` letting the opening brace go survived the first run - no probe had a GUID with one brace -
  and is killed by the assertion added for it. The `MODEL_TABLE_STYLE` thrown when a built-in is no
  `a:tblStyle` is reachable only through a corrupted catalogue, which the drift test forbids.
- **`pnpm check`** green before the commit that claims the sub-phase.

## What the release said

Pull request #66 merged as `e02129c` on CI run 36534856875 with every check registered and complete —
`check`, the candidate gate, the round trip, both font jobs, CodeQL and the website build — and
main's own run 36535538484 green on the merge. The canary before, 36534876056, was green against the
previous release. Release run 36536202313 published `@pptx-studio/model` 0.3.0,
`@pptx-studio/validate` 0.3.0, `@pptx-studio/render-svg` 0.6.4, `@pptx-studio/render-dom` 0.3.5,
`@pptx-studio/writer` 0.1.4 and `@pptx-studio/cli` 0.4.2 in 3 min 34 s, version commit `5e57de5`.
The deploy that followed, 36536560614, waited 40 s for `cli@^0.4.2` and was live 4 min 58 s after
it started; the site's validate page lists thirty-three rules with `V032` and `V033`, and its model
reference lists `tableStyleOf`, `builtinTableStyle` and `BUILTIN_TABLE_STYLES`. The canary after,
36536825726, installed the release and was green.

## Deviations from the plan

- **The inline `a:tableStyle` is resolved here, not in 4.3.** C8 measured that PowerPoint reads only
  its `@styleId`, so there is nothing for 4.3 to parse into the cascade.
- **Two firewall rules**, not in the plan: `V032` answers ADR 0056's open question, and `V033`
  exists because C8 measured eight forms as repairs.
- **No `Document.tableStyles`.** The plan expected the package's part to feed resolution; it does not.
- **Three layouts moved.** `tools/ground-truth/model/tables/` became `tables/grid/` beside the new
  `tables/styles/`, and `builtin-text-styles.ts` became `builtin/text-styles.ts` beside
  `builtin/table-styles.ts`: in each case a second sibling would otherwise have shared the first's
  prefix. The text-styles generator writes through `repoPath` now and re-emits its table byte for byte.
- **Legal files corrected.** `LEGAL.md` and `NOTICE` credited libgdiplus for the pattern tiles, which
  ADR 0022 measured from PowerPoint instead; `NOTICE` lacked the Tier B theme disclosure ADR 0009
  said it had. Both are fixed, and the package copies follow.
- **The rule count left the comments.** Every source comment that stated it had drifted (twenty-six,
  twenty-eight and twenty-nine were all still in the tree); each lost the number and was cut to the
  comment caps. Documentation says thirty-three, and the site's validate demo reads `RULES.length`.

## Open questions

1. **The grid's text.** C8's cells are empty, so the default grid's text colour and font — and
   whether a table that resolves to nothing draws its text the way No Style, Table Grid would — are
   4.3's to measure, with the rest of the cascade.
2. **What the writer adds.** PowerPoint writes a definition for each style applied in the session and
   never for one merely named. When 4.4 lets a user pick a style, whether the export should write
   its definition the way PowerPoint does is a question about other readers, not about PowerPoint.
3. **`p:style`'s own lenient defaults.** A `p:style` reference with no `@idx` reads as 0 and a
   `fontRef` with none as `minor`; the table-style references here throw instead, as the schema
   requires. Neither default on `p:style` has been measured.
