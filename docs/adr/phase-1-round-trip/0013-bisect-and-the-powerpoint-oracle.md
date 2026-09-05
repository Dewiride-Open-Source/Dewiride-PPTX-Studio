# 0013 — `cli bisect`, and asking PowerPoint a question it does not want to answer

Date: 2026-08-29
Status: **accepted** — 135 changes narrowed to one removed element in 5 runs, against the real PowerPoint
Sub-phase: 1.5 — `cli bisect` + the PowerPoint loop

---

## Context

The plan gives 1.5 one sentence of design and one of purpose:

> Given a deck PowerPoint repairs, iteratively revert parts and subtrees to
> original bytes until the smallest reproducing change is found. This _is_ the
> debugger for this project. Scripted open-in-PowerPoint verification via COM.

And the appendix says why it has to exist at all:

> Passing all of this is necessary, not sufficient. PowerPoint rejects some
> schema-legal markup for unenumerated reasons.

`@pptx-studio/validate` covers the twenty-nine failures we know how to describe.
This sub-phase is for the rest, which are the interesting ones, and for which
PowerPoint's entire diagnostic channel is one sentence naming no part, no
element and no reason.

---

## The finding that decides the whole design

**PowerPoint's automation interface repairs silently and reports success.**

`Presentations.Open` has no repair parameter. `Presentations.Open2007` does —
`OpenAndRepair`, and Microsoft documents its default as **`msoTrue`**. There is
no documented way to make `Open` behave differently. Separately,
`Application.DisplayAlerts` documents `ppAlertsNone` — which is the value at the
start of every session — as: _"If a macro encounters a message box, the default
value is chosen and the macro continues."_ The default answer to "PowerPoint
found a problem with content in deck.pptx" is Repair.

Put together: the obvious script reports `opened = true` on exactly the files
this project exists to avoid producing. Our own `open-in-powerpoint.ps1` from
sub-phase 1.1 is that script. It can tell a refusal from a success and it cannot
see a repair at all, which is the failure Gate 1 is written against.

Passing `OpenAndRepair:=msoFalse` turns the silent repair back into a catchable
error. Measured on a deck built for it — `a31-embedded-fonts` with its
`<Default Extension="fntdata"/>` deleted, the plan's own canonical repair bug:

| Call                      | Result                                             |
| ------------------------- | -------------------------------------------------- |
| `OpenAndRepair:=msoFalse` | fails, `HRESULT 0x80CB8002`                        |
| `OpenAndRepair:=msoTrue`  | opens, 3 slides, shape counts `[5, 3, 4]` — intact |

Same file, same session, opposite answers. The second row is what every previous
script in this repository would have reported.

### What that says about `corpus/reject`

`tools/corpus/tiers/rejects/fixtures.ts` carried a standing caveat: the seventeen
minimal packages had never been opened in PowerPoint, only the larger probe
packages their markup came from. They have now been, and the answer is worth
recording because it corrects a word:

- **All seventeen fail with `OpenAndRepair` off.** The minimal fixtures do
  reproduce a failure, which is what was owed.
- **All seventeen open with it on**, repaired down to one slide holding one
  shape — two for `r17`.

So the minimal fixtures are **repairs, not outright refusals**, while several
entries describe themselves as a "whole-package refusal". Those words are left
alone: they record the larger packages, and the two can legitimately differ.
Nothing here has reproduced PowerPoint's "could not open the file" path.

The corpus decks were swept the same way: **all 51 open with `OpenAndRepair`
off.** They do not merely open — they open without needing repair, which is a
materially stronger statement than anything measured before.

---

## Decisions

### The atoms are changes, not elements — this is isolation, not simplification

The distinction is Zeller's (_Simplifying and Isolating Failure-Inducing Input_,
IEEE TSE 2002) and it decides everything downstream. Simplification starts from
one failing input and cuts it down; it has to guess what a smaller input looks
like, and most of its guesses are not well formed. Isolation starts from a
passing configuration and a failing one and reduces the _difference_.

We are always in the second case: a deck PowerPoint repairs came from a deck it
opened, and we have both. So a configuration is a subset of the changes between
them, applied to the original. Three consequences, and all three matter:

- **Every configuration is well formed.** The splice unit is a whole node
  replaced by the whole node the other package has in that position, so no
  candidate can invent unbalanced markup. Not one oracle run is spent on a file
  no editor could have produced. Asserted, by parsing every candidate the search
  produced.
- **`ddmin`'s precondition holds by construction** — and is still checked. The
  empty configuration must pass and the full one must fail; when either does
  not, saying so is the most useful thing the command can do.
- **The answer is a sentence about your edit**: "one of these 135 changes", not
  "something about OOXML".

The search is `ddmin` applied level by level down the change tree, which is
Misherghi and Su's HDD (_HDD: Hierarchical Delta Debugging_, ICSE 2006). Moving
down a level is free: applying every child of a change reproduces that change
exactly, because the regions between the children are the regions that compared
equal — so the next level is known to fail without spending a run to prove it.

### The delta is over ZIP entries, because `[Content_Types].xml` is not a part

`PartStore.partNames` excludes the content-type stream by design, and it is
right to. But a missing `<Default Extension="fntdata"/>` in it is the canonical
"PowerPoint found a problem with content" bug — the one the plan cites, the one
pandoc shipped, and the one measured above. **A bisector built on parts could
not vary the single best-documented repair prompt there is.**

Working at the entry level also means nothing has to know what an entry holds.
Whatever parses as XML is descended into — parts, `.rels`, `docProps` and the
content-type stream alike — and whatever does not is one atom, whether it is a
PNG, an OLE2 compound file or markup too damaged to read. There is no
content-type table to keep in step, and no list of which parts are interesting.

### Children are matched from both ends, not by position

The first version paired children by index. On the flagship case it reported
`element /[Content_Types].xml /Types` — the whole stream — because deleting one
`<Default>` takes `<Types>` from seventeen children to sixteen, and after that
nothing lines up.

Matching a common prefix and a common suffix first, and treating only the middle
as changed, turns that into:

```
removed  /[Content_Types].xml /Types/Default[3]
  -  <Default Extension="fntdata" ContentType="application/x-fontdata"/>
  +  (nothing)
```

When the two middles are the same length they pair one to one, which is the
ordinary edit-in-place case and still resolves down to a single attribute. When
they are not, the whole middle is one change. That is coarse if a part had
several _separate_ insertions and exact for one, which is what a writer
produces. It is a real limit and it is written down rather than papered over.

### Three oracles, and three verdicts rather than two

`validate` (the default) asks the twenty-nine rules: no PowerPoint, milliseconds
per run, and what CI uses. `powerpoint` asks the real thing over COM.
`command` runs anything else, with `{}` standing for the candidate's path.

Every oracle returns `fails`, `passes` or **`unresolved`**, and the third is not
decoration. A run that answered neither question — a timeout, a COM server that
would not start, a command that does not exist — must not be read as either. As
"passes" it would let the reducer discard the change that matters and then blame
an innocent one; as "fails" it would let it keep everything. It is counted,
reported, and never reduced towards. The PowerShell harness exits `2` for its
own failures precisely so this stays distinguishable from `1`.

### Synchronous, including the subprocess

Every oracle we have is either pure computation or a subprocess the caller
blocks on anyway, and `spawnSync` blocks perfectly well. Making the reducer
asynchronous to serve a caller that does not exist would put a Promise in the
middle of a browser package and change `main()`'s signature for nothing.

### What the harness will not do to the machine it runs on

Read-only, no window, nothing written back. `AutomationSecurity` is set to
`msoAutomationSecurityForceDisable`, because the property **defaults to Low —
which enables all macros** — and this script's entire job is opening files that
are wrong on purpose, including a macro-enabled one. If PowerPoint is already
running it attaches to that instance and never quits it; it only quits one it
started itself.

---

## What building it found

### Four unit tests that passed for no reason

The first draft of `bisect.test.ts` built its fixtures with a single difference
between the two packages. `ddmin` returns immediately on a set of one, so four
tests asserted the right answer without the search ever running — including the
ones for the run cache and for the `maxRuns` ceiling, which by construction
could never have been exercised. Every reduction test now runs against a delta
of dozens, and the file says why in its header.

### The ceiling threw away the progress it had made

Hitting `maxRuns` returned the level the last completed `ddmin` had produced,
which for a single-entry delta is "the whole entry" — true, and useless. It now
records every failing configuration on the way down, so a run stopped one test
short of the answer returns the smallest set it reached. On the test case that
is 20 changes instead of 1 entry.

### One edit, two rules

Deleting the `fntdata` `Default` trips `V002`, which names that invariant — and
also `V001`, because five `.fntdata` parts are now left resolving to no content
type at all. Both are pinned. A report showing only the famous one would
understate what is wrong with the file.

### Truncating both sides of a difference from the front is useless

Two versions of a 4 kB element that differ at character 3 000 render as the same
sixty-eight characters twice, which reads as a bug in the bisector. The window
is centred on the first character where they actually differ, using
`firstDifference` from `@pptx-studio/xml` — written for 1.4's round-trip report
and turning out to be the same problem.

---

## Consequences

`pnpm check` green. The three deliberate breakages the plan asks for, one per
category the firewall recognises, each buried in three dozen harmless changes:

| Break                                    | Localized to                              | Rule   |
| ---------------------------------------- | ----------------------------------------- | ------ |
| `<Default Extension="fntdata"/>` removed | `[Content_Types].xml` `/Types/Default[3]` | `V002` |
| two children swapped in the slide master | two changes under one `a:xfrm`            | `V010` |
| one relationship id renumbered           | `presentation.xml.rels` `…/@Id`           | `V006` |

- **135 changes to 1, in 5 oracle runs, against the real PowerPoint**, in 13
  seconds including COM startup. The `validate` oracle reaches the same answer
  in the same number of runs.
- **The `--write` package reproduces the failure exactly**: `0x80CB8002` with
  repair off, opens with all three slides with it on, and differs from the
  original in exactly one entry.
- **The swap resolves to two changes, and that is the right answer**, not a
  failure to reduce: putting either element back leaves the other somewhere the
  schema allows. 1-minimal is not the same as smallest, and the report does not
  claim otherwise.
- **Bisecting each of the 51 corpus decks against our own export finds nothing
  to bisect** — 1.3 and 1.4 restated at this granularity, and the baseline that
  makes a real bisection meaningful when one is ever needed.
- **All 51 decks open in PowerPoint with `OpenAndRepair` off**; all 17 reject
  fixtures fail with it off and open repaired with it on.

### What is not done

- **The CI gate and the badge are 1.6.** This runs by hand.
- **No bisection has been driven by a failure we did not plant.** Every break
  here was deliberate, because the writer does not currently produce one — which
  is the good news and also the reason this tool is untested against the case it
  was built for.
- **Several separate insertions in one part degrade to one coarse change.**
  Matching from both ends handles one contiguous edit exactly; a longest-common-
  subsequence alignment would handle several, and is not written.
- **`ddmin` is quadratic in the worst case** and the PowerPoint oracle costs
  about a second and a half a run. `--max-runs` bounds it, and says when it bit.
