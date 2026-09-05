# @pptx-studio/validate

**The repair firewall.** Twenty-nine rules a `.pptx` must not break, checked
before any bytes are handed over, each finding carrying the part and the XPath.

Apache-2.0 · browser and Web Worker only, no Node · part of
[PPTX Studio](https://github.com/Dewiride-Open-Source/Dewiride-PPTX-Studio).

```ts
import { assertValid, formatReport, validatePackage } from '@pptx-studio/validate';

// Diagnosing a file somebody sent you.
const report = validatePackage({ bytes });
console.log(formatReport(report, { explain: true }));

// The export gate: throws unless the package is safe to hand over.
assertValid({ store, baseline, bytes: store.write() });
```

## Why this exists

PowerPoint emits no diagnostic log. When it refuses a file the message names no
part, no element and no line — "PowerPoint could not open the file", or
`0x80070570`, "the file or directory is corrupted and unreadable". When it
_repairs_ one it says even less, and the user is left holding a deck that has
quietly lost something.

So this is not a schema validator with a nice report. It is the only feedback
loop that exists, and it runs on every export in development and in production.

## Twenty-nine rules, in two halves

Roughly half come from ECMA-376: content-type coverage, `xsd:sequence` child
order, `minOccurs`, the four identifier ranges. Those are derivable, and one of
them literally is — the ordering table is generated from the Transitional
schemas and checked against 194 148 elements PowerPoint wrote.

The other half cannot be derived from anything. Each is the record of a package
built with **one** change in it, opened in PowerPoint 16.0.20326, and declined:

|        |                                                                                                |
| ------ | ---------------------------------------------------------------------------------------------- |
| `V022` | a `p:ph type="hdr"` on a slide. The other seven content types all open.                        |
| `V023` | a geometry guide referenced but never defined. `ST_GeomGuideName` is an unconstrained token.   |
| `V024` | a `c:strLit` inside a series `c:tx`, two elements away from where the literal forms are legal. |
| `V025` | a `p:control`, in all eight forms tried. An empty `p:controls` is accepted.                    |
| `V026` | a `cs:chartStyle` with thirty of its thirty-one entries. No chart style at all is fine.        |
| `V019` | a master id colliding with a layout id. They are one number space and no schema says so.       |
| `V020` | a `p:cNvPr/@id` between 2147483648 and 4294967294. 4294967295 opens; it is minus one.          |
| `V009` | a slide without exactly one layout relationship; a `cx:chartSpace` with no `.rels`.            |
| `V004` | a percent-escape of an _unreserved_ character in a part name. `0x808D1005`.                    |
| `V012` | an `a:ahXY` directly under `a:custGeom`, without its `a:ahLst` wrapper.                        |

Every rule carries a `why` recording what was tried and what opened, so whoever
eventually contradicts one knows what they are contradicting.

## It answers "did _we_ break it", not "is this valid"

A fatal finding in a file the user imported ten seconds ago is information. The
same finding in a file they just edited is a bug in us, and handing over the
bytes would give them a repair prompt with no explanation.

Refusing both would be the strict-looking choice and it is the wrong one: a deck
with one pre-existing defect — a dangling image relationship, common in files
that have been through three other tools — could then be opened here and never
saved again. The editor would be refusing to give somebody back their own file
over a problem it did not cause.

So every finding carries an `origin`, and only `introduced` ones refuse an
export. It is computed rather than judged per rule: the rules run a second time
against the package **as it was opened** and the two reports are differenced.
That second pass happens only when the first found something fatal, so a clean
export pays nothing for it.

## Nothing is silently skipped

Three rules compare against the package as it was opened, and one reads the
archive rather than the part store. Given neither, they do not run — and the
report lists them under `skipped` with the reason, because a preservation rule
that reports nothing because it had nothing to compare against looks exactly
like one that found nothing. Parts that would not parse land in `problems` for
the same reason.

## How it is verified

Two suites that check opposite things, and neither substitutes for the other.

**It fires.** `validate.test.ts` breaks a minimal deck twenty-nine ways, one per
rule, and asserts each rule reports the right part and the right XPath.

**It is quiet.** `tools/corpus/suites/validate.test.ts` runs all fifty-two committed
corpus decks through it and asserts no fatal finding. Every deck in that corpus
opens in PowerPoint, so a rule that fires on one is a bug and there is no third
possibility. Three rules were narrowed by exactly that run: `V012` was reporting
`mc:AlternateContent` and other extension markup the ordering table simply has
never read; `V020` was reporting shape ids reused across the two branches of one
`mc:AlternateContent`, which are alternatives and may collide; and it found a
real defect in our own `a20-tables`, which wrote a table style's fill without its
`a:fill` wrapper in thirteen places.

**And it fires on the right things.** `corpus/reject/` holds seventeen packages
PowerPoint refuses, one measured finding each, and `C-REJECT` asserts each is
caught by the rule its manifest names. Four rules have no instance anywhere in
the good corpus — every file containing one is a file PowerPoint refuses — and
without those fixtures they would be enforced entirely on trust.

## What passing does not mean

Necessary, not sufficient. PowerPoint rejects some schema-legal markup for
reasons nobody has enumerated, and this package knows the ones we have found.
The strongest defence is architectural — _never synthesize markup we did not
read_. The second strongest is `cli bisect` (sub-phase 1.5), which reduces a deck
PowerPoint refused until the smallest reproducing change is left, and whose
output is how the measured rules above got here.

## API

|                                            |                                                                            |
| ------------------------------------------ | -------------------------------------------------------------------------- |
| `validatePackage(options)`                 | check a package; returns a `Report`, never throws for a finding            |
| `assertValid(options)`                     | the same, but throws `ValidateError` when something we introduced is fatal |
| `formatReport(report, options)`            | the report as text, grouped by part                                        |
| `RULES`, `ruleById`, `RULE_IDS`            | the table, with `title`, `why`, `evidence` and `severity`                  |
| `xpathOf`, `elementLocation`, `lineColumn` | locations, for a caller building its own report                            |

`ValidateOptions` takes `store` (the package about to be written), `bytes` (the
archive, for the two rules that read it), `baseline` (the package as it was
opened) and `rules` (a subset). On an export, `store` must be the store the
writer emitted from and not one re-opened from the written bytes — a re-opened
store thinks every part came from the archive, and `V027` would then have no
history to ask about.
