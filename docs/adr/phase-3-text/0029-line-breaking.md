# 0029 — Line breaking

- **Status** accepted
- **Sub-phase** 3.3
- **Experiment** T3 — 2221 probes across twelve packages, PowerPoint via `TextRange2` over COM
- **Fixture** [`corpus/ground-truth/line-breaks.json`](../../../corpus/ground-truth/line-breaks.json)
- **Code** `packages/text/src/lines/break.ts`, `packages/text/src/lines/kinsoku.ts` (generated)
- **Supersedes, in part** the plan's 3.3, which specified UAX#14 via the `linebreak` package

## Context

The plan says:

> **3.3 — Line breaking.** UAX#14 via `linebreak`, then DrawingML tailoring: suppress intra-word
> breaks unless `latinLnBrk=1` (ECMA says the default is true; Word/PowerPoint behave as if false),
> `p:kinsoku` `invalStChars`/`invalEndChars` as a backward-walking post-filter, `hangingPunct`
> excluding a trailing hangable's advance from the fit test.

Three of those four clauses turn out to be right about the _phenomenon_ and wrong about the
_mechanism_, and the fourth — the UAX#14 base — is wrong outright. This ADR records what was measured
instead, and why the deviation is a correction rather than a shortcut.

## The method

`TextRange2.Lines(i, 1).Text` returns the characters on line _i_. So unlike T2, which had to infer a
rule from a float, T3 reads the answer directly: **line 1 is the break position**.

That turns one string into a complete measurement by way of a width sweep. A greedy breaker sets line
1 to the longest prefix that both fits the box and ends at an opportunity, so a box sized to hold
exactly _i_ code points reports back _the largest opportunity at or before i_. Sweeping _i_ from 1 to
the end of the string enumerates the whole opportunity set, one shape per _i_, with nothing left to
infer.

Box widths come from Chromium's `measureText`, measured per prefix, with each box placed at the
midpoint between prefix _i_ and prefix _i+1_. T2 put the two engines a median of 0.15% apart and half
a character advance at 24pt is several points, so the midpoint is safe by two orders of magnitude —
and the margin is recorded per probe rather than assumed. A `wrap="none"` control per string checks
the estimate against PowerPoint's own width: over 288 comparisons the two agree to a **median of
0.019%**, p95 0.121%, max 0.202%.

### The first pass proved less than it looked like

It is worth recording, because the fix is the reason the result is worth anything.

When a box is too narrow for any opportunity, PowerPoint breaks at exactly as many code points as
fit. So a string whose _only_ candidate opportunities are the ones under test reads `line 1 = fitMax`
at every width — which is equally consistent with "every position breaks" and "no position breaks".
`alpha/beta/gamma` measured as an unbroken staircase and said nothing at all about the solidus. Six
probes were in that state.

The fix was to give every such string a guaranteed opportunity near the start: `aa ` and then the
material under test. Once a real break exists at position 3, every later reading of `line 1 = 3`
_proves_ that nothing in between is an opportunity, because the greedy rule would have preferred it.
The question turns from "is this consistent with" into "is this refuted by".

Two other probes were rebuilt for the same reason. The `strictFirstAndLastChars` deck originally
stated the standard kinsoku sets — so the toggle had nothing to toggle, and could only ever report
"no difference". And the kinsoku sweep was run a second time with `hangingPunct="0"`, because a
character that hangs is _preferred_ by the greedy rule whether or not the position before it is
forbidden, which made the reading ambiguous on exactly the characters where it mattered.

## Decision

### 1. Not UAX#14

PowerPoint's opportunity set is not UAX#14 tailored; it is a much smaller rule that UAX#14 is a
superset of. Scored over the same 2212 non-Thai probes:

| reading                                         | fits          |
| ----------------------------------------------- | ------------- |
| **measured**                                    | **2212/2212** |
| kinsoku applied to every break, as ECMA implies | 2200          |
| no hanging punctuation                          | 2157          |
| no emergency break                              | 2085          |
| kinsoku off entirely                            | 2010          |
| break only after spaces                         | 1898          |
| a faithful UAX#14 reading                       | 1878          |

A line may end at position _k_ when `k-1` is a space and `k` is not; when `k-1` is a hyphen-minus, en
dash, em dash or soft hyphen; when either neighbour is East Asian; or when `latinLnBrk="1"`. A
NO-BREAK SPACE or NON-BREAKING HYPHEN on either side suppresses it, and for East Asian positions the
kinsoku sets do.

Specifically **not** opportunities, each proven by a fallback to an earlier break rather than merely
unobserved: after `/` or `\`; anywhere in `http://a.example/b` or `C:\dir\file.txt`; at a ZERO WIDTH
SPACE; at a comma, point, currency sign or percent inside a number; before an em dash.

**Consequence:** no `linebreak` dependency, no Unicode line-break table, and no `LineBreak.txt`
fetched from anywhere. The classes that matter are the space, four characters that break after
themselves, two that glue, East Asian script (from the `\p{Script=…}` escapes the engine already
carries, plus two ranges for CJK punctuation and the width forms), and the kinsoku sets. This is the
one place where the measurement made the work _smaller_.

### 2. `latinLnBrk` absent behaves as `false`

The plan predicted this and T3 settles it. On `ab antidisestablishmentarian` across 24 box widths, a
paragraph stating nothing lays out identically to one stating `0` and differently from one stating
`1` at every width. ECMA-376 gives the attribute a default of `true`.

`@hangingPunct` absent behaves as **`true`** — ECMA does not say. `@eaLnBrk` is **inert**: all three
of absent, `1` and `0` lay out identically on a nine-ideograph string at every width. It is accepted
and ignored, and the interface says so, because a caller passing it should not quietly get a
different rule than they think.

### 3. Kinsoku is built in, and the file usually cannot reach it

This is the largest deviation from the plan, and the falsifier that found it was designed in from the
start: a deck stating **no** `p:kinsoku` at all.

| comparison                              | probes differing |
| --------------------------------------- | ---------------- |
| standard sets stated vs none stated     | **0**            |
| strict off, stated sets = the built-ins | 0                |
| strict off, novel sets                  | 10               |
| strict off, **empty** sets              | 8                |
| a Latin kinsoku on English text         | 0                |
| the same with strict off                | 0                |

So: the built-in list applies **unless** a package both states its own sets _and_ sets
`strictFirstAndLastChars="0"`, and then the file's sets **replace** the built-ins entirely. There is
no hard core — with empty sets, strict off and hanging off, every position in every East Asian probe
became a break, including before an ideographic comma. An intermediate reading that the comma and
full stop were hard-coded survived two decks and was refuted by the third.

A renderer implementing the plan as written — read `p:kinsoku`, apply those sets — would produce the
right answer on the rare deck that carries the element and **no kinsoku at all** on the overwhelming
majority that do not.

The filter is also **gated**: it reaches only opportunities the East Asian rule created. A package
forbidding `b` at the start of a line and `a` at the end changes nothing about English text. The
ungated reading, which is what ECMA supports, scores 2162/2179.

### 4. The built-in sets are measured, not transcribed

Each of 100 candidates was placed between two runs of ideographs — which break freely — and the two
positions either side were read back: **71 may not begin a line, 19 may not end one**, and six
ordinary characters included as controls all came back unrestricted, which is what makes the other
readings mean anything.

The widely repeated version of these sets, including the comment in this repository's own
`a10-rtl-cjk` corpus generator, is **wrong in four places**: it lists `¢` (U+00A2), `£` (U+00A3), `¥`
(U+00A5) and `＃` (U+FF03) as restricted, and PowerPoint restricts none of them — while restricting
the fullwidth `￡` (U+FFE1) and `￥` (U+FFE5) that the list omits. Halfwidth and fullwidth currency
signs were confused for each other.

The table is generated into `packages/text/src/lines/kinsoku.ts` by
`tools/ground-truth/text/line-breaks/write-tables.ts`, and `break.test.ts` re-derives it from the fixture, so it
cannot drift from the measurement.

### 5. The fit test has three rules, not one

- **Trailing spaces are never charged**, however many, and regardless of `hangingPunct`. A box sized
  for four code points returns `"aaaa "` — five.
- **`hangingPunct` (default on) exempts one trailing comma or full stop**, and exactly six characters
  hang: `,` `.` `、` `。` `，` `．`. Nothing else does — not the closing brackets, not the small kana,
  not the prolonged sound mark, all of which are start-forbidden and force a fallback instead. The
  plan's phrase "a trailing hangable's advance" was right; the set is much smaller than it sounds.
- **A soft-hyphen break pays for the hyphen it draws.** At a box sized for exactly the code points up
  to and including a soft hyphen, PowerPoint refused the break and fell all the way back to the
  space. The soft hyphen has no advance of its own, so the only thing that can have overflowed is the
  hyphen it rendered. This is the one place a break costs more than the text it contains, and it was
  found as a single miss in an otherwise-perfect score.

### 6. Two ranges of "East Asian" that Unicode's script property does not give

`isEastAsian` is the gate on both the ideograph break rule and the kinsoku filter, and the script
escapes the engine already carries — `Han`, `Hiragana`, `Katakana`, `Hangul` — do not cover
everything that behaves East Asian. Two ranges are added, and both were survivors of the mutation
sweep before they were measured:

- **CJK Symbols and Punctuation** (U+3000–U+303F). Every other East Asian probe has an ideograph
  next to the character under test, so the classifier had only ever been asked about `、` in
  company. `aa 、bb。cc` has none: measured, the position after the ideographic comma **is** an
  opportunity and the one before the full stop is **not**, so the classifier has to reach these
  characters on their own.
- **Halfwidth and Fullwidth Forms** (U+FF00–U+FFEF). Unicode calls fullwidth `Ａ` `Script=Latin`, so
  nothing but the range makes it East Asian. Measured: `aa ＡＢＣbb` breaks between every pair of
  fullwidth letters and not between the two ASCII ones that follow.

### 7. Emergency breaking is unconditional

When no opportunity exists at or before the fit limit, the line takes exactly as many code points as
fit. Measured on a 25-letter word with no opportunity anywhere: it broke at every one of 24 box
widths in **all three** `latinLnBrk` variants, including `"0"`. The same word with a space in front of
it never broke internally in the `0` and absent variants, so the rule is "force only when there is no
alternative", not "force when narrow".

A forced break still takes the spaces that follow it. Measured on `aa 、bb。cc` in a two-character
box, where the ideographic comma may not begin a line so no opportunity was available at all:
PowerPoint returned `aa ` — three code points — rather than `aa`. Hanging a trailing space is a
property of spaces, not of the break that produced them. This was the single miss in an otherwise
perfect score, and it only showed up because the kinsoku rules made a probe with a space in it run
out of opportunities.

## Incidental findings

**A `p:kinsoku` whose `@lang` is not East Asian makes PowerPoint refuse the package.** The
`latin-custom-lang` deck differs from `latin-custom` in exactly one attribute — `lang="en-US"` rather
than `ja-JP` — and PowerPoint reports _"The file or directory is corrupted and unreadable"_ and
declines to repair it. This is a validator rule for 1.2 and a writer constraint: it is schema-legal
markup that PowerPoint rejects, which is exactly the class of failure `cli bisect` exists for.

**Arabic prefix widths are not monotone.** `مرحبا` measures 45.3pt where `مرحب` — one letter shorter —
measures 51.0pt, because the fifth letter joins the fourth into a shorter connected form. The first
implementation binary-searched for the widest fitting prefix; the single Arabic probe caught it. The
search is now a scan, bounded at twice the box width, and the assumption is written down where it is
relied on.

**PowerPoint segments Thai with a dictionary.** It breaks `สวัสดี` from `ชาวโลก` — two words — and
nowhere else in a twelve-character string. The rule has no dictionary, no JavaScript library carries
one, and the plan's risk table already says Thai may over-run. It is held out of the scored set and
given its own section, with the miss count pinned so the gap fails the suite if it grows: 5 of 9
probes.

## Verification

`pnpm check` green end to end. The centrepiece is a **replay**: the fixture carries the prefix widths
that sized every probe box, so the whole experiment runs again in the suite with no browser and no
PowerPoint, and the implementation has to produce the same line 1 on all 2212 scored probes.

**Mutation sweep: 31 of 32 killed.** Four of the first sweep's five survivors were missing
assertions, and two of those were missing _measurements_ rather than missing tests — the sweep is
what said so:

- "NO-BREAK SPACE stops gluing" survived because its suppressing behaviour was generalised from the
  non-breaking hyphen. Measured instead: one probe variant, question closed, mutant dead.
- "CJK punctuation / the width forms stop being East Asian" survived because every East Asian probe
  had an ideograph as a neighbour. Three new probe strings, both mutants dead, and one of them
  turned up the trailing-space rule for forced breaks above.
- "the end of the text is never a candidate" and "a failed soft-hyphen candidate ends the scan" were
  ordinary missing assertions and now have them.

The one survivor is **genuinely equivalent**: bounding the trailing-space scan at the start of the
line rather than at the start of the text. Every break, chosen or forced, absorbs the space run that
follows it, so no line but the first can begin with a space — and for the first the two bounds are
the same number. The bound stays because it is a helper's job not to read another line's characters,
whatever the caller's invariants happen to be today.

## Consequences

- `@pptx-studio/text` gains `break.ts` and a generated `kinsoku.ts`, and still has **no runtime
  dependencies**. The plan's `linebreak` dependency is not taken, so `packages/text` remains as
  dependency-free as `opc` and `xml`.
- 3.4's autofit ladder can measure and break with the same engine, which is the property the plan's
  bet rests on.
- 3.6 will need the hanging set for alignment: a hanging character is outside the measure, so a
  centred or right-aligned line must ignore it. `LineBox.measuredEnd` carries that.
- `packages/validate` should gain a rule for `p:kinsoku/@lang` in 1.2.

## Open questions

1. **Is the kinsoku gate on script or on language?** `latin-custom` states `lang="ja-JP"` sets naming
   Latin characters on `en-US` runs, and they have no effect — which both readings explain. The
   `lang="en-US"` variant that would separate them is refused by PowerPoint outright, so the question
   may not be answerable from the file at all. Implemented as a script gate, which needs no language
   plumbing and reproduces every measurement.
2. ~~Does a NO-BREAK SPACE suppress the character-level break under `latinLnBrk="1"`?~~ **Closed.**
   It does, on both sides, exactly as the non-breaking hyphen does. It was a generalisation until a
   surviving mutant said so, and then it was measured: one more probe variant, and both the mutant
   and the question died.
3. **Which of the two positions around an em dash does a _wider_ context open?** Measured only with
   the dash between two Latin words. B2's two-sidedness may reappear next to a space.
4. **What happens to a break opportunity that falls inside a `a:br` or between two runs?** T3 measured
   single-run paragraphs throughout. 3.8 needs the answer when the two renderers are asked to agree.
5. **Is the greedy rule greedy at every width, or does PowerPoint ever balance?** Every probe here is
   a first line. A paragraph whose last line is one word is where a balancing breaker would differ.
6. **Does `latinLnBrk="1"` break between a letter and a combining mark?** The Thai probes suggest not,
   but they are confounded with the dictionary.

Carried from 2.10, 2.11, 3.1 and 3.2: the last-line height term (3.6), the `a:blipFill` scope gap,
and Gate 2's LibreOffice heatmap.
