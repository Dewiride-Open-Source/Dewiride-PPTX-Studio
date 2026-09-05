# 0031 — Bullets, fields and script runs

Sub-phase 3.5. Status: accepted. Supersedes nothing; extends
[0027](./0027-the-text-cascade.md), [0028](./0028-measurement-and-the-line-model.md) and
[0030](./0030-autofit.md).

## The problem

Three questions the file does not answer.

**What does `a:buAutoNum type="romanLcPeriod"` draw?** ECMA-376 gives each of the 41
`ST_TextAutonumberScheme` values a one-line gloss and no sequence. It does not say what follows `z`,
which character stands for zero in the East Asian schemes, or what happens at 900.

**What number does it draw?** A six-item numbered list PowerPoint authors itself carries **no
`startAt` anywhere**. Every number on every slide is computed at layout time, by a rule that exists
nowhere in writing.

**Which typeface draws each character?** A run may name four — `a:latin`, `a:ea`, `a:cs`, `a:sym` —
and PowerPoint picks one per character. Nothing in the file says which.

## The method: read it, do not fit it

Every earlier experiment here fitted a model to a measurement. T5 mostly does not need to, because
**`Slide.Export(path, "EMF")` records PowerPoint's own GDI calls**, and a bullet is drawn by
`ExtTextOutW` with the characters in it. So the answer comes back as the string itself — `"(iv)"`,
`"a."` — along with the face, the size and the colour it was drawn in. There is no threshold and no
candidate to enumerate for identification, which is what made T3 a stronger experiment than T2 and
is the same advantage here at a much larger scale.

PowerPoint has no SVG converter installed and refuses that filter by name; EMF and WMF both work,
and EMF is the one with 32-bit coordinates and `ExtTextOutW` rather than the ANSI-only original.

Three instruments, failing differently:

| instrument                | answers                               | cannot answer            |
| ------------------------- | ------------------------------------- | ------------------------ |
| the EMF's text records    | the characters, face, size and colour | where anything landed    |
| `Paragraphs(i).BoundLeft` | positions in points, exactly          | what the characters were |
| `TextRange.Text`          | a field's value in any locale         | anything about a bullet  |

The third earns its place: four of the sixteen field locales are complex scripts, which PowerPoint
**shapes before drawing**, so their EMF records hold glyph ids rather than characters. The same is
true of two autonumber schemes, and those are recorded as an open gap rather than guessed at.

**The authoring step settled the vocabulary before a probe existed.** `PpNumberedBulletStyle` was
walked from 0 to 47; 0 to 40 were accepted and each wrote its own `a:buAutoNum/@type` into a saved
package. So the 41 scheme names and their order are PowerPoint's own, and `tools/ground-truth/text/bullets/analyse.ts`
re-reads that package and throws if the table disagrees.

## What was measured

6849 probes across 44 packages, 624 slides exported. Twelve questions, each scored against every
reading a person would plausibly implement, and in each case exactly one candidate fits everything:

| question                                 | rows | measured | nearest rival |
| ---------------------------------------- | ---- | -------- | ------------- |
| what a scheme renders                    | 5097 | 5097     | 5087          |
| the numbering pass                       | 15   | 15       | 13            |
| the private-use mapping for `buChar`     | 21   | 21       | 20            |
| which face draws the bullet              | 20   | 20       | 18            |
| what size it is drawn at                 | 15   | 15       | 14            |
| what colour it is drawn in               | 7    | 7        | 5             |
| what box a picture bullet occupies       | 11   | 11       | 9             |
| where the text starts on a bulleted line | 108  | 108      | 90            |
| where a wrapped line starts              | 6    | 6        | 5             |
| which script slot draws a character      | 21   | 21       | 20            |
| which face when the slot is absent       | 47   | 47       | 46            |
| whether a field shows its cached text    | 290  | 290      | 0             |

## The decisions

### 1. Autonumber is a numeral system and a decoration, and both are measured

The 41 names are compositional — nineteen numeral systems by five decorations — so the split is a
function rather than a 41-row table nobody could check. Four things a careful implementation gets
wrong:

- **Alphabetic numbering is not bijective base-26.** After `z` comes `aa`, then `aaa` at 53: the
  letter is `(n−1) mod A` repeated `floor((n−1)/A)+1` times. The bijective reading scores
  **1396 of 5097**.
- **The East Asian zero is U+25CB WHITE CIRCLE**, not U+3007 IDEOGRAPHIC NUMBER ZERO. They are
  indistinguishable on screen and a table of Chinese numerals gives the wrong one. 5001 of 5097.
- **The three East Asian families have three different rules.** Simplified keeps the ten-form to 99
  (26 is 二十六), Traditional only to 19 (26 is 二六), Japanese and Korean not at all (10 is 一○).
  One shared rule scores 5069.
- **The Wingdings circled numbers wrap at ten**; the Unicode circled digits run to twenty and then
  fall back to ASCII. The two behave oppositely, which is why both were probed to 4000.

### 2. The alphabets carry three measured parameters each, and one of them is an anomaly

`STARTS` was eighteen well-chosen values, and on eighteen values the alphabetic schemes look simple —
right at 27, 50, 100 and 400, wrong at 900, 1000, 3999 and 4000, where the repeat count collapses
from 35 to 5. Four anomalies is not a rule and neither is the model they break, so `alphaLcPeriod`
was swept **one value at a time to 1300** and the others to 800.

- **The counter wraps.** At 781 the Latin count resets to one, so the modulus is 780 — which is
  26 × 30, a Latin number that the Thai and Devanagari alphabets inherited even though their own
  maximum counts come out at 20, 49 and 23 rather than 30. Hebrew wraps at **392**, which no
  arithmetic on 22 produces.
- **Thai has an off-by-one nobody else has.** Its bands begin at 42, 82, 123, 164 — that is 41k for
  every band after the second — while Latin's begin at 26k+1 and Devanagari's at 16k+1. It is a
  measured parameter because nothing in the shape of the rule predicts it.
- **Hebrew fills with the last letter**, not the current one: 26 is tav-dalet where Latin would give
  dalet-dalet.
- **Thai uses 41 of its 44 consonants**, skipping U+0E03, U+0E05 and U+0E06. Only walking shows that.

None of this is reachable by a real deck. It is here because a rule stated for `startAt` has to hold
across `ST_TextBulletStartAtNum`, and 4000 was in the original sweep whether or not anybody wanted it
to be.

### 3. A symbol bullet is mapped through the ANSI codepage, not the code point

`buChar char="•" buFont="Wingdings"` draws **U+F095**, a filled circle. Adding 0xF000 to U+2022 gives
U+F022, an envelope. The rule is: take the character's **Windows-1252 byte** and add 0xF000 — and
U+2022 is byte 0x95 in Windows-1252 and nothing at all in Latin-1.

This matters more than its one line suggests, because `•` is by a distance the most common bullet
character there is. The low-byte reading scores 20 of 21 and the Latin-1-only reading 19 of 21; both
are wrong on exactly the character every deck uses.

A text face is left alone: `char="§"` in Arial draws a section sign. Which faces are symbol faces is
a **list of the five this experiment measured**, not a rule, because the general answer is a (3,0)
symbol subtable in the font's `cmap` and reading that is 8.1's job.

### 4. `a:buFont` is inert for an autonumber; `a:buClr` and `a:buSzPct` are not

Measured on twenty probes, including ones where `buFont` named Wingdings and the run named Courier
New: **a number is drawn in the first run's face, always**. PowerPoint's own UI writes
`buFont="+mj-lt"` on every list it numbers, and that attribute changes nothing.

The obvious generalisation — that all three decorations are decorative for a number — is wrong.
`a:buClr` reaches an autonumber and so does `a:buSzPct`. That is why the three merge as three slots
and not as one.

"First run" is load-bearing and was probed with two runs of disagreeing sizes in both orders: the
bullet follows the **first**, not the largest, and not the paragraph's `a:defRPr`.

### 5. The numbers are not in the file; the run is what has to be reconstructed

A paragraph's number is `startAt` plus the count of preceding paragraphs in the same run, where a run
is **consecutive paragraphs at the same level sharing a scheme and a `startAt`**, deeper levels pass
through it, and a shallower one ends it.

The two that decide it, and both are one probe each:

- `1, 2, [1, 2], 3, 4` — the outer list resumes at three after a nested one, so the walk **skips**
  deeper paragraphs rather than stopping at them.
- `1, [1], 2, [1]` — the inner list restarts every time it is re-entered, because the outer paragraph
  between the two inner ones ends the inner run.

And one that is easy to miss: `startAt="7"` followed by three plain paragraphs renders **7, 1, 2, 3**,
not 7, 8, 9, 10. Repeat the `startAt` on every paragraph and it is 7, 8, 9, 10. An absent `startAt`
compares equal to `1` for this purpose, which is why the model keeps them apart in the type and equal
in the comparison.

### 6. Where the bullet goes: two terms a plausible implementation omits

`textLeft = max(marL, −indent, bulletLeft + advance)` with `bulletLeft = max(0, marL + min(0, indent))`.

- **A positive `indent` does not move the bullet.** Only a hanging one does — hence `min(0, indent)`.
  Eighteen probes at `indent="12"` put the text exactly where `indent="0"` does.
- **The bullet is clamped to the frame.** A hanging indent deeper than `marL` would put it left of
  the text box; PowerPoint pins it at zero and the text follows, which is why `marL="0" indent="-36"`
  starts the text at 36 rather than at the bullet's own width.

A wrapped line is at `marL` exactly — no bullet, no clamp — and was scored as a separate question.
`defTabSz` has no effect: the gap after a bullet is an advance and not a tab.

A picture bullet is **0.7 × the font size** tall and follows its source's aspect ratio, confirmed at
three sizes and three ratios. Letterboxing it into a square puts the text in the wrong place on
anything that is not square.

### 7. Script runs, read rather than inferred

The EMF makes this the cheapest family here. GDI is asked for one face at a time, so a run whose
characters resolve to different faces comes out as several records, each naming its face. Private use
goes to `a:sym`; Hebrew, Arabic, Thai and Devanagari to `a:cs`; kana, Han, Hangul, fullwidth forms and
CJK punctuation to `a:ea`; everything else — including Greek, Cyrillic, general punctuation, currency
and maths — to `a:latin`.

The two readings worth naming: **`a:cs` is not "everything non-Latin"** (that scores 8 of 21), and
**`a:sym` is used, for the private-use area only** (never using it scores 20 of 21, which is close
enough to look right on a whole deck and wrong on every Wingdings run in it).

An absent slot falls **straight to `a:latin`**, measured on the rows where PowerPoint drew in a face
the probe actually named — the decisive one being a private-use character with `a:sym` dropped, drawn
in Georgia directly.

## What the plan got wrong, and what replaces it

The plan specified for fields: _"14 reserved `a:fld` types with `Intl.DateTimeFormat` keyed on
`@lang`, falling back to the cached text."_ Three of those four clauses did not survive.

- **There are eighteen reserved types, not fourteen.** `datetime1` through `datetime13`,
  `datetimeFigureOut`, `slidenum`, ECMA's bare `datetime`, and — undocumented in any enumeration this
  project has read — **`uaqdatetime1` and `uaqdatetime2`**, which are the Umm al-Qura calendar.
  PowerPoint's own UI writes them for `PpDateTimeFormat` 15 and 16.
- **`Intl.DateTimeFormat` does not reproduce PowerPoint.** Scored against 97 measured readings it
  agrees on **25**. What PowerPoint does is format with **Windows' own per-locale patterns**, which
  is why `datetime3` is "5 September 2026" in one locale and "05/09/26" in the next, and why
  `datetime12` in German is `"12:56 "` — a pattern ending in an AM/PM designator that German defines
  as the empty string.
- **The cached `a:t` is not what a reader displays.** 290 readings of decks whose caches were
  overwritten with `#stale#` showed the real value every time; one reopened four minutes after it was
  written showed `12:04 PM` where the file said `12:02 PM`. The cache is a fallback, never the answer.
- **The fallback clause is right**, and it is what an _unreserved_ type gets: a field typed
  `notAReservedType` rendered its cached text exactly.

So the patterns are a **measured table**, generated by `tools/ground-truth/text/bullets/write-tables.ts` from readings in
sixteen locales, and `renderField` runs a three-rung ladder: `slidenum` from the slide's position, a
reserved type with a measured pattern recomputed, everything else from the cache with a diagnostic.

### Deriving the patterns needed three guards, and each was paid for by a wrong answer

A pattern is derived by substituting the components of the instant the deck was read and then
**re-rendering it and comparing against the string it came from**. That round trip alone is not
enough, because one instant cannot tell a real component from a coincidence:

1. **A numeric token must match a whole run of digits.** The Umm al-Qura year 1448 derived as `h448`
   — the hour was one o'clock, `h` renders "1", and re-rendering put the 1 back.
2. **The finished pattern must hold no digits.** A leftover digit is a component frozen into the
   pattern, which is what happens for every non-Gregorian calendar.
3. **A literal letter is refused outright.** ICU says Korean's afternoon marker is `PM`; Windows
   writes 오후, so the substitution never happens, the marker survives as literal text, and the
   pattern round-trips perfectly while saying "afternoon" at nine in the morning.

The third guard is conservative and drops cells that are probably right — the Chinese year, month and
day markers and the Russian abbreviation for "year". **172 cells across 14 locales** survive it, all
provably safe; 118 readings are recorded as gaps with the reason. Closing them needs Windows' own NLS
tables or a second measurement at a different time of day, and both are outside 3.5.

## What is deliberately not here

- **`arabic1Minus` and `arabic2Minus` throw.** They are real schemes and PowerPoint shapes them into
  glyph ids before drawing, so the EMF holds indices into a font and no text. The repeat structure is
  still visible and is kept in the fixture; the letters are not. Guessing them from the Unicode block
  would produce a sequence that renders plausibly and has never been checked against anything.
  Reversing the `cmap` is 8.1's.
- **`packages/model` does not depend on `packages/text`.** The bullet is described twice —
  `BulletSize` in the model says what the file states, `BulletSize` in `text` says what a renderer
  needs — following the precedent `line-model.ts` set for `Spacing`. Coupling them would also have
  meant a `package.json` change.
- **A multi-character `a:buChar`.** PowerPoint draws `char="ab"` as `"DE"` and `char="xy"` as `"[\"`
  — every character shifted down by 29, consistently across four probes and six characters. Its own
  writer never produces more than one character, ECMA describes one, and reproducing an offset that
  looks like a defect would be reproducing a defect. The measurement is in the fixture; the
  implementation renders the string as written and says so.

## Open questions

1. **The two Arabic alphabetic schemes**, above. 8.1.
2. **118 date-pattern cells**, above. Needs Windows NLS data or a second measurement at another hour.
3. **Why 780, and why 392?** 780 is 26 × 30 and 392 is nothing recognisable × 22. Both are measured
   to fit perfectly and neither is explained.
4. **Thai's bias.** Measured, unexplained, and the only alphabet that has it.
5. **`hindiAlphaPeriod` above 768** truncates: twelve readings where the scheme's own full stop never
   appears and the length wanders between 47 and 49 for a value the rule says is constant. Excluded
   from scoring by an objective filter — the decoration the scheme name promises is absent — and kept
   in the fixture.
6. **Whether `a:buChar`'s mapping is the font's charset or its `cmap`.** The list of five symbol faces
   is what was probed. PowerPoint writes `charset="2"` on the `a:buFont` it produces, which may be the
   real signal; these probes wrote no `@charset` and the mapping happened anyway, so the font itself
   is being consulted. 8.1.
7. **Alignment.** `ind-algn-*` shows a centred paragraph moves the bullet with the text, so the bullet
   is part of the aligned line — but the probes carry a marker run that contaminates the width, so
   nothing quantitative is claimed. 3.6 owns alignment.
8. **A substituted face may be rescaled.** Courier New has no Thai, and the Thai probe came back in
   Leelawadee UI at 1.4 times the size it asked for. Recorded, not modelled. 3.7.

## Verification

- 6849 probes, 44 packages, 624 slides. Nine packages repaired, all of them deliberately hostile:
  four out-of-range `startAt` values and five out-of-range `buSzPct` values, one probe per package so
  each refusal names its own cause.
- 5658 scored rows across twelve questions, every one fitting exactly one candidate.
- 59 tests replaying the fixture, including the refuted alternatives.
- **Mutation sweep: 41 mutants, 40 killed.** The survivor — walking the string by code unit instead of
  by code point — is genuinely equivalent, because every range in the script tables is inside the
  Basic Multilingual Plane. It is documented as such in `script-runs.ts` rather than deleted.
- `pnpm check` green end to end.
