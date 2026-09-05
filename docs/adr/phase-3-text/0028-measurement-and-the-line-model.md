# ADR 0028 — Measurement and the line model

**Sub-phase 3.2.** Status: accepted. Adds `packages/text`, the first package at layer 1 since 2.8.
Confirms the plan's central claim, and finds four rules the plan does not state — one of which
changes the line count of any deck whose line spacing is not a whole percent.

Code: `packages/text/src/{line-model.ts, measure.ts, errors.ts}`.
Measurement: `corpus/ground-truth/text-metrics.json`, from
`tools/ground-truth/{text-metrics.ts, tools/ground-truth/text/metrics/build-deck.ts, tools/ground-truth/text/metrics/read.ps1, measure-in-browser.ts, tools/ground-truth/text/metrics/analyse.ts}`.

---

## What this sub-phase had to decide

The plan states, without evidence, that PowerPoint's line height is a font-independent **1.2 × font
size**. Everything downstream rests on it: line breaking (3.3) needs to know how many lines fit,
autofit (3.4) searches a ladder of sizes against a box, and anchoring (3.6) needs a block height.
A wrong constant here is wrong on every line of every slide.

It is also the kind of claim that is cheap to confirm and worthless to confirm. So the probe table
was built the other way round: **a wrong model has to predict a different number.**

### The design, and why the increment rather than the height

`TextRange2.BoundHeight` is the line box plus whatever the first and last lines carry. Read once it
is a sum nobody can decompose. `Lines(i, 1).BoundTop` is the top of line _i_, so the difference
between two of them is the **baseline-to-baseline advance directly**, with every constant term
already cancelled — and cancelled whether or not the constant is what anyone expects.

Extra lines are made with `a:br`, not with extra paragraphs, because a paragraph carries
`spcBef`/`spcAft` and a measurement that needs two things zeroed to be true is a measurement of two
things. A smaller `paraLines` set does use paragraphs with the spacing explicitly zeroed, so that
"a line is a line however it was made" is checked rather than assumed. It is: 43.2pt either way, at
every face and size.

Thirteen typefaces, chosen because their vertical metrics disagree violently — Impact is tall and
condensed, Verdana wide with a huge x-height, Courier New a monospace with a small one. A
fourteenth entry names a face that does not exist; it is not padding, and what it found is below.

**2091 probes across five packages**, every one opened by PowerPoint without a repair.

---

## The plan's claim is right, and stronger than it says

| what                          | measured                                                  |
| ----------------------------- | --------------------------------------------------------- |
| line advance at 100%          | exactly `1.2 × size`, all 13 faces, all 7 sizes           |
| worst spread across the faces | **0.00003pt**                                             |
| a one-line block              | `1.2 × size` — there is no first-line padding term at all |
| where the block starts        | the frame top **exactly**, at every spacing from 50%–300% |

The sizes include a fractional 10.5pt, because `@sz` is hundredths of a point and every
implementation that rounds it to a whole point somewhere still passes every whole-point test.

This is why `lineAdvance` takes no font argument: there is nowhere to pass one. It is also the whole
reason `line-height: normal` must never reach a slide — a browser resolves it from `hhea.lineGap` on
macOS and `OS/2.usWinAscent` on Windows, so the same markup gives different line counts per machine.

---

## Four rules the plan does not state

### 1. Percent multiplies the line box; exact points does not

`a:lnSpc/a:spcPct val="150000"` on an 18pt run is **32.4pt**, not 27pt. The percentage multiplies
the 1.2 line box, not the font size. `a:spcPts` is used **raw**, with no 1.2 anywhere.

Nothing in the file distinguishes these readings and they differ by 20% on every line.

| model                                                 | fits    |
| ----------------------------------------------------- | ------- |
| percent × 1.2 × size; exact points as given           | 463/463 |
| percent × 1.2 × size; exact points also scaled by 1.2 | 397/463 |
| a flat 1.2 × size, ignoring `a:lnSpc`                 | 244/463 |
| percent × size                                        | 66/463  |

### 2. The percentage is rounded half up to a whole percent, before use

`@val` is in thousandths of a percent. PowerPoint quantises it: **99.5% lays out as 100%, 100.4% as
100%, and 100.5% as 101%.**

The half-percent points found the quantisation but could not name it — round-half-up and ceiling
both send 99.5 up and 100.5 up. A second sweep at a fifth of a percent separated them, measured on
the advance between two baselines rather than through any model:

| model                            | fits    |
| -------------------------------- | ------- |
| round half up to a whole percent | 371/371 |
| ceil to a whole percent          | 365/371 |
| floor to a whole percent         | 359/371 |
| the value as given               | 353/371 |
| round to a tenth of a percent    | 353/371 |

This is not a curiosity. PowerPoint's own dialog writes whole percentages, but other producers write
`val="106667"` for "1.07 lines" — which lays out as 107%, not 106.667%.

### 3. `@spc` is absolute points, applied after every character including the last

Not em-relative: the same `@spc` widened a ten-character string by the same number of points at 12pt
and at 32pt. And it is applied _after_ every character, so ten characters get ten gaps — 18/18
against 0/18 for the between-characters reading. Chromium's `letterSpacing` does exactly the same,
which is why it is used directly rather than corrected.

### 4. `@kern` is an inclusive minimum size, and `kern="0"` means never

ECMA-376 says `@kern` is "the minimum font size at which character kerning occurs", and every
implementation reaches for a boolean anyway.

| reading                        | fits  |
| ------------------------------ | ----- |
| `kern > 0 && sz >= kern`       | 12/12 |
| `kern > 0 && sz > kern`        | 11/12 |
| non-zero means on, at any size | 8/12  |
| present means on               | 5/12  |

The inclusive/exclusive pair differ on exactly one probe — a 12pt run with `kern="1200"`, which _is_
kerned — which is why the table crosses the threshold with the size instead of sitting beside it.
`kern="0"` meaning never is not in ECMA at all.

Only Arial has kern pairs for the test string; Georgia and Consolas showed no difference at any
setting. Two of three faces being silent is why three were measured.

---

## Three things found while measuring, that the experiment was not looking for

### `BoundWidth` includes a trailing paragraph mark

`TextRange2.BoundWidth` is wider than the string it covers, by exactly the face's **space advance**,
because the range includes the paragraph mark. Compared raw against a browser it reports a **5–20%**
disagreement that is entirely an artefact of the reader.

This was caught by the `@spc` probes: a ten-character string widened by _eleven_ times the spacing.
Two extra probes per face (`H` and `HH`) measure the mark without needing to know what it is —
`HH − H` is one `H` advance with the mark cancelling, and `H` minus that advance is the mark. Every
advance in the fixture has it subtracted.

Recorded because a later sub-phase re-reading `BoundWidth` and not knowing this would conclude the
renderer is 7% wrong and go looking in the wrong place.

### A file that states no `@kern` still gets one

ECMA says an omitted `@kern` kerns at every size. It does not: an 8pt run in a file stating no
`@kern` anywhere came back **unkerned**, and a 12pt one kerned.

The reason is 3.1's tenth source. A master that declares no `p:txStyles` gets PowerPoint's built-in
ones, and those carry `kern="1200"`. So the attribute is never really absent, and a resolver must
read `@kern` **through the cascade**; `undefined` at the measurement layer means the cascade
genuinely yielded nothing, and only then does the ECMA default apply.

### PowerPoint substitutes an unknown Latin face with Calibri

The fourteenth face in the table does not exist. Of fourteen faces there are **thirteen distinct
advance rows**, and the shared one is the absent face measuring exactly as **Calibri**.

`document.fonts.check` returned **true** for that face. This is the plan's own warning about 3.7
arriving early, and confirms the design it prescribes: presence has to be detected by an advance
fingerprint, because the browser's own answer is not one.

---

## How close a browser gets

The deliverable is a measurer over `OffscreenCanvas.measureText`, so the number that matters is how
far Chromium is from PowerPoint on the same face, size and string. Over **364 comparisons** on
installed faces, with the paragraph mark removed:

|        |                         |
| ------ | ----------------------- |
| median | **0.149%**              |
| p95    | 0.702%                  |
| max    | 11.1% — Cambria at 12pt |

Twelve of thirteen faces have a mean under 0.37% and a worst under 1.6%. The outlier is small-size
hinting: PowerPoint grid-fits 12pt Cambria and Chromium does not, and no amount of correctness here
closes that.

The plan's risk table promises "a tracked, improving number — never pixel-perfection". This is that
number, it is in the fixture, and the suite fails if it regresses.

Two things had to agree in _kind_ or the measurer could not be built on a canvas at all, and both
do: Chromium applies `letterSpacing` after every character including the last, and the faces
Chromium kerns are the faces PowerPoint kerns.

---

## The last line, which is committed as data and not as a rule

A block is `H(n) = (n − 1) × advance + lastLineHeight` — 463/463. Every line box **except the last**
measures exactly the advance. The last one does not, and it is the only place in this sub-phase
where a font-dependent term appears.

The best model found is

```
advance <= 1.2 × size ? max(advance, 0.75 × advance + descent) : 0.75 × advance + descent
```

with `descent` a per-face, per-size constant. It fits **1548 of 1579** one-line probes, against
1027 for `max(advance, 0.75 advance + descent)` and 897 for "the advance itself". All **31 misses
are Courier New**, in the 101%–106% band — and Courier New is the one face whose `descent` is
exactly `0.3 × size`, which is the single value that makes the two branches of the rule meet at
100% instead of stepping. The rule is degenerate exactly where it fails.

`descent` is **not** any quantity `TextMetrics` reports. Chromium's `fontBoundingBoxDescent` is
quantised to whole pixels — 0.250, 0.222, 0.219 em for Arial at 12/18/32pt — while the measured
term is a clean 0.2279 em at all three. Identifying it needs to read the font's own tables, which is
sub-phase 8.1's SFNT reader, and which this repository may not do to a system font on this machine.

So it is committed as a **measured table with its score and every probe it misses**, clearly
separated from the fitted rules, and `blockHeight` takes it as an argument rather than guessing. A
plausible guess here is exactly the silent wrong answer the engineering standards forbid. Anchoring
(3.6) is where it has to be finished.

> **Closed in 3.4, two sub-phases early.** Autofit's fit test is a comparison against this number,
> so [0030](./0030-autofit.md) had to measure it: `0.75 × advance + b × size`, floored at the advance
> while the advance is inside `C × size`, with the 0.75 identical across six faces and `b` and `C`
> properties of the typeface. Arial's `b` is 0.227619, which is the 0.2279 em estimated above. The
> conclusion that no `TextMetrics` quantity supplies it stands and is now measured at 4096px rather
> than inferred from quantised readings: Arial and Verdana report `fontBoundingBoxDescent` 0.0026
> apart while their coefficients differ by 0.0203. `blockHeight` still takes the term as an
> argument, for the reason given above.

---

## What is deliberately not here

- **Line breaking.** 3.3. Every probe deck sets `wrap="none"` on purpose: if the shape width could
  break a line, every width here would be a measurement of the breaker.
- **Autofit**, 3.4. **Bullets and fields**, 3.5. **Anchors and insets**, 3.6.
- **`packages/text` has no consumer yet.** That is the plan's staging, not speculative generality:
  3.3 through 3.8 all read it, and it carries 498 tests of its own.
- **Script runs and complex text.** The strings are Latin. CJK, bidi and Thai are 3.3 and 3.5, and
  the plan already says Thai may over-run.

---

## Verification

- **2091 probes**, five packages, all opened by PowerPoint with no repair.
- Every committed rule scored against its alternatives, with the analysis **throwing rather than
  emitting a fixture** two candidates fit or none does. It threw twice while being written: once
  when a least-squares fit absorbed the Courier New disagreement it was supposed to expose, and once
  when the half-percent sweep left round-half-up and ceiling tied.
- The advance is font-independent to a worst spread of **0.00003pt** across 13 faces.
- **`pnpm check` green**: layering (12 packages), corpus (6 manifests, 86 entries, 83 files),
  format, lint, five typecheck projects, 13 builds, **roundtrip 52/52 decks / 1474 parts**, 25
  pkg:qa, **77 test files / 2693 tests**.
- **Mutation sweep: 23 of 23 killed.** One survived the first run — a measurer that computes the
  kerning flag correctly and then sets `fontKerning = 'normal'` regardless — because every kerning
  test checked the predicate in isolation and none drove it through the canvas. That was a missing
  assertion, not an equivalent mutant, and the test that closes it also pins the _magnitude_ of the
  switch against PowerPoint's own two readings.

---

## Open questions

1. **What is `descent`?** A clean per-face ratio (Arial 0.2279 em, Courier New 0.3000, Verdana
   0.2075, Consolas 0.2571) that matches no `TextMetrics` value. 8.1's SFNT reader can answer it for
   a font the user supplies. Until then 3.6 uses the table.
2. **Why does Courier New keep the `max` branch to 106%?** The residual above. It is one face in one
   6% band, and the rule is degenerate there, but "degenerate" is a description and not an
   explanation.
3. **Is the paragraph mark always the space advance?** Measured true for 14 faces at 4 sizes, worst
   spread 0.007 em. Not measured for a face with no space glyph.
4. **Does the 1.2 factor survive a mixed-size line?** Every probe sets one size per line. A line
   whose runs differ in size is 3.3's problem and is not measured here.
5. **Does `@spc` interact with justification?** Only `algn="l"` was probed.
6. **Is `spcPts` also quantised?** The percentage is; exact points fitted at every value tried, but
   no value below 6pt or off a whole point was tried.
