# 0030 — Autofit

- **Status** accepted
- **Sub-phase** 3.4
- **Experiment** T4 — 3521 probes across twenty-eight packages, PowerPoint driven and read over COM
- **Fixture** [`corpus/ground-truth/autofit.json`](../../../corpus/ground-truth/autofit.json)
- **Code** `packages/text/src/lines/autofit.ts`
- **Closes** the question [0028](./0028-measurement-and-the-line-model.md) left open about the last
  line of a block, three sub-phases before it was due

## Context

The plan says:

> **3.4 — Autofit, two modes.** _View:_ apply the stored `@fontScale`/`@lnSpcReduction` verbatim so
> a PowerPoint-authored file pixel-matches. _Edit:_ discard and re-run the discrete 14-step ladder,
> rounding `pt × scale` to whole points **before** measuring. Never emit a `fontScale` outside the
> ladder — PowerPoint re-snaps it and text visibly jumps. An `autofitDirty` flag per `txBody` keeps
> view-only decks round-tripping byte-identically. **`spcFirstLastPara` defaults false** — discard
> the first paragraph's `spcBef` and the last's `spcAft`, or every centred/bottom-anchored box
> shifts.

Unusually for this project, almost all of that survives. The two modes are right, the ladder is
real, the rounding to whole points is real, and `spcFirstLastPara` does default false. What the
measurement adds is the arithmetic the plan could not have known: the ladder has fifteen rungs and
not fourteen, the rounding has an exception, `@lnSpcReduction` does not do what its name suggests,
and the fit test compares against a height whose last term nobody in this repository had yet
measured.

## The method

`a:normAutofit` is unusual among the things this project has had to measure, because PowerPoint
will _write the answer down_. Drive `TextFrame2.AutoSize` to `msoAutoSizeNone` and then back to
`msoAutoSizeTextToFitShape`, save the deck, and the `@fontScale` and `@lnSpcReduction` PowerPoint
chose are in the file, in its own handwriting. No rendering to interpret, no threshold to infer.

Two design decisions did the rest of the work.

**One short word per paragraph.** The obvious probe — prose in a box, shrink the box, read the
scale — cannot be scored without a line-breaking model, because the text height depends on the line
count, which depends on the scale, which is the unknown. Any error in our breaker would land on the
ladder and be indistinguishable from it. With one three-letter word per paragraph nothing can wrap
at any rung, so the line count is the paragraph count exactly, and the fit test collapses to
arithmetic over numbers the file states. `Lines().Count` is recorded for every probe and asserted
against the count the probe was built for, so that is checked rather than assumed.

**Sweep the box, not the content.** Adding a paragraph moves the required scale in steps of a whole
line — far too coarse to tell 92.5% from 92%. Moving the box height by a point with the content
held still moves it by `1 / (lines × 1.2 × size)`, which at twenty lines of 18pt is 0.23% of scale.
Every height at which the emitted scale steps up is a bracket on that rung tight enough to name a
quarter of a per cent, and 1235 such boxes were swept a point at a time.

### The trigger is a window, and this is worth saying out loud

The first run measured nothing and looked like a result. Every one of twenty probes came back
`<a:normAutofit/>`, unshrunk, with `TextFrame2.AutoSize` reading back as `msoAutoSizeTextToFitShape`
and `BoundHeight` reporting a text block nearly twice its box. Re-editing the text changed nothing.

The presentation had been opened with `WithWindow:=msoFalse`. **Autofit runs in PowerPoint's layout
path, and the layout path does not exist without a window.** A headless COM harness measures nothing
here and reports it as "PowerPoint never shrinks" — which is a plausible-looking finding, and false.

The window may be minimised: the emitted scales were checked against a run with it restored and are
identical, so the sweep does not throw windows at whoever is at the keyboard.

## Decision

### 1. The ladder is fifteen rungs, and its order is measured

| #   | `@fontScale` | `@lnSpcReduction` |     | #   | `@fontScale` | `@lnSpcReduction` |
| --- | -----------: | ----------------: | --- | --- | -----------: | ----------------: |
| 1   |       100000 |                 0 |     | 9   |        70000 |             20000 |
| 2   |       100000 |             10000 |     | 10  |        62500 |             20000 |
| 3   |        92500 |                 0 |     | 11  |        55000 |             20000 |
| 4   |        92500 |             10000 |     | 12  |        47500 |             20000 |
| 5   |        92500 |             20000 |     | 13  |        40000 |             20000 |
| 6   |        85000 |             10000 |     | 14  |        32500 |             20000 |
| 7   |        85000 |             20000 |     | 15  |        25000 |             20000 |
| 8   |        77500 |             20000 |     |     |              |                   |

The font scales step by 7.5 points of a percent from 100 down to 25. There is nothing below 25%:
text that does not fit there is left overflowing.

The order is **font scale descending, and within one font scale the smaller reduction first** — and
that is a measurement rather than a reading of the table. At 32pt the rounding of `size × scale`
makes rung 6 _taller_ than rung 5: `round(32 × 0.85) = 27` at 90% spacing is 24.3pt a line, against
`round(32 × 0.925) = 30` at 80% spacing giving 24.0pt. Twenty-two boxes of `ladder-b` fit both, and
PowerPoint took the smaller text every time. "The tallest rung that fits" scores 3066 of 3262.

**The enumeration is complete, not merely unrefuted.** Under an exact-point `a:lnSpc` the reduction
changes nothing, so every rung at one font scale has the same height and the _first_ one in the list
wins. Sweeping such a box a quarter-point at a time gives one step per font scale and reads off the
smallest reduction each carries: 0 at 100% and 92.5%, 10000 at 85%, 20000 everywhere below. A
`(85000, 0)` rung would have won at 189.5pt and did not; a `(100000, 20000)` rung would have won
across a 23-box stretch of the main sweep and did not.

### 2. `round(size × fontScale)` to a whole point — except at 100%

18pt at 92.5% lays out at **17pt**, not 16.65. Halves go up: 18pt at 25% is 4.5 and draws the
advance of 5pt, and 28pt at 37.5% is 10.5 and draws 11pt. Round-half-to-even scores 3225 of 3262,
refuted by five rungs of a 20pt sweep where the product lands exactly on a half.

The exception is the one a plausible implementation gets wrong. **A `@fontScale` of exactly 100% —
or none at all — leaves the size alone, fraction and all.** Eleven lines of 10.5pt fit a 139pt box,
which needs 12.6pt a line and not the 13.2 that rounding to 11pt would give. Rounding
unconditionally scores 3249 of 3262 and costs half a point on every line of every deck that uses a
fractional size, which is every deck whose author ever pressed the shrink-font button.

### 3. `@lnSpcReduction` is subtracted from the percentage, and exact spacing ignores it

Two separate findings, both contrary to the obvious reading of the attribute's name.

**Percentage spacing.** The reduction is subtracted from the `a:lnSpc` percentage and the difference
is quantised to a whole percent _afterwards_. At single spacing that is the same number as
multiplying the advance by `1 − r`; at 150% it is four per cent different, and the quantisation
order matters as soon as either value is a fraction of a percent — 149.6% reduced by 12.5% is 137%,
not 137.5%. Multiplying scores 2859 of 3262 and quantising first scores 3200.

**Exact-point spacing.** `a:lnSpc` stated as a length **ignores the reduction entirely**. 30pt stays
30pt at every rung. The consequence is worth stating plainly: a paragraph with exact line spacing
cannot be made shorter by autofit at all, and PowerPoint still walks the whole ladder for it,
producing a staircase of eleven rungs inside four points of box height.

The result is floored at **one per cent** of the line box, not at zero. A 15% spacing reduced by 20%
is −5%, and PowerPoint draws it at 0.096pt a line at 8pt and 0.480pt at 40pt.

### 4. The last line of a block is not the advance

This is the term [0028](./0028-measurement-and-the-line-model.md) took as a parameter and said a
plausible guess would be exactly the silent wrong answer this repository forbids. Sub-phase 3.6 was
going to measure it. The fit test needs it three sub-phases early, so T4 measured it.

```
lastLine = advance ≤ C × size  ?  max(advance, 0.75 × advance + b × size)
                               :  0.75 × advance + b × size
```

`0.75` is the same for every face to six decimal places — six faces, seven sizes, two line spacings
each, fitted by least squares and coming back as `0.750000` with a spread of two parts in a million.
`b` and `C` are properties of the typeface:

| face            |         `b` |      `C` |
| --------------- | ----------: | -------: |
| Tahoma          | 0.205318787 |      1.2 |
| Verdana         | 0.207355431 |      1.2 |
| Arial           | 0.227618991 |      1.2 |
| Georgia         | 0.231489897 |      1.2 |
| Times New Roman | 0.234424476 |      1.2 |
| Courier New     | 0.300000058 | 1.269842 |

Worst residual across all six, held-back rows included: 0.0043pt.

**Why 3.2 could not see this.** At single spacing the floor branch wins and the last line is exactly
the advance, so a block of _n_ lines measures `n × advance` and the two readings are the same
number. Every probe in T2 was at or near single spacing. Under a stated `a:lnSpc` they come apart,
and `n × advance` scores 2725 of 3262.

**Why the floor is not `min(advance, 1.2 × size)`.** That looks like the same rule and is refuted by
three probes: 24pt Arial under a 30pt exact spacing draws its last line at 27.96pt, below both the
30pt advance and the 28.8pt box. It scores 3204 of 3262.

**Why `C` is a field and not the constant 1.2.** Five faces cannot tell them apart. Courier New's
threshold, read off a sweep that holds the spacing at an exact length and moves the size through it,
is bracketed in `[1.269841, 1.276596)` — nowhere near 1.2. Taking 1.2 everywhere scores 3252 of 3262.

**A browser cannot supply either constant, and this is a real limitation.**
`TextMetrics.fontBoundingBoxDescent` is the obvious candidate: it is within 0.0003 of `b` on Courier
New and within 0.0013 on Tahoma and Verdana. It is 0.0157 out on Arial, and — decisively — Arial and
Verdana report descents 0.0026 apart while their measured `b` differs by 0.0203, so no function of
the descent alone produces both. `fontBoundingBoxAscent + fontBoundingBoxDescent` puts Verdana
(1.2153) above Courier New (1.1328) where `C` puts it below. Both were measured in Chromium 151 at
4096px, where the whole-pixel quantisation of these numbers is four decimal places below the ratio.

So `packages/text` carries the six measured faces as data, `faceLineMetrics()` **throws** for a face
it does not know, and `APPROXIMATE_FACE_METRICS` exists under a name that has to be typed. The
approximation is bounded: `b` spans 0.0947 across six faces and multiplies the size of exactly one
line in a block, so at 18pt the worst case is 1.7pt on a block of any length — and it is zero
whenever the spacing is inside the floor, which includes every paragraph at single spacing.

### 5. Paragraph spacing, and `spcFirstLastPara`

`a:spcPct` is a fraction of the **scaled 1.2 line box**, so it shrinks with the font; `a:spcPts` is
used **raw** and does not follow the font scale. Reading the percentage as a fraction of the font
size scores 3186 of 3262; scaling the exact lengths scores 3146.

`spcFirstLastPara` **defaults false** — the first paragraph's `spcBef` and the last's `spcAft` are
dropped. Defaulting it true scores 3254 of 3262. Absent and `"0"` were measured separately and are
identical.

One observation that belongs to 3.6 rather than here: with the flag on, `BoundHeight` grows by a
`spcBef` but `BoundTop` does not move and the first line still sits exactly on the frame top. The
extra space is somewhere other than above the first line. The _height_ is what the fit test needs
and the height is measured; where the space goes is an anchoring question.

### 6. The fit test is against the box minus the insets

`bodyHeight = shape height − tIns − bIns`. Testing against the shape height scores 3254 of 3262. The
API takes `bodyHeightPt` rather than a rectangle so that forgetting to subtract them is a thing a
caller has to do deliberately.

### 7. A stored scale is applied exactly as written

The plan's view mode is right, and the falsifiers land:

- A deck stating `fontScale="25000" lnSpcReduction="20000"` on two words in a box that holds twenty
  renders at 4.8pt a line. A recomputation would give 21.6.
- A deck stating `fontScale="100000"` on twelve lines in a box that holds two renders overflowing.
  Opening the file does not shrink it.
- A `fontScale` of 45000, 65000 or 90000 — none of them rungs — is applied as written and not
  snapped.

The ladder constrains what we _write_. It does not constrain what we _read_.

### 8. `a:spAutoFit` adds just under one per cent

```
shape height = 1.009765625 × (text height + tIns + bIns)
```

181 probes across four faces bracket the constant in `[1.009764816, 1.009766650]`. It scales the
insets as well as the text, so it cannot be a font metric — it is slack PowerPoint adds so the text
it just measured definitely fits. The package uses `517/512`, which is inside the bracket: a dyadic
rational inside a measured range is a better guess at what a program computed than the range's
midpoint. That is a guess about the _form_ of a number whose _value_ is measured, and it is listed
below as open.

PowerPoint also declines to resize when the target is close to the current height — boxes from 87 to
90pt were all left alone against a target of 87.244. That hysteresis is deliberately not implemented:
`requiredShapeHeight` answers how tall the shape should be, and the decision to leave a shape alone
belongs to the command that resizes it.

## Incidental findings

- **`Lines(i).BoundTop` and `BoundHeight` disagree under paragraph spacing.** With
  `spcFirstLastPara="1"` the block is a `spcBef` taller than the distance from the first line top to
  the last, so the reported line positions do not sum to the reported block height. The fit test
  follows `BoundHeight`, which is what PowerPoint itself compares against.
- **The advance floor is one per cent of the line box**, not zero and not a fixed length. Measured
  at five sizes from 8pt to 40pt.
- **Twelve `normAutofit` scales in corpus deck `a11` are now testable rather than decorative.**
  [0009](../phase-1-round-trip/0009-the-corpus.md) says explicitly that `a11` is not the source of truth for the ladder
  and that 3.4 would measure it. It did.

## Verification

- **3262 of 3262**, one model, every rival refuted. The score combines both instruments: 2699
  emitted rungs and 563 block heights. Neither settles the rule alone — three readings of the last
  line fit every rung and only one fits the blocks.
- Nearest rivals: the floor threshold at 1.2 everywhere (3252), `spcFirstLastPara` defaulting true
  (3254), the fit test ignoring the insets (3254), rounding the size even at 100% (3249).
- The constants in section 4 were fitted on the view decks, which contain no rungs at all, and then
  used to predict 2699 rungs. That is out of sample, not a curve fit.
- 28 packages, every one opened with no repair, and every probe laid out the line count it was built
  for.
- The wrap deck is not scored here — its line count is a consequence of the rung and predicting it
  needs the breaker the analysis must not import. 54 of 54 chose a rung that holds the text
  PowerPoint laid out, and for 46 the rung above provably overflows without a breaker. The package
  test replays all 54 with the real engine.
- `packages/text/src/lines/autofit.test.ts`: 45 tests, replaying every swept box and every block reading.
- **Mutation sweep: 34 of 34 killed, no survivors.**

## Consequences

- `blockHeight` in `line-model.ts` still takes `lastLineHeight` as a parameter, and now there is a
  function to compute it. The two are deliberately not merged: `line-model.ts` is arithmetic over
  what a file states and knows nothing about typefaces, and `lastLineHeight` needs per-face data.
- `@pptx-studio/text` still has no runtime dependencies.
- 3.6 inherits a smaller question than it was given: not "what does the last line measure" but
  "where does the space go".
- 3.7 inherits a concrete one: the substitution table needs `b` and `C` per face, and they cannot be
  read from a browser. `packages/fonts`' SFNT reader in 8.1 is the next place to look.

## Open questions

1. **What are `b` and `C`?** Both are almost certainly derived from the font's vertical metrics, and
   three candidates from `TextMetrics` are refuted above. 8.1 can test them against the real `hhea`
   and `OS/2` tables. Until then, six faces are data and everything else is an explicit
   approximation.
2. **What is `1.009765625`?** Measured to seven significant figures and unexplained. It is not a
   font metric — it scales the insets.
3. **The spAutoFit hysteresis.** The band was observed (87 to 90pt against a target of 87.244) but
   not characterised. It does not affect what height to compute, only whether to apply it.
4. **A body whose last paragraph differs from the others.** Every probe held the paragraphs
   identical, so "one last line for the whole body, taken from the last paragraph" is the only
   reading consistent with what was measured and is still a generalisation. A body mixing sizes or
   line spacings between paragraphs is unmeasured.
5. **Mixed run sizes within one line.** Inherited from 3.2 and untouched here: the line model takes
   one size per line, and which size a mixed line uses is unmeasured.
6. **Whether the ladder is the same in other PowerPoint builds.** Everything here is 16.0.20326 on
   Windows. The corpus deck `a11` carries twelve scales from an unrelated authoring session and they
   are all on the ladder, which is weak evidence that it is stable.
