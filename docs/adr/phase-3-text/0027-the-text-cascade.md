# ADR 0027 — The text cascade

**Sub-phase 3.1.** Status: accepted. Opens Phase 3. Supersedes nothing; corrects one member of a
type 2.9 declared in advance, and refutes four claims the plan makes.

Code: `packages/model/src/{text.ts, parse-text.ts, resolve-text.ts, builtin-text-styles.ts}`.
Measurement: `corpus/ground-truth/text-cascade.json`, from
`tools/ground-truth/{text-cascade.ts, build-text-deck.ts, read-text.ps1, analyse-text.ts}`.

---

## What this sub-phase had to decide

The plan names nine places a text property can come from and an order for them, and observes that no
two implementations agree on that order. It does not say where the order came from. So the whole of
this sub-phase is one question asked nine ways: **given a run whose `a:rPr` declares nothing, which
of the nine answers?**

### The measurement is a ladder of packages, not a rendering comparison

A resolved font size is a number PowerPoint will tell you. `Font.Size` on a run with an empty
`a:rPr` is the answer to the entire cascade, reported to a quarter of a point, with no bitmap and no
antialiasing in the way. Give each of the nine levels a size no other level has, and the number that
comes back **names the level that won**.

So: nine packages, each declaring one level fewer than the last, and a tenth reading for each on a
shape that is not a placeholder. Sizes four points apart and none of them 18, because 18 turns out
to be what a silent package resolves to and a probe that landed on it by coincidence would be
unreadable. Two of the nine levels — `p:defaultTextStyle` and the theme's `a:objectDefaults` — live
outside any sheet, which is why this is nine packages and not nine slides.

39 packages, 133 probes, 129 usable. Two packages were repaired on open (a slide may carry neither a
`sldImg` nor an `hdr` placeholder) and their four probes are excluded from every score.

---

## The order, scored rather than confirmed

Confirming the plan's order would have been worth almost nothing. The ladder was built so that a
_wrong_ order predicts a different number at some rung, so the analysis enumerates **all 40320
permutations** of the eight named sources and scores each against the nine placeholder rungs.

| what was scored                                                    | result                            |
| ------------------------------------------------------------------ | --------------------------------- |
| permutations, with a placeholder assumed to read all eight sources | **0 of 40320 fit**                |
| permutations, once the two it ignores are dropped                  | 56 fit, **1 distinct read order** |

The one order that fits:

```
a:rPr  >  a:pPr/a:defRPr  >  shape a:lstStyle  >  layout ph a:lstStyle  >  master ph a:lstStyle  >  p:txStyles
```

which is the plan's order for its first six levels. The interesting half is the two it drops.

---

## What the plan got wrong

### 1. The walk has two termini, and they are not two rungs of one ladder

A shape whose inheritance chain ends at a type with a bucket reads that bucket **and stops**. A
shape with no bucket reads `p:defaultTextStyle` **and stops**.

Measured directly: rung seven of the ladder declares a `p:defaultTextStyle` at 20 points and nothing
else, and the body placeholder came back at 28 — PowerPoint's own built-in body size — while the
shape beside it that is not a placeholder came back at 20. And a `sldNum` placeholder in a package
declaring both `p:otherStyle` at 12 and `p:defaultTextStyle` at 20 came back at **20**.

A body placeholder ignores a `p:defaultTextStyle` that is right there in the same package. That is
not an ordering; it is a fork.

### 2. The `@type` at the _end_ of the chain picks the bucket, not the shape's own

2.9 measured that the first hop matches on `@idx` alone, so a slide `title` at idx 0 really does
land on a layout `body` at idx 0. When it does, it reads `p:bodyStyle`.

Four crossed probes, and no other reading produces all four:

| slide `@type` | layout `@type` | measured | bucket |
| ------------- | -------------- | -------- | ------ |
| `title`       | `body`         | 24pt     | body   |
| `body`        | `title`        | 54pt     | title  |
| `title`       | `sldNum`       | 18pt     | none   |
| `sldNum`      | `body`         | 24pt     | body   |

### 3. `dt`, `ftr`, `sldNum` and `hdr` reach no bucket at all

The plan assigns them `p:otherStyle`. Measured across all sixteen `ST_PlaceholderType` values, one
package each: the twelve content types resolve to their bucket and those four fall through to
`p:defaultTextStyle`.

### 4. Nothing measurable reads `p:otherStyle`

ECMA calls it "the text style for all other text". It is not read by a shape that is not a
placeholder — on the slide, on the layout, or on the master, all three probed — nor by the four
types the plan assigned to it, nor by any of the sixteen. Every probe that could have reached it
came back at the floor instead. This is a negative finding and is recorded as one; see the open
questions.

### The bucket rule, scored over the eight readings anyone could hold

Three independent choices — whose `@type` selects the bucket, whether the header-and-footer trio
reads `p:otherStyle`, and whether a shape with no bucket does — give eight models. Over 37 probes:

| model                                                                       | fits        |
| --------------------------------------------------------------------------- | ----------- |
| **chainEnd / hf-none / plain-none** (measured)                              | **37 / 37** |
| slide / hf-none / plain-none                                                | 33 / 37     |
| chainEnd / hf-other / plain-none                                            | 32 / 37     |
| slide / hf-other / plain-none                                               | 29 / 37     |
| chainEnd / hf-none / plain-other                                            | 21 / 37     |
| slide / hf-none / plain-other                                               | 17 / 37     |
| chainEnd / hf-other / plain-other                                           | 16 / 37     |
| **slide / hf-other / plain-other** (the plan's, and every implementation's) | **13 / 37** |

The two halves were then put back together and re-scored as one function against every situation-
bearing probe: **55 / 55**. A model that fits each half separately and not the two together has been
fitted twice rather than measured once, and the analysis throws rather than emitting a fixture it
cannot fit.

---

## Three more things the file says and PowerPoint does not read

- **`a:objectDefaults/a:spDef/a:lstStyle`.** The plan's eighth level. A package declaring a size
  there and nowhere else resolves to the floor. (It is plausibly the defaults applied to a _newly
  inserted_ shape rather than a resolution level, which would explain it; either way it is not one.)
- **`a:defPPr`**, the element before `a:lvl1pPr` in every `CT_TextListStyle`. Probed on a shape's own
  `a:lstStyle` and on `p:defaultTextStyle`; neither changed anything. It is parsed and preserved, and
  never resolved against.
- **`p:otherStyle`**, above.

---

## The tenth source, which the plan does not name

Rungs seven, eight and nine all reported **28 points** for a body placeholder — a size nothing in
any of those packages declares. There is a source below every source the plan names, and it is
PowerPoint's own: when a master declares no `p:txStyles` at all, PowerPoint substitutes its whole
built-in set.

Read out at all nine levels for a title, a body and a shape that is not a placeholder, and committed
as `packages/model/src/builtin-text-styles.ts`, generated from the fixture:

| bucket | sizes, levels 1–9   | `marL`            | `indent` | bullet |
| ------ | ------------------- | ----------------- | -------- | ------ |
| title  | 44, then 18         | 0                 | 0        | no     |
| body   | 28, 24, 20, then 18 | 228600 + 457200·n | −228600  | yes    |
| other  | 18                  | 457200·n          | 0        | no     |

**It is substituted wholesale, not merged underneath.** The distinction cost one more package and
matters on every partially-specified master: a body placeholder whose `p:bodyStyle/a:lvl1pPr`
declares only a `marL` came back at 18 points with no indent, not at the built-in body's 28 with its
hanging one. So when `p:txStyles` is present there is no per-property backstop beneath it, and what
remains is a floor of 18 points, `marL` 0, `indent` 0.

That floor is the fourth refutation: **ECMA-376 gives `@marL` a default of 347663 and `@indent` one
of −342900, and PowerPoint uses neither.** A parser that materialises them at parse time gives every
paragraph a 27-point hanging indent nothing asked for, and — worse — blocks the value it should have
inherited, because a default that has been written down cannot be told from a declaration. Measured
three ways, including one probe taking `marL` from `p:bodyStyle` and `indent` from the layout
placeholder and getting both.

---

## Two things the plan got right, and one it did not have to say

- **The cascade is per property.** Five levels declaring one property each produced a 24-point bold
  italic underlined struck-through run, where a nearest-level-takes-all reading gives an 18-point run
  that is only struck through.
- **`ST_Percentage` has two spellings.** `val="150000"` and `val="150%"` produced identical line
  spacing inside a Transitional part, as 2.6 found for the colour transforms that share the type.
  `parseInt("150%")` is 150, which is 0.15% line spacing, and the slide collapses to a line.
- **The levels are wholly independent, in both directions.** A layout placeholder declaring only
  `a:lvl2pPr` leaves the first and third to the master; and — the case that had to be measured
  separately — a layout placeholder declaring only `a:lvl1pPr` does **not** shadow the second and
  third. That second probe exists because a mutation that fell back to `a:lvl1pPr` survived the whole
  suite: the difference had never been measured, and `level-cross` looks like the experiment that
  would have caught it but is not, because there the nearer style declares the level being asked for
  and both readings answer the same way.

---

## The correction to 2.9's `Origin`

2.9 declared the full `Origin` union in advance, guessing at the members the text cascade would
need, so that 3.1 would add cases rather than widen a type every consumer had switched on. The guess
was right about five of them and wrong about one: `themeObjDefaults` named a level that never fires,
and is removed. `run`, `paragraph` and `builtin` are added. Declaring the union early was still the
right call — one member changed, and no consumer had to.

---

## What is deliberately not here

`a:bodyPr` is kept as the element it was rather than modelled. Anchors, insets and vertical text are
3.6 and autofit is 3.4, and half a typed model of `a:bodyPr` would be a worse version of the one that
arrives then. Bullets are 3.5: the built-in table records _whether_ PowerPoint drew one, because that
is what 3.5 will need to reproduce, and the glyph and its font are in the fixture rather than in the
table. Hyperlinks on a run are not parsed; they carry a relationship and belong with the sheet that
resolves it.

---

## Verification

- `pnpm check` green: layering, corpus, format, lint, five typecheck projects, 12 build tasks,
  roundtrip 52/52 decks and 1474 parts, 23 pkg:qa, **75 test files and 2196 tests**.
- 91 of those tests are this sub-phase's unit suite, and every one is written from the fixture: each
  case carries the situation its package put PowerPoint in, and the test rebuilds that situation out
  of XML and runs it through this package's own resolver.
- **The honest check is the corpus, not the unit suite.** The probe packages ask one question each
  and contain none of the markup nobody designed - a `p:txBody` with no paragraphs, an `a:rPr`
  carrying a gradient, a master declaring seven of nine levels. So `tools/corpus/text-cascade.test.ts`
  sweeps every paragraph and every run of every sheet of all 52 committed decks: 4 assertions over
  4209 resolutions, none throwing, every size, margin and indent inside a range a slide could hold.
  Its origin histogram is pinned as a set with floors, because a corpus that stopped exercising a
  level would otherwise go on passing:

  ```
  txStyles 1453   schemaDefault 899   run 887   masterPh 556
  shape 362   defaultTextStyle 64   layoutPh 50   paragraph 2
  ```

  `builtin` is zero and asserted to be, which is the point of it: every PowerPoint-authored master
  writes `p:txStyles`, and that is exactly why the built-in styles needed a probe package rather than
  a deck.

- The refutations are tests, not prose. The plan's bucket rule is re-scored inside the suite and
  asserted to fit 13 of 37; the `parseInt` reading of a percentage is asserted wrong; the ECMA
  defaults for `marL` and `indent` are asserted not to be what resolves.
- **Mutation sweep: 26 of 26 killed, no survivors.** Two of the twenty-six survived the first run,
  both of them the fallback-to-`a:lvl1pPr` reading, and the fix was a new _measurement_ rather than a
  new assertion — see `level-shadow` above.

---

## Open questions

1. **What reads `p:otherStyle`?** Nothing this experiment could construct. Notes slides and the
   handout master are the untested candidates, and both are out of scope until 12.x. Until then the
   element is parsed, preserved and never consulted.
2. **`Font.Name` comes back empty when no level declares a typeface.** A shape reaching a
   `p:txStyles` bucket that names no `a:latin` reports `''` through the object model, while the same
   shape reaching the built-in styles reports the theme's minor font — the built-in styles name
   `+mn-lt` and an explicit bucket that says nothing leaves the typeface genuinely unresolved. What
   PowerPoint _draws_ in that state is a 3.7 question and is not answered here.
3. **`a:objectDefaults` may be an insert-time default rather than dead markup.** The measurement says
   only that it is not a resolution level. Worth one probe in 5.x, when inserting a shape becomes a
   thing this project does.
4. **A `@lvl` outside `ST_TextIndentLevelType` is clamped to 8 rather than refused**, on the grounds
   that PowerPoint still opens the file. Not measured — no probe was built for it, because a deck
   PowerPoint repairs answers a question about the file PowerPoint wrote.
5. **The built-in styles were measured on one build** (16.0.20326). They are Office's stock master
   text styles and are unlikely to move, but nothing here would notice if they did.

Carried from 2.11: selection enters groups immediately; a locus is sampled at a fixed 25 points; no
chrome for a wholly non-finite geometry; `data-axes` unmeasured against a real `custGeom`. Carried
from 2.10: the `path="rect"` gradient drawn as an ellipse; `a:grpFill` on a rotated child; the
content bounds of a rotated leaf; a group rotated inside a non-uniformly scaled group; the tile phase
of a pattern under `grpFill`; which colour map an inherited master shape resolves against; whether a
layout shows its master's placeholders.

**Scope gap, re-flagged for the fourth time:** `a:blipFill` is unassigned in the plan though Gate 2
asks for cropped image fills. It has a place in the `Fill` union, a place in the renderer's `switch`,
and nothing behind either. A picture on a dropped deck renders as its outline.
