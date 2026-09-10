# 0044 — What the runner's own fonts said

**Sub-phase 3.10 · accepted · supersedes the T14 section and open questions 1 and 2 of [0043](0043-what-3-10-left-open.md)**

[0043](0043-what-3-10-left-open.md) built experiment T14 and said the finding was not in that record
because the job had never run. It has now. It went red, and it was right to: it found **three
defects in shipped code**, refuted the premise its own gate was built on, and refuted a claim in
0042 that everything since has been resting on.

79 faces from `/usr/share/fonts` on `ubuntu24` image `20260831.293.1`, Chromium 151.0.7922.34,
3420 comparisons of which 2832 gated. One file skipped, `wqy-zenhei.ttc`, for being a collection.

## Chromium does not measure the same way on two operating systems

**All 3420 browser widths are whole pixels**, and so are all 158 face-box values. On Windows, where
T13 ran, the same browser returns `172.7999725341797`.

It is per glyph, not a rounding of the total — `round`, `floor` and `ceil` of our total reproduce
the browser on 604, 586 and 603 of 2832 gated rows, while the error scales with the glyph count from
10 to 1349. FreeMono settles the rule outright, having one advance (600/1000 em) for every glyph:

| px                              | 8        | 12       | 16       | 32       | 100 | 1000 |
| ------------------------------- | -------- | -------- | -------- | -------- | --- | ---- |
| observed `(browser − reader)/n` | +0.20001 | −0.19999 | +0.40001 | −0.19999 | 0   | 0    |
| `round(0.6·px) − 0.6·px`        | +0.2     | −0.2     | +0.4     | −0.2     | 0   | 0    |

**This refutes 0042.** Its width finding — _"the advance is linear in size to within 1.04e-4 px,
which is the finding the whole approach rests on: the browser does not hint it"_ — was measured on
seven fonts built by `tools/ground-truth/lib/truetype.ts`, and a font that carries no hinting program
cannot be hinted. The claim was true of the probes and never generalised. A table reader can be
exact against an unhinted face on Windows and cannot be exact against a real face on Linux, and
nothing in T13 could have told the difference.

## So the gate was asserting something no reader can do

T14's threshold was `1e-4` relative, justified by T13: the reader reproduces the browser exactly, so
a residual on a real face is a missing feature rather than noise. With a whole-pixel browser that
premise is false and the gate can never pass, which makes it useless in both directions.

The replacement is derived from the mechanism rather than chosen to fit. `W_browser = Σ round(aᵢ)`
contributes at most half a quantum per glyph; `W_reader = Σ (trunc₁⁄₆₅₅₃₆(aᵢ) + round₁⁄₆₅₅₃₆(kᵢ))`
contributes at most `1/65536 + 1/131072`. Both are per glyph:

```ts
widthTolerance(glyphs) = glyphs * (ADVANCE_QUANTUM / 2 + 1 / 65536 + 1 / 131072);
```

It cannot be tightened: Lato-ThinItalic on `latin-spread` at 1000px attains it exactly — 40 glyphs,
delta exactly 20.000000 px, every advance a half pixel on a 2000-unit em rounding the same way. The
worst row on any clean face is 0.79998 of it. The smallest **failing** row is 1.00622, so the margin
at the boundary is 0.6%, and that is why the only slack above `glyphs/2` is the reader's own step
(4.58e-5 of the bound) rather than any round number.

The quantum is asserted rather than assumed. A run whose widths are not whole pixels fails with
`advance-quantum` naming the first fractional width, because the derivation does not hold under a
subpixel browser and the bound would be four orders looser than that browser needs.

**What the gate cannot see is now printed in its own summary**, because it is a trap for whoever
reads it next: at a whole-pixel quantum the four readings that differ only below 1/65536 px —
`shipped`, `exact`, `advanceRounded`, `wholeStringTruncated` — are indistinguishable, scoring
433/434/433/434 of 3420. Only `unkerned` separates, because a missing kern is a fraction of the
width rather than of a pixel. T13 remains the only experiment that can separate the other four.

That also answers 0043's second open question differently from how it was asked. It expected a
2048-em face to separate the four face-box readings; 41 of the 79 faces are 2048-em and they
separate nothing, because the browser reports the box in whole pixels too. `round` fits 76/79 by
construction. The question is not closed, it is **differently open**: it needs a browser that
reports the box fractionally.

## Defect 1 — the reader skipped Extension lookups, and silently lost kerning

A GPOS `LookupType` of 9 is Extension Positioning, wrapping a real subtable behind a 32-bit offset.
`gposKerning` had `if (u16(r, lookup) !== 2) continue;`, so every one was skipped.

Eighteen Lato faces carry their `kern` feature that way. On every gated row of all eighteen the
`shipped` and `unkerned` readings were **identical to the last decimal** — no kerning found at all —
while the browser was 8.1% narrower on the kern-pair sample. Silently: no error, no warning, just
Latin text measured about 8% wide, which moves every line break and misfires autofit.

It was reproduced before it was fixed, which is the only reason the fix is trustworthy. A twelfth
probe reaches the same `PairPos` through `ExtensionPosFormat1`, and against the **unmodified** built
reader:

```
gpos-only        kernBetween(A,B) = -200
gpos-extension   kernBetween(A,B) = 0        <-- the defect
```

Chromium honours the wrapped pair — the two probes agree at all six sizes on all six strings — so the
reader was wrong rather than strict. The opposite outcome would have refuted the whole fix.

Rescored over both axes, which table wins and whether type 9 is followed:

| reading                                                                                      | fits      |
| -------------------------------------------------------------------------------------------- | --------- |
| GPOS when it yields a pair, else the `kern` table, **following one type 9 Extension lookup** | **72/72** |
| GPOS `PairPos` only, following one Extension lookup                                          | 66/72     |
| GPOS when it yields a pair, else the `kern` table, type 2 only _(the shipped bug)_           | 66/72     |
| the `kern` table when present, else GPOS                                                     | 66/72     |
| the legacy `kern` table only                                                                 | 54/72     |
| neither                                                                                      | 48/72     |

Exactly one level is followed; the specification forbids an Extension targeting another.

## Defect 2 — `renderDeck` died on a machine holding 53 usable faces

The example job, which installs the published packages from npm, failed with
`CLI_NO_FACE: no face on this machine can stand in for "Calibri Light"` — with 53 faces indexed.
`resolve` tried the family, its substitute, then `LAST_RESORT_FAMILIES`, and returned `undefined`.

ADR 0033 built `FaceUse{asked, drawn, substituted}` precisely so the answer could be a report rather
than an exception. The last resort was the one place that design was not applied. The chain now ends
in the shortened forms of the asked-for name, each also through the substitution table
(`Calibri Light` → `Calibri` → `Carlito`), and finally the first indexed family in code-unit order.
`CLI_NO_FACE` survives with its code, reachable only when the library is genuinely empty.

The shortened name is tried **after** Calibri/Carlito, not before, because ADR 0033 measured
PowerPoint's own last resort at 22/22 and `fontStack()` puts those two families in the browser's CSS
stack. A name heuristic firing first would make the two renderers fall back differently, which is the
invariant `faces.ts` exists to keep.

Two things fell out of it. The directory walk **now sorts by name** — it had been in `readdir` order,
which is ext4's hash order on the runner and NTFS's collation here, so which file claimed a family
slot was the filesystem's choice. And a blank `name` ID 16 no longer claims a family, since it would
sort before every real name and become the machine-wide last resort.

## Defect 3 — the face box on IPA Gothic, still open

`ipag.ttf` reports `880/120` in the browser and `879.88/195.80` from the reader: **75.8 px on the
descent**, 151 times what one rounding can cost. The ascent agrees — `1802/2048 × 1000 = 879.88`
rounds to 880 — so only the descent is wrong, and 120 px back into units is 245.8, which the
whole-pixel interval puts at 245 or 246 against `usWinDescent`'s 401.

It is **two font files, not three**: `ipag.ttf` and `fonts-japanese-gothic.ttf` carry the same
SHA-256, so one of them is counted twice; `ipagp.ttf` is distinct.

Three explanations were available and the run refutes two of them. The browser does **not** cap the
box at the em: 56 of the 79 faces report a box larger than the em and the reader agrees with every
one, DejaVu Sans among them at 1164 px against `usWin`'s 1164.1. Linux does **not** refuse `usWin`
generally, for the same reason. What survives is that Chromium read a different table here, whose
pair happens to sum to exactly the em.

The obvious suspect is the font, and it is not. T13's `split` already carries this shape — three
pairwise-distinct descenders and a `usWin` box of **1.2× the em**, where these CJK faces are 1.076×
— and Windows Chromium reports 900/300 for it, which is `usWin`. So the font properties in play are
not sufficient to move Chromium off `usWin`, which is evidence the difference is the **platform**
rather than the face. It is not proof: `split` has never been measured under FreeType.

The reader is deliberately unchanged. `metricsOf` still answers exactly what it answered — that
reading is 24/24 on Windows and this run is not evidence against it there — but `FaceMetrics` now
carries the `hhea`, `usWin` and `sTypo` pairs and the `fsSelection` bit alongside the chosen one, so
the next run can score which table the browser read instead of reporting an unexplained miss.

## Verification

- **`pnpm check` green**, all of it.
- **T13 rescored over twelve probes** and every earlier answer held: face box 24/24, kerning 72/72,
  widths 432/432, ideographic baseline 32/32. Regenerating the fixture into a scratch path
  reproduces the committed one byte for byte.
- **Mutation: 7 tried on the Extension walk, 6 killed and 1 named equivalent; 12 tried on the
  substitution chain, 12 killed after two genuine survivors; 13 tried on the gate, 13 killed; 14
  tried on the face-box scoring, 14 killed.**
  The two survivors on the chain were missing assertions — nothing covered a blank `name` ID 16, and
  the shortening test's library had no family sorting before `Segoe`, so the last resort was
  answering for it.
- **The gate was driven red and then green on the real artifact**, not on a fixture: red at 231
  failures over the run as it stands, still red at 20 with the Lato faces removed, green only once
  every face carrying a reader defect is out.
- The reader's own experiment refused to lie in between: run against the unfixed reader,
  `analyse.ts` **threw and wrote no fixture** — `the best reading … fits only 420/432`.

## Deviations

1. **`packages/cli/src/render/measure.test.ts` was edited from two tracks at once.** A twelfth probe
   makes four fixture-derived counts false, and holding the new probe out of the grid to keep a
   literal true would have been the softened gate the standards forbid.
2. **`pnpm build` reported `FULL TURBO` with changed sources in the tree, twice, leaving a stale
   `dist`.** Since `tools/ground-truth/**/analyse.ts` scores `packages/cli/dist`, a stale `dist`
   means scoring the previous reader. `--force --no-daemon` built correctly. This is a hole in the
   evidence chain of every ground-truth experiment and it is an open question below.

## Open questions

1. **Which table Linux read the face box out of.** 75.8 px of descent on two IPA Gothic files,
   narrowed to "a table whose pair sums to the em" but not settled. The next run of the job scores
   five readings against the candidate pairs and the answer is whichever fits every face. **The gate
   stays red until then, on purpose.** Two failures cover it and neither can be silent while the
   reader is wrong: `face-box` fires when the shipped reading misses a face and no reading fits them
   all, and `face-box-rival` fires when any rival fits strictly more faces than the shipped one —
   without the second, a run where `hhea` fits 79/79 and shipped fits 76/79 would go green on the
   very run that answers the question.
2. **Whether Extension lookups explain FreeSerif.** `FreeSerif.ttf` and `FreeSerifBold.ttf` also
   showed `shipped` identical to `unkerned` on every Latin, Cyrillic and Greek row, and 2.9–3.4%
   narrower in the browser — but unlike Lato they **do** kern the Arabic and Devanagari samples, so
   whatever is missed there is not the whole face. The Extension fix may or may not cover it. The
   next run of the job answers this, and it is the first thing to read in it.
3. **What `renderSlide` writes into `font-family`.** The exported SVG says
   `font-family="Calibri Light"` with no fallback list, so a browser opening that file picks its own
   face while our measurement used another. That is the one place the two renderers can still
   disagree about what was drawn.
4. **Turbo's cache said `FULL TURBO` over changed sources.** Until that is understood, no experiment
   scoring `dist` can be trusted without `--force`.
5. **The subpixel regime is untested against real data.** The gate refuses a run whose widths are
   not whole pixels rather than guessing; only an authored test row covers that path. If Chromium
   turns on subpixel positioning for this image, CI goes red with `advance-quantum` and the bound has
   to be derived again on the record.
6. **The shaping ratios are recorded, and 0042's third open question is answered as a number rather
   than closed**: `latin-ligatures` 1.010274 median over 74 faces, `arabic` 1.043478 over 14,
   `devanagari` 1.018119 over 9. The reader over-measures because it sums unshaped advances; nothing
   here makes it stop.
