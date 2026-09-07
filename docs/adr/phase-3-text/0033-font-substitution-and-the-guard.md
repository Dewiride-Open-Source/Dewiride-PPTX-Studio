# ADR 0033 — Font substitution and the guard

**Sub-phase 3.7.** Status: accepted. Adds `packages/text/src/fonts/` and
`packages/model/src/resolve/typeface.ts`. Six questions, each fitting exactly one candidate
reading. Two of the plan's own instructions for this sub-phase turn out to be wrong, and one of
them is wrong in a way that renders every missing-font deck in the wrong typeface.

Code: `packages/text/src/fonts/{presence.ts, substitute.ts, report.ts}`,
`packages/model/src/resolve/typeface.ts`.
Measurement: `corpus/ground-truth/font-substitution.json`, from
`tools/ground-truth/fonts/substitution/{probes.ts, build-deck.ts, read.ps1,
measure-in-browser.ts, analyse.ts}`.

---

## What this sub-phase had to decide

The plan asks for three things: a substitution table with a `metricCompatible` flag, detection "via
`document.fonts.check` plus a 3-glyph advance fingerprint", and a panel listing every substituted
typeface. Its verification is "a deck referencing 20 absent fonts reports all 20".

None of that says what a _correct_ substitution is. Two machines are involved and they disagree:
the one that authored the deck, where PowerPoint chose something when the named face was absent,
and the one rendering it now, where a browser will choose something else. Getting the second to
agree with the first is the whole job, and neither end of it was measured.

So T7 asked both. **159 probes across 8 packages**, every one opened by PowerPoint without a
repair, read three ways: the `LOGFONTW` face name in an EMF export of every slide, `TextRange2`'s
own resolved font name, and `BoundWidth` — compared only against another probe read the same way,
so the reader's own bias cancels.

### The instrument, and its one caveat

`Slide.Export(path, "EMF")` records the GDI calls PowerPoint makes, and
`EMR_EXTCREATEFONTINDIRECTW` carries the `LOGFONTW` it built. That is **what PowerPoint asked
for**, after its own substitution and before GDI's font mapping and font linking. The distinction
is visible in the readings: `Helvetica` comes out as a record naming `Helvetica`, at Arial's exact
width, because PowerPoint passed the name through and the Windows font mapper resolved it. So the
EMF answers "what did PowerPoint decide" and the width answers "what did it get", and this ADR is
careful about which of the two each finding rests on.

---

## PowerPoint draws every absent Latin face in Calibri

Twenty names no machine has, varied along the axis a font mapper might read — `Zzz Probe Serif`,
`Zzz Probe Mono`, `Zzz Probe Mincho`, `Zzz Probe Gothic`, `Zzz Probe Condensed` and fifteen more.
All twenty came back Calibri.

| reading                          | fits  |
| -------------------------------- | ----- |
| always Calibri                   | 22/22 |
| the theme's own minor Latin face | 20/22 |
| the style word in the name       | 17/22 |
| always Arial                     | 0/22  |

The runner-up is the one worth naming. "It falls back to the theme's minor font" is right on 20 of
22 rows **because the theme's minor font is Calibri** in the package those rows come from. The two
rows that separate them are in a package whose theme names an absent face for both collections:
there the theme reading predicts the absent name and the measurement says Calibri.

**And the substitution is metric-preserving, exactly.** Four strings at three sizes, twelve
comparisons, and the absent face's `BoundWidth` equals Calibri's to the last digit in all twelve —
against 1 for Arial and 0 for Times New Roman and Segoe UI. PowerPoint does not scale a substitute
to the metrics of the face it replaced; it has none to scale to.

### Of the three hints on `CT_TextFont`, only `@charset` is read

`a:latin` carries `@panose`, `@pitchFamily` and `@charset` beside `@typeface`, and a file can
therefore describe a font it names but does not embed. Twenty-one probes held the name constant and
varied one hint at a time.

| reading                         | fits  |
| ------------------------------- | ----- |
| only `@charset` is read         | 21/21 |
| all three are ignored           | 19/21 |
| `@panose` picks the family      | 14/21 |
| `@pitchFamily` picks the family | 12/21 |

Nine `@pitchFamily` values spanning every GDI family and both pitches, and five PANOSE strings
copied off the faces whose shape they describe — including Courier New's on a probe that also sets
`pitchFamily="49"`, which is as hard as a file can push for a monospace — moved the answer not at
all. Every one came back Calibri.

`@charset` moves it:

| `@charset`       | drawn in  |
| ---------------- | --------- |
| 0 (ANSI)         | Calibri   |
| 2 (Symbol)       | Calibri   |
| −95 (Greek)      | Calibri   |
| −78 (Arabic)     | Arial     |
| −128 (Shift-JIS) | Yu Gothic |

That table is **data, not a rule**. Five points is enough to refute "the hints are ignored" and not
enough to state what the mapping is; `SUBSTITUTES` does not encode it, and it is an open question
below.

### `+mj-lt` names the major collection, and substitution happens after

| reading                                          | fits |
| ------------------------------------------------ | ---- |
| `+mj-` to `a:majorFont`, `+mn-` to `a:minorFont` | 8/8  |
| always the minor collection                      | 5/8  |
| the two collections swapped                      | 2/8  |

Read against four themes: one naming present faces, two naming a present Latin face with empty
`a:ea`/`a:cs`, and one naming an absent face in all four slots. The last is what fixes the order of
operations — the reference is followed first and the result is then substituted, rather than the
token itself being treated as a missing font.

`resolveTypeface` returns `undefined` for a collection entry the theme leaves as the empty string,
which the stock Office theme does for `a:ea` and `a:cs`. That is not the same as the theme naming
nothing, and the two must not collapse: see the open question below.

### `TextRange2.Font.Name` reports the file's name, not the drawn one

159 of 159, against 75 for the reading that it reports what was drawn. The object model resolves a
theme reference and stops; it never mentions the substitute. So a diagnostics panel cannot be built
on it, and — more useful — the name to key a report on is the one in the file, which is what
`fontReport` takes.

---

## The browser half: two of the plan's instructions are wrong

### `document.fonts.check` is not a detector

It was asked about 38 families and said yes to all 38, including the 20 that do not exist. Scored
as a detector it is right on 18 of 38 — every present family, and no absent one. This confirms what
3.2 stumbled into with a single absent face and closes it with a number.

The reason is in the specification rather than in Chromium: `check` reports whether the matching
`FontFace` objects **in the document's font set** are loaded, and a locally installed face is not
in that set. Nothing matches, nothing is unloaded, the answer is yes. It stays useful for the faces
we register ourselves in 8.6 and is useless for everything else.

### Three anchors, not one

A family is absent when it measures identically to a generic stacked behind it. Which generic
matters more than it looks:

| reading                                  | fits  |
| ---------------------------------------- | ----- |
| equal to all three anchors               | 38/38 |
| equal to `monospace` alone               | 37/38 |
| equal to `serif` alone                   | 36/38 |
| equal to `sans-serif` alone              | 36/38 |
| the bare family differs from the default | 36/38 |
| `document.fonts.check`                   | 18/38 |

Chromium on Windows resolves `monospace` to Consolas, `serif` to Times New Roman and `sans-serif`
to Arial. Any single anchor therefore declares its own resolution missing — the `monospace` reading
is wrong about exactly one family and that family is Consolas. All three together are wrong about
none, and a family would have to be metrically identical to three different faces at once to fool
it.

The rule that follows is honest about its limits. `isFontAvailable` answers **true** for
`Helvetica` on this machine, because Windows maps it to Arial and the text will lay out as Arial.
That is not a false positive: the guard detects every substitution that changes layout and cannot
detect one that does not, which is the only claim an advance-based instrument can support.

### The fingerprint is a word, not three glyphs

The plan says "a 3-glyph advance fingerprint". The search was run exhaustively over a pool of 62
characters plus one word, asking for the smallest set that gives as many distinct values as the
whole pool does.

**No set of three or fewer single characters reaches it.** `Hamburgefonstiv` alone does. The reason
is a pair on this machine: Segoe UI and Leelawadee UI share their Latin design and returned the
same advance for every one of the 62 characters, differing only over the word — by 1.3px at 100px.
A three-character fingerprint would have called a Leelawadee UI run Segoe UI.

Under the chosen fingerprint the 18 present families fall into 15 classes, and the three collisions
are all real: `Arial = Helvetica`, `Times = Times New Roman`, `Courier = Courier New`. Two names for
one file. That is exactly what `metricCompatible` is supposed to mean, so `metricsAgree` is the
measurement of it, and it is the check `verifySubstitute` runs.

### And it is only comparable at one size

Advances are **not** linear in size: the worst deviation between a 100px advance and 6.25 times the
16px one is **11.1%**, from hinting at small sizes. `FINGERPRINT_PX` is fixed at 100 for that
reason, and two fingerprints taken at different sizes must never be compared.

---

## What the code does with all this

`fontStack(family)` is the load-bearing output, and it exists because of one measured
disagreement: **PowerPoint's last resort is Calibri and Chromium's is its own `serif`.** A renderer
that sets `font-family: "Missing Face", sans-serif` and lets the browser decide draws the deck in
Times New Roman where its author saw Calibri — different advances, different line breaks, a
different number of lines in every autofitted box. So the stack ends where PowerPoint would have:

```
"Missing Face", "Substitute", "Calibri", "Carlito", sans-serif
```

`fontReport` returns one finding per distinct typeface with a run count, and grades each as
`present`, `substituted` or `missing`. `metricCompatible` is a claim the table's compiler makes —
Liberation Sans was designed to Arial's advances, Carlito to Calibri's — and `verified` says
whether this machine could check it. An unverified claim is reported as unverified rather than as
true. Nothing in the table is asserted as measured that was not measured: the three pairs above are
the only ones this machine could confirm, and it confirmed them.

---

## What went wrong

**The first fingerprint pool was too small to falsify anything.** Eight characters were measured,
Segoe UI and Leelawadee UI agreed on all eight, and the analysis was ready to report a three-glyph
fingerprint that reached the ceiling — because the ceiling was itself computed from those eight.
Widening the pool to 62 characters plus a word moved the ceiling and refuted the answer. A search
whose search space is also its scoring rule proves nothing.

**The empty-theme-collection question was unscoreable as first built.** `+mn-cs` against a theme
whose `a:cs` is `""` had nothing to be compared with, so two probes were added that name the
theme's Latin face in the slot by hand — which is what "an empty entry falls back to Latin"
predicts — and then a second package with a different Latin face, so the reading could not be a
property of Verdana. The question is still not answered; see below.

**One mutant survived the first sweep** — flipping a present face's `metricCompatible` to `false`.
Nothing asserted that a face drawn as itself is compatible with itself, which is the row a
diagnostics panel shows most often. That was a missing assertion, not an equivalent mutant. Second
sweep: **29 mutants, 29 killed**.

---

## Deviations from the plan

- **The code is in `packages/text/src/fonts/`, not a new `packages/fonts`.** The guard is
  measurement — it needs the same `OffscreenCanvas` route `runs/measure.ts` uses, and two routes to
  a measurement are two measurements. Phase 8's `packages/fonts` is about font _files_ and depends
  on nothing; creating it now would also have meant a workspace change this sub-phase did not need.
- **The theme reference resolves in `model`, not in `text`.** `themeFontRef` and `FontScheme`
  already live there and the two packages do not depend on each other.
- **`document.fonts.check` is not part of the detector.** The plan pairs it with the fingerprint;
  it scores 18 of 38 and contributes nothing the fingerprint does not.
- **The fingerprint is one word, not three glyphs**, for the reason above.

---

## Open questions

1. **What is the `@charset` mapping?** Five values were measured and three of them agree. Whether
   the substitute is chosen by script, by a per-charset default face, or by the Windows font-link
   table is not established, and nothing in the code depends on it.
2. **What does an empty `a:ea`/`a:cs` in a theme resolve to?** Measured and not settled. With a
   Verdana theme, Thai text under `+mn-cs` came out 63.6pt wide naming Angsana New, while the same
   slot naming the Latin face came out 78.75pt — so it is _not_ a fallback to Latin. But with a
   Courier New theme the East-Asian answer moved from MS Gothic to Yu Gothic, which makes the
   choice a property of the Latin face after all. This is Windows font linking, it is
   machine-specific, and a browser cannot reproduce it. `resolveTypeface` returns `undefined` and
   the caller decides.
3. **Is the substitute Calibri, or is it the shell's UI font?** Every reading here comes from one
   Windows 11 machine with Office 365. A second machine — particularly a non-English install —
   would say whether `POWERPOINT_LAST_RESORT` is a constant or a system setting.
4. **Bold and italic were not probed.** A substituted face's synthetic bold may not match
   PowerPoint's, and `@b`/`@i` change advances.
5. **The `metricCompatible` claims for the seven open clones are unverified here**, because none of
   those fonts is installed. 8.8 bundles them and can settle every row at once.
6. **Non-Latin absence is not covered.** `a:ea` and `a:cs` fall through to Windows font linking,
   which has no browser equivalent; the report treats every slot the same way, and a CJK deck on a
   machine without a CJK face will be reported accurately and rendered approximately.
7. **`packages/validate` has no rule for a `@panose` that is not 20 hex characters**, nor for a
   `@charset` outside a signed byte. Both are cheap and neither exists.
