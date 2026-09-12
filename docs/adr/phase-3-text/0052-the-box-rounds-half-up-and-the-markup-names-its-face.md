# 0052 — The box rounds half up, and the markup names its face

**Sub-phase 3.10 · accepted · closes open question 3 of
[0044](0044-what-the-runners-own-fonts-said.md), open question 2 of
[0045](0045-the-face-box-belongs-to-the-rasteriser.md), and the two carried entries
[0044](0044-what-the-runners-own-fonts-said.md) raised; asks the macOS question of
[0045](0045-the-face-box-belongs-to-the-rasteriser.md) without yet answering it**

[0045](0045-the-face-box-belongs-to-the-rasteriser.md) left the face box half-measured: which
table each rasteriser reads was settled, but which whole pixel the browser rounds the box to was
not, because every T13 probe is a 1000 em measured at 1000 px and lands on a whole pixel under any
rounding. [0044](0044-what-the-runners-own-fonts-said.md) left the CLI's markup naming a face it
had not measured in. Both are closed here, and two things are put on the record that were only
ever asserted: that the T13 fixture holds on a Windows machine other than this one, and that Turbo
had served a stale `dist`.

## Two probes off the pixel

`split-2048` is the three-way split on a 2048 em, with `usWin` 1902/451: at 1000 px that is
928.7109375 over 220.21484375, one fraction each side of the half, so floor, ceil and round-half-up
each predict a different pair. `split-2000` puts `usWin` 1857/400 at exactly 928.5 over 200, with
an even integer part, so half-up says 929 and half-to-even says 928. Nothing else changes: the
builder's 1000-unit literals are glyph outlines and `OS/2` averages, and neither feeds
`fontBoundingBox` or `hmtx`.

Chromium 151.0.7922.34 on this machine reported **929/220** and **929/200**. Table and rounding
are scored as one model, 24 of them over 28 boxes, and the tie rule does the rest — the twelve
1000-em probes tie every rounding, so the two new ones are the whole separation:

| rounding, on `usWin` or `sTypo` under bit 7                           | fits      | separated by                |
| --------------------------------------------------------------------- | --------- | --------------------------- |
| **round half up, `floor(x + 0.5)`**                                   | **28/28** |                             |
| round half to even                                                    | 27/28     | `split-2000` ascent         |
| ceil                                                                  | 27/28     | `split-2048` descent        |
| floor                                                                 | 26/28     | both ascents                |
| half up, borrowing one from the ascent where the descent rounded down | 26/28     | `split-2048`, both sides    |
| exact                                                                 | 25/28     | all three fractional values |

The last row but one is Blink's own Linux hack, compiled only under subpixel positioning; it is
scored so that it is refuted on DirectWrite rather than assumed away, and T14's whole-pixel Linux
runner has never shown it either.

Two mutants on the experiment prove the design carries the answer and not the script: dropping
`split-2000` makes `analyse.ts` refuse with half-up and half-to-even tied at 26/26, and moving
`split-2048`'s descent onto a whole pixel (451 → 512) makes it refuse with half-up and ceil tied
at 28/28.

**The no-`BASE` ideographic baseline is the rounded descent.** `split-2048` carries no `BASE`
table on purpose, so its fallback is read off a descent that is not a whole pixel, and the browser
answered −220 at 1000 px and −22 at 100 px. Question 4's readings now take the size, and
"the `BASE` `ideo` coordinate of the `DFLT` script, else the face box descent, rounded as the box
is" fits 36/36 where the reading that shipped — the same, unrounded — fits 34/36. The winner is
derived from question 1's answer rather than enumerated, because half-up, floor and half-to-even
all round 220.21 the same way and would have tied.

The width and kerning questions rescore unchanged at 504/504 and 84/84; `split-2000` is the first
DirectWrite width evidence on the em Lato has, and the shipped quantiser fits it.

### What changed in the reader

`packages/cli/src/render/measure.ts` rounds the box the way the browser probe reads it: at
`FACE_BOX_PX`, half up, handed back as a fraction of the em, and the no-`BASE` ideographic fallback
is that rounded descent. `sfnt.ts` stays in exact font units, because T14 scores `FaceMetrics` in
units. The mutation pass on the reader: 7 tried, 6 killed — floor, ceil, exact, half-to-even and
the unrounded fallback each die on a named probe — and one equivalent, `1000` written for
`FACE_BOX_PX`, which is pinned instead by asserting the fixture was measured at `FACE_BOX_PX`, so
the day the constant moves the fixture has to be re-measured. A rounding written into `metricsOf`
in units dies on `split-2048`.

## The markup names the face it was measured in

The browser measurer and both emitters use the bare quoted family, so in a browser measurement and
drawing agree through one name. Only the CLI path measures in a resolved face the SVG never named:
`font-family="Calibri Light"` over line origins computed for Carlito.

`TextOptions.cssFamilyFor` is the hook, called once per piece with the run's family, weight and
style, defaulting to the quoted family so a caller that passes nothing emits byte-identical markup
— the fidelity baseline renders through that path and does not move. The CLI answers from the
resolved-face cache the measurer already filled, so it costs no second lookup and cannot name a
face other than the one that measured: `fontStack(asked, drawn)` leads with the drawn face, then
the asked one, then the substitution table, Calibri, Carlito, `sans-serif`, deduplicated.
`Calibri Light` drawn in Carlito reads `"Carlito", "Calibri Light", "Calibri", sans-serif`.

Drawn-first is the order because it is the only one that is true on the machine that made the
file: a viewer that has the drawn face draws exactly what was measured; one with only the asked
face draws the author's face on positions computed for a stand-in; one with neither lands where
PowerPoint would. There is no `substituted` branch — on a direct hit the drawn face is the asked
one and the list is `fontStack(asked)`.

One thing had to move for this to be safe. On the last-resort path `drawn` is a raw `name` string,
and a name carrying a quote or a control character sorts before every letter in code-unit order —
so the one face the markup cannot carry was the one the index would pick as the machine-wide last
resort. `claim` now skips a name the CSS shorthand refuses, as it already skipped a blank one.

Mutation: 8 tried, 8 killed — the hook ignored, the engine dropping the option, the default
returning the stack, the order swapped, the dedupe removed, `drawn` omitted, weight and style not
passed to the resolve, the guard removed. The weight-and-style mutant is invisible in the markup
(`pickIn` falls through all four slots) and dies only on a resolve count.

## Windows re-checks T13; macOS is asked

The T13 fixture was measured on one Windows 11 machine, and every reader test re-derives from it.
`probe-fonts` in `ci.yml` re-takes the measurement on `windows-latest`, scores it into the
committed fixture's path, formats it, and requires `git diff --exit-code` to be silent. Playwright's
Chromium is pinned by the lockfile, so the only free variable is Windows Server's DirectWrite, and
a diff there is a finding. It is in the release's `REQUIRED` list: it depends only on the commit.

`real-faces-macos` is the same T14 job on `macos-latest`, dispatch only and gating nothing.
`shippedBoxFor('darwin')` stays the DirectWrite reading — the run scores the assumption
[0045](0045-the-face-box-belongs-to-the-rasteriser.md) made, and `face-box-rival` names the table
if it is wrong; the summary now says CoreText is "scored against the DirectWrite reading it is
assumed to share" rather than calling the runner DirectWrite. Its first run is expected red on
`advance-quantum` — CoreText positions subpixel — and the artifact carries the face-box table
regardless, since `score()` runs every gate before `analyse.ts` throws.

So that the table question cannot depend on which faces a runner ships, both real-faces jobs now
index T13's probe fonts as a second `--font-dir`, written by `build-fonts.ts`: `split` separates
the three tables by 100 px on any rasteriser. Through T14's own pipeline on this machine the
DirectWrite reading fits the fourteen probes 14/14 with none separating, the FreeType composite
7/14 — the shape a macOS run will show whichever way CoreText reads. On Linux this also measures
`split` under FreeType, which 0044 noted had never been done. The T14 variants table now scores
the shared `BOX_ROUNDINGS` rather than its own four, so a Linux run records half-up against
half-to-even on the 2000-em Lato faces.

## Turbo did not serve a stale `dist`

[0044](0044-what-the-runners-own-fonts-said.md) recorded `FULL TURBO` over edited sources, fixed by
`--force --no-daemon`. Turbo 2.10.11's own schema marks `daemon` deprecated and unused by
`turbo run`, so `--no-daemon` was a no-op and `--force` did the rebuild — which says nothing about
why the hit happened. Three trials, each: append a blank line to `packages/cli/src/render/sfnt.ts`,
`--dry=json`, build, revert, build.

| step                            | task hash          | cache |
| ------------------------------- | ------------------ | ----- |
| baseline                        | `fd4aaa760d1308ad` | hit   |
| after the edit, trial 1         | `0fa97f0b68a6fdc7` | miss  |
| after the revert                | `fd4aaa760d1308ad` | hit   |
| after the same edit, trials 2–3 | `0fa97f0b68a6fdc7` | hit   |

The input hash of the edited file changed on every trial and the task hash with it. The hits on
trials 2 and 3 are on bytes Turbo had already built, and a hit restores `dist/**` from the cache,
which is the right `dist` for those bytes. Not reproduced. The most likely reading of 0044 is that
one: an edit, a build, a revert, and the same edit again, with the restore mistaken for a stale
replay. The carried entry closes as a convention rather than a fix: every CI job that scores
`packages/cli/dist` builds with `--force`, `gate.test.ts` asserts it, and the documented T13
commands carry it.

## Verification

`pnpm check` green. T13 rescored over fourteen probes: face box 28/28, kerning 84/84, widths
504/504, ideographic 36/36, and `analyse.ts` refused both probe mutants above. Mutants: 7 on the
reader (6 killed, 1 equivalent and pinned), 2 on the experiment (both refused), 8 on the markup
(8 killed), 5 on the CI jobs against `gate.test.ts` — the diff step removed, the job moved to
Ubuntu, dropped from `REQUIRED`, the macOS job running on push, `--force` dropped — all killed.

## Deviations

1. **`tools/ground-truth/lib/box-rounding.ts` is shared by two experiments.** T13 scores it and
   T14 records it, and a rounding named in one that the other spelt differently would be two
   findings about one fact.
2. **`build-fonts.ts` is a new experiment role** — `build-deck.ts`, for fonts — in
   `tools/ground-truth/fonts/metrics/`.
3. **T14's tolerance and quantum are as 0045 left them.** A macOS run re-derives the quantum on
   the record, not here.

## Open questions

1. **macOS**, until the dispatch runs. The reading it names is the finding, and a `coretext`
   backend is the follow-up if the table differs.
2. **A fractional `BASE` coordinate.** Every probe's coordinates are whole pixels at both baseline
   sizes; whether Chromium rounds one that is not is unmeasured, and the reader hands it back
   exact.
3. **The Linux borrow under subpixel positioning.** Refuted on DirectWrite and never seen on the
   whole-pixel CI runner; a Linux desktop with subpixel positioning on is neither.
4. **The browser renderer's own exported SVG** still carries a bare family, deliberately, to keep
   the fidelity baseline still; a file saved from the studio and opened elsewhere has the problem
   the CLI no longer has.
5. **A deck rendered on one platform and viewed on another** has two answers, as 0045 said. The
   markup now says which face the positions were computed for, which is the most a string can do
   without carrying outlines.
