# ADR 0035 — The fidelity harness

**Sub-phase 3.9.** Status: accepted. Supersedes nothing.

---

## The decision in one line

**Two questions, kept apart, because only one of them has an exact answer.** How far our raster is
from PowerPoint's is _reported_ and never gated, because the two differ everywhere and any pass/fail
on that number would need a tolerance. Whether our own raster changed is _gated_, on a SHA-256, with
no tolerance to loosen.

---

## What this sub-phase had to decide

The plan specified: "digest-pinned container, repo-checked-in `fonts.conf`, `odiff`,
`fidelity/scores.json`, PR delta table", with a determinism warning about Calibri falling back to
Liberation Sans on a GitHub runner, and an acceptance criterion of "re-run twice on different
machines, identical scores".

Three of those five are downloads, which this repository's Hard rule 2 forbids without the owner
saying yes: a LibreOffice container image, the `odiff` binary, and the metric-compatible fonts a
`fonts.conf` would alias to. So the first question was whether the sub-phase is blocked. It is not,
and the design that avoids all three is better on its own merits rather than merely cheaper.

---

## The instruments

- `tools/ground-truth/render/fidelity/read.ps1` — PowerPoint 365 (16.0 build 20326) draws all 149
  slides through COM and exports each as a PNG, **twice**.
- `tools/ground-truth/render/fidelity/analyse.ts` — refuses a colour-managed PNG, reduces both
  exports through the browser, commits the first and records how far the two disagree.
- `tools/fidelity/` — the harness: the metric, the browser, the environment probe, the report.
- `corpus/ground-truth/render/fidelity/` — the committed oracle, 50 files, 1.84 MiB.

The corpus is 50 committed `.pptx` decks and **149 slides**. The two `.pptm` decks are excluded:
PowerPoint prompts on opening a macro-enabled file unless macro security is lowered, and a capture
that depends on a machine's security policy is not one anybody else can reproduce.

---

## The findings

### PowerPoint does not rasterise a slide identically twice

The check that was supposed to settle this on day one fired on day one. Exporting every slide twice
in the same process and comparing bytes, **28 of 149 differed**.

The first reading was wrong and worth recording: these are palettised PNGs, and a differing palette
order changes the bytes while decoding to identical pixels. So the check was rewritten to compare
_decoded pixels_. The 28 remained. It is the picture that varies, not the container.

But the magnitude decides what it means:

|                                                  |                             |
| ------------------------------------------------ | --------------------------- |
| slides differing between two consecutive exports | 28 / 149                    |
| worst `meanBp` drop                              | **4 bp** (99.96% agreement) |
| worst `maxD`                                     | 47 / 255                    |
| worst cells touched                              | 245 / 8160 (3%)             |
| slides whose _reduced grid_ was identical anyway | 3                           |

Three slides differ in raster digest and agree completely once reduced — the difference is inside a
cell. This is antialiasing jitter at edges, not a different picture. It is concentrated on SmartArt,
pictures, CJK and the non-16:9 decks.

**This is the finding that shapes the whole design.** An oracle that disagrees with itself cannot be
a gate. Had the gate been "score against PowerPoint", the harness would have been flaky from its
first run, and the standard response to a flaky gate is to widen it until it stops complaining —
which is exactly the move this repository forbids. So the oracle became a _reported_ number carrying
a published noise floor, and the gate moved to the one comparison that is exact.

### Our own raster is exact

Measured before relying on it: four decks, five fresh `chromium.launch()` calls, SHA-256 of the full
padded RGBA. **Identical every time**, and identical again with and without the pinned Skia flags.
The SVG string is byte-identical too.

So the gate is `sha256(raster) === recorded`, with no tolerance anywhere in the pass/fail decision.

The pinned flags — `--force-color-profile=srgb`, `--disable-lcd-text`,
`--disable-font-subpixel-positioning`, `--font-render-hinting=none`, `--disable-remote-fonts` — were
measured to change nothing on win32-x64. They are kept as insurance for a Linux runner, where
FreeType hinting and LCD filtering are real and version-dependent. That they matter there is
**unverified**, and this record says so rather than implying they were shown to.

### The score is reproducible to the basis point

The plan wanted two machines. One exists. The honest substitute is two independent full captures —
PowerPoint reopened, all 50 decks re-exported, everything re-reduced:

|                                                      |                        |
| ---------------------------------------------------- | ---------------------- |
| oracle grid files byte-identical across two captures | **49 / 50**            |
| corpus mean, capture A vs capture B                  | **9719 bp vs 9719 bp** |
| per-slide scores that moved                          | **0 / 146**            |
| our own raster digests across both runs              | identical              |

### The one file that differed is a clock, and the jitter check cannot see it

`a09-fields` is the deck of `a:fld` date and time fields, and `datetime11`, `datetime12` and
`datetime13` are the time-of-day variants. Two captures minutes apart render different text; the
difference localises to four small regions, worst 17 cells, `maxD` 9.

The double-export check is blind to this **by construction**: both exports of a slide happen within
the same second, so the clock has not moved. The measured noise floor is therefore the noise floor of
the _rasteriser_, and not of the _content_. One slide's oracle encodes a moment in time, and the
score for that slide compares a stale clock against a live one over a few dozen cells. At corpus
scale it is 0 bp, which is why it is recorded here rather than worked around.

### Where we actually stand against PowerPoint

**Corpus mean 9719 bp** — 97.19% cell agreement over 146 scored slides. The furthest ten:

| slide           | meanBp | what it is                             |
| --------------- | -----: | -------------------------------------- |
| a23-smartart-01 |   5677 | SmartArt, which is sub-phase 4.6       |
| b04-table-01    |   7653 | tables, which are 4.1-4.4              |
| b06-smartart-01 |   7699 | SmartArt                               |
| a23-smartart-03 |   8038 | SmartArt                               |
| a23-smartart-02 |   8051 | SmartArt                               |
| a03-fills-02    |   8199 | **fills, which are already built**     |
| a22-chartex-01  |   8286 | ChartEx, which is 10.1                 |
| a37-mce-01      |   8302 | `mc:AlternateContent`                  |
| a37-mce-02      |   8601 | `mc:AlternateContent`                  |
| b09-picture-01  |   8642 | `a:blipFill`, the known 2.11 scope gap |

Eight of the ten are content types no phase has built yet, which is the harness working: it ranks
what is missing. `a03-fills-02` is the row that is not explained by scope, and it is open question 3.

---

## What was built

### The metric

A raster reduces to an integer grid: 8x8-pixel cells, BT.601 luma at 8-bit precision, chroma
decimated a further 2x2. A 16:9 slide is drawn 960 wide, padded to whole cells, and reduces to
120x68 = 8160 luma cells and 60x34 chroma cells each.

Per luma cell, `d = max(|ΔY|, |ΔCb|, |ΔCr|)` — the largest of the three, never their sum or mean. A
wrong hue at equal luma is a real defect, and averaging it against two agreeing channels is how it
becomes invisible. `(0,131,0)` and `(255,0,0)` share both `Y = 77` and `Cb = 85` and differ only in
`Cr`; the test suite uses exactly that pair.

```
meanBp = 10000 - round(10000 * Σd / (255 * N))
```

integer, half up, so two machines cannot round it differently. **There is no tolerance constant and
no deadband**: a single pixel one level off in one cell moves the number, and there is therefore no
constant anybody can later widen.

The reduction is `tools/fidelity/metric/reduce.ts`, written with no free variables so that
`page.addScriptTag` can send its own `toString()` into the browser. Our SVG and PowerPoint's PNG are
decoded by the same Chromium through the same `<img>` and `drawImage`, and reduced by the same
source. There is consequently **no PNG decoder in this repository**, no gamma argument, and no
second decode path to keep in agreement.

### The environment gate

The plan's answer to the Calibri trap was to vendor Carlito and alias to it. That pins a _cause_, and
only the causes somebody listed. This pins the _effect_, and needs no download:

- the advance of `Hamburgefonstiv` at 100px — 3.7's own fingerprint, reused rather than reinvented
- SHA-256 of the 8-bit alpha of a pangram drawn in the face

The advance alone is not enough by construction: being indistinguishable by advance is precisely what
"metric-compatible" means, so a clone would pass. The coverage hash is what separates them. Either
disagreeing throws `FID_FONT_METRICS_CHANGED` or `FID_FONT_COVERAGE_CHANGED` naming the family. A
machine whose fonts have moved does not score; it stops.

This also detects a _different version of the same-named font_, which a fontconfig alias never would.

### What the corpus actually asks for

The probe reports that the corpus draws in **Aptos**, **Aptos Display**, `PptxStudio Alpha` and
`PptxStudio Bravo`, and this machine has none of them. Aptos is Office's current default and no open
clone exists; 3.7 already falls back to Carlito, which is also absent here. So even the Windows
baseline scores substituted text against PowerPoint's own substitution — two different fallbacks
compared to each other. The report names every substituted face for exactly this reason.

---

## Verification

- `pnpm fidelity` — 146 slides scored, 3 not drawn, corpus mean 9719 bp, gate PASS.
- The three not drawn are the `a:bodyPr/@vert="eaVert"` gap recorded as 0034's open question 5. A
  direction the renderer refuses is reported as a gap, never scored as zero: scoring a slide we did
  not draw would put a number on something that never happened.
- **Metric mutation sweep: 12 mutants, 12 killed**, control green. Every BT.601 coefficient, both
  rounding rules, the chroma clamp, the chroma decimation rounding, the `max` over channels, the
  connectivity, the region floor and the region ranking.
- The first pass of that sweep killed 8 of 12. All four survivors were **missing assertions, not
  equivalent mutants** — the toys had uniform chroma inside every chroma cell and a `Σd` where half-up
  rounding never changed the integer. The cases that bite were derived and added.
- Storage: 1.84 MiB of oracle in 50 files, largest 135 KiB against a 512 KiB per-file cap. `corpus/`
  totals 7.0 MiB against a 12 MiB cap.

---

## Deviations from the plan

| plan item                           | disposition         | why                                                                                                                                                                                                                                                                                                                                                                                              |
| ----------------------------------- | ------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| digest-pinned LibreOffice container | **dropped**         | LibreOffice is not the format's author. Optimising a score against its raster is transliteration performed by an optimiser rather than by copy-paste, which `LEGAL.md` forbids in the slower form. It also lacks Calibri, so it does not solve the problem it was proposed for, and it is a second drifting variable in a design whose priority is not drifting. And it is a forbidden download. |
| repo-checked-in `fonts.conf`        | **dropped**         | Vendoring fonts is a download, and bundling the OFL set is scheduled as 8.8. The font environment is measured and pinned per `envId` instead of configured.                                                                                                                                                                                                                                      |
| `odiff`                             | **dropped**         | A forbidden install, and a native binary is a determinism risk in a harness whose whole point is determinism. Replaced by an integer, dependency-free metric with a 12-mutant sweep over it.                                                                                                                                                                                                     |
| `fidelity/scores.json`              | **kept**            | Written every run, gitignored, byte-stable: fixed key order, no timestamps, no host names, integers only.                                                                                                                                                                                                                                                                                        |
| PR delta table                      | **kept, reordered** | Leads with the changed-slide count and the worst slide. A mean over 149 slides cannot move for a defect on one of them, so putting it first would train reviewers to ignore the report. It is labelled "provenance, not signal".                                                                                                                                                                 |
| "two machines, identical scores"    | **replaced**        | One machine exists. Two independent full captures, 49/50 grid files byte-identical and 0/146 per-slide scores moved, is the strongest form available here — and it tests something the original did not: that the _oracle_ is reproducible, not merely that the harness is.                                                                                                                      |
| the score is the gate _(implicit)_  | **replaced**        | The gate is the SHA-256 of our own raster. The score reports and ranks. This is what removes every tolerance from the pass/fail decision, and it is forced by the measurement that PowerPoint disagrees with itself.                                                                                                                                                                             |
| Turborepo inputs                    | **not done**        | `pnpm fidelity` runs directly and is not a Turbo task, so there is no cache to poison yet. Wiring it belongs with the CI workflow, which cannot exist until a Linux baseline is recorded. Open question 1.                                                                                                                                                                                       |

---

## Open questions

> Questions 3, 4 and 6 are answered in
> [0041](0041-what-the-harness-was-not-gating.md), which also records a hole in
> the gate that none of these named. 1, 2, 5 and 7 stand.

1. **There is no `expected.linux-x64.json`, so CI cannot run the gate.** The baseline for a platform
   has to be recorded on that platform, and this machine is Windows. Until a one-off
   `workflow_dispatch` run records it, `pnpm fidelity` is a developer gate and not a CI one. The
   `FID_ENV_UNKNOWN` throw is what stops a Linux run from silently recording its own baseline and
   calling it agreement.
2. **The pinned Chromium flags are unverified where they matter.** They change nothing on win32-x64.
   Whether they are sufficient on a Linux runner is exactly what question 1 would answer.
3. ~~**`a03-fills-02` scores 8199 bp and nothing in the plan explains it.**~~ **Answered, and
   neither way round.** Slide 2 is the image-fill probe slide and image fills were not drawn at
   all; [2.12](../phase-2-geometry-and-paint/0036-image-fills.md) took it to 9378 and
   [2.14](../phase-2-geometry-and-paint/0038-gate-2.md) to 9638. The per-region diff this asked for
   is [0041](0041-what-the-harness-was-not-gating.md), and it finds a third bad shape on the slide
   that 0038 missed.
4. ~~**The noise floor measures the rasteriser, not the content.**~~ **Measured** in
   [0041](0041-what-the-harness-was-not-gating.md): two exports 91 seconds apart move 31 cells at
   maxD 6 and 0 bp, against the rasteriser's 245 cells at maxD 47. The content term does not move
   the floor.
5. **The score is blind to compensating sub-cell errors** — text one pixel left and a rule one pixel
   right inside the same 8x8 cell. This is intrinsic to any spatial pooling and is not eliminated;
   the exact raster gate still catches it, because the gate does not pool.
6. ~~**`maxD` and `meanBp` are reported per slide, but nothing yet localises _where_ on a slide.**~~
   **Built** in [0041](0041-what-the-harness-was-not-gating.md): every differing cell is charged to
   the shape that painted it, and Gate 2's regions name theirs.
7. **149 slides is not the world.** Every deck in it is one we wrote.
