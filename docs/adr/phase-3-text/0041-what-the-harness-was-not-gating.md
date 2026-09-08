# 0041 — What the harness was not gating, and where the difference is

- **Status** accepted
- **Sub-phase** 3.9, reopened
- **Supersedes** open questions 1, 3, 4 and 6 of
  [0035](0035-the-fidelity-harness.md); leaves 2, 5 and 7 open
- **Code** `tools/fidelity/blame.ts`, `tools/fidelity/record.ts`,
  `tools/fidelity/fidelity.ts`, `tools/gate2/gate2.ts`
- **Measurement** one deck exported twice across a minute boundary, against
  `corpus/decks/a09-fields.pptx`

## Context

[0035](0035-the-fidelity-harness.md) built the harness and left seven open
questions. Three were answerable, and going after them turned up something
neither of them named: **the gate had a hole in it, and had been passing over
three slides it held no record of.**

## The gate was silently ungating three slides

The gate is `sha256(raster) === recorded`, per slide. It decided that in one
line:

```ts
changed: was !== undefined && was.rasterSha256 !== raster.rasterSha256,
```

`was === undefined` — a slide the baseline has never heard of — is therefore
**not changed**, and a slide nothing has a record of passes. The baseline was
last written at Gate 2 with 152 slides. [0039](0039-upright-glyphs-in-vertical-text.md)
then made `b02-layouts-10` and `b02-layouts-11` draw, and
[0040](0040-the-wordart-column.md) made `a10-rtl-cjk-03` draw. Neither re-recorded.
So for two sub-phases the harness reported **PASS on 155 slides while gating
152**, and said nothing.

A gate whose whole claim is that it has no tolerance cannot have an implicit one
for slides it has not seen. `assertBaselineCovers` now throws
`FID_BASELINE_INCOMPLETE` naming them, and the reverse case — a recorded slide
that no longer draws — fails as `FID_RENDER_CHANGED` rather than quietly
shrinking the set.

## Four declared failures could not happen

`FidelityErrorCode` declared fourteen codes and threw ten. The four that could
not happen were not spare: each named a real property nothing was checking.

| code                          | what it claimed                         | now                                                                                         |
| ----------------------------- | --------------------------------------- | ------------------------------------------------------------------------------------------- |
| `FID_RECORD_IN_CI`            | a baseline is never recorded in CI      | thrown; `process.env.CI` is read                                                            |
| `FID_RECORD_UNJUSTIFIED`      | a re-record states its reason and count | thrown; `--why` and `--expect <n>` are required to overwrite                                |
| `FID_RASTER_NONDETERMINISTIC` | our raster agrees with itself           | thrown; `--record` rasterises every slide twice                                             |
| `FID_ORACLE_NONDETERMINISTIC` | PowerPoint's export is stable           | **deleted** — `analyse.ts` deliberately records jitter instead, so the code contradicted it |

`--record` is the only way to turn a red gate green: it _is_ the escape hatch,
and it had no lock on it at all. It now needs a written reason and the exact
number of slides it will move, so a re-record that quietly carries a second
change fails on the count rather than landing. `--expect 3` is what re-recorded
the baseline above.

**A test now asserts every declared code is thrown somewhere**, reading the
sources of `tools/fidelity`, `tools/gate2` and the experiment, and skipping
`errors.ts` and the tests themselves so that declaring a code or quoting it in an
assertion does not make it look live. Re-adding
`FID_ORACLE_NONDETERMINISTIC` fails it by name.

Two more recorded-and-unchecked facts closed with it: `Environment.chromium` was
written into every baseline and never compared, so a Skia change arrived as an
unexplained `FID_RENDER_CHANGED`; it now throws `FID_BROWSER_CHANGED`. And
`servedUrl` — the check that a URL stays inside the two trees the page may read —
was exported and bypassed at all four call sites, which now use it.

## The one record a runner may make

The locks above collided with the only way to get a Linux baseline: a baseline
has to be recorded on the platform it gates, the only Linux this project has is a
GitHub runner, and `FID_RECORD_IN_CI` refuses to record there.

Resolved by looking at what the lock is for rather than by relaxing it. The
danger is a run that **rewrites** the record and calls the result agreement, and
that danger needs a record to rewrite. A platform's _first_ baseline overwrites
nothing, gates nothing until a person commits it, and therefore cannot turn a red
run green — the same distinction `--why`/`--expect` already draw locally, where a
first record needs no case made for it.

So `--bootstrap` writes a platform its first baseline and nothing else:

- in CI a bare `--record` is still refused, with or without a reason and a count;
- `--bootstrap` is refused the moment that platform has a baseline, on **any**
  machine, so it cannot be reached for when a gate goes red;
- the job is `workflow_dispatch` only, never on a push or a pull request, and it
  uploads rather than commits. The commit is the review.

Verified by running the runner's exact command here against a platform with its
baseline removed: it wrote all 155 digests **byte-identical** to the recorded
ones, and both refusals fire.

`record-fidelity-baseline` in `.github/workflows/ci.yml` is that job. It uploads
`corpus/ground-truth/manifest.json` beside the baseline, because `--record` also
claims the new fixture there and `pnpm corpus` refuses a file under `corpus/`
that no entry claims.

### What the Linux baseline turned out to say

Recorded on `ubuntu-latest` by that job, 155 slides. Two things in it are worth
keeping:

- **The runner has none of the faces the corpus asks for.** It reports seven
  families requested and six missing — Aptos, Aptos Display, Calibri, Calibri
  Light and both `PptxStudio` faces — and every one of the seven, the "present"
  one included, measures the same 698.046875 advance for `Hamburgefonstiv` at
  100px. They are all resolving to one fallback. The Linux score against
  PowerPoint is therefore much weaker evidence than the Windows one, and the
  report's substituted-faces list is what says so.
- **Only 2 of 155 rasters, and 2 of 155 SVG strings, match the Windows
  baseline.** Text measurement moves the layout, not just the pixels. That is
  the case for a per-platform baseline stated as a number rather than assumed.

The browser is the _same_ build on both — `HeadlessChrome/151.0.7922.34` — because
Playwright is pinned in the lockfile. So `FID_BROWSER_CHANGED` will fire on a
lockfile bump, which is when it should, and not on crossing platforms.

### CI had been red on every push, and it was the step order

Dispatching that job ran the whole workflow, and `check` failed on Linux at
**Lint**, on `apps/studio/src/export.ts`, with `'Report' is an 'error' type that
acts as 'any'`. It had been failing on every push to `main` since at least
2026-09-07 and I had not looked.

`pnpm lint` and `pnpm typecheck` are type-aware, and a workspace import resolves
through the package's `exports` to `dist/index.d.ts`. I had ordered **Build after
both**, so on a clean checkout they resolve nothing and report every use of a
cross-package type as `any`. It never showed up here because a `dist` from an
earlier build is always lying around. `pnpm check` had the same order, so a fresh
clone would have failed identically.

Reproduced locally rather than guessed at: removing `packages/validate/dist`
gives three errors in that file, and putting it back removes them. Build now runs
before Lint in both, and the steps `pnpm check` had that CI did not — structure,
references, fidelity, Gate 2 — are in the job.

## Where the difference is

0035's open question 6: "`maxD` and `meanBp` are reported per slide, but nothing
yet localises _where_ on a slide. A reviewer gets a number and a slide, not a
region and a shape name."

The join was cheaper than expected, because both halves already existed:
`layoutSlide` and `flatten` compose group transforms into slide EMU and are
already loaded in the harness page. `drawSlide` now returns every placed shape
scaled to raster pixels, and `boxOf` turns each into the axis-aligned box a
rotated rectangle occupies.

**Attribution is by frame in paint order, and has no floor.** Each differing cell
is charged to the last shape whose box covers it; the per-shape totals therefore
add up to the slide's own `sumD` and nothing is filtered out before a reader sees
it. Ranked per cell the shape owns rather than by total difference — otherwise a
slide-sized diagram nobody has built yet outweighs every real defect in the
corpus put together and is the whole table.

`regionsOf` was already live in Gate 2's page but printed bare cell coordinates;
`regionShapes` names them, and both reports now say what is wrong rather than
where.

### What it found on its first run

Ranked by agreement over each shape's own area, the corpus's worst-drawn shapes
are not the worst-scoring slides. Four things nobody had localised:

| shape                            | slide              | meanBp over its own area | maxD | what it is                                                                      |
| -------------------------------- | ------------------ | -----------------------: | ---: | ------------------------------------------------------------------------------- |
| `blipFill stretch fillRect`      | `a03-fills-02`     |                 **8272** |  159 | **not in any open question**                                                    |
| `dkVert` / `narVert` / `narHorz` | `a03-fills-03`     |                 **8316** |   71 | three of the 54 pattern tiles                                                   |
| `gradFill path rect`             | `a03-fills-01`     |                 **9116** |   69 | [0022](../phase-2-geometry-and-paint/0022-fills.md) open question 3, quantified |
| `gradFill path shape`            | `a03-fills-01`     |                     9678 |   28 | 0022 open question 2, quantified                                                |
| `Vector with a raster fallback`  | `a25-svg-blips-01` |                 **5906** |  197 | SVG blips, which are 10.7                                                       |

The two 0022 questions had been recorded and never measured; they now have
numbers. `blipFill stretch fillRect` is the interesting one:
[0038](../phase-2-geometry-and-paint/0038-gate-2.md)'s open question 4 named
`a03-fills-02`'s residual as "its right column — `blipFill tile flip` above
`blip biLevel lum`", which is what a heatmap shows. There is a **third** bad
shape in the middle of that slide at maxD 159, and eyeballing the picture had
missed it twice.

That is the argument for the feature in one row: the reviewer was looking at the
right slide and still read the wrong answer off it.

## The clock is smaller than the rasteriser

0035's open question 4: the noise floor is measured by exporting each slide twice
inside the same second, so a time-of-day `a:fld` has not moved between them and
the recorded floor is the rasteriser's alone.

Measured directly. `a09-fields` exported twice, 91 seconds apart, 10:36:50 and
10:38:21, so the minute changed:

|                        |                                                  |
| ---------------------- | ------------------------------------------------ |
| slides 1 and 3         | raster **byte-identical**                        |
| slide 2                | 31 cells of 8160, **maxD 6**, meanBp still 10000 |
| where                  | 3 regions, inside `Fields 1` and `Fields 2`      |
| the rasteriser's floor | maxD 47 over 245 cells, 4 bp                     |

**The content term is an order of magnitude below the rasteriser term and does
not move the floor.** The question is answered rather than worked around: the
recorded floor is honest as it stands.

One thing it does _not_ bound. Ninety-one seconds moves a minute field and
cannot move a date field, so `datetime1`–`datetime10` are untouched here. Their
drift is not periodic — it grows as the capture recedes — and stays recorded as
0038's open question 3.

## Verification

- **CI runs the gate.** `pnpm fidelity` is a step of the `check` job, against
  `expected.linux-x64.json` recorded on the same runner, alongside `pnpm
structure`, `pnpm references` and `pnpm gate2` — four steps `pnpm check` had
  and CI did not.
- **`pnpm references` earned its place on its first CI run**: `tools/bench/README.md`
  linked to `CLAUDE.md`, which is gitignored, so the link resolved here and
  nowhere else. A check that only ever runs on the machine that wrote the file
  cannot see that class of break at all.

- `pnpm fidelity`: 155 slides scored, 0 not drawn, corpus mean **9768 bp**, gate
  PASS — and now gating all 155 rather than 152.
- **20 of 21 mutants killed.** Paint order, the raster clip, the partial cell,
  the `covered` denominator, both rankings, the region limit, the turned box's
  spread and its centre, every record lock including all three `--bootstrap`
  branches, both directions of a baseline gap, and the half-up rounding.
- Two of those assertions exist because the first sweep did not have them: the
  parser accepted `--bootstrap` with nothing asserting it, and the CI branch
  spelled a condition the guard above it had already made redundant, which is
  now written the short way rather than tested twice.
- The survivor removes the **vertical** clamp in `ownersOf` and is equivalent,
  not a missing assertion: with the column already clamped, an out-of-range row
  always yields a flat index outside the array, which `Int32Array` drops.
  Removing the **horizontal** clamp is the one that corrupts — it wraps a shape
  off the right edge onto the next row — and that mutant dies.
- 3265 → **3298 tests**, 91 → 93 files.
- The shape probe changes no markup: all 155 raster digests are unchanged across
  it, which is the same check 0038 used to prove the `drawSlide` extraction.
- `pnpm check` green.

## Deviations

| plan item                 | disposition  | why                                                                                                                                                                                                                                                                                                                                     |
| ------------------------- | ------------ | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Turborepo inputs          | **rejected** | Turbo hashes declared _files_. The environment gate keys on installed fonts and the browser, which are properties of the machine and not files, so a font change could not invalidate a cache entry and Turbo would replay a stale PASS - the exact poisoning the plan raised inputs to prevent. Not caching it prevents it absolutely. |
| localise with `regionsOf` | **both**     | 0035 assumed regions were the answer. Regions need a floor and a floor is a chosen number; shape attribution needs none, so it leads and regions name their shapes beside it.                                                                                                                                                           |

## Open questions

1. **Carried from 0035, still open: the pinned Chromium flags are unverified
   where they matter.** There is now a Linux baseline, but it was recorded with
   the flags on and nothing has recorded one with them off, so whether they
   change anything on a runner is still unmeasured. Two dispatches of
   `record-fidelity-baseline` either side of `PINNED_ARGS` would settle it, and
   the answer decides whether the flags are insurance or ceremony.
2. **A runner image bump will break the environment gate, correctly and
   inconveniently.** The Linux baseline pins the fonts `ubuntu-latest` had on
   the day it was recorded. When GitHub changes that image's font package the
   gate throws `FID_FONT_COVERAGE_CHANGED` and someone has to dispatch a
   re-record. That is the gate working; it is also a maintenance cost nobody has
   paid yet, and no measurement here says how often it falls due.
3. **`blipFill stretch fillRect` is wrong and nothing owns fixing it.** 2.12 built
   image fills and 2.14 fixed two of the blip effects; this is a third defect on
   the same slide, found by attribution rather than by eye, and it needs a probe
   before it needs a patch.
4. **Attribution is by the frame, not by what the shape painted.** A `noFill`
   shape sitting over another still claims its cells, and text that overflows its
   frame is charged to whatever is underneath. Both are visible in the report as a
   name that does not fit the difference, and neither is measured.
5. **`(no shape)` conflates two very different things** — a background we drew
   differently, and a `p:graphicFrame` we did not draw at all. `a22-chartex-01`
   charges 99.6% of its difference to `(no shape)` for the second reason, and the
   report does not say which it is.
