# 0042 — Rendering without a browser, and the first release

**Sub-phase 3.10 · accepted**

Publish `@pptx-studio/render-svg` and `pptx-studio render`. The plan's one line about this
sub-phase — _"server-side PPTX thumbnailing with no LibreOffice is an adoptable product on its
own"_ — turned out to hide a question nobody had asked: **how does a renderer measure text where
there is no canvas?**

Everything else was already reachable from Node. Geometry, fills, strokes, effects and images are
pure functions over the model, and an image's size comes out of its own header rather than a
decoder (0036), so `renderSlide` has always been a synchronous string function. I checked before
designing anything:

```
GEOMETRY-ONLY ok, bytes = 204
WITH TEXT THREW: TextError | no OffscreenCanvas in this environment
```

204 bytes for a slide of a text-cascade deck. Geometry-only is not a thumbnailer; it is a blank
rectangle with a background. So the sub-phase is really about the measurer, and the rest is
packaging.

## The verb had already promised no browser

`main.ts` has listed the unbuilt verbs since 0.8, each with the sub-phase that brings it. Mine said:

> `render` render slides to SVG or PNG **without a browser** (sub-phase 3.10)

I wrote that and then had to honour it. Three ways out, and only one survives:

|                                |                                                                                                                                                     |
| ------------------------------ | --------------------------------------------------------------------------------------------------------------------------------------------------- |
| **Ship without text**          | Refuted by the 204 bytes above.                                                                                                                     |
| **Drive a headless browser**   | Playwright is ~300MB and a peer dependency the pitch explicitly disclaims. It also makes the CLI unusable in the container where it is most wanted. |
| **Read the font's own tables** | What the rest of this ADR is about.                                                                                                                 |

PNG is a different matter and I dropped it: rasterising means shipping a rasteriser, and the whole
claim of the verb is that it needs nothing but Node. The help text now says SVG, and the README
says why. **This is a deviation from the plan and from my own earlier wording.**

## A second measurement engine, and why that is allowed

`packages/text/src/runs/measure.ts` says, in as many words, that _nothing anywhere else in the
package is allowed to measure by a different route_. A table reader is exactly such a route.

The rule holds because the reader is **outside** that package, in `@pptx-studio/cli`, which is layer
6 and the one place `node:*` is legal. It has to be: font discovery reads the filesystem. Putting it
in `packages/text` would have broken that package's stated invariant and dragged Node into a browser
package at the same time.

What makes two engines safe is not the package boundary, though. It is that I measured what the
first one does instead of assuming it.

## T13 — asking the browser, in writing

The repository's rule is to ask the format's author before writing code. Here the author of the
answer is Chromium, and the difficulty is that **a real font cannot be used to ask**: Arial's
`hhea.ascender` and its `OS/2.usWinAscent` are both 1854, so a browser reading either would look
identical. Every implementation that "works on Arial" is untested.

So `tools/ground-truth/lib/truetype.ts` builds fonts that disagree with themselves on purpose, and
`tools/ground-truth/fonts/metrics/measure-in-browser.ts` loads them into Chromium as data URIs under
family names that exist nowhere on earth. Seven fonts, six strings, six sizes from 8px to 1000px.

### The face box is `usWin`, unless bit 7 says otherwise

One font with `hhea` 800/−200, `usWin` 900/300 and `sTypo` 700/−100 — three distinct pairs. Chromium
reported **900/300**. The same bytes with `fsSelection` bit 7 (`USE_TYPO_METRICS`) set reported
**700/100**.

| reading                                             | fits      |
| --------------------------------------------------- | --------- |
| `usWin`, or `sTypo` when `fsSelection` bit 7 is set | **14/14** |
| `OS/2.usWinAscent` / `usWinDescent`                 | 12/14     |
| `OS/2.sTypoAscender` / `sTypoDescender`             | 12/14     |
| `hhea.ascender` / `descender`                       | 10/14     |

`hhea` is the reading everybody writes first and it is never right. Bit 7 matters more than the
table suggests: most fonts shipped since about 2015 set it, so a reader that ignores it is wrong on
the modern half of any font directory. The face box is what 3.8 divides between ascent and descent
to place the baseline, so getting it wrong moves every line of text.

### GPOS beats the legacy `kern` table

Three fonts: one with a `kern` table only, one with a GPOS `kern` feature only, one with both saying
**different numbers** — −200 and −100. Chromium took −100.

| reading                                  | fits      |
| ---------------------------------------- | --------- |
| GPOS when present, else the `kern` table | **42/42** |
| GPOS `PairPos` only                      | 36/42     |
| the `kern` table when present, else GPOS | 36/42     |
| the legacy `kern` table only             | 30/42     |
| neither                                  | 24/42     |

The third font is the only probe that separates the top three, and it is the one that matters: the
legacy table is far simpler to parse, so "read `kern`, and GPOS if there is no `kern`" is the
natural shortcut. It is wrong on every font carrying both, which includes most Microsoft core fonts.

### The advance is linear, and quantised at 1/65536 px

The question behind the question: **does the browser hint advances?** If it did, no table reader
could ever match it. It does not — the exact-float reading is never more than **1.04e-4 px** out,
which is four orders of magnitude under the 0.149% median by which the same browser disagrees with
PowerPoint (T2). That is the finding the whole approach rests on.

Chasing the residual gave an exact rule:

| reading                                             | fits        |
| --------------------------------------------------- | ----------- |
| advance truncated, kern rounded, both at 1/65536 px | **252/252** |
| advance and kern both truncated                     | 244/252     |
| the whole string quantised once                     | 168/252     |
| advance and kern both rounded                       | 154/252     |
| exact float, no quantisation                        | 88/252      |

Implementing it costs three lines and buys a test that asserts **equality** rather than a tolerance,
which is worth much more than the 1e-4 px it recovers. Truncation and flooring both fit, because
every advance in these fonts is positive; the fixture says so rather than picking one.

### A seventh font, added because a mutant survived

The first mutation run left five survivors, one of which was _"prefer the format 4 cmap over format
12"_. It survived because every probe font had only a format 4 subtable — the preference had nothing
to choose between. So I added `astral`, a font carrying both, mapping U+20000, which format 4 cannot
address at all. A reader that takes the first subtable it recognises now answers "no glyph".

That is the mutation pass doing its job: it did not find a bug in the code, it found a **claim with
no evidence behind it**, and the fix was a new measurement rather than a new assertion.

## The reader

`packages/cli/src/render/` — the tables (`sfnt.ts`), which file is which family (`faces.ts`), the
two probes `render-svg` asks for (`measure.ts`), and the verb (`render.ts`).

`sfnt.ts` reads `head`, `hhea`, `OS/2`, `name`, `cmap` (formats 4, 12, 6 and 0), `hmtx`, `GPOS`
`PairPos` (formats 1 and 2, with both coverage and both class-definition formats) and `kern`
(format 0). It reads a `ttcf` collection and an `OTTO` CFF font, because advances live in `hmtx`
whatever the outlines are. It decodes no outline and never will; this is not a font library.

`faces.ts` indexes the platform's font directories plus any `--font-dir`, which are searched first
so a caller can override a face without installing one. A file that will not parse is skipped rather
than fatal — a real font directory holds bitmap-only files, damaged files and files that are not
fonts, and one of them must not stop a deck rendering. Substitution goes through
`@pptx-studio/text`'s own table, so the two renderers cannot fall back differently, and the choice is
**reported** rather than hidden:

```
1 slide(s) at 1920x1080 -> thumb.svg
2 typeface(s) from 2 indexed face(s), 1 substituted
  Calibri Light -> Calibri
```

`measure.ts` implements `TextMeasurer` and `FaceBoxProbe`, which `render-svg` already accepted as
injectable — the seam was there before this sub-phase needed it, which is the strongest evidence
that bet 4 was structured right.

## What it cannot do

**No shaping.** Latin, Greek and Cyrillic are correct, and so is CJK, whose advances do not depend
on context. Arabic, Devanagari and the other scripts that need a shaper will measure **wide**,
because the reader sums unshaped advances and a shaper would have joined and substituted them. This
is a real limit, it is in the README, and the browser renderer does not have it.

**No PNG**, as above.

## What running the release path found

Same lesson as 3.9, and I did not learn it hard enough the first time: **code that has never been
run is not working code, it is code that has never been run.**

- The **Release workflow has failed on every push to main**, six times, since it was written in 0.1.
  `changesets/action` renamed all four of its inputs in v2.1.0 and fails outright on the old ones. I
  pinned v2.1.1 and kept v2.0 argument names.
- **No published package carried a licence or a notice.** Every manifest had
  `files: ["dist", "README.md"]`. Apache-2.0 §4(a) requires a copy of the licence to travel with the
  work and §4(d) requires the NOTICE text — and ours is not decorative: it carries Apache POI's
  attribution for the 187 preset shape definitions in `@pptx-studio/geometry`. npm includes a
  `LICENSE` at a package root regardless of `files`; it does **not** do the same for `NOTICE`.
  `pnpm legal` (`tools/repo/check-legal.ts`) now copies both into every publishable package and
  fails if either drifts from the root.
- **Publishing `render-svg` means publishing ten more packages.** Its closure is `model`,
  `geometry`, `paint`, `text`, `opc`, `xml`; the CLI adds `census`, `validate`, `writer`. That is not
  a choice — a dependency has to be resolvable. `render-dom` is the only package held back, through
  changesets' `ignore`, because nothing in this sub-phase asks anyone to depend on it.

## Verification

- **`pnpm check` green**, all of it, including the new `pnpm legal` step.
- **19 of 19 mutants killed, no survivors.** The first run killed 14; the five survivors were one
  missing measurement (the cmap preference, above) and four weak assertions, all of which asserted a
  property that a wrong implementation also has — comparing two renders' aspect ratios to each other
  passes for `height = width` too. Every one is now pinned to a number.
- **252 of 252 widths reproduced exactly**, and the suite re-derives them from
  `corpus/ground-truth/font-metrics.json`, which carries the font bytes Chromium was shown. A test
  written off `measure.ts` would have passed whatever `measure.ts` said.
- The suite never reads a font off the machine running it. Probe fonts are written to a temporary
  directory and indexed with `system: false`, so the widths under test do not depend on what
  happens to be installed.
- End to end, on `a07-text-cascade.pptx`: text drawn, baselines placed, substitution reported, and
  the 42 code points the probe fonts genuinely lack named rather than silently drawn at zero width.

## Deviations from the plan

1. **No PNG.** Named above.
2. **The font-table reader is in `@pptx-studio/cli`, not in a `fonts` package.** Phase 8.1 plans an
   SFNT reader in `@pptx-studio/fonts` for a different job — EOT, `fsType`, embedding — and _"nothing
   from a later phase gets built early"_. What is here is the minimum 3.10 needs to keep a promise
   3.10 already made, in the package that needs it. 8.1 should read it and decide what to lift.
3. **Workspace dependencies were added to `@pptx-studio/cli`** — `model`, `render-svg`, `text`. No
   external package was added and nothing was downloaded.

## Open questions

1. **The ideographic baseline is the descent.** Correct for every face without a `BASE` table `ideo`
   entry, which is every Latin face and what Chromium reports. A CJK face with one would disagree,
   and vertical text (3.9) is where it would show. No probe font carries a `BASE` table.
2. **How far the reader is from Chromium on a real face is unmeasured.** It is exact on seven fonts
   built for the purpose. Real faces bring GPOS class-based kerning, contextual features and
   `hmtx` tables with thousands of entries. The measurement to make is the reader against
   `measureText` on whatever faces a Linux runner has, in one CI job.
3. **No shaping**, as above. Bounding how wide Arabic measures would at least turn a limit into a
   number.
4. **The advance quantisation is `trunc` or `floor`, and no probe separates them.** A font with a
   negative advance would.
5. **`@pptx-studio` is unclaimed on npm**, so nothing here has actually been published. The
   workflow, the changeset and the licence files are ready; the scope is not.
