# 0043 — What 3.10 left open, and what installing it found

**Sub-phase 3.10 · accepted · supersedes open questions 1 and 4 of [0042](0042-rendering-without-a-browser.md)**

0042 published eleven packages to npm at 0.1.0 with provenance and left five open questions. This
record closes two of them, opens the experiment that will close two more, and fixes an API that
only became visibly wrong once something outside this repository called it.

The thread running through all of it is the one 0042 already named and did not finish learning:
**code that has never been run is not working code, it is code that has never been run.** 3.9 found
it in CI, 3.10 found it in the release workflow, and it was still true of the packages themselves.

## Nothing had ever installed what we published

`pnpm pkg:qa` is `publint --strict && attw --pack . --profile esm-only`. Both inspect the _shape_ of
a tarball. Neither installs one and neither executes one. `apps/studio` is `private: true` with
every dependency `workspace:^`, so it resolves through the pnpm link farm and never sees the
registry at all.

So on the day 0042 was accepted, the sum of evidence that `npm i @pptx-studio/render-svg` worked was
that the tarballs had the right file names in them.

`examples/nextjs-studio` is the thing that closes it: a Next.js application with one tool per
published package, installed **from the registry** with npm, outside the workspace.

`examples/` needed no exclusion syntax to sit outside the link farm. `pnpm-workspace.yaml` globs
`packages/*` and `apps/*`, so a third top-level directory is off the farm by default — which is the
entire reason the sample lives there rather than under `apps/`.

**No lockfile is committed**, deliberately. A lockfile would pin the sample to the resolutions of
the day it was written and stop it ever being a canary again; CI installs fresh, so every run
re-resolves `^0.1.0` against the registry.

### What the install proved

| claim                                | measurement                                                                                                                                    |
| ------------------------------------ | ---------------------------------------------------------------------------------------------------------------------------------------------- |
| the closure resolves publicly        | eleven packages install at 0.1.0                                                                                                               |
| they are tarballs, not links         | every `node_modules/@pptx-studio/*` is a real directory; CI fails on any symlink                                                               |
| `workspace:^` rewrote correctly      | `^0.1.0` across all eleven                                                                                                                     |
| `catalog:` rewrote correctly         | `fflate` at `^0.8.3`                                                                                                                           |
| the 0.5 gate holds through a tarball | **194 of 194 XML parts re-emit byte-identical** across the six shipped decks                                                                   |
| the firewall and the writer hold     | export clean past all 29 rules; `b02-layouts.pptx` out at 44 044 bytes and opened in the real PowerPoint on this machine with no repair prompt |

### `render-dom` is the twelfth package, and it is not published

0042's fifth open question — that the trusted-publisher binding lives on npmjs.com and a twelfth
package will fail its first release — stopped being hypothetical. `@pptx-studio/render-dom` is at
`0.0.0` and sits in changesets' `ignore`, so `mountSlide` and `mountOverlay` cannot be installed at
all. The sample draws through `slideNode` / `renderSlide` and `serializeSvg` instead.

Nothing was lost, and that is worth recording rather than assuming: `shapeOverlay`, the 2.11 debug
overlay, lives in `render-svg`, not `render-dom`.

## Two things a consumer hits that this repository's own suite cannot

Neither is a defect in a package. Both were invisible until something outside the tree called in.

**Measuring during render breaks server-side rendering.** `createCanvasMeasurer()` throws
`TEXT_NO_CANVAS` where there is no `OffscreenCanvas`, so a server pass produces no lines and the
first client pass produces real ones — a hydration mismatch React refuses. Measurement has to happen
in an effect. The core suite runs in a real Chromium on purpose (0001), and a real Chromium always
has a canvas, so no test here could have found this.

**An unpainted placeholder is unclickable.** A PowerPoint placeholder with neither fill nor line
emits `<path fill="none" stroke="none">`, and SVG's default `pointer-events: visiblePainted` gives
an unpainted path no pointer events at all. Text compounds it: `renderSlide` draws a shape's text in
a `<g data-text>` that is a **sibling** of its geometry rather than a child, so a click on a word
never reaches `[data-shape]`. An editor over this renderer has to hit-test both and override
`pointer-events`; the sample does, and says why.

## `renderDeck` took the verb's options, and now takes the drawing's

The sample's thumbnail route wanted SVG strings. Against 0.1.0 it had to write:

```ts
renderDeck(new Uint8Array(body), {
  slide: null,
  width: 1600,
  fontDirs: [],
  systemFonts: true,
  text: true,
  out: null,
  json: false,
  quiet: true, // CLI-shaped; meaningless here
});
```

Eight required fields, three of which `renderDeck` never reads. A field that is accepted and ignored
is a lie the compiler enforces: a caller setting `json: true` gets no JSON and no error.

`RenderDeckOptions` is the five things a drawing has an opinion about — `slide`, `width`,
`fontDirs`, `systemFonts`, `text` — every one optional, so `renderDeck(bytes)` is a whole call.
`RenderOptions` keeps its name, `extends RenderDeckOptions` with `out`, `json` and `quiet`, and is
what `runRender` takes.

| rival                                             | why not                                                                                             |
| ------------------------------------------------- | --------------------------------------------------------------------------------------------------- |
| make `out`/`json`/`quiet` optional, one type      | saves three lines of eight and leaves three properties the function ignores                         |
| `Omit<RenderOptions, 'out' \| 'json' \| 'quiet'>` | no name a caller can import; hides the problem from the signature without removing it               |
| nested `{ render: …, out, json, quiet }`          | honest, and noisier at the only place it is built; `extends` says the same thing                    |
| a twelfth package, `render-node`                  | it needs `node:fs` for font discovery, and `cli` is already the one package where `node:*` is legal |

Two things fell out that were not the point but were provably wrong:

- `RENDER_DEFAULTS` is deleted. `width: 1920` was written in three places and read in none of them by
  `renderDeck`. `DEFAULT_WIDTH` is now the only place `1920` appears, and the help text interpolates
  it.
- A width that is not a positive whole number throws `CLI_WIDTH`. Under 0.1.0 only the CLI validated
  `--width`, which was safe while the CLI was the only caller; `renderDeck(bytes, { width: 0 })`
  painting an empty `<svg width="0">` is the silent fallback the engineering standards ban.

The evidence that this is the shape rather than a preference is inside the package already:
`renderSlide`, `slideNode` and `indexFonts` all take an all-optional bag with `= {}`. `renderDeck`
was the only library entry point that did not, and the only one whose options were written by
reading `parseArgs` output.

**One mutant survives and is named rather than worked around.** `options.systemFonts !== false` →
`=== true` lives, because asserting that `systemFonts` defaults to true means asserting a scan of
this machine's font directories, which hard rule 1 forbids. It is killed by the `real-faces` job
below, on a runner where scanning is not somebody's private disk.

**The sample stays on 0.1.0.** It installs from the registry, so its two call sites keep compiling
against the published eight-field shape and will go quiet rather than red when 0.2.0 publishes. The
canary stops covering the current API on that day, which is the `carried` entry in the plan.

## The ideographic baseline came from the wrong place

0042 open question 1 said the ideographic baseline is the descent, "correct for every face without a
`BASE` table `ideo` entry, which is every Latin face", and that no probe font carried one.

Four probe fonts carry one now, and **the shipped reading was wrong**.

**The answer: the `BASE` horizontal axis' `ideo` coordinate of the `DFLT` script; where there is
none, the face box descent. 30/30, alone at the top.**

| reading                                                                      | fits      |
| ---------------------------------------------------------------------------- | --------- |
| the `BASE` `ideo` coordinate of the `DFLT` script, else the face box descent | **30/30** |
| the same, else `OS/2.usWinDescent`                                           | 28/30     |
| `DFLT` or `latn`, else the face box descent                                  | 26/30     |
| `DFLT` where it falls inside the em, else the face box descent               | 26/30     |
| the face box descent                                                         | 22/30     |
| the `ideo` coordinate of the **run's** script, else the face box descent     | 20/30     |
| the `icfb` coordinate of the `DFLT` script                                   | 18/30     |
| `hhea.descender`                                                             | 10/30     |
| zero, the alphabetic baseline                                                | 0/30      |

Three of the four findings were not anticipated:

**The script is `DFLT`, never the run's.** The obvious reading — a table that names a coordinate per
script must be consulted with the script you are drawing — scores 20/30. `base-table` carries `DFLT`
−450, `latn` −410 and `hani`/`kana` −380; measuring U+4E00 returned −450. Chromium reads `DFLT` on
the hanging baseline too (620, where `latn`'s is 630).

**There is no fall-through.** `base-no-dflt` names `latn`, `hani` and `kana` with coordinates and no
`DFLT`. Chromium ignored the table whole — ideographic −300, the box descent, and hanging 720, which
is Blink's own 0.8 × ascent fallback. Without that font, "`DFLT`", "`DFLT` else `latn`" and "`DFLT`
else the first script named" would have tied at the top and `analyse.ts` would have refused to
commit any of them.

**A coordinate outside the em is handed back as written.** `base-outsize` sets `ideo` to −1200 on a
1000 em; Chromium reports −1200 and does not clamp. So `packages/text`'s `ideographicOffset` sees
1.2, falls outside `(0, 1)` and answers 0, and the CLI reader answers 0 for the same reason —
parity measured rather than asserted. Had Chromium clamped, that guard would have been a divergence.

The no-`BASE` fallback is the **face box** descent (`usWin`, or `sTypo` under bit 7), not
`hhea.descender`: `split` gave −300 and `split-usetypo` −100, and "`DFLT` else `hhea`" scores 18/30.

Eleven probe fonts now, and the three original questions were re-scored against all of them: face
box 22/22, kerning 66/66, widths 396/396.

## `trunc` versus `floor` is not an open question

0042 said "the advance quantisation is `trunc` or `floor`, and no probe separates them. A font with
a negative advance would."

**That font cannot exist.** `hmtx.advanceWidth` is a `UFWORD` — `uint16` — against `lsb`'s `FWORD`,
and `advancesOf` reads it with `u16`. `quantiseAdvance`'s argument is `units × px / em` with
`units ≥ 0`, `px > 0` guarded and `em > 0`, so it is provably in `[0, ∞)`, where `Math.trunc` and
`Math.floor` are the same function.

The distinction is foreclosed by the format rather than unmeasured by the experiment. Two tests pin
it: one patches a fixture font's own `hmtx` bytes to `0xFFFF` and asserts 65535, which is what kills
an `i16` mutant given no probe font exceeds 32767; the other walks the fixture's whole grid
asserting every quantised argument is non-negative and that the two functions agree.

## T14 — the same reader, on faces nobody built for it

0042's second and third open questions ask how far the reader is from Chromium on a _real_ face, and
for a number rather than a caveat on Arabic. Both are one job, because a Linux runner has real faces
and some of them cover Arabic.

`tools/ground-truth/fonts/real-faces/` is the experiment. `metrics/` asks with fonts built to
disagree with themselves; `real-faces/` asks with the faces a machine has — class-based GPOS
kerning, GSUB ligatures, `hmtx` tables with thousands of entries, and a `unitsPerEm` that is not a
round 1000. Eleven samples over seven scripts at six sizes.

Latin, Cyrillic, Greek, Hebrew and CJK are gated at `1e-4` relative: a fifteenth of the 0.149 %
median by which the same browser disagrees with PowerPoint (T2), and about 0.004 px on a ten-glyph
string at 8 px, far below the ~0.5 px that moves a line break. Arabic, Devanagari and `liga` are
**recorded, not gated** — the reader sums unshaped advances and a shaper does not, and that ratio is
the answer to question 3.

Nothing is committed as a fixture. A digest over whatever fonts a runner carries would go red the
week GitHub bumps its font package; the number belongs in an ADR, dated, and the run belongs in an
artifact.

No gated sample contains an `f`. `liga` is on by default in `measureText`, and `fi`/`ff`/`fl` are
the substitutions a Latin text face applies unasked, so a gated `f` would score the reader against a
shaper it never claimed to be.

**The finding itself is not in this record.** The job has never run on Linux. What is established
is that the harness measures, that it cannot reach a system font directory — `--font-dir` is
required and there is no fallback — and that every gate goes red on command.

## Six files git could not diff

Chasing the baseline reading meant reading `measure.ts` repeatedly, and `grep` kept skipping it.
Six tracked source files carried a **raw NUL byte** rather than an escape — separators in a cache
key and two join keys, four assertions that an EOT name field is NUL-terminated, one XML string a
test expects the parser to reject, and one inside a comment.

Every use is right; the spelling was not. Git decides a blob is binary by looking for a NUL in its
first 8000 bytes, so three of the six had no diff, no blame and no merge, and `grep` and
`Select-String` silently skip all six. For a repository whose comment rule bans history on the
grounds that **the diff and git history are the record**, a source file the diff cannot read is a
hole in the record.

They are `\u0000` now, and `\0` in the comment. The strings are identical bytes, which the suite
proves rather than the change asserting it: `fixtures.test.ts` passes only if the escape produces a
real NUL, and `review.test.ts` still gets `ERR_MALFORMED_XML` from an XML declaration with one in
it.

## Verification

- **`pnpm check` green**, all of it.
- **T13: 11 probe fonts, 514 observations, every rival scored.** Chromium 151.0.7922.34 accepted all
  eleven; `analyse.ts` refuses a fixture where two readings tie, and one mutant proved that by
  making it refuse — writing the `BASE` coordinate negated dropped the winner to 22/30 and no
  fixture was written.
- **Mutation: 9 tried on the reader, 9 killed; 8 tried on the T14 gate, 8 killed after one genuine
  survivor.** That survivor was a missing assertion, not an equivalent mutant: nothing pinned the
  case where a rival _ties_ the shipped reading, which is exactly what "unique top scorer" forbids.
- **The T14 gate has been seen red.** A green artifact was doctored by one part in a thousand on a
  single reading and `analyse.ts` exited 1 naming the sample, the delta in px, in font units and
  relative, and the threshold it passed.
- **The API fix was typechecked from outside.** A throwaway consumer compiled against the built
  `dist/index.d.ts` under `strict` + `exactOptionalPropertyTypes`; the 0.1.0 eight-field literal now
  fails with `TS2353` naming `out`.
- Guards on the baseline probes held: `alphabeticBaseline` was 0 on all 30 rows, and every Han row's
  width was exactly `advance × px / upem`, so no installed face answered for a probe.

## Deviations

1. **`packages/cli/src/render/errors.ts` gained `CLI_WIDTH`.** `RenderError` will not accept a code
   outside its union, so no other file can declare it.
2. **`score.ts` is a seventh experiment file role**, for an experiment that is also a CI gate: the
   verdict is a pure function with its own suite, because a gate whose scoring is inline is a gate
   nobody can test.
3. **T14 skips font collections.** `.ttc` / `.otc` and any file yielding more than one face are
   counted as skipped. The reader itself handles `ttcf` (0042); the job does not, so on an image
   where Noto CJK ships as a `.ttc` the `cjk` sample will have no covered face.
4. **The changeset is `minor`, not `major`.** The change is breaking, and changesets reads `major`
   literally: it would publish 1.0.0.

## Open questions

1. **The T14 numbers themselves.** Face count, the per-script table and the shaping ratios are all
   unmeasured until the job's first run on `ubuntu-latest`. A red A3 most likely means the reader
   skips a type-9 Extension GPOS lookup — `sfnt.ts` has `if (u16(r, lookup) !== 2) continue;`, so a
   face wrapping its `kern` lookups that way kerns as zero. A red A5 means the face box differs on
   Linux. Both are findings, and both are fixed in the reader, never in the threshold.
2. **The face box readings are not separated by synthetic fonts.** Every probe has `unitsPerEm` 1000
   and integer metrics, so exact, round, floor and ceil all tie at 4/4 and the honest answer is
   `null`. Only a 2048-em face — `1901 × 1000 / 2048 = 928.22265625` — separates them, and only a
   real font directory has one.
3. **A CJK face with a `BASE` table is still unmeasured against PowerPoint.** T13 measured Chromium.
   What PowerPoint puts on the ideographic baseline is a different question, and 3.6's vertical text
   is where a disagreement would show.
4. **The trusted-publisher binding**, unchanged from 0042: it is a setting on npmjs.com, nothing in
   this tree asserts it, and `render-dom` will fail its first release until someone adds it there by
   hand.
5. **The sample pins `^0.1.0`.** On the day 0.2.0 publishes it stops exercising the current API and
   starts exercising the previous one, silently. Bumping it is a deliberate act, and it is carried in
   the plan rather than left to be noticed.
