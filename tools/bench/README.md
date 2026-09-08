# bench

Sub-phase 0.8's verification: **parse a 200 MB / 300-slide deck in a browser tab
and record the timing.**

There was no such deck to hand, and `CLAUDE.md` says not to go looking for one on
the machine this runs on. So it is generated, from a recipe,
deterministically — the same recipe writes the same bytes, which is what lets
[`corpus/bench/manifest.json`](../../corpus/bench/manifest.json) pin a SHA-256
against a file far too large to commit.

## Run it

```bash
pnpm build                                          # the page bundles dist/

node tools/bench/make-deck.ts media-200mb  <dir>/media-200mb.pptx
node tools/bench/make-deck.ts xml-heavy    <dir>/xml-heavy.pptx
node tools/bench/make-deck.ts small        <dir>/small.pptx

node tools/bench/run-bench.ts --decks <dir> --repeat 5 --out corpus/bench/results.json
```

Add `--screenshot page.png` for a picture of the rendered explorer. The benchmark
exercises the rendering path but never looks at it, and a picture is the only
thing that catches a report which computed correctly and drew nothing.

Building the 188 MiB deck takes about four seconds. The benchmark launches the
Chromium that `pnpm test` already uses; nothing is installed.

To poke at a deck by hand instead:

```bash
node tools/bench/serve.ts --decks <dir>     # http://127.0.0.1:5173
```

…then drop a `.pptx` on the page. `--headed` on `run-bench.ts` does the same
thing while the benchmark drives it.

## The files

| file            | what it is                                                |
| --------------- | --------------------------------------------------------- |
| `deck.ts`       | the recipes and every byte of markup in them              |
| `make-deck.ts`  | CLI: recipe in, `.pptx` and a SHA-256 out                 |
| `zip-stream.ts` | a ZIP32 writer that streams to a file descriptor          |
| `png.ts`        | valid PNGs of a chosen size, made of incompressible noise |
| `serve.ts`      | a cross-origin-isolated static server for `apps/studio`   |
| `run-bench.ts`  | drives the real page in Chromium and prints the numbers   |
| `bench.test.ts` | the generator, and the census checked against it          |

## Three things worth knowing before trusting a number

**The media must be incompressible.** A deck padded with zeroes deflates to
nothing, and the ZIP reader then never does the work a real 200 MB file makes it
do. Every image here is noise, stored rather than deflated — DEFLATE over 190 MB
of noise costs minutes and makes the output very slightly larger.

**Two deck shapes, because one number is not enough.** 184 of `media-200mb`'s
188 MiB never reach the XML layer at all, so it measures the ZIP reader and says
nothing about the tokenizer. `xml-heavy` is 4.4 MiB on disk and 87 MiB inflated,
and it is the one whose throughput predicts anything about phase 2 onwards.

**Measure the stages separately, and measure them warm.** "Parse took N seconds"
adds together four things with wildly different costs, only one of which is ours
to improve: inflating is fflate's, decoding UTF-8 is the platform's, tokenizing
is this project's, counting is the census's. `run-bench.ts` reports all four,
and the stage pass runs _after_ the census runs so nothing is charged for JIT
warm-up. A cold first pass makes the tokenizer look like the bottleneck whether
or not it is, which is exactly the mistake that sends someone to optimise the
wrong loop. It happened here: the first version of the census scanner ran at
10 MiB/s because it built a `uri + ' ' + local` key for every element in the
package, and the number looked like a slow tokenizer.

## Opening a generated deck in PowerPoint

Not part of `pnpm check` — a test that needs a licensed application cannot gate
a pull request. Run by hand, the same way sub-phase 0.7 did:

```powershell
$app = New-Object -ComObject PowerPoint.Application
$app.DisplayAlerts = 1                 # ppAlertsNone; a repair prompt is modal
$pres = $app.Presentations.Open($path, -1, 0, 0)   # ReadOnly, Untitled, no window
$pres.Slides.Count
$pres.Close(); $app.Quit()
```

`$app.DisplayAlerts = 1` is not optional. A repair prompt is a modal dialog, and
a modal dialog blocks the COM call rather than returning an error.

Both `small` and `media-200mb` open in PowerPoint 365 with no repair prompt.
That is worth more than it looks: every byte of those decks was written by
`deck.ts`, so it is early evidence for the writer in sub-phase 1.3.
