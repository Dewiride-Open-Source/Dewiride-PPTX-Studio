# 0008 — `inspect`, the Worker boundary, and what a 200 MB deck actually costs

Date: 2026-08-27
Status: accepted
Sub-phase: 0.8 — `cli inspect` + Worker shell + benchmark. Closes Gate 0.

---

## Context

The plan's 0.8 is three things at once: a Node CLI that describes a deck, a
browser page with the parse Worker boundary in it, and a number — "parse a
200 MB / 300-slide deck in a browser tab and record the timing".

The number is the part that could not be produced. There is no 200 MB deck on
this machine, and `CLAUDE.md` forbids going to look for one. So the deck is
generated, and most of the work in this sub-phase went into making a generated
deck a thing a measurement can be trusted against.

---

## Decisions

### The census is a package, not a feature of the CLI

Both consumers need the same answer, and neither can own it. `pptx-studio
inspect` is Node and could never run in a tab; the browser explorer runs in a
Web Worker and could never import Node.

So `@pptx-studio/census` — layer 1, browser runtime, depending on `opc` and
`xml` and nothing else. The CLI adapts. This is a package the plan's monorepo
layout does not name, and it is the only structural deviation in this sub-phase.
Two things argue for it beyond today's need: sub-phase 1.1's `corpus/manifest.json`
requires a `features[]` array on every entry, and something has to compute that
without a human curating it; and the same object is what an "unsupported
content" panel in phase 3 will read.

### A census is plain JSON, and no tree ever leaves the Worker

`XNode` carries a `parent` pointer on every node and holds the entire decoded
source alive through every node's byte offsets. Structured-cloning one across
`postMessage` would not be sending a message, it would be sending a second copy
of the deck. The rule is enforced by a test that walks the whole result object
looking for `parent`, `children`, `qname` or `source`.

The result is also restricted to the intersection of structured clone and
`JSON.stringify`: no `Map`, no `Set`, no `undefined`, absence spelled `null`.
Structured clone would carry a `Map` and `JSON.stringify` would silently drop
it, and the two consumers would then disagree about what a census is.

### Scan over tokens, not over trees

A census of a 200 MB deck is a few kilobytes of counters. Building a tree to
produce them would make the working set grow with the deck for no benefit:
counting how many `a:tbl` a package holds does not need any node to know its
children.

What that costs is real and worth stating: **a token stream is not a tree, so
nothing objects to `<a><b></a>` or to a part that simply stops.** The first
version of the scanner reported a truncated slide as perfectly healthy, which is
the wrong answer from a tool whose entire job is describing a file that is
already suspect. Balance is now checked against a stack of open names, and an
unbalanced part becomes a `PART_NOT_WELL_FORMED` problem.

### Features are keyed to namespaces, and to the phase that renders them

Every lookup resolves the prefix in the scope the element was written in, so
`zz:tbl` is a table when `zz` is bound to DrawingML and `a:tbl` is not one when
`a` is bound to something else. `mc:Ignorable` and `mc:Choice/@Requires` are
resolved the same way, because they hold prefixes and the same prefix resolves
to two different URIs in different parts of one corpus.

Each feature carries the sub-phase that implements it. That is what makes the
output a schedule instead of a disclaimer — dropping a deck on the page and
reading back "SmartArt 7 (phase 4.5), ChartEx 0, ink 0" says exactly what stands
between that file and a faithful render.

### `apps/studio` is framework-free, and Next.js is deferred

The plan puts a Next.js shell here. It is not built, for two reasons and one
constraint.

The reasons: `packages/react` does not exist until sub-phase 5.7, so a React
shell today would wrap nothing; and the thing 0.8 has to establish is the Worker
boundary, which is more honest in plain DOM — nothing underneath it can quietly
start depending on a framework, which is the contract every package in this
repository is held to anyway.

The constraint: adding Next, React and their types is roughly three hundred
megabytes of new registry downloads and three new catalog entries, and
`CLAUDE.md` rule 2 says not to install anything without being asked. Everything
new in this sub-phase resolves offline from packages already in the store.

The one line that changes when the shell does grow a framework is called out in
`apps/studio/src/client.ts`:

```ts
new Worker(new URL('./worker.js', import.meta.url), { type: 'module' });
```

That expression is the canonical form — Vite, webpack 5 and Turbopack each
recognise it and rewrite it to their own emitted chunk, and a plain bundler
leaves it alone to resolve against the module's own URL at runtime, which is why
it works here with no plugin at all.

---

## What was measured

Chromium 151.0.7922.34, headless, on this machine. Cross-origin isolated, so the
page is served with `Cross-Origin-Opener-Policy: same-origin` and
`Cross-Origin-Embedder-Policy: require-corp`. Best of five runs, in a Web
Worker, driving the same code path a dropped file goes through.

| deck          | archive   | entries | parts | slides | inflated  | best        | median  |
| ------------- | --------- | ------- | ----- | ------ | --------- | ----------- | ------- |
| `media-200mb` | 188.5 MiB | 1616    | 995   | 300    | 192.0 MiB | **0.347 s** | 0.504 s |
| `xml-heavy`   | 4.4 MiB   | 1316    | 695   | 300    | 87.4 MiB  | 4.808 s     | 5.138 s |
| `small`       | 2.0 MiB   | 145     | 93    | 20     | 2.2 MiB   | 0.037 s     | 0.040 s |

**The plan's envelope is comfortably met.** A 188.5 MiB, 300-slide deck is fully
read — archive opened, every relationship walked, every XML part inflated,
tokenized and counted — in about a third of a second in a browser tab.

### Where the time goes

Measured after the census runs, so nothing is charged for JIT warm-up:

| stage        | `media-200mb` | `xml-heavy` |
| ------------ | ------------- | ----------- |
| inflate      | 74 MiB/s      | 108 MiB/s   |
| decode UTF-8 | 806 MiB/s     | 640 MiB/s   |
| tokenize     | 43 MiB/s      | 38 MiB/s    |

The tokenizer is the floor, at **349 ns per token** on the media deck and 384 ns
on the XML-heavy one. That is not a character-scanning cost — it is allocation:
`xml-heavy` produces 5.94 million tokens, each an object with an attributes
array of further objects, and V8 spends the time making and collecting them. If phase 12 ever needs the parser to
go faster, the lever is a token representation that allocates less, not a
cleverer scanning loop.

### Two decks, because one number would have been misleading

184 of `media-200mb`'s 188 MiB never reach the XML layer at all. It is a
measurement of the ZIP reader and says nothing about the tokenizer. `xml-heavy`
is the inverse — 4.4 MiB on disk, 87 MiB inflated, a 20:1 ratio across the whole
archive — and it is the deck whose throughput predicts anything about phase 2
onwards. It also sits usefully close to the decompression budget from sub-phase
0.2: 87 MiB against a 1 GiB ceiling, so a deck an order of magnitude denser than
this one would be refused rather than swallowed.

### Transferring the deck costs nothing, and proving that took a second attempt

The archive is transferred, not structure-cloned. Handing 188.5 MiB to the
worker takes **1.0 ms** — less than the 3.9 ms the 2 MiB deck happened to take
on the same run, which is the shape of a number that does not depend on file
size at all.

The first attempt to measure this reported 197 ms and would have been quietly
wrong. It timed from `postMessage` to the worker's _first response_, and the
first response was a progress message — so the number included opening the
archive and walking six hundred relationship parts, and reported the lot as the
cost of `postMessage`. The worker now posts an `accepted` message before doing
any work at all, which is the only way the claim is falsifiable.

---

## Findings

### `DOMParser` is not available in a Web Worker — measured, not cited

The second architectural bet rests partly on this, and it is now checked on
every run, in the environment the claim is about, and shown on the page:

| in the worker         | Chromium 151 |
| --------------------- | ------------ |
| `DOMParser`           | **absent**   |
| `XMLHttpRequest`      | present      |
| `OffscreenCanvas`     | present      |
| `TextDecoderStream`   | present      |
| `crossOriginIsolated` | true         |

`OffscreenCanvas` being present matters for sub-phase 3.2, which puts text
measurement in a Worker and has no fallback if it is not there.

### `measureUserAgentSpecificMemory()` is exposed to the document but not to the Worker

On a cross-origin-isolated page in Chromium 151,
`performance.measureUserAgentSpecificMemory` is defined on the main thread and
**undefined inside a dedicated worker**, although the specification lists worker
scopes among its contexts.

This matters because it is the only measurement that includes `ArrayBuffer`s:
`performance.memory` counts the JS heap and would report a 188 MiB archive as
costing nothing at all. Phase 12.2's image-cache budget therefore has to be
measured from the document, or by a different technique entirely. Recorded here
so it is not rediscovered.

The probe reports _why_ it has no number rather than returning a bare `null` —
"not defined here" and "SecurityError, the context is not isolated" are different
problems with different fixes, and a silent `null` is indistinguishable from
having forgotten to ask.

### PowerPoint opens a 300-slide deck this project wrote, with no repair prompt

Every byte of `media-200mb.pptx` was written by `tools/bench/deck.ts`. Opened
through COM in PowerPoint 365 (build 16.0.20326):

```
opened   true      slides 300     shapes 5191    sections 6
fonts    ProbeAlpha embeddable=msoTrue embedded=msoTrue
```

Charts, SmartArt, tables, notes slides, sections, custom shows, comments,
`mc:AlternateContent`, animation timelines and an embedded font, and no repair
prompt on any of it. That is early evidence for the writer in sub-phase 1.3, and
worth more than the same result on a twenty-slide file.

**5191 is also exactly the number of shape ids the generator allocated**, and
274 against 274 on the twenty-slide deck. Two counts of the same thing — one
from the code that wrote the file, one from the application that read it — and
they agree on both.

`Font.Embedded` reporting `msoTrue` is not evidence the font data loaded — sub-phase
0.7 established that it means only that a `p:embeddedFont` entry matched the
typeface. Nothing here contradicts 0.7's conclusion that PowerPoint will not
read an uncompressed EOT.

### The census found a real defect in the generator on its first run

`ppt/commentAuthors.xml` was written and given a content type but never related
from `ppt/presentation.xml`, so it was an orphan part. The census reported one
unreachable part and named it. That is the tool doing its job on the first file
it was ever pointed at, and it is why the generator's own test now asserts that
no relationship dangles and no part is unreachable.

### The first scanner was three times slower than it needed to be, and it looked like the tokenizer

The obvious shape for a namespace scope is `Map<prefix, uri>`. Per element that
meant hashing a sixty-character URI to reach the namespace's counters and
building a `uri + ' ' + local` key to look up the feature table — one string
allocation for every element in the package. On `xml-heavy`, with 3.6 million
elements, the census ran at 10 MiB/s and the tokenizer looked like the problem.

Binding a prefix to an object that already holds the counters and its slice of
the feature table took it to 22 MiB/s against a 46 MiB/s tokenizer, which is the
right ratio: the census now costs about half again what tokenizing costs, and
the remaining floor is genuinely the tokenizer's.

The general lesson is the one the harness now enforces: measure the stages
separately, and measure them warm, or the first cold pass sends you to optimise
the wrong loop.

---

## Verification

- `pnpm check` green: layering, format, lint, typecheck, build, publint/attw, tests.
- `packages/census` — 25 tests in real Chromium, including a JSON/structured-clone
  round trip and a sweep that fails if any `XNode` field appears in the result.
- `packages/cli` — 12 tests in Node. `main` takes its streams and returns an
  exit code instead of writing to `process.stdout` and calling `process.exit`,
  so the whole command line is exercised in-process.
- `tools/bench` — 15 tests, of which five check the deck generator and the census
  against each other. The generator counts what it put in; the census counts what
  it finds; the two are written independently, so making them agree is a real
  check on both. A census verified only against its own fixtures verifies nothing.
- The 200 MB deck opened in PowerPoint 365 by hand, through COM.

`vitest` grew a third project, `cli`, running in Node. The core project runs in
real Chromium and now excludes `packages/cli` explicitly: the CLI is the one
package that is Node by design, and its first `node:fs` import throwing under
Chromium is the browser project working as intended rather than a reason to
loosen it.

---

## What is deferred

- **The Next.js shell.** Needs an install decision. Nothing about the Worker
  boundary changes when it arrives; `client.ts` names the one line that does.
- **A feature-census entry for threaded comments** (`p188:`) and for the modern
  comment parts. The classic `p:cmLst` form is detected; the 2018 form is not,
  because there is no deck here that has one.
- **`corpus/bench/results.json` is a record, not a fixture.** Nothing asserts on
  the timings, because a slower laptop is not a regression. When sub-phase 3.9
  builds the fidelity harness with its digest-pinned container, these numbers
  can move into something CI compares.
- **The census reads content types, not magic bytes.** A part whose `Override`
  lies about its type is scanned as whatever the content type claims. That is
  the right default for a describing tool and the wrong one for a validator,
  which is sub-phase 1.2's problem.
