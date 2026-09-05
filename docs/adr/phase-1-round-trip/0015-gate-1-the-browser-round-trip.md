# 0015 — Gate 1: a deck out of a browser, and into PowerPoint

Date: 2026-09-02
Status: **accepted** — a real download from a real page opens in real PowerPoint with repair off
Gate: 1 — the preserving package, demonstrated end to end

---

## Context

The gate is one sentence:

> a live demo that opens a deck with charts + SmartArt + animations + OLE +
> macros, re-saves it, and the downloaded file opens in real PowerPoint with
> **zero repair prompt and zero visible difference**.

Sub-phases 1.1 to 1.6 built every piece of that. What none of them built is the
sentence itself, and the difference is not decoration. Three things in it are
outside everything the writer's own suite can reach:

- **A deck with all five at once.** Tier A isolates features on purpose, so the
  corpus had `a21-charts`, `a23-smartart`, `a17-animations`, `a26-ole` and
  `a32-macros` and nothing holding more than one of them.
- **A browser.** `packages/writer` has run its suite in real Chromium since 0.1,
  so "the writer works in a browser" was never the open question. The open
  question was the _path_: a `File` a user chose, a Worker, a `Blob`, an object
  URL and a download.
- **"Zero repair prompt."** Which is a question only one program on this machine
  can answer, and only when asked in the one way that makes a repair an error.

---

## Decisions

### A kitchen-sink deck, and it breaks the tier's rule on purpose

`a43-kitchen-sink.pptm`: ten slides, four chart parts, three diagram data
models, three OLE objects, three `p:timing` trees and a VBA project, in one
macro-enabled package.

Every other Tier A probe is about one thing so that a failure names it, and that
remains right — `cli bisect` exists because narrowing after the fact is
expensive. But the gate asks a different question from the five decks it would
otherwise be answered by, and the difference is real. Four failure modes live
only in the combination:

- `[Content_Types].xml` now carries four `Default` entries — `xlsx`, `emf`,
  `vml`, `bin` — from three feature families, so the part a single missing
  `Default` corrupts is shared. That is the canonical repair prompt (pandoc
  #11492), and no single-feature deck can produce a package where two families
  compete for it.
- The main part carries the `macroEnabled` content type **while** four chart
  parts carry theirs. The two are independent and nothing else in the corpus
  says so.
- The `p:timing` trees target `p:graphicFrame` shapes rather than the `p:sp`
  shapes `a17` targets. A `p:spTgt/@spid` naming a frame is exactly the case
  where renumbering `p:cNvPr/@id` on export silently unhooks an animation, and
  nothing had one before.
- `ppt/vbaProject.bin` and `ppt/embeddings/*.xlsx` are typed by two `Default`s
  that a package holding only one of them never has to reconcile.

**It is built by merging, not by re-authoring.** The chart, diagram and OLE
parts come out of `a21`, `a23` and `a26`'s own `build()` output and the VBA
project out of `a32`'s, so the deck cannot drift from the decks it is made of,
and markup that already passes an independent census keeps passing it. The
merge asserts that the four part lists do not collide rather than believing it —
`ppt/charts/colors1.xml` against `ppt/diagrams/colors1.xml` is the near miss.

Every slide of all three source decks comes across, and that is a constraint
rather than a choice: a probe deck's parts are exactly the parts its slides
reach, so a subset of the slides would leave chart parts unreachable and
`probes.test.ts` asserts the reachability graph has no orphans.

Relationship ids needed no reconciliation, which is the thing everyone expects
to be the problem. An rId is an `xsd:ID` scoped to one `.rels` part, so `rId2`
on the chart slide and `rId2` on the OLE slide are unrelated names in unrelated
files. `p:cNvPr/@id` would have been a real problem, and is not one either, for
the same reason: slides arrive whole and nothing is merged onto them.

### The demo saves two ways, and the second one is the point

Gate 1's wording is satisfied by a no-op, and satisfying it by construction is
the whole architecture. But a page that can only emit the bytes it read never
runs the half of the writer that matters — the half where one part is serialised
afresh and every other part is streamed out of the source archive still
compressed. `rewritten: 1, streamed: 54` is the claim, and a demo that never
rewrites anything cannot make it.

So there is a second button, and it makes the smallest edit in the package: the
text of `cp:lastModifiedBy` in `docProps/core.xml`. Nothing on any slide moves,
so "zero visible difference" stays true; PowerPoint shows the value in File →
Info, so a human can check it; and it goes through `applyEdit`, which is the
only sanctioned way to change a tree here and hands back its own inverse.

It is **not** a preview of Phase 5. There is no command, no history and no
model — `packages/model` does not exist. One XML edit, written where the demo
needs it and nowhere else.

### The page keeps the way back to the bytes, not the bytes

The deck is _transferred_ into the Worker, which detaches it, so after an
inspection the page no longer holds the archive — by design, since a second copy
of a 200 MB file on the main thread is the exact cost the transfer avoids.

A `File` can be read again and a URL can be fetched again. What is kept is
therefore a `read()` closure, not a buffer, and the worker holds nothing between
requests. The alternative — keeping the bytes alive in the worker on the chance
somebody presses Save — would make every inspection cost the memory of an export
that may never happen.

### Export is a second request type, not a flag on the parse

They are separate acts: a page can inspect a deck ten times and never export it.
More concretely, exporting needs the bytes a second time, and the first set was
detached by the transfer that delivered it.

### The harness downloads the file rather than reading it out of the page

`pnpm gate1` drives the real page in the Chromium `pnpm test` already uses,
presses Save through the same automation hook `tools/bench` uses for the
benchmark, and captures Playwright's `download` event. It would have been easier
to have `page.evaluate` return an array of bytes, and it would have tested less:
a stale `URL.revokeObjectURL` cancels the download, a missing `download`
attribute navigates instead of saving, and a `Uint8Array` view onto a larger
buffer writes the wrong bytes. None of those three fails a unit test of the
export, and all three are in the path between `exportDeck` and a file on disk.

### `OpenAndRepair` off, and the harness quits the PowerPoint it started

1.5's finding, restated because it is the finding this gate turns on: the flag
defaults to **on**, and under automation the repair is silent. A file PowerPoint
would prompt about opens successfully, already repaired, and reports success.
With it off the repair is a catchable error. There is deliberately no
`--allow-repair` on `pnpm gate1`: "can PowerPoint make something of it" is not
the question, and a flag that softened this is a flag somebody reaches for on a
bad day.

The oracle attaches to a PowerPoint that is already running and never quits that
one, because it is the user's. It does start one when there is none, and the
harness now quits that, because an invisible instance holding a file handle is
what somebody hits an hour later and cannot explain.

### The browser half is in CI; the PowerPoint half cannot be

`apps/studio/src/export.test.ts` runs in Chromium on every `pnpm test`, as a new
`apps` project beside `core`, `cli` and `tools`. It fetches the committed corpus
over Vitest's own dev server — which is the one thing a package test cannot do,
and `packages/validate/src/testing/deck.ts` already says why: browser mode has
no filesystem, so the writer's suite works on synthetic packages.

The PowerPoint half stays a command run by hand, for the reason ADR 0014
established: no hosted runner has Office on it.

---

## What building it found

### The demo's own file is not byte-identical to the corpus deck, and should not be

Same length, 112 differing bytes, every one of them at offset +4 of a header:
the ZIP **version needed to extract**. `tools/ground-truth/lib/zip.ts` writes 2.0 on
every entry; `packages/opc`'s writer writes 1.0 for a stored entry, which is
what APPNOTE specifies — 2.0 is the floor for DEFLATE.

Both are correct and every reader accepts both. It is a new member of the class
ADR 0012 named when it refused to make the round-trip gate a byte comparison
(entry order, deflate level, DOS timestamps, attribute order), found this time
on a deck we generated ourselves with two of our own writers. A byte-equality
gate would have been red here, at Gate 1, on a file PowerPoint opens without
complaint.

### A comment in the round-trip suite had never added up

The breakdown beside the pinned totals read "the 43 is wider than the 24 parts
under `/ppt/media/` ... five `.fntdata`, four embedded workbooks, eleven
`docProps/thumbnail.jpeg` and one `vbaProject.bin`" — which is 45, against a
total of 43. Nothing would ever have said so, because the breakdown was prose
beside three numbers rather than a fourth number. Recounted: 23 media, 5 fonts,
5 workbooks, 11 thumbnails, 2 VBA projects, 46 in all.

### `macros` came off the single-probe list

`C-COV` tracks which census keys rest on one deck, and `a43` is the corpus's
second `vbaProject.bin`. Fourteen became thirteen. That is most of what a second
probe for a feature is worth, and it is why the coverage test asserts the list
rather than its length.

---

## Consequences

`pnpm check` green. `pnpm gate1` on 2026-09-02, PowerPoint 365 build
16.0.20326:

```
gate1: a43-kitchen-sink.pptm
  PASS  none    55 parts, 0 rewritten, 55 streamed, 29 rules / 0 blocking
        reopened clean; PowerPoint: opened, repair false, 10 slides
        differs in: nothing
  PASS  stamp   55 parts, 1 rewritten, 54 streamed, 29 rules / 0 blocking
        reopened clean; PowerPoint: opened, repair false, 10 slides
        differs in: /docProps/core.xml
```

- **Zero parts rewritten on a no-op export**, through a browser, on a package
  holding charts, SmartArt, animations, OLE and macros.
- **PowerPoint reports the same ten slides and the same shape counts** —
  `[7, 2, 2, 3, 2, 2, 2, 3, 3, 3]` — for the downloaded file as for the
  original. That is "zero visible difference" measured rather than eyeballed,
  as far as an automation interface can measure it.
- **The corpus is 52 decks, 1474 parts**: 875 XML, 553 relationship graphs, 46
  by SHA-256. The badge says `52/52`.
- **`pnpm test` is 57 files, 1451 tests**, including a new `apps` project — the
  export path asserted in Chromium against the committed corpus on every run.

### What is not done

- **Nothing renders.** The page shows a deck's internals and hands it back; a
  slide is still not drawn anywhere. Phase 2 starts there.
- **"Zero visible difference" is measured through COM**, which reports slide and
  shape counts and not pixels. A pixel comparison needs the fidelity harness,
  which is 3.9. Two files that PowerPoint opens with the same structure could in
  principle still differ visually; nothing here would catch it.
- **One deck.** The gate names five features and one deck now has all five, but
  `pnpm gate1` has been run against `a43` and not against the other fifty-one.
  `pnpm roundtrip` covers all of them, without PowerPoint.
- **The stamp is the only edit.** There is no command layer, no history, and
  nothing that touches a slide. That is Phase 5, and this is deliberately not a
  down payment on it.
- **No user deck has been through this.** Every file in the corpus was written
  by this repository or by PowerPoint on this machine, on request. The first
  real-world deck will find something.
