# 0014 — The CI round-trip gate, and a badge that cannot lie

Date: 2026-08-29
Status: **accepted** — 51/51, committed as a Shields endpoint and re-measured on every run
Sub-phase: 1.6 — CI gate

---

## Context

The plan gives 1.6 one line:

> **1.6 — CI gate.** Round-trip job on every PR; badge. _Verify:_ badge reads 50/50.

By the time this was reached the round trip was already gated three times over:
`tools/corpus/roundtrip.test.ts` asserts 51/51 from source, `pnpm test` runs it,
and `pnpm check` runs `pnpm test`. So the sub-phase adds no coverage. What it
adds is a **number in public**, and that turns out to be a different problem
with its own failure mode: a badge is a claim made to people who will never run
the tests.

---

## Decisions

### The number is committed, not published from a run

The usual way to get a live number onto a README is a workflow that writes to a
gist or pushes a commit. That needs a token, a bot identity, and write
permission on a workflow that otherwise needs none — all to publish a fact that
is **already determined by the contents of the repository**. The corpus is
committed. The writer is committed. The round trip is deterministic. Nothing
about a CI run decides the answer.

So the answer is committed too, as a Shields endpoint document at
`.github/badges/roundtrip.json`, and `pnpm roundtrip` recomputes it and fails
when the two disagree. The badge cannot claim 51/51 unless the last run to touch
it measured 51/51, and `permissions: contents: read` stays true across the whole
workflow.

The failure mode this closes is the ordinary one: somebody removes a deck, or
adds five, and the README quietly keeps saying what it said last year. Editing
the file to say something false fails the gate; changing the corpus without
regenerating it fails the gate. The only way to move the number is to move the
corpus and then say so.

### It goes red, never amber

Fifty of fifty-one is not 98% working. It means some package does not survive
being read and written, which is the one thing this architecture is built to
guarantee. A graded colour would invite reading the badge as a score.

### A job of its own, in front of the Chromium install

`check` already runs `pnpm roundtrip`, so the new job gates nothing that was not
gated. It is still worth its minute: it is the number the README carries, and
behind a `playwright install --with-deps chromium` the one result anybody wants
from a pull request arrives last. This is the same reasoning that already puts
`layering` and `corpus` first _inside_ `check`, applied one level up.

It builds four packages rather than seven — `--filter=@pptx-studio/writer...` —
and never touches a browser. The build is needed at all only because nothing
links `@pptx-studio/*` into `tools/`, so the script loads the writer from
`dist`, exactly as `write-census-keys.ts` does and for the same reason.

### Two routes to the same number, on purpose

`pnpm roundtrip` reaches the round trip through `dist`. `tools/corpus/badge.test.ts`
reaches it through Vitest's alias to **source**, and asserts the committed badge
against what that produces. A build that does not match the sources it came from
therefore shows up as two different numbers rather than as one confident wrong
one.

### The badge says what it measures, and stops there

It says the decks round-trip. It does **not** say they open in PowerPoint, and
the distinction is the whole of Gate 1. No hosted runner can ask: GitHub's
Windows images ship Visual Studio's Office _development_ workload — `Workload.Office`,
`Component.TeamOffice` — and not one Office application. There is nothing on a
runner that can open a `.pptx`.

That question belongs to `bisect --oracle powerpoint` from 1.5, on a machine
with PowerPoint, run by hand. The narrow claim is stated three times — in the
job summary, in the README section the badge links to, and in the report module
— because a badge that gets read as "verified against PowerPoint" would be the
most misleading artefact in the repository.

### What the gate checks about itself

Three quiet failures get assertions, because none of them fails anything on its
own:

- the badge going stale — the committed file against a fresh measurement;
- the job being reorganised away — `ci.yml` is parsed and asserted to have a
  `roundtrip` job that runs `pnpm roundtrip` on `pull_request`;
- the gate drifting out of `pnpm check` — the script order is asserted too,
  including that `build` comes before `roundtrip`, since the script needs `dist`.

---

## What building it found

### The workflow was one flag short of leaking its token

`actions/checkout` writes the `GITHUB_TOKEN` into `.git/config` unless
`persist-credentials: false` is set, where every later step — and everything
those steps install — can read it. This workflow installs a dependency tree and
runs it. Nothing here pushes, so the credential had no reason to exist past the
checkout. Both jobs set it now, and a test asserts every `actions/checkout` in
the file does, alongside the SHA pinning the workflow already had.

### `corpus/` could not hold the badge

The obvious home was `corpus/roundtrip.json`. `C008-orphan` fires on any file
under `corpus/` that no manifest claims — the rule that exists so an unlicensed
deck cannot be dropped in unnoticed — and it was right to fire. The badge is not
corpus data; it lives in `.github/badges/`.

### The deck list comes from the manifests, not the directory

`readdirSync` would count a file no manifest claims, which is precisely a deck
whose licence `pnpm corpus` has never checked. Putting that into the number on
the front page is the one way this badge could become a legal problem rather
than a factual one, so the count reads the same manifests `roundtrip.test.ts`
does.

---

## Consequences

`pnpm check` green. The gate is one command, locally and in CI:

```
roundtrip: 51/51 deck(s), 1419 part(s) (840 xml, 536 rels, 43 binary)
```

- **The badge reads `51/51 decks`**, which is the plan's verification with the
  corpus at the size it actually reached.
- **The per-collection table lands in the pull request's job summary** — 41
  `decks`, 9 `authored`, 1 `written` — so a reviewer sees which collection moved
  without opening a log.
- **1419 parts, 840 / 536 / 43**, identical to what 1.4 pinned, reached through
  `dist` rather than through source.
- **No workflow in this repository has write access to anything.**

### What is not done

- **Gate 1 is not claimed.** It asks for a live demo that opens a deck with
  charts, SmartArt, animations, OLE and macros, re-saves it in the browser and
  downloads a file that opens clean in PowerPoint. Every piece of the export
  path is proven; the page that drives it is not built.
- **Nothing in CI opens PowerPoint**, and nothing in CI can. The strict
  `OpenAndRepair` sweep from 1.5 was run by hand on this machine and is not
  repeated on any push.
- **The badge URLs resolve once the repository is public.** They point at
  `raw.githubusercontent.com` on `main`, which is the point — the badge is
  served from the committed file, not from a run.
