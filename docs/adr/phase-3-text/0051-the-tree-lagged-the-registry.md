# 0051 — The tree lagged the registry

**Sub-phase 3.10 · accepted · closes open question 3 of
[0046](0046-nothing-gated-the-thing-being-published.md), open question 1 of
[0048](0048-every-package-proves-it-can-publish.md) and [0050](0050-what-the-release-actually-published.md)
on this repository's side, and the carried entry [0043](0043-what-3-10-left-open.md) raised**

Reading the nine records of 3.10 against the tree they describe found the tree saying less than the
registry in five places. The root README said nothing was on npm while twelve packages were attested.
The example pinned `@pptx-studio/cli` at `^0.1.0`, which a caret below 1.0 can never resolve to the
0.2.0 its own source calls, so `npm install` in that directory as committed broke `tsc`. The plan's
note said the 2048-em face-box question had its answer while the plan's own carried entry said it
did not. Twelve published versions carried no provenance and no deprecation, so a consumer pinning
one could neither verify it nor be told why. And the canary's comment said a failure "becomes an
issue somebody owns" above a step that wrote an annotation to a page nobody reads.

Each is [0047](0047-ten-packages-were-never-attested.md)'s defect in a different file: a property of
the published thing asserted from memory of the workflow that produced it. This record makes each
one something the tree reads or writes instead.

## A monitor with no reader

`canary.yml` gains a second job, `report`, that `needs: published` and runs `if: always()`. On a red
install it creates a `canary` issue, or comments on the one already open with the run URL and the
date; on a green run it closes whatever is open. A cancelled run does nothing. It is its own job
because `issues: write` must never sit beside code installed from the registry — the isolation
[0048](0048-every-package-proves-it-can-publish.md) built for `id-token`, for a different
permission — and it needs neither a checkout nor an install: `gh` is on the image and resolves the
repository from `GH_REPO`. The `published` job loses the `GITHUB_TOKEN` it held and never used.

`gate.test.ts` reads the shape rather than trusting it: the report job needs the install and runs
on `always()`, opens on `failure` and closes on `success`, holds exactly `{ issues: write }` with no
`uses:` step and no line that runs npm, pnpm or npx; the install job has no `issues` permission and
no token; the workflow stays at `contents: read`. The install detector matches an invocation at the
start of a line, not the word — the issue body says "from npm" — and is itself asserted to see the
install job, so a blinded detector fails rather than passes vacuously.

## Unverifiable is not the same as unmarked

`attestations.ts` read the `latest` of each package and counted. It now reads the whole packument
once — one request, and the minutes-long lag after a name is first created, which
[0046](0046-nothing-gated-the-thing-being-published.md) hit in `newcomers.ts`, cannot reach a daily
monitor — and gives every version a standing: `provenance`, `deprecated`, or `neither`. Provenance
comes first, because a deprecated version a consumer can still verify is verifiable; `deprecated`
is a non-empty string, since npm lifts a deprecation by emptying it. The canary fails on any
`neither`, and on a `latest` that is not `provenance`, because a caret lands there.

`--require-all` is gone, on [0049](0049-the-right-to-publish-cannot-be-monitored.md)'s reasoning: a
gate whose strictness is a flag is off by default, and the canary was its only caller.

What this cannot do is deprecate. That is one `npm deprecate` per version with a one-time password
each — the ten at 0.1.0, `cli@0.1.0` and `render-dom@0.0.0`, twelve in all, and a range saves
nothing because each package has exactly one such version. Until they are run the canary reads
twelve `neither` and opens the issue above, which is the loop demonstrating itself. Caret safety was
already on the record: the published `cli@0.2.0` and `render-dom@0.1.0` ask `^0.1.0` and resolve to
the attested `0.1.1`, so no consumer following a range reaches a deprecated version.

## The example's manifest is written, not maintained

A check that the example's ranges are _satisfied_ by the workspace would have an ordering defect:
the release runs the candidate gate after `changeset version`, so on a minor bump no caret satisfies
both CI, where the tree holds the old versions, and the release, where it holds the new. So the
manifest is written at the one moment versions change and checked as proof of it.

`exampleRanges` writes `^<version>` for every in-scope dependency — exactly what `pnpm pack` writes
for `workspace:^`, measured in [0046](0046-nothing-gated-the-thing-being-published.md) — and
`staleRanges` refuses any range that is not that string, a patch behind included, because the
manifest is written by the release and a hand-edit is the thing being caught. `candidate.ts` reads
the versions off what it packed and refuses before it installs anything; `pnpm version-packages`
runs `example.ts` between `changeset version` and the lockfile. Its first run is the bump this record carries: eleven ranges, `^0.1.1` for the ten
and `^0.2.0` for `cli`.

The carried entry "The registry canary pins ^0.1.0" is removed. Both its claims are false now: the
canary installs `latest` since 0046, and the manifest is written by the release and refused by the
gate when it is not.

## What the tree said

- The root README: twelve packages on npm with an npm badge; the ten Phase 3 sub-phases as rows
  with Gate 3 open; `cli/` naming all five verbs and a `render` example; `pnpm check` described in
  the order `package.json` runs it, `legal`, `structure`, `references` and `pkg:qa` included; the
  example described as the candidate gate and the `latest` canary use it.
- `packages/cli/src/main.ts` named `fidelity` as a verb it has never had, in a fourteen-line
  comment; `tools/repo/layering/layers.ts` did the same in its role string; the help text lost its
  apostrophe; `render-dom`'s tarball was the one of twelve without `CHANGELOG.md`;
  `tools/repo/adr-index.ts` said "ninety-six" records of a plan with 109 sub-phases.
- `tools/release/` held eleven hand-authored files and this record adds three, so it splits by
  subject at the twelve-file cap rather than the thirteenth: `candidate/` is the gate on this
  commit's tarballs, `registry/` is every question only npm can answer, and `gate.test.ts` about
  the workflows stays at the top. Every reference moved with it; the ADRs before this one keep
  their old paths, being records.

## Verification

`pnpm check` green. The scratch install, the canary's issue and the deprecations are proved by CI
and by a person rather than here, for the reason 0046 gives: this machine does not reach the
network unasked.

What CI said, once this landed on `main` at 55f3073: the schedule fired first, run 34689427313,
and its `report` job opened issue #23; a dispatch forty seconds later, 34689455583, took the other
branch and commented "Still red". The candidate gate, 34688374783, passed its first run on the
ranges `example.ts` had written. Release 34692318833 then published `cli@0.3.0`,
`render-svg@0.2.0`, `text@0.2.0` and `render-dom@0.1.1` over OIDC; canary 34692582931 installed
that `latest` set, typechecked, smoked, built, and had `npm audit signatures` verify every one,
failing only on the twelve versions that read `neither`. The twelve were deprecated by hand, and canary
34699970846 read 16 with provenance, 12 deprecated, none `neither`, went green, and its `report` job
closed #23 with the run URL: the loop has now run every branch it has.

| mutant                                     | killed by                                                            |
| ------------------------------------------ | -------------------------------------------------------------------- |
| `report` loses `needs: published`          | reports from a job of its own, after the install has finished        |
| `report` loses `always()`                  | the same                                                             |
| open and close conditions swapped          | opens on red and closes on green, not the other way round            |
| `issues: write` moved onto `published`     | gives the job that installs from the registry nothing to write with  |
| workflow permissions gain `issues: write`  | the same                                                             |
| `report` gains a checkout                  | holds the right to write issues where nothing from the registry runs |
| `report` runs `npm install`                | the same                                                             |
| the install detector blinded               | gives the job … nothing to write with, after one genuine survivor    |
| attestation test inverted                  | one version: has provenance / is neither                             |
| deprecated checked before provenance       | reads an attested and deprecated version as verifiable               |
| an empty string counts as deprecated       | reads an emptied deprecation as no deprecation                       |
| `latest` from key order                    | marks latest from the tag, not from the order                        |
| `unverifiable` drops the latest clause     | a latest that is only deprecated still fails                         |
| the message omits the version              | an older version that was neither, named with its version            |
| `staleRanges` accepts any satisfying caret | asks for exactly what pack writes                                    |
| `staleRanges` ignores an unpacked name     | refuses a range on a package this run did not pack                   |
| `exampleRanges` rewrites off-scope         | ignores everything outside the scope, after one genuine survivor     |

Seventeen tried, seventeen killed, after two genuine survivors. The detector had folded `uses:`
steps into the same predicate as `run:` lines, so the checkout action satisfied "the detector still
sees the install" on its own; they are two predicates now. And the off-scope test used `0.0.1` as
its untouched value, which was the literal the mutant wrote; it uses a version nothing else names.
Both were missing assertions, and neither was called equivalent.

`candidate.ts` not calling `staleRanges` survives every unit test, as `candidate.ts` not reading
the lockfile did in 0046: the CI `candidate` job on a hand-broken example is the test, and this
record's own first run — every committed range stale until `example.ts` wrote them — is it.

## Deviations

1. **No semver library.** The rule is the exact caret string, which needs none; a satisfies check
   would have needed one and been wrong in the release.
2. **`--require-all` dropped** from `attestations.ts` without a replacement flag.
3. **The split reached files this record did not otherwise touch** — `newcomers.ts`,
   `publishers.ts`, `oidc.ts`, `bootstrap.ts` and their tests moved unchanged but for one import.

## Open questions

1. **A revoked publisher is still invisible between releases**, from
   [0049](0049-the-right-to-publish-cannot-be-monitored.md), and the preflight still mints tokens
   it never uses, from [0048](0048-every-package-proves-it-can-publish.md); npm has no read-only
   form of either question.
2. **One package manager, one Node version** in the scratch install, from
   [0046](0046-nothing-gated-the-thing-being-published.md).
3. **The canary issue has one label and one title**, so two distinct failures on one day share an
   issue. The run URL in each comment is what tells them apart.
4. **A canary inside the registry's replication window reads half a release.** Run 34692440841,
   dispatched two minutes after the release, was served `render-svg@0.2.0` — which asks
   `text@^0.2.0` — minutes before it was served `text@0.2.0`, and the install failed on
   `ETARGET`. The window closed on its own; a canary dispatched by hand after a release should wait
   for `attestations.ts` to list every new version as `latest` first.
