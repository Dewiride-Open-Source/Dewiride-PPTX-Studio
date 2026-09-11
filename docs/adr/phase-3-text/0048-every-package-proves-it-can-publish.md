# 0048 — Every package proves it can publish, before any of them do

**Sub-phase 3.10 · accepted · closes the open questions of
[0047](0047-ten-packages-were-never-attested.md)**

[0047](0047-ten-packages-were-never-attested.md) recorded that ten of twelve published packages carry
no provenance, and left two things unanswered: whether those ten have a trusted publisher at all, and
that nothing here runs `npm audit signatures`. Both are answered here, and the ten are attested.

## The ten did have publishers

`npm trust list` needs a one-time password per package, so nothing automated can ask it and the
registry exposes no public endpoint — 0047 was right about that, and the packument's
`_npmUser.trustedPublisher` does not help, because it records how a _version_ was published rather
than what the package is configured to accept. Asked by hand, once, for all twelve: every one carries
a GitHub publisher on this repository's `release.yml`, with `publish, stage publish`, under twelve
distinct config ids.

So nothing was missing. The one-time token job simply never asked for provenance, which is exactly
what 0047 concluded and is now measured rather than inferred.

## Knowing it once is not knowing it

A trusted publisher can be revoked or migrated with nothing in this repository changing, and an
OTP-gated answer cannot be asked again by a machine. An answer good only on the day it was given is
the shape of defect 0047 is about.

npm trades a GitHub ID token for a publish token **one package at a time**:

```
POST /-/npm/v1/oidc/token/exchange/package/@pptx-studio%2fxml
```

Measured against the live registry: that path answers 400 without an `Authorization` header and 401
with a bad one, and the same path without a package on the end is 404. There is no bulk form. So the
right to publish is a per-package fact, `changeset publish` asks for it one package at a time, and a
package whose publisher is missing fails partway through a release that has already shipped others.

`tools/release/publishers.ts` asks for every one of them while nothing has been published. It runs
twice, with the gate-and-monitor split [0046](0046-nothing-gated-the-thing-being-published.md)
established:

- In the release, `--require-pending` — the packages whose version in the tree is not already on the
  registry, which is exactly the set `changeset publish` would send, and after `changeset version`
  is every bump.
- In the canary, `--require-all`, daily, so a revoked publisher is found before a release needs it.

**An unknown answer is not a refusal.** Only 401, 403 and 404 count as "not you"; any other status
throws. Reporting a 500 as a missing publisher would refuse a release npm would have accepted, and a
gate that fails for reasons that are not about the commit is the thing 0046 removed.

**The job that mints those tokens installs nothing.** It is its own job in `canary.yml`, with
`id-token: write` and two pinned actions, and no registry install anywhere in it — third-party code
running beside that permission could mint a publish token for every package in the scope. The job
that does install from the registry has no `id-token` at all. Both are asserted.

## Verifying is not counting

0047's `attestations.ts` reads `dist.attestations` and counts. `npm audit signatures` checks the
signature and the attestation, which is what a consumer would actually run, and 0047 left it out
because the candidate gate installs tarballs and a tarball carries no attestation by construction.

The canary is where it belongs, and it gets a project of its own that depends on the published
packages at `latest`. Not the example's tree: that would audit several hundred third-party packages,
and a gap in one of those is not ours to answer for and would make this red for a reason nobody here
can fix.

## The ten

One changeset, ten patches. `updateInternalDependencies` is `patch`, so changesets carries `cli` and
`render-dom` along as dependents and all twelve publish over OIDC, which attests automatically.
Nothing about the code changes; the changelog entry says so.

## Verification

`pnpm check` green. 7 mutants tried, 7 killed.

| mutant                                              | killed by                                               |
| --------------------------------------------------- | ------------------------------------------------------- |
| the release stops asking npm whether it may publish | requires the preflight among the release's steps        |
| the preflight moves after the publish it guards     | requires it to run before `changeset publish`           |
| a token is put back on the publish path             | refuses `NODE_AUTH_TOKEN` anywhere in the release       |
| the canary asks only about the pending release      | requires `--require-all` there, not `--require-pending` |
| the token-minting job loses `id-token`              | requires `id-token: write` on that job                  |
| the token-minting job installs something            | refuses an install step beside the minting permission   |
| the registry-installing job gains `id-token`        | requires that job to have none                          |

Eleven unit tests cover the addressing and the decision: the audience comes from the registry host,
the scope separator is escaped, the exchange names one package, the audience appends to an issuer
that already has a query string, the pending set is matched on version rather than name, and a
refusal blocks only a package this release would publish.

The network half is proved by CI rather than locally, for the same reason 0046's scratch install is:
it needs a workflow that can mint an ID token, and this machine cannot.

## Open questions

1. **The 0.1.0 versions stay unattested.** Attesting is per version, so every `0.1.0` remains on the
   registry unverifiable, and a consumer pinning one gets exactly what they got before. Deprecating
   those versions would say so in the one place a consumer looks; not done here.
2. **The preflight mints tokens it never uses.** They are short-lived and never read, but a check
   whose only implementation is to acquire the credential is heavier than a check should be. npm has
   no read-only form of the question today.
3. **One rasteriser's answer.** Unrelated to this record and still open from
   [0045](0045-the-face-box-belongs-to-the-rasteriser.md): macOS is measured by neither platform.
