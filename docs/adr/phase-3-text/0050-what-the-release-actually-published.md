# 0050 — What the release actually published

**Sub-phase 3.10 · accepted · corrects a mechanism claim in
[0048](0048-every-package-proves-it-can-publish.md)**

Twelve of twelve packages carry a provenance attestation. `tools/release/attestations.ts` reads the
registry and answers `12 attested, 0 without, 0 not published`, which closes what
[0047](0047-ten-packages-were-never-attested.md) opened.

## Ten published, not twelve

[0048](0048-every-package-proves-it-can-publish.md) said `updateInternalDependencies: patch` would
carry `cli` and `render-dom` along as dependents, so all twelve would publish. It did not, and the
reason is worth keeping: **`workspace:^` records no version.** A dependent whose internal ranges are
the workspace protocol has nothing for `changeset version` to rewrite, so no dependent needed a bump
and only the ten holding changesets published. `cli` stayed at 0.2.0 and `render-dom` at 0.1.0.

The count still reached twelve because those two were already the attested ones. The outcome 0048
predicted was right; the mechanism it gave for the outcome was not.

Nothing is left stale by that. The published `cli@0.2.0` asks for `^0.1.0`, which resolves to the
attested `0.1.1`, so a consumer installing it today gets verifiable packages without `cli` being
republished at all.

## The preflight, measured on both sides

The same script, the same twelve packages, two workflows, three hours apart:

| asked from    | answer                                     |
| ------------- | ------------------------------------------ |
| `canary.yml`  | 0 of 12, each `404 ... package not found`  |
| `release.yml` | 12 of 12, `may publish from this workflow` |

That is [0049](0049-the-right-to-publish-cannot-be-monitored.md) as a controlled experiment rather
than an inference. The only variable was which workflow asked, and a trusted publisher names one.

## Verification

Release run `34580266504` on `be83b2d`; the version commit is `1b68a8a`. Every step green, including
the preflight's first real run and the candidate gate on the bumped tarballs.

## Open questions

1. **The 0.1.0 versions stay unattested**, as [0048](0048-every-package-proves-it-can-publish.md)
   said they would — attestation is per version. A consumer pinning exactly `0.1.0` still gets a
   package `npm audit signatures` cannot verify. Every caret range in every published manifest
   resolves past it, so this only reaches someone who pinned deliberately.
2. **A revoked publisher is still invisible between releases**, from
   [0049](0049-the-right-to-publish-cannot-be-monitored.md), and nothing here changes that.
