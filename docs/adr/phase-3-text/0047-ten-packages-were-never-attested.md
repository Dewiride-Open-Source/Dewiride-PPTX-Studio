# 0047 — Ten packages were never attested

**Sub-phase 3.10 · accepted · supersedes the provenance claim of
[0042](0042-rendering-without-a-browser.md)**

[0042](0042-rendering-without-a-browser.md) ends its account of the first publish with _"Eleven
packages at 0.1.0, each with a provenance attestation."_ `docs/plan/phases.json` said the same. The
registry disagrees:

| package                     | latest | attestation |
| --------------------------- | ------ | ----------- |
| `@pptx-studio/cli`          | 0.2.0  | provenance  |
| `@pptx-studio/render-dom`   | 0.1.0  | provenance  |
| the other ten, all at 0.1.0 | 0.1.0  | **none**    |

Two of twelve. The two are exactly the two published since, over OIDC through a trusted publisher.
Every package from the original release carries nothing a consumer can verify.

## Why

0042 reasoned correctly that provenance cannot be generated off a laptop and that the first publish
of an empty scope had to use a token, and then assumed the token job produced attestations. It does
not follow. **npm enables provenance automatically only on the OIDC path.** With `NODE_AUTH_TOKEN`
set, the CLI takes the legacy route, and provenance has to be asked for explicitly — `--provenance`,
`NPM_CONFIG_PROVENANCE=true`, or `publishConfig.provenance`. It is supported with a token and it is
not the default. The deleted one-time job did not ask.

0042 knew the first half of this. Its own words: _"with `NODE_AUTH_TOKEN` set the npm CLI takes the
legacy route and silently skips OIDC, and a silently unsigned release is worse than a failed one."_
That sentence is the reason `release.yml` may never grow a token path. It is also, read once more,
a description of what the token job itself did — and the conclusion drawn two lines later was that
eleven packages were attested.

## The defect is that nothing asked

This is [0046](0046-nothing-gated-the-thing-being-published.md)'s thesis again, one level out. There,
nothing executed a published package. Here, nothing read one. A property of the published artifact
was asserted from the shape of the workflow that produced it, held for a month, and appeared in a
record and a status field — and one HTTP request to the registry falsifies it.

`tools/release/attestations.ts` reads `dist.attestations` for the latest version of every public
package and reports which carry provenance. `--require-all` makes it an assertion. It runs in
`canary.yml`, because whether the published artifacts are attested is a question about the world
rather than about the commit, and that is what the canary is for.

It is deliberately not a release gate. Provenance is created _during_ publish, so a gate could only
check the previous release — which is precisely the circularity 0046 removed.

## What it does not fix

The ten are still unattested and this record does not change that. Attesting them means republishing
each at a patch version over OIDC, which needs a trusted publisher configured for each; only `cli`
and `render-dom` are known to have one, because only those two have published that way. A release
that bumps a package whose publisher is missing fails on that package, and `changeset publish` goes
one at a time. Worth doing, worth checking first, and not done here.

## Open questions

1. **Whether the other ten have trusted publishers at all** is unmeasured. `npm trust list` answers
   per package and needs an OTP each time, so nothing in CI can ask it; the registry exposes no
   public endpoint for it. Until one of them publishes over OIDC, it is unknown.
2. **`npm audit signatures` is the consumer-side check** and nothing in this repository runs it. It
   would verify the attestations rather than merely counting them, and the candidate gate's scratch
   consumer is the obvious place — except that it installs from tarballs, which have no attestations
   by construction.
