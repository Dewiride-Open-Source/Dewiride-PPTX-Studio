# 0049 — The right to publish can only be asked by the workflow that holds it

**Sub-phase 3.10 · accepted · supersedes the canary half of
[0048](0048-every-package-proves-it-can-publish.md)**

[0048](0048-every-package-proves-it-can-publish.md) put the OIDC preflight in two places: the release,
for the packages it is about to publish, and the canary, daily, for all twelve. The canary half
cannot work. Its first dispatch said so — `0 of 12 can publish over OIDC`, every one of them
`404, OIDC token exchange error - package not found`.

## What that 404 means

An npm trusted publisher names a repository **and a workflow file**, and all twelve name
`release.yml`. The exchange validates the ID token's workflow claim against that configuration, and
when nothing matches it answers 404 with `package not found` — which is not what happened. The
packages exist; `npm trust list` shows a publisher on each of them, measured the same day. The status
is an authorization failure wearing a lookup's clothes, and it is a known trap rather than a quirk of
this repository.

The natural experiment was already on the record. `cli@0.2.0` and `render-dom@0.1.0` published over
OIDC against these same twelve configurations, from `release.yml`. The only variable this run changed
was the workflow asking.

## Why there is no fix worth having

Attaching a second trusted publisher for `canary.yml` would make the check pass. It would also hand a
scheduled, unattended workflow — one whose whole job is to install the public registry into itself —
the right to publish all twelve packages. That job was isolated with no install and two pinned actions
precisely so nothing it ran could reach that permission. Buying the answer with the permission inverts
the reason for asking.

So the question is not monitorable. **A credential can only be verified by exercising it, and only its
holder can exercise it.** The release asks, from the workflow that legitimately holds the right, at the
one moment the answer is load-bearing.

That is a narrower guarantee than 0048 claimed, and it is the true one: a publisher revoked between
releases is found by the next release, which fails before publishing anything rather than halfway
through. The daily warning 0048 promised was never available at a price worth paying.

## What the gate lost, and what it kept

`--require-all` had no caller left. `--require-pending` went with it, for a better reason: a gate whose
strictness depends on a flag is off by default, and a step that lost the flag would have exited 0 on
every refusal while still looking like a gate. `publishers.ts` now has one behaviour — report every
package, and fail on any that this release would publish and npm will not accept.

The refusal message says what a 404 there actually means. The operator reading it is mid-release, and
the wording would otherwise send them hunting for a package that is sitting on the registry.

`npm audit signatures` keeps its place in the canary, where it works: the dispatch that refuted the
other half reported `13 packages have verified registry signatures` and `2 packages have verified
attestations`, which is the pre-release state, measured rather than assumed.

## Verification

`pnpm check` green. The canary dispatch is run `34570631490`; its example job failed only on the
attestation assertion, which is the ten this release attests.

| mutant                                              | killed by                                         |
| --------------------------------------------------- | ------------------------------------------------- |
| the release stops asking npm whether it may publish | requires the preflight among the release's steps  |
| the preflight moves after the publish it guards     | requires it to run before `changeset publish`     |
| a token is put back on the publish path             | refuses `NODE_AUTH_TOKEN` anywhere in the release |
| the canary asks a question it cannot be told        | refuses `publishers.ts` anywhere in the canary    |
| any canary job gains `id-token`                     | refuses that permission on every job in the file  |

## Open questions

1. **A revoked publisher is invisible until the release that needs it.** That release fails safely,
   having published nothing, but the operator finds out at the worst moment rather than the cheapest.
   npm exposes no read-only form of the question and no endpoint a monitor could hold.
2. **`npm trust list` is still the only full answer**, and it needs a one-time password per package,
   so it stays a thing a person does by hand and nothing here can schedule.
