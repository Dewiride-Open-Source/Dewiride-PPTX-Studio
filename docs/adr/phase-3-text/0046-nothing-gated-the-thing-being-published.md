# 0046 — Nothing gated the thing being published

**Sub-phase 3.10 · accepted**

Three defects reached npm at 0.1.0: kerning silently dropped behind an Extension GPOS lookup,
`renderDeck` refusing to substitute on a machine holding 53 usable faces, and a face box read from
the wrong table. [0044](0044-what-the-runners-own-fonts-said.md) and
[0045](0045-the-face-box-belongs-to-the-rasteriser.md) record the defects. This records why they
were publishable, which is a different question and the more important one.

Then the fix for them could not be released, because the gate that should have caught them had been
wired so that it blocked its own repair.

## Nothing had ever executed a published package

Three things looked like they covered this and none of them did.

- **`pnpm pack:qa` is `publint --strict && attw --pack .`** Both read the tarball's manifest —
  `exports`, `types`, `files`. Neither resolves a specifier at run time and neither executes a line
  of the package. publint's own docs are explicit that it is static analysis over built output.
- **`apps/studio` is `private: true` with `workspace:^` dependencies**, so it resolves through the
  pnpm link farm. It exercises the source tree, not the artifact.
- **`pnpm check` runs the suites**, all of which import from source for the same reason.

So the packages went to the registry having never been installed anywhere, and every defect above is
the kind that only appears once something resolves them through package boundaries and runs them.

The `example` job was added afterwards to close exactly this, and it half-closed it: it does install
real tarballs with npm outside the workspace. But it installs them **from the registry**, so it can
only ever report on a version that is already published. It is one release behind by construction,
and it has no power to stop the release it is behind.

## A monitor wired into a gate socket

`release.yml` refuses to publish unless CI on that commit concluded `success`. That is the right
instinct and the wrong reading, because CI contained the `example` job. So:

```
release   -> requires CI green on this commit
CI        -> red, because `example` fails
example   -> installs @pptx-studio/cli@^0.1.0 and runs the broken published code
the fix   -> requires a release
```

The dispatched run failed in twelve seconds with `CI on a2051ce is 'failure'`. It could never have
done anything else.

The sharp framing is not "circular dependency" but **the gate was not a function of the commit**.
The same commit could be green today and red tomorrow with nobody touching it, because one of its
inputs was the state of a public registry. That disqualifies it as a gate independently of the
deadlock; the deadlock is only the loudest symptom. A gate asks _is this commit fit to ship_ and
must be answerable from the commit. A monitor asks _is the world healthy right now_ and can only be
repaired by a new release. The `example` job was the second wearing the clothes of the first.

The job's own comment already said so — _"it can go red for reasons that have nothing to do with the
commit, and a gate that does that stops being read"_ — and I put it in the run the release gate
reads anyway.

## What replaces it

**A gate on the candidate.** `tools/release/candidate.ts` packs every publishable package, installs
the tarballs into a scratch consumer, refuses anything that did not come from one of them, and runs
the example against the result. It depends on the commit and on nothing already published, so the
release can require it. It would have caught all three 0.1.0 defects.

**Packing is not the hard part; pinning the edges is.** `pnpm pack` rewrites `workspace:^` and
`catalog:` exactly as publish does — measured, not assumed: `@pptx-studio/cli`'s eight
`workspace:^` dependencies come out of the tarball as `^0.1.0`. Which is the problem. Once rewritten
they are ordinary semver ranges, and an ordinary semver range resolves from the registry.

The dangerous case is the package **nobody bumped**. Changesets versions only packages with a
changeset and their dependents, so in a release of `cli` alone, `xml` stays at 0.1.0 and `cli`'s
packed manifest asks for `^0.1.0`. The registry has exactly that. The install succeeds, the suite
passes, and it passed against the published artifact rather than the candidate — no error, no
warning. (The bumped package is safe by accident: `^0.2.0` matches nothing published, so it fails
loudly.)

Two things close it, and the gate does both:

1. Every candidate appears in the scratch project's `dependencies` **and** its `overrides`, with a
   byte-identical `file:` specifier — npm rejects an override that differs from a direct dependency.
   Overrides reach transitive edges, which is the whole point.
2. `@pptx-studio:registry` points at `http://127.0.0.1:1/`, so a lookup for our own scope has
   nowhere to go. A fallback becomes a connection error instead of a quiet success. Real
   dependencies still come from npmjs.org, as they must.

Then the lockfile is read back and every `@pptx-studio/*` entry must be `resolved` from a `file:`
URL, not a symlink, with `integrity` equal to the sha512 of the tarball this run packed.

**The monitor keeps its job, in `canary.yml`,** on a schedule and a dispatch. It still answers a real
question no gate can — did the _publish_ go wrong, rather than the code — and it can no longer block
anything. It installs `latest` rather than the example's own ranges, because a caret below 1.0 is
bounded by the minor: `^0.1.0` can never resolve to 0.2.0, so a monitor honouring the manifest would
report on the version each release replaces and need a hand-edit after every one.

**The release gate names jobs** instead of reading the run's conclusion, which folds in every job in
the file. A job absent from the list cannot block a publish. A job named in the list that did not run
is a failure, not a pass — without that, a rename makes the gate pass vacuously, which is worse than
the deadlock it replaced.

**The gate runs again inside the release, after `changeset version`.** Before that step the packed
ranges are the old ones; only afterwards do the tarballs carry the versions a consumer will really
resolve. It is the last thing between the code and a registry that cannot be undone.

## A release cannot create a package npm does not have yet

A trusted publisher is a setting on an npm package, and [npm/cli#8544] is open, so there is no way
to attach one to a name the registry does not hold. `npm trust` says the same in as many words: the
package must already exist. So the first publish of a new name cannot go out over OIDC at all.

`@pptx-studio/render-dom` is that package — `0.0.0`, in changesets' `ignore`, E404 on the registry.
While it is ignored nothing tries to publish it. The hazard is the moment somebody lifts the ignore:
`changeset publish` publishes one package at a time, so the release would ship some packages, reach
the new name, fail to authenticate, and stop — leaving a half-published release that cannot be
rolled back.

`tools/release/newcomers.ts` runs before `changeset publish` and asks the registry which of the
packages this release would publish are new. If any are, it refuses and prints the bootstrap: a
one-time granular token scoped to `@pptx-studio`, `npm publish --provenance --access public`, attach
the trusted publisher, revoke the token, then lift the ignore. Those steps are printed in full
rather than linked, because whoever reads them is mid-release and they are not guessable.

Verified against the live registry in all three states: with `render-dom` ignored, 11 packages and 0
new; with the ignore lifted and the name absent, 12 and 1, naming it, exit 1; and after the
bootstrap, 12 and 0.

**Ask the dist-tags endpoint, not the packument.** Running this for real found that
`registry.npmjs.org/@scope%2fname` answers 404 for a while after the publish that created the
package — 404 on the packument and 200 on both `/0.0.0` and `/-/package/@scope%2fname/dist-tags`, at
the same second. npm builds the aggregated document asynchronously, and the minutes after a first
publish are the one moment this check is asked about, so the endpoint that lags is the wrong one to
ask. It would have refused a release for a package that already existed.

`@pptx-studio/render-dom` has been through it: published once by hand at `0.0.0` to create the name,
trusted publisher `2dc44547-f099-464d-bce9-b94d2ca1b44b` attached to this repository's `release.yml`,
and the name lifted out of `ignore`. Its first OIDC release carries provenance; the `0.0.0` that
created it does not, which is the cost of npm having no way to pre-register.

[npm/cli#8544]: https://github.com/npm/cli/issues/8544

## Verification

`pnpm check` green. 7 mutants tried, 7 killed.

| mutant                                             | killed by                                                  |
| -------------------------------------------------- | ---------------------------------------------------------- |
| overrides dropped from the scratch manifest        | pins every candidate as both a dependency and an override  |
| a registry resolution is accepted                  | fails the silent case: resolved from the registry          |
| a symlink is accepted                              | fails a symlink, which is the link farm leaking in         |
| integrity is not compared                          | fails a tarball whose bytes are not the ones packed        |
| a packed package that never installed is ignored   | fails a package that was packed and never reached the tree |
| the `file:` specifier keeps backslashes            | writes a file: specifier with forward slashes              |
| the release gate stops requiring the candidate job | requires the gate that consumes the candidate tarballs     |

The scratch install itself is proved by CI rather than locally: it installs from the network, which
this machine does not do without being asked.

## Deviations

No Verdaccio. A local registry is the higher-fidelity option and is what Babel, Storybook and Nx
use — the consumer then resolves by name and range exactly as a user would. It is also a new
dependency, and `overrides` plus a dead scope registry closes the same hole with tooling already
present. If the peer-dependency surface ever grows, revisit it.

## Open questions

1. **One package manager, one Node version.** The scratch install is npm on the repo's pinned Node.
   pnpm's isolated layout is what catches an undeclared dependency that npm's hoisting hides, and
   the two disagree about `exports` resolution often enough to be worth a matrix. Not done here.
2. **`@pptx-studio/render-dom` is still unpublished.** The guard above refuses to half-publish it,
   which is not the same as shipping it: the bootstrap needs a credential this repository does not
   hold, so it stays a human step.
3. **The canary has no owner.** A monitor nobody reads is worse than no monitor. It annotates the run
   on failure; it does not yet open an issue or notify anyone.

[npm/cli#8544]: https://github.com/npm/cli/issues/8544
