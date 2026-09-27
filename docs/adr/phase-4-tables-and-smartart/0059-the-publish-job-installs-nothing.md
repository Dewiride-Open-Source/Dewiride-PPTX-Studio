# 0059 — The publish job installs nothing

Date: 2026-09-27
Status: **accepted** — the right to publish sits in a job that installs nothing and publishes only
tarballs it has held to digests recorded before any consumer code ran. The job that installs the
registry holds nothing it could publish or push with.

**Maintenance after 4.1 · supersedes the single-job release that
[0046](../phase-3-text/0046-nothing-gated-the-thing-being-published.md) and
[0048](../phase-3-text/0048-every-package-proves-it-can-publish.md) describe, and applies 0048's own
rule to it.** 0048 wrote the rule for the canary: "the job that mints those tokens installs nothing …
third-party code running beside that permission could mint a publish token for every package in the
scope."

`release.yml` broke that rule in one job. That job held `id-token: write` and `contents: write`, and
its step "Consume the exact tarballs that are about to be published" did two things:

- a lockfile-less, age-ungated `npm install` of the website's whole dependency tree;
- then ran that tree: typecheck, lint, tests, smoke, `next build`.

[0057](0057-transitive-ranges-float.md) showed how fast that tree moves: a patch two levels down
reached every consumer within hours. A release run in that window executes whatever arrived beside
the publish identity. So did every install script in that tree, since npm runs them by default.

## Two jobs, and why moving the permission is not enough

Moving `id-token: write` into a second job is the obvious half. The other half is what that job
publishes.

The tarballs are packed on the runner that then installs and runs the website. Code from that tree
could rewrite them on disk before upload, and `publish` would put valid provenance on whatever it was
handed. A job boundary protects a token, not an artifact built beside the threat.

So `prove` fixes what it will hand on before the threat can run:

1. It versions and commits, still with `persist-credentials: false`.
2. `candidate.ts --pack-only` packs every publishable package and writes `packed.json`: each file
   name and its sha512.
3. The step writes two job outputs:
   - `commit`, the version commit's id;
   - `digest`, the sha256 of `packed.json`.
4. It uploads the tarballs, `packed.json` and a `git bundle` of the version commit. Artifacts are
   immutable, and the job's token cannot delete one.
5. Only then does `candidate.ts --packed` hold every tarball to `packed.json` and install the
   website around them.

`publish` needs `prove`, holds `contents: write` and `id-token: write`, and installs nothing: no
`pnpm/action-setup`, no `pnpm`, no `npm install`, no website.

1. It checks out `main` and refuses if `main` has moved since `prove` ran on it.
2. It downloads the artifact and holds it to what `prove` recorded:
   - `packed.json` against `digest`;
   - the bundle's commit against `commit`;
   - then fast-forwards.
3. `publishers.ts` asks npm, per package, whether this workflow may publish. It is unchanged, and its
   imports are `node:` and repository files only.
4. `publish.ts` holds each tarball to `packed.json` and skips a version npm already serves, so a retry
   is clean. It runs `npm publish <tarball> --access public --provenance` over OIDC, with no
   `NODE_AUTH_TOKEN`, and tags each version `<name>@<version>` as changesets would.
5. It pushes the version commit and its tags.

Trusted publishing binds a publisher to the workflow file, not to a job
([0049](../phase-3-text/0049-the-right-to-publish-cannot-be-monitored.md)), so the split changes nothing
npm checks.

## What this also closes

`changeset publish` re-packed each package itself: the gate proved one set of tarballs and the
registry received another, built by the same code a step later. `publish` now sends the bytes the
gate installed. `--provenance` is passed explicitly, so a publish that cannot attest fails before it
publishes rather than shipping unverifiable, which is how ten packages shipped in
[0047](../phase-3-text/0047-ten-packages-were-never-attested.md).

## What it does not close

`prove` still runs the root toolchain from `pnpm install --frozen-lockfile`. That toolchain is
lockfile-pinned, 24-hour age-gated (`minimumReleaseAge: 1440`), and its install scripts are denied
unless listed. That is the trust every CI job already extends to it. A compromised build tool could
still produce a hostile tarball that `publish` would faithfully ship. Provenance does not answer
that, and nothing short of a separately built, reproducible artifact would.

## Verification

- `tools/release/gate.test.ts` asserts the layout, 4 new cases:
  - the job that installs has `contents: read`, `actions: read`, no `id-token`, and a checkout
    without credentials;
  - the job with `id-token` needs `prove`, runs no install, no `candidate.ts` and no pnpm setup, and
    is the only job with `id-token`;
  - pack, upload and consume come in that order, and pack records both outputs;
  - `sha256sum --check` runs before `publishers.ts`, on `needs.prove.outputs`.

  The existing gate and OIDC cases now read the `prove` and `publish` jobs.

- `readPackedRecord` refuses a file name that is not a bare `.tgz`, and `tamperedTarballs` names the
  file whose bytes moved: 9 cases.
- Both scripts are run on a directory whose tarball moved after it was packed, with npm pointed at an
  unreachable registry. `candidate.ts --packed` and `publish.ts` each exit non-zero, naming the file.
- Mutants, 9/9 killed: in `release.yml`, `pnpm install` in `publish`, `id-token` on `prove`, checkout
  credentials kept, the digest check dropped, consume before upload; the file-name refusal; a
  `tamperedTarballs` that reports nothing; each script's digest check skipped.
- Locally:
  - `pnpm candidate tmp/rel --pack-only` packed 12 and wrote `packed.json`.
  - One byte appended to a tarball made both `candidate.ts --packed` and `publish.ts` refuse.
    `publish.ts` was never run untampered here, where it would reach `npm publish` with this
    machine's credentials.
  - The bundle handoff ran in two scratch clones: `git bundle create origin/main..HEAD`, `verify`,
    `fetch … HEAD`, `FETCH_HEAD` equal to the recorded commit, fast-forward.
- The first real run through both jobs is the release of
  [0058](0058-what-code-scanning-found.md)'s fixes, and is recorded below once it has happened.

## Open questions

1. `npm publish <tarball> --provenance` over trusted publishing has not been exercised in this
   repository until that first run.
