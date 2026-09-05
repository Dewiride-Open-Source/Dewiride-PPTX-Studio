# ADR 0001 — Toolchain for the PPTX Studio monorepo

- **Status:** Accepted
- **Date:** 2026-08-26
- **Sub-phase:** 0.1

## Context

PPTX Studio is a browser-first, ESM-only TypeScript monorepo that publishes a family of
`@pptx-studio/*` packages to npm. The single hardest constraint in the plan is architectural,
not functional: the core packages **must not** be able to touch Node APIs, must not depend
upwards through the layer graph, and must not acquire a dependency on React. Convention decays;
CI does not. So the toolchain is chosen for what it can _mechanically prevent_, not for speed.

Versions below were read from the npm registry on 2026-08-26, not from memory.

## Decisions

### TypeScript 6.0.3, **not** 7.0.2

TypeScript 7.0.2 (the Go-native compiler, GA 2026-07-08) is 8–12× faster on full builds. We are
not adopting it yet, for one concrete reason:

> `typescript-eslint@8.68.0` declares `peerDependencies.typescript: ">=4.8.4 <6.1.0"`.

TypeScript 7.0 shipped without a stable programmatic API (expected in 7.1), and typescript-eslint's
type-aware rules are built directly against that API. Installing both makes pnpm refuse the peer
range; forcing it crashes inside `typescript-estree`. Losing ESLint would cost us the three bans
that are the entire point of this sub-phase.

The migration cost of waiting is close to zero: TypeScript 7.0 is specified to match **6.0's**
type-checking and command-line behaviour, so 6.0.3 → 7.x is a version bump, not a code change.

**Migration trigger (write it down or it never happens):** bump to TypeScript 7 the day
`typescript-eslint` publishes a release whose `peerDependencies.typescript` range admits `^7`.
Until then, `typescript` is pinned exactly at `6.0.3` in the catalog.

Rejected alternative: TypeScript 7 + `oxlint` (+ `oxlint-tsgolint`, which is genuinely stable and
covers 59 of 61 type-aware rules). Oxlint is materially faster, but this repo's load-bearing lint
rules are `no-restricted-imports`, a bespoke top-level-await rule, and import-boundary zones —
custom-rule authoring and esquery selectors, which is exactly where the ESLint ecosystem is deeper.
Oxlint may be added later as a fast pre-filter; it is not a replacement.

### ESLint 10.9.1, flat config, `eslint.config.mjs` (not `.ts`)

ESLint 10 removed eslintrc entirely, so flat config is not a choice. We write the config in plain
`.mjs` rather than `.ts` deliberately: `eslint.config.ts` requires `jiti >= 2.2.0` as a loader,
which adds a transpile step and a dependency to the one file that must never fail to load.

Also note: ESLint 10 resolves `eslint.config.*` starting from **each linted file's directory**,
not the cwd. A stray config in a package directory would silently take over that package. There is
exactly one config file in this repo, at the root, and `tools/layering/check-layering.mjs` will
grow a rule to keep it that way if it ever becomes a problem.

### pnpm 11.6.0 + workspace catalogs

All shared versions live in the `catalog:` block of `pnpm-workspace.yaml`; no package pins a shared
version itself. The `catalog:` protocol is rewritten to a concrete range at `pnpm publish`/`pack`
time, so consumers never see it.

pnpm 11 changed several defaults that matter here and are being kept as-is:

| Setting                   | Default      | Why we keep it                                              |
| ------------------------- | ------------ | ----------------------------------------------------------- |
| `minimumReleaseAge`       | `1440` (24h) | Supply-chain lag window. Every version we pin is far older. |
| `blockExoticSubdeps`      | `true`       | No git/tarball subdependencies sneaking into the tree.      |
| `verifyDepsBeforeRun`     | `install`    | `pnpm build` cannot run against a stale `node_modules`.     |
| `optimisticRepeatInstall` | `true`       | Cheap.                                                      |

`allowBuilds` replaces `onlyBuiltDependencies` / `neverBuiltDependencies` / `ignoredBuiltDependencies`
/ `ignoreDepScripts`. Postinstall scripts are **deny-by-default**; each entry in `allowBuilds` is an
explicit, reviewed decision.

### Turborepo 2.10.12 — for `build` only, for now

`turbo run build` with `dependsOn: ["^build"]` and declared `outputs`. Lint, typecheck and test are
deliberately **root-level single-process** commands: with two packages there is nothing to
parallelise, and a cache key that is subtly wrong is worse than no cache. Turbo's scope expands in
3.9, where the fidelity harness genuinely needs `inputs` declarations (corpus dir + font dir +
container digest) to avoid stale cached PASSes on the test that defines the product.

### Vitest 4.1.11, **browser mode** for every core package

This is enforcement, not preference. Under jsdom, `node:fs` resolves and a Node API leak into a
core package is invisible until a user loads it in a browser. Under real Chromium it throws.

Vitest 4 moved the browser provider into a separate package and changed the config shape:

```ts
browser: { provider: playwright(), enabled: true, headless: true, instances: [{ browser: 'chromium' }] }
```

`test.projects` replaces the deprecated workspace file. Two projects: `core` (browser, Chromium) and
`tools` (Node) — the tools scripts are Node-side by definition and must not run in a browser.

### tsdown 0.22.14

ESM-only output, `dts: true`, `platform: 'browser'` for core packages. Its peer range already admits
TypeScript 5, 6 and 7, so it does not block the TS 7 migration. `publint` and `attw` run against the
packed tarball, not the source tree.

We are **not** using tsdown's `exports.devExports`. It writes dev exports pointing at `src/` and
moves the real ones to `publishConfig`, which pnpm honours but **npm does not** — and the release
path may have to be `npm publish` (see below). Instead, cross-package resolution during development
goes through `paths` in the root `tsconfig.json` and matching Vite aliases: tests and typecheck read
source, builds and publishes read `dist`.

### Release: Changesets, and publish via `npm`, not `pnpm`

npm trusted publishing (OIDC) is GA and, when used, emits SLSA Build Level 3 provenance
automatically — `--provenance` is no longer needed, and `NODE_AUTH_TOKEN` must **not** be set or the
CLI silently falls back to the legacy token path.

Caveat found during research: pnpm 11 replaced the npm CLI publish fallback with a native
implementation, and OIDC publishing that worked under pnpm 10 now fails with a 404
(pnpm/pnpm#11513). Until that is fixed, the release workflow installs and builds with pnpm but runs
the publish step through the npm CLI (>= 11.5.1). The release workflow is a stub until 3.10; this
note exists so the trap is not rediscovered then.

Trusted-publisher configurations created after 2026-05-20 must explicitly select at least one
allowed action when registered on npmjs.com.

## Consequences

- One extra TypeScript major to cross later, on a schedule we control, in exchange for keeping
  type-aware linting from day one.
- Core tests need a real Chromium (~150 MB via `pnpm exec playwright install chromium`). Accepted:
  that browser _is_ the specification for a browser-first renderer.
- The layer graph is enforced at the **manifest** level by `tools/layering`, because manifests are
  what ship. Import-level enforcement rides along in ESLint as a second net.

## What the first install actually taught us

Five things that only surfaced by running the thing, recorded so they are not rediscovered:

1. **`minimumReleaseAge` bit on day one, and correctly.** `turbo@2.10.12` was published on
   2026-08-25 at 18:38Z, inside the 24-hour window, and pnpm refused it. The fix is to pin
   `turbo@2.10.11`, not to lower the guard. This is what the setting is for; the first time it fires
   is the wrong moment to decide it is inconvenient.
2. **`baseUrl` is deprecated in TypeScript 6 and removed in 7.** `tsc` errors outright unless
   `ignoreDeprecations` is set. Since TypeScript 4.1, `paths` entries resolve relative to the
   tsconfig that declares them, so `baseUrl` is redundant as well as doomed — dropped, and the
   `paths` targets are written as `./packages/...`.
3. **typescript-eslint's project service needs somewhere to put loose config files.**
   `vitest.config.ts` and `packages/*/tsdown.config.ts` sit outside every auto-discovered tsconfig's
   `include`, so they fail to parse with "not found by the project service". They are listed in
   `allowDefaultProject` for linting and typechecked separately by `tsconfig.node.json` — which is
   deliberately _not_ named `tsconfig.json`, because the service auto-discovers that filename and
   would then report the same file as belonging to two projects.
4. **`attw` needs `--profile esm-only`.** Its default profile reports the node10 resolution failure
   and the CJS-resolves-to-ESM warning that follow inevitably from being ESM-only. `publint
--strict` passes unmodified; the profile flag keeps every other attw check while accepting the
   one consequence we chose.
5. **`verifyDepsBeforeRun: install` makes `pnpm <script>` unusable for negative testing.** Breaking
   a manifest to prove a guard fires means pnpm tries to install the broken manifest first. Invoke
   the checker as `node tools/repo/layering/check-layering.ts` in those cases.

## Deferred

- **pnpm 11.24.0 is available; we are pinned at 11.6.0** via the `packageManager` field. Worth
  taking, but as a deliberate bump with the lockfile regenerated, not as a drive-by.
- **A second linter.** `oxlint` + `oxlint-tsgolint` would be materially faster and works on
  TypeScript 7 today. Revisit as a pre-filter alongside ESLint, not as a replacement, once there is
  enough code for lint time to be noticeable.
