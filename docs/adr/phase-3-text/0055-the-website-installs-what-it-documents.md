# 0055 — The website installs what it documents

Date: 2026-09-14
Status: **accepted** — the site builds from the registry's tarballs under the path GitHub Pages
serves it at, walks in Chromium before it uploads, and deploys after a release rather than inside
one; two mutants killed, one killed once the smoke test was made to count, one shown
equivalent

**Sub-phase 3.12.** Before phase 4, a public site: the product in plain language, a live demo of
every package, documentation for each, the owner's own decks droppable onto it, on GitHub Pages
for nothing, and rebuilt after every release from what npm serves — never from the workspace, so
the site is at most one release behind the tree and never ahead of what a visitor can install.
`examples/nextjs-studio` was the seed: eleven routes, dark only, one server route, and — the part
worth keeping — the consumer the release candidate gate (ADR 0046) and the daily canary (ADR 0049)
already ran. This is the record of turning that example into the site without losing the contract.

Code: `website/`, `tools/release/candidate/{candidate,website,consumer}.ts`,
`tools/repo/plan-status.ts`, `tools/corpus/manifest/website.ts`, `.github/workflows/website.yml`,
`.github/workflows/canary.yml`, `tools/release/gate.test.ts`. Nothing under `packages/` changed
and nothing is published, so there is no changeset.

---

## The example becomes the site

`git mv examples/nextjs-studio website`, at the top level. ADR 0043's reason for the example's
location still holds: outside `packages/*` and `apps/*` it is off the pnpm link farm, and
`pnpm-workspace.yaml` has no exclusion syntax, so `npm install` in it fetches real tarballs. An
`examples/` directory with one child was a rename wearing a folder, and the one child is no longer
an example of anything — it is the site.

What the site is made of, by directory, each named for what it holds (hard rule 4; `src/design/`, at ten
files, is the largest):

| directory                  | what                                                                                                                                                              |
| -------------------------- | ----------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `src/app/`                 | the routes: `/`, `/docs/[[...slug]]`, `/demos/[demo]`, `/playground`, `/search.json`, the 404, sitemap and robots                                                 |
| `src/landing/`             | the seven blocks of the front page: hero, showcase, principles, packages, install, measured, honesty                                                              |
| `src/demos/<package>/`     | one directory per package, `demo.tsx` the entry; `registry.ts` is the list, `loaded.tsx` the lazy map, `embed.tsx` the frame the docs and the demo pages share    |
| `src/playground/bench.tsx` | the render-dom viewer with every demo's panel as a tab beside it                                                                                                  |
| `src/deck/`                | the one deck every page reads, its parsed document, the picker, the sample list, and `worker/` for census, validation and export                                  |
| `src/design/`              | ten primitives the demos are built from; the tokens are in `src/app/globals.css`                                                                                  |
| `src/docs/`                | the Fumadocs wiring: the source, the MDX components (`ApiTable`, `Demo`, `ErrorCodes`, `PackageHeader`, `Roadmap`), the static search dialog and its lazy wrapper |
| `src/reference/`           | the API tables, read at build from `generated/` (gitignored)                                                                                                      |
| `src/status/`              | `plan.json`, the projection of `docs/plan/phases.json` that `tools/repo/plan-status.ts` writes and `pnpm references` checks                                       |
| `src/site/`                | facts about the site: `publicUrl`, the twelve packages, the installed versions, the navigation, `SiteError`                                                       |
| `src/shell/`               | the brand, the footer, the version badge; the header, theme toggle and mobile drawer are Fumadocs' `HomeLayout`                                                   |
| `content/docs/`            | 30 MDX pages: getting started, concepts, twelve package pages, ten guides, the CLI, the errors, the status, the FAQ                                               |
| `snippets/<package>/`      | the docs' code as thirteen `.ts` files with `//#region example`, compiled by `npm run typecheck` and pulled into MDX with `<include>`                             |
| `prerender/`               | what the build renders ahead of time: `decks.mjs` (the CLI over the sample decks) and `reference.mjs` (each installed `dist/index.d.ts` into JSON)                |
| `deploy/`                  | `serve-out.mjs` mounts `out/` at the base path; `check-export.mjs` walks it in Chromium                                                                           |
| `public/decks/`            | eleven corpus decks, byte for byte, with a manifest of digest, licence, source and slide count                                                                    |

Turbopack is the bundler (Next 16's default), Tailwind 4 the styles, Fumadocs 16 the docs chrome,
and the palette is the renderer's own debug-overlay colours from 2.11 so the site and the tool it
documents look like one thing. Fonts are IBM Plex through `@fontsource`, below.

## Three consumers, one manifest

`website/package.json` names a registry range for each of the twelve packages, and
`pnpm version-packages` rewrites those ranges to the versions it is about to publish
(`tools/release/candidate/website.ts`, formerly `example.ts`). Three things install from that
manifest and run the same four scripts, and the only thing that differs between them is where the
packages come from:

| consumer                                | when                                  | the packages come from                                                                   | runs                                                              |
| --------------------------------------- | ------------------------------------- | ---------------------------------------------------------------------------------------- | ----------------------------------------------------------------- |
| the candidate gate (`ci.yml`, ADR 0046) | every CI run, required by the release | this commit's tarballs, `pnpm pack`ed into a scratch copy with the scope's registry dead | `typecheck`, `lint`, `smoke`, `build`                             |
| the canary (`canary.yml`, ADR 0049)     | daily, and on dispatch                | whatever `latest` is, the ranges rewritten to it first                                   | `typecheck`, `lint`, `smoke`, `build`, then the attestation check |
| the site (`website.yml`, this ADR)      | after a release, on a push, on a PR   | the ranges in the manifest, from the registry, after waiting for it to serve them        | `typecheck`, `lint`, `smoke`, `build`, `check-export`, deploy     |

The gate proves the tarballs work before they exist on npm; the canary proves the registry still
serves what a consumer can build; the site proves the same with a base path, in a browser, and
then serves the result. `lint` is new to the gate and the canary: the site's `eslint.config.mjs`
carries React 19's `react-hooks` rules and `typescript-eslint`'s type-checked set, and a demo
that sets state in an effect is a bug a visitor sees as a flicker, not a build failure.

`candidate.ts` copies `website/` to the scratch consumer with a filter for what a local install
or build leaves behind (`node_modules`, `package-lock.json`, `.next`, `out`, `public/rendered`,
`src/reference/generated`), so the gate's build renders the decks and reads the reference from
the candidate tarballs, not from a stale local run. On Windows the gate starts `npm` and `pnpm`
through `cmd.exe /d /s /c`, because both are `.cmd` shims that `spawn` cannot start without a
shell and `shell: true` prints a deprecation for the arguments it would have to quote.

## Why the release does not run the deploy

`release.yml` pushes the version commit with `GITHUB_TOKEN`, and GitHub starts no workflow for a
push made with that token. The evidence is on `main`: commit `9625cca` (`chore: release`, the
fourth release) has zero workflow runs —
`gh api repos/…/actions/runs?head_sha=9625cca` answers `total_count: 0` — so a `push` trigger on
`website/**` never sees the commit that changes the ranges. Five ways to run the deploy after a
release were on the table:

| way                                          | what it costs                                                                                                                                                                                                                                               |
| -------------------------------------------- | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| a job in `release.yml`                       | `pages: write` and `id-token: write` on the workflow that publishes; a deploy that installs from the registry seconds after the publish, which ADR 0054 measured as the moment the registry is least likely to serve it; and a red deploy failing a release |
| `push` to `main` on `website/**`             | never fires for the version commit — `9625cca`                                                                                                                                                                                                              |
| `repository_dispatch` from the release       | the same rule: an event raised with `GITHUB_TOKEN` starts no run; a PAT would fix it and is a second secret to rotate                                                                                                                                       |
| a schedule                                   | a site hours behind by the clock rather than minutes behind by the event                                                                                                                                                                                    |
| **`workflow_run` on `Release`, `completed`** | fires on the release's completion whatever token it used; asks the release for nothing; the deploy gates on `conclusion == 'success'` so a release that did not finish deploys nothing                                                                      |

`website.yml` also runs on a push to `main` that touches `website/**` (a docs change deploys
without a release), on a pull request that touches it (build and walk, no upload), and on
dispatch. `tools/release/gate.test.ts` holds the shape: the site's job names are never in the
release's `REQUIRED` list and never in `ci.yml`; `workflow_run` names `Release`; `release.yml`
never names `website.yml`; the deploy job has no checkout, nothing installed, and exactly
`pages: write` and `id-token: write`; nothing in the workflow names `NODE_AUTH_TOKEN` or
`publishers.ts`. The last two are ADR 0048's isolation for two more permissions: the OIDC token
the deploy holds cannot publish, because npm matches a publisher to `release.yml` by file name and
answers 404 to any other workflow (ADR 0049), but a checkout beside it would still be a place for
third-party code to run with it.

## Waiting for the registry

ADR 0054 measured the lag between a publish and the registry serving it: 38 seconds for `cli`
0.3.1, minutes for `model` 0.1.4, and the packument leading the tarball. A deploy that installs
the moment the release completes is the canary's first red run every time. So the build job asks
first: for each `@pptx-studio/*` range in the manifest, `npm view <spec> version --prefer-online`
every 20 seconds until it answers, 15 minutes in all, `--prefer-online` so the loop is not
answered by its own packument cache. Then `npm install --prefer-online`, retried every 20 seconds
for 5 minutes on exactly `ETARGET` or `E404` — the class "npm does not have that version yet" —
and failing at once on anything else, so a real install error is not five minutes of retries. The
numbers for the first release after this lands go in the follow-up below.

## The lockfile

None, and `npm install` rather than `npm ci`, as ADR 0043 and ADR 0046 decided for the example:
the release rewrites the ranges before it publishes, so a lockfile could never be refreshed inside
the release; the gate and the canary both need a fresh resolution to mean anything; and every
off-scope dependency in the manifest is an exact pin, so what varies between runs is the twelve
packages and their transitive `fflate`. Dependabot watches `website/` weekly, grouped, and ignores
`@pptx-studio/*` because the release owns those ranges.

## Fonts

`@fontsource-variable/ibm-plex-sans` and `@fontsource/ibm-plex-mono`, imported in
`globals.css`, self-hosted from the export. Not `next/font/google`: it fetches from Google on
every build — the gate, the canary and the deploy — and when the fetch fails it **silently
substitutes a system face**, which is precisely the failure this repository built a font guard to
refuse (ADR 0033). No font binaries in the tree either; the OFL stays at a package boundary.

## The CLI at build time

The example's one server piece was `POST /api/thumbnails` over `@pptx-studio/cli`, and a static
export has no server. `prerender/decks.mjs` runs `renderDeck` and the `inspect` and `validate`
verbs over `public/decks/*` in `prebuild`, into `public/rendered/` (gitignored): every slide as
SVG, `inspect.txt`, and `render.json` with the `FaceUse` report. The CLI page shows exactly what
the deploy runner rendered, labelled with its platform and Node version and the exact command,
and the same call runs from this commit's tarballs inside the candidate gate — a proof the POST
route only ever made at run time, on whoever's machine. Eleven decks, 144 slides, 6.1 s on this
machine at 1280 px.

**A machine with no fonts substitutes everything.** `renderDeck` leaves `systemFonts` at its
default, so it indexes the build machine's real faces; a runner that ships none draws every
typeface in a fallback and `render.json` says so. The CLI page shows the report either way.

**Eleven decks, not twelve.** The plan named `a08-bullets` for the sample set; its second slide
uses `arabic1Minus`, which ADR 0054 records as `TEXT_AUTONUMBER_UNMEASURED` — the renderer
refuses a scheme it has no measurement for rather than guess — so `renderDeck` throws and the
smoke test goes red. The deck stays in the corpus and out of the site until 3.5's open question
closes.

## The reference, read from the installed types

Hand-written API tables drift: the example's render-dom page documented a `resize()` that
ADR 0054 deleted. `prerender/reference.mjs` opens each installed package's `dist/index.d.ts` with
the TypeScript compiler API, follows the `//#region src/<file>` markers tsdown leaves, sibling
chunk re-exports and cross-package re-exports, and writes one JSON per package with each export's
kind, signature, first JSDoc sentence and — for the error class — its code union. The site reads
those on the server at build (`readFileSync`, never an import), so nothing generated reaches the
client bundle or the type check. What the installed versions export:

| package      | exports | source files | error codes |
| ------------ | ------- | ------------ | ----------- |
| `opc`        | 77      | 12           | 31          |
| `xml`        | 129     | 15           | 13          |
| `geometry`   | 80      | 11           | 5           |
| `paint`      | 173     | 19           | 21          |
| `text`       | 162     | 18           | 18          |
| `model`      | 144     | 18           | 26          |
| `render-svg` | 92      | 16           | 7           |
| `render-dom` | 20      | 4            | 2           |
| `validate`   | 40      | 6            | 3           |
| `writer`     | 53      | 8            | 3           |
| `cli`        | 28      | 7            | 7           |

The snippets are the same idea for the prose: thirteen `.ts` files under `snippets/`, in the
site's `tsconfig`, so a renamed export fails `npm run typecheck` in the gate before the publish and
in the canary against `latest`; the MDX pulls them in with fumadocs-mdx's `<include>` and the
`#region` selector.

## The base path

GitHub Pages serves a project site at `/<repository>/`. Next prefixes what `next/link` and its
own chunk loader emit with `basePath`, and nothing else — a `fetch('/decks/x.pptx')` goes to the
origin's root and answers 404. So `BASE_PATH` is one environment variable: `''` locally, in the
gate and in the canary; `/${{ github.event.repository.name }}` in the deploy. `next.config.ts`
validates it (`''` or `/<segment>`), feeds it to `basePath` and to `NEXT_PUBLIC_BASE_PATH`, and
`publicUrl(path)` in `src/site/base-path.ts` is the one way to a file under `public/`; it throws
`SITE_PUBLIC_PATH` on a path that does not start with `/`. The deploy job asks `configure-pages`
where Pages serves the repository and refuses to deploy a build made for any other path.

## What the walk found

`deploy/check-export.mjs` serves `out/` under the base path with `node:http` and drives it in
Chromium: `/` renders; `/demos/census/` reaches the Worker and a census comes back;
`/demos/render-svg/` fetches a deck and draws a `[data-shape]`; `/demos/render-dom/` mounts a live
slide; `/playground/` mounts the viewer and runs a census under a tab; `/demos/cli/` shows an SVG
the CLI rendered; `/docs/` has a sidebar, `/search.json` is 200 and Ctrl+K finds "started"; the
first-load script of four routes is under budget; an unknown path gets the 404 page. Every 4xx,
every failed request and every console error fails it. Three things the walk found that a passing
build did not:

- **The Worker is a bootstrap and four chunks.** `new Worker(new URL('./worker.ts',
import.meta.url), { type: 'module' })` under Turbopack becomes an 849-byte bootstrap
  (`turbopack-worker-*.js`) that reads its chunk list from its own URL and loads four chunks —
  `opc`, `xml`, `census`, `validate`, `writer` and the protocol, 70 kB gzipped in all — under the
  base path. Turbopack also copies the raw `worker.ts` into `_next/static/media/` and references
  it from the module that resolved the URL; nothing fetches it. The walk's census assertion is
  what proves the chunks load where Pages serves them.
- **The search client shipped on every page.** Fumadocs' `RootProvider` takes the dialog
  component, and importing it statically put Orama and the dialog into the landing page's first
  load: 374 kB gzipped. `search-lazy.tsx` wraps it in `next/dynamic` and the landing page is 218.
  The same pass found that importing `@pptx-studio/<name>/package.json` for the version badge is
  not tree-shaken by Turbopack, so `installedVersion` reads the file on the server at build.
- **A build made on Windows 404s its own prefetches.** Next's segment cache writes
  `__next.__PAGE__.txt` into nested directories on Windows (vercel/next.js#92339), so the browser's
  prefetch of each link answers 404. The walk tolerates exactly that URL shape on `win32` only;
  on the Linux runner the same answer is a real failure.

## Measurements

| claim                                           | measurement                                                                                                                           |
| ----------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------- |
| first-load script, gzipped: `/`                 | 218 kB (budget 230)                                                                                                                   |
| `/docs/packages/opc/`                           | 239 kB (budget 250) — Fumadocs' shell and the MDX runtime over the same React                                                         |
| `/demos/render-dom/`                            | 202 kB (budget 230)                                                                                                                   |
| `/playground/`                                  | 203 kB (budget 230)                                                                                                                   |
| the largest chunks                              | Shiki's Oniguruma wasm, 623 kB raw / 232 kB gzipped, and its grammars, 199 kB gzipped — loaded by a code block, never in a first load |
| the Worker                                      | 849-byte bootstrap + 4 chunks, 70 kB gzipped                                                                                          |
| `out/`                                          | 35 MB, 893 files; `search.json` 758 kB; `rendered/` 144 SVGs for eleven decks                                                         |
| `next build`, this machine                      | 16.8 s to compile from a clean cache, 39 s end to end warm including `prebuild`; the walk 12 s                                        |
| `prerender/decks.mjs`                           | 11 decks, 144 slides, 6133 ms at 1280 px                                                                                              |
| the candidate gate, this machine                | 219 s end to end: pack, install, typecheck, lint, smoke, build                                                                        |
| the candidate gate in CI, before                | 59 s (run 34764116456, `the packages this commit would publish`, typecheck + smoke + build of the example)                            |
| the candidate gate in CI, after                 | measured on the pull request, in the follow-up                                                                                        |
| the canary in CI, before                        | 42 s (run 34763792565)                                                                                                                |
| the site's dependencies                         | 21 runtime (12 of them `@pptx-studio/*`), 15 dev; no lockfile                                                                         |
| the branch                                      | 215 files, +13232 −4043; 164 of them under `website/`                                                                                 |
| Lighthouse on `/`, a docs page, `/playground/`  | measured against the live site, in the follow-up                                                                                      |
| publish to served, the first release after this | measured by the registry wait step's log, in the follow-up                                                                            |

## Verification

- **Every commit:** `pnpm structure && pnpm references && pnpm corpus && pnpm format:check &&
pnpm lint && pnpm typecheck`, `pnpm exec vitest run --project tools` (981 tests, 37 files), and
  in `website/`: `npm run typecheck && npm run lint && npm run smoke && npm run build` with
  `BASE_PATH` unset and set to `/Dewiride-PPTX-Studio`, then `npm run check-export` against the
  prefixed build — thirteen assertions green on the final tree.
- **The gate, locally:** `node tools/release/candidate/candidate.ts tmp/candidate` packs twelve,
  installs them against a dead registry, and typecheck, lint, smoke and build pass from the
  tarballs in 219 s. It must be given a long path: Turbopack refuses a project under an 8.3 short
  name (`JAGDIS~1`) with "leaves the filesystem root", so the session scratchpad cannot host it and
  CI's `$RUNNER_TEMP` can.
- **Mutants**, each applied to a copy-aside file and the check re-run:
  1. `publicUrl` returns the path unprefixed → `check-export` red on `404 /decks/a46-hundred-slides.pptx`
     and `404 /rendered/a46-hundred-slides/slide-001.svg`, then the census never arrives. **Killed.**
  2. the search client's `from` unprefixed → `/search.json` still 200 by direct request, then
     `404 /search.json` from the dialog and no result for "started". **Killed.**
  3. `revalidate = false` removed from the index route → the export still writes `/search.json`
     as static, from a clean cache. **Equivalent**: under `output: 'export'` either
     `dynamic = 'force-static'` or `revalidate` satisfies the route check, and removing both
     fails the build with "not configured on route /search.json". The redundant line is
     deleted. The first run of this mutant passed for the wrong reason — Turbopack's incremental
     build reused the page-data collection — so a build mutant here starts from `rm -rf .next`.
  4. the smoke test's deck filter narrowed to `.pptx` → the `.pptm` deck silently skipped, ten of
     eleven decks smoked, exit 0. **Survived** the first version of the smoke test. It now reads
     `manifest.json` as its list, holds the directory to it and each deck's slide count to the
     manifest's; the same mutant then fails with `11 named, 10 present`. **Killed.**
- **`tools/release/gate.test.ts`** holds the deploy's shape, above, and the canary
  test no longer names a job by its old literal: it asserts that no canary job name is a CI job
  name.
- **`pnpm corpus`** rule `C020-website-copy`: each file under `website/public/decks/` is a
  byte-identical copy of the corpus deck its manifest entry names, and every deck present is
  claimed.
- **`pnpm references`** fails when `website/src/status/plan.json` is not what
  `tools/repo/plan-status.ts` projects from `docs/plan/phases.json`.
- **`pnpm check`** green on the commit that claims this sub-phase.

## What the deploy said

The first run is the merge of this pull request. Its run id, the registry wait, the curls of `/`,
`/docs/packages/opc/`, `/decks/b02-layouts.pptx`, the Worker chunk and `/search.json`, a deck
dropped on `/playground/` and its export opened in PowerPoint, and Lighthouse on three routes go
here in the follow-up, `docs(3.12): what the deploy said`, as ADR 0054 did for its releases.

## Deviations from the plan

- **Eleven sample decks, not twelve**: `a08-bullets`, above.
- **The first-load budgets are 230/250/230/230 kB**, not the plan's 180/220. The plan's numbers
  were guesses; the measurement is 218/239/202/203 and the budgets sit just above it so a
  regression of a chunk's worth fails.
- **`revalidate = false` is gone**, and the plan's mutation for it was wrong: removing it does not
  fail the export. Recorded as equivalent, above.
- **The smoke test counts against the manifest**; the plan's fourth mutant showed that it did not.
- **`concepts` is one page**, not a directory of six; the eight sections read better as one
  scroll with a table of contents than as six pages of two paragraphs.
- **Ten design primitives, not eighteen**, and three shell files, not six: Fumadocs' `HomeLayout`
  supplies the header, the theme toggle and the mobile drawer, and a primitive with one caller was
  its caller.
- **`eslint-config-next` is not used.** It crashes under ESLint 10 (`scopeManager.addGlobals`),
  so the config takes `@next/eslint-plugin-next` and `eslint-plugin-react-hooks` directly.
- **The gate spawns through `cmd.exe` on Windows** and uses `spawnSync` with a status check
  rather than `execFileSync`, whose types do not carry `windowsVerbatimArguments`.
- **The OG image is a checked-in PNG**, generated once from the icon; Next's
  `opengraph-image` route exported it without an extension.

## Open questions

1. **The npm cache without a lock.** `setup-node`'s cache is keyed on a lockfile and there is
   none, so every deploy resolves and downloads afresh. A cache keyed on the manifest's digest
   would save the download and not the resolution, which is the part that has to be fresh.
2. **A custom domain.** `BASE_PATH` would become `''` and `configure-pages` would report it; the
   deploy's assertion is written so that this is a one-line change to the workflow's `env`.
3. **A hundred slides in the strip is the site's ceiling.** The viewer's strip mounts what
   `IntersectionObserver` sees and no more, but a visitor's 300-slide deck still parses on the
   main thread and the stage mounts one slide at a time; virtualization of the strip and a
   Worker-side parse are sub-phase 12.1's.
4. **What 12.6 still owes.** This is the docs site as far as phase 3's packages go. Gate 12 asks
   for 200 slides, first paint under 3 s and a published fidelity score; 12.6 owes those numbers,
   and the pages for phases 4 to 11's features as each lands.
5. **Lighthouse, the first deploy and the registry wait** are in the follow-up, not here.
6. **The `pull_request` build has no Chromium cache.** `npx playwright install --with-deps
chromium` runs on every build; a cache keyed on the Playwright version would remove it.
