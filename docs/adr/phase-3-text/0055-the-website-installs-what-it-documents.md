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

| claim                                                 | measurement                                                                                                                                                                                                                                                                                |
| ----------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| first-load script, gzipped: `/`                       | 218 kB (budget 230)                                                                                                                                                                                                                                                                        |
| `/docs/packages/opc/`                                 | 239 kB (budget 250) — Fumadocs' shell and the MDX runtime over the same React                                                                                                                                                                                                              |
| `/demos/render-dom/`                                  | 202 kB (budget 230)                                                                                                                                                                                                                                                                        |
| `/playground/`                                        | 203 kB (budget 230)                                                                                                                                                                                                                                                                        |
| the largest chunks                                    | Shiki's Oniguruma wasm, 623 kB raw / 232 kB gzipped, and its grammars, 199 kB gzipped — loaded by a code block, never in a first load                                                                                                                                                      |
| the Worker                                            | 849-byte bootstrap + 4 chunks, 70 kB gzipped                                                                                                                                                                                                                                               |
| `out/`                                                | 35 MB, 893 files; `search.json` 758 kB; `rendered/` 144 SVGs for eleven decks                                                                                                                                                                                                              |
| `next build`, this machine                            | 16.8 s to compile from a clean cache, 39 s end to end warm including `prebuild`; the walk 12 s                                                                                                                                                                                             |
| `prerender/decks.mjs`                                 | 11 decks, 144 slides, 6133 ms at 1280 px                                                                                                                                                                                                                                                   |
| the candidate gate, this machine                      | 219 s end to end: pack, install, typecheck, lint, smoke, build                                                                                                                                                                                                                             |
| the candidate gate in CI, before                      | 59 s (run 34764116456, `the packages this commit would publish`, typecheck + smoke + build of the example)                                                                                                                                                                                 |
| the candidate gate in CI, after                       | 93 s (run 34779051444): lint, and a build that renders eleven decks and reads the reference                                                                                                                                                                                                |
| the site's build job in CI, on the pull request       | 130 s (run 34779051486): the registry answered every range in 5 s, install 34 s, typecheck 4, lint 10, smoke 3, build 27, Chromium 26, the walk 9                                                                                                                                          |
| the canary in CI, before                              | 42 s (run 34763792565)                                                                                                                                                                                                                                                                     |
| the site's dependencies                               | 21 runtime (12 of them `@pptx-studio/*`), 15 dev; no lockfile                                                                                                                                                                                                                              |
| the branch                                            | 215 files, +13232 −4043; 164 of them under `website/`                                                                                                                                                                                                                                      |
| Lighthouse 13.4.1, median of three, before the review | mobile `/` 76 · `/docs/packages/opc/` 71 · `/playground/` 56 for performance, 96–97 accessibility, 100 best practices, 100 SEO; desktop 100 / 92 / 85, 96 accessibility. What failed and what changed is under "What the review of the site found"; the numbers after are in the follow-up |
| publish to served, the first release after this       | 2 min 24 s: validate 0.1.2 published at 04:42:03Z, the `workflow_run` build started as the release finished, the registry answered every range on its first ask 22 s after the publish, the install needed no retry, and the site was live at 04:44:27Z (run 34807019459)                  |

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

## What CI said

Pull request #38. The `website` workflow's first run anywhere was on the pull request, run
34779051486: the registry served every one of the twelve ranges on the first ask, the install took
34 s with no retry, and the walk passed all thirteen assertions on Linux with the same first-load
numbers as this machine — 218/239/202/203 kB — and without the Windows prefetch tolerance, which
is what shows the tolerance is scoped right. The runner has no fonts: every deck rendered with
every typeface substituted, as the CLI page says it will. The candidate gate passed from the
branch's tarballs in 93 s.

`check` failed twice on the same test and nothing else: 3.11's border-probe raster walk in
`packages/render-svg/src/render.test.ts`, sixteen renders through the `<img>` path, took 18.2 s
and then 18.1 s against browser mode's 15 s default, where main's last run had it at 7.2 s.
The whole job was 1.4× slower than that run — the package build 8.7 → 12.3 s, the suite's
transform 16.9 → 25.1 s — so it was the runner, and the heaviest raster test crossed the cap
first. The walk now has 60 s; its assertions are unchanged.

## What the deploy said

Pull request #38 merged as `88028d2` on CI run 34779982487. The push to `main` touched
`website/**`, so the first deploy was run 34780332460 on the `push` trigger: the build job 140 s
(the registry answered every range in 7 s, install 36 s, build 29 s, the walk 9 s), the deploy job
9 s (`configure-pages` reported `/Dewiride-PPTX-Studio`, the assertion against `BASE_PATH`
passed, `deploy-pages` 5 s). GitHub Pages was enabled with `build_type: workflow` and the
repository's homepage set to the site by `gh api` before the merge; the `github-pages`
environment admits `main` only.

The live site, asked with curl:

| path                                         | answer                                                                            |
| -------------------------------------------- | --------------------------------------------------------------------------------- |
| `/`                                          | 200 `text/html`, 102 098 bytes                                                    |
| `/docs/`, `/docs/packages/opc/`              | 200, 97 956 and 310 933 bytes                                                     |
| `/demos/census/`, `/playground/`             | 200, 38 982 and 35 905 bytes                                                      |
| `/search.json`                               | 200 `application/json`, 757 738 bytes                                             |
| `/decks/b02-layouts.pptx`                    | 200 `application/vnd.openxmlformats-…`, 45 876 bytes — the corpus original's size |
| `/rendered/a46-hundred-slides/slide-001.svg` | 200 `image/svg+xml`, 2295 bytes                                                   |
| `/og.png`, `/sitemap.xml`, `/robots.txt`     | 200, 59 753 / 4745 / 105 bytes                                                    |
| `/no-such-page/`                             | 404, the site's own page, 23 109 bytes                                            |

The live site, driven in Chromium from this machine: `/playground/` reached network idle in
5.5 s and mounted its first `[data-shape]` at 5.6 s, the Census tab answered at 7.5 s — the
100-slide default deck, cold, over the network — with 68 responses and none at or above 400; the
Worker bootstrap `turbopack-worker-2ru9m5gbh1na6.js` and the deck both came from under
`/Dewiride-PPTX-Studio/`; on `/docs/packages/opc/` Ctrl+K found a page over the live index. Then
a deck the site does not ship, `corpus/authored/b05-chart.pptx` (authored in PowerPoint, with a
chart the site preserves and does not draw), through the picker's file input: opened and mounted
in 100 ms, Census and Validate answered, "Export unchanged" streamed every part and the download
was 66 130 bytes from a 67 962-byte source. `packages/cli/scripts/powerpoint-oracle.ps1` opened
that download in the real PowerPoint with `OpenAndRepair` off: `ok: true, repair: false`, two
slides, two shapes each.

This first deploy ran on a push whose ranges the registry already served; the registry wait
after a release, and Lighthouse, are measured under "What the review of the site found", below.

**What Dependabot found.** The new `website/` entry opened four pull requests within the hour;
the one that bumps TypeScript to 7 fails `npm install` with `ERESOLVE` on a peer range, which is
the red the design wants — and its run (34780499037) showed the install step reporting **success**
on that failure. The retry loop piped npm through `tee` and, with no `pipefail`, read `tee`'s exit
code; the step after it failed on a missing module instead. The loop now redirects npm to the log
and reads npm's own status, and `gate.test.ts` runs the step's script under bash with a fake `npm`:
one that exits 0 passes the step, one that prints `ERESOLVE` and exits 1 fails it at once with no
retry. With the pipe put back, that test is red.

## What the review of the site found

The site's first review, on the live page, found the editor unusable as one: a shape had to be
selected and then edited through a form on the right, and a click on a colour swatch produced a
box headed "Refused". Checked on the live page before anything changed: on the default deck's
first slide the second shape — `Deck title`, which declares `<a:noFill/>` — answered the red
swatch with "This shape declares no a:solidFill - its fill is inherited, from the placeholder
chain or the theme style matrix. The Inheritance tool says which." That refusal was mine, at the
demo level, on the theory that writing a local override onto a shape that inherits is a layout
decision rather than a gesture. PowerPoint writes exactly that override when a placeholder is
dragged or recoloured, and `SCOPE.md` names the target in one line: Canva's simplicity with
PowerPoint's inheritance model. The page had the 0.6 primitive — three XML edits that know their
inverse — behind a form, and called it an editor.

### What the editor is now

Direct manipulation on the slide, over five gestures, each an XML edit on the slide's own part:

| gesture                                                           | what is written                                                                                                                                                                                                                                                                                                                                           | as PowerPoint writes it                                                                                                                                                                                                  |
| ----------------------------------------------------------------- | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| drag, arrow keys                                                  | `a:off/@x @y` on the shape's own `a:xfrm`; a shape that inherits its transform is given one at the frame it resolved to, `insertInOrder` placing it first in `p:spPr`; a group's child moves in the group's units (`ext / chExt`); a graphic frame through its own `p:xfrm`                                                                               | a dragged placeholder gets a local `a:xfrm`; the layout keeps everything else                                                                                                                                            |
| colour, from eight swatches or the native picker                  | an `a:srgbClr/@val` when the shape's own fill is already one; otherwise an `a:solidFill` inserted where the schema puts it and whichever fill was there — `noFill`, `gradFill`, `pattFill`, a scheme colour — removed, an `a:alpha` carried over; a connector's colour is its `a:ln`'s                                                                    | a recoloured placeholder gets a local `a:solidFill`; transparency survives a colour change                                                                                                                               |
| text colour, from the same swatches                               | an `a:solidFill` in every run's, break's and field's `a:rPr` and in every `a:endParaRPr` the body has, set when it is already an `a:srgbClr`, otherwise inserted where the schema puts it in place of whatever fill was there; a run without an `a:rPr` is given one, first among its children; an empty paragraph without an `a:endParaRPr` is given one | a font colour over the whole text: every `a:rPr` gets the `a:solidFill`, `lang` and `b` untouched, the `a:endParaRPr` too, and a paragraph with runs and no `a:endParaRPr` is not given one — measured 2026-09-14, below |
| double-click, Enter, or the toolbar: type where the text is drawn | one entry per paragraph, split at soft breaks; an unchanged paragraph is not touched; a changed one keeps its `a:pPr` and its first run's `a:rPr` and the rest of its runs collapse into that one; a new paragraph copies the last one's properties; an `a:br` carries the run's properties                                                               | typing over a selection that spans runs keeps the first run's formatting                                                                                                                                                 |
| Delete                                                            | `removeChild` of the shape element; the inverse puts it back at the same index                                                                                                                                                                                                                                                                            | —                                                                                                                                                                                                                        |

A plan is a generator of batches, and each batch is applied before the next is planned, because
a later insertion has to see what an earlier removal did to the indices; `applyAll` rolls the
earlier batches back when a later one is refused, so a refused gesture leaves the tree as it found
it. A drag or a typing session is one undo step: the session opens a gesture, the edits' inverses
accumulate into it, and `end` pushes it. After every edit the touched part is serialised into
the store and that one sheet is parsed again with `parseSheet`, its `parent` and `theme`
carried from the sheet it replaces, so a keystroke re-reads one slide's XML and not the package.

The text box is a `contentEditable` in a `foreignObject` placed in the same transform as
render-svg's text layer — points, rotated with the shape — with the face, size, weight and colour
read off the first `<tspan>` the renderer emitted and the insets and anchor from the resolvers;
the drawn text beneath it is hidden while it is open. Paragraphs are `<div>`s and soft breaks
`<br>`s, which is what the browser's Enter and Shift+Enter write — except that under
`white-space: pre-wrap` Chromium writes Shift+Enter as a newline character in the text node,
which the first browser session caught as a literal newline inside `a:t` in the export; the
reader splits text nodes on newlines as well. Escape puts the text back through the gesture's
inverses.

What is selectable is what the slide owns: a shape the layout or master draws answers a click with
"comes from the layout; editing it there is phase 7", and when a layout shape carries the same
`cNvPr/@id` as one of the slide's own the click is settled by a point-in-frame test on the
slide's shape. render-svg draws nothing for a `p:graphicFrame`, so the stage draws every graphic
frame the slide owns as a labelled dashed box — "Chart 1 · chart, preserved and not drawn yet" —
that can be selected, moved and deleted, which is `SCOPE.md`'s frame-edit tier. A picture's
colour control is off with "A picture keeps its image"; a group's with "colour the shapes inside
it"; a frame's with "preserved as it is". The toolbar has four controls: Fill, Text, Edit,
Delete. The status line names the shape in every step: "Filled Deck title #E8453C", "Coloured
the text of Deck title #2FA869", "Undid: Deleted Accent panel".

The FAQ's "Why is move refused on some shapes?" and the guide's "two refusals worth keeping" taught
the same refusal and are rewritten to say what a gesture writes; the demo's name is "Editor";
the export walk now selects the title shape and colours it from the toolbar on every deploy.
The review's other finding was that the demos could not be understood; every demo's header and
its card on the index now open with one plain sentence on what the page shows and a "Try:" line
naming the first thing to do on it, and the blurbs no longer speak in the packages' own terms
("the tree with its byte offsets, an edit and its exact inverse, byte-identical re-emission" is
"the XML of one part, edited in place and written back without touching a byte you did not
change").

### What the second review found

The toolbar's one colour control coloured the fill, and a visitor who clicked it on the title
wanted the words: the screenshot came back with the title on an orange panel and the text still
white, and "no way to change text color". The control is now two — Fill and Text — with the same
swatches, each carrying the colour as drawn.

What PowerPoint writes for a font colour was asked in writing before the planner was: a
presentation authored over COM with a title placeholder, a text box of two paragraphs with a bold
run and a soft break, and an empty trailing paragraph, saved before and after
`TextRange.Font.Color.RGB` over the whole text. The slide after carries `<a:solidFill><a:srgbClr
val="E8453C"/></a:solidFill>` inside every `a:rPr` — the bold run's `b="1"` and every
`lang` kept, the bold run's earlier red replaced — inside the `a:br`'s `a:rPr`, and inside the
empty paragraph's `a:endParaRPr`; a paragraph that has runs and no `a:endParaRPr` is not given
one. That is the row in the table above, and the planner writes exactly it.

Driving the toolbar from the keyboard found a second defect: a picked swatch unmounts under the
pointer, focus fell to the document, and Ctrl+Z after a colour did nothing until the slide was
clicked again — the same after Delete from the toolbar and after a typing session closed. Focus
now goes back to the slide when a swatch is picked, a shape deleted or the text box closed, and a
key pressed on a toolbar control is that control's alone, so Enter on the Fill button opens its
swatches and not the text box.

Reading the toolbar's swatches back over the session found a third: every one was near black,
because the site formatted `Rgba` channels as bytes and paint's are 0..1, so `#FFFFFF` came out
as `#010101` — the fill's swatch had been wrong since the toolbar was written, and the text box's
fallback luminance test could never call a background light. The site now formats through paint's
own `toHexColor`, and the swatches read white, `#4472C4` for the accent panel and `#2FA869`
after the pick.

The selection outline carried a square at each corner, which is the universal promise of a resize
handle, and dragging one did nothing. The outline is now a plain rectangle: resize and rotate are
sub-phase 5.3's, measured against PowerPoint rather than written as an `a:ext` edit here, and a
control that is drawn is a control that works.

### What the editor found in the packages

**V028 mistook the extent of a transform for an extension.** The first placeholder dragged on
`b02-layouts` exported nothing: `ERR_VALIDATION_FAILED`, `V028 at
/p:sld/p:cSld/p:spTree/p:sp[1]/p:spPr/a:xfrm/a:ext - <a:ext> is not markup this part arrived
with`. The rule holds `mc:AlternateContent` and every extension opaque, by local name so that
vocabularies never read are covered too, and it matched any element spelled `ext` — the extent
inside `a:xfrm` is spelled `ext` as well, and nothing had written a new `a:xfrm` before.
An `ext` is now an extension only under an `extLst`; the new test writes a transform onto a
placeholder that inherited one and holds V028 silent, and against the old rule it fails. Pull
request #45, `@pptx-studio/validate` patch, released before this change to the site was merged,
because the site installs from npm and a deploy carrying the new editor with the old rule would
refuse that download.

### Verification

`npm test` in `website/` — `node --test` over Node 24's type stripping, so the planners are
tested with no runner installed — runs every gesture over every slide the site ships: 144 slides,
1181 shapes, 1088 text bodies, 1120 declared transforms and 61 inherited positions; 1090 `p:sp`
whose declared fill is 165 inherited, 451 `noFill`, 190 solid, 183 gradient, 89 pattern, 9 blip
and 3 group. Each test applies the plan, checks what it wrote — every touched element in schema
order by `outOfOrderChildren`, the new `a:xfrm` first in `p:spPr`, the old fill gone, the
transparency carried, the first run kept, an unchanged paragraph still clean — parses the sheet
again, applies the inverses and holds the part to its original bytes. Eleven tests, 1.9 s;
fifteen with the text colour, 8 s. The text colour's tests count what the decks' runs carry —
1,071 `a:rPr` with no fill of their own, 190 with a scheme colour, one with an RGB, 13 breaks, 91
fields, 179 empty paragraphs — and, because every run the site ships carries an `a:rPr`, make
the case without one by stripping them from a shape and holding what the planner writes back.
Five mutants of the planner: breaks skipped, the paragraph end skipped, a missing `a:rPr` not
written, an empty paragraph's `a:endParaRPr` not written — each failed a test; the fifth,
dropping the fast path that sets an existing `a:srgbClr` in place of inserting a new one, is
equivalent in what it writes and survived.

In Chromium over the served export, on the default deck: a click on `Deck title` shows its
toolbar; the red swatch turns its `<a:noFill/>` into `#E8453C`; a double-click opens the text
box with "Northwind Analytics", typing draws live, Enter makes a paragraph and Shift+Enter a soft
break, Ctrl+Enter commits, and Escape after further typing puts the text back; a 100 × 20 px drag
moves the subtitle 100 × 20 px; Delete removes the accent panel and Undo restores it;
Shift+Arrow nudges ten points; Download hands over 164 225 bytes with one part rewritten and 220
streamed untouched. Then the shapes the default deck lacks: `b02-layouts`'s title, which
inherits its position, dragged, retyped and coloured; `a43-kitchen-sink`'s chart frame selected,
moved 60 px, deleted and undone; `a06-lines`'s connector coloured through its line, `#44546A`
to `#2FA869`; `a44-transforms`'s group child dragged 50 × 30 px on screen, which is the drag
divided by the group's child scale; `b09-picture`'s picture moved with its colour control off.
The seven downloads opened in the real PowerPoint through `powerpoint-oracle.ps1` with
`OpenAndRepair` off: every one `ok: true, repair: false`, with its slide and shape counts.
The moved placeholder's download was made on the candidate gate's build — the site from this
commit's tarballs, the fixed rule inside — because the registry served the old one.

The text colour in Chromium over the export: the default deck's title from `#FFFFFF` to
`#2FA869` in every `<tspan>`, `b02-layouts`'s title and subtitle from the inherited `#000000`,
`a43-kitchen-sink`'s from `currentColor`; Ctrl+Z back and Ctrl+Y forward; a fill after the text
colour leaves the text's; and the three downloads opened in the real PowerPoint with no repair,
which then read the colour back over COM as `#2FA869`, type RGB, on every coloured shape and
`#FFFFFF`, type scheme, on the untouched subtitle.

### What Lighthouse said

Measured before the review's changes, so that this is the baseline: 13.4.1, from the scratchpad
against the live site with Playwright's Chromium, three runs per route and form factor, the
median reported. Mobile is Lighthouse's default — a Moto G Power with a 4× CPU slowdown and
simulated 3G.

| route                 | mobile perf · a11y · best · SEO             | desktop              |
| --------------------- | ------------------------------------------- | -------------------- |
| `/`                   | 76 · 96 · 100 · 100 (TBT 1.08 s, LCP 2.2 s) | 100 · 96 · 100 · 100 |
| `/docs/packages/opc/` | 71 · 96 · 100 · 100 (TBT 1.95 s)            | 92 · 96 · 100 · 100  |
| `/playground/`        | 56 · 97 · 100 · 100 (LCP 4.8 s, TBT 1.4 s)  | 85 · 96 · 100 · 100  |

Three defects behind the numbers, each with a change:

- **Contrast.** `--fg-faint` was 3.29:1 on the dark surface and 3.75:1 on the light one, and
  white on the accent `#3b8fd4` was 3.46:1 — small text needs 4.5:1. The tokens are now
  `#5f6a7a` (light, 4.70:1 on the lightest surface it sits on) and `#7a8594` (dark, 4.89:1),
  and what sits on a filled accent is a token of its own, `--on-accent`: white in light (6.20:1),
  the canvas black in dark (5.75:1), which the Fumadocs primary foreground reads too.
- **An unnamed image.** Fumadocs' built-in GitHub link is an `<svg role="img">` with no name;
  the link is declared in the site's own navigation with a labelled mark.
- **Main-thread blocking on the simulated phone.** The landing page opens the real 100-slide deck
  on the main thread; the docs page hydrated its embedded demo eagerly, with the deck parse, on a
  page whose reader has not scrolled to it. A docs-page embed now mounts when it is within a
  screen of the viewport; the landing page's cost is the deck, and a Worker-side parse is
  sub-phase 12.1's.

The numbers after are measured on the deploy that carries these changes, below.

### What the deploy said after the review

Pull request #46 merged as `84c94a3`; its deploy, run 34807495414, built and walked in 129 s
and deployed in 10 s. The live page, driven the same way as the export: the full session on the
default deck, then `b02-layouts`'s title — the placeholder whose download V028 had refused —
dragged, retyped, coloured and downloaded; both downloads opened in the real PowerPoint with no
repair, 100 slides with four shapes on the first and eleven with two.

The registry wait after a release, measured on the first `workflow_run` deploy: the release that
published validate 0.1.2 finished at 04:42:13Z, the website run started at the same second, the
registry step answered all twelve ranges in 5 s — validate 0.1.2 served 22 s after its publish, on
the first ask — the install took 28 s with no retry, and the site was live 2 min 24 s after the
publish. The 15-minute ceiling and the 20-second poll never came into it.

Lighthouse on that deploy, same matrix as the baseline:

| route                 | mobile perf · a11y · best · SEO                  | desktop                   |
| --------------------- | ------------------------------------------------ | ------------------------- |
| `/`                   | 76 · **100** · 100 · 100 (TBT 1.10 s, LCP 2.0 s) | 100 · **100** · 100 · 100 |
| `/docs/packages/opc/` | **77** · 96 · 100 · 100 (TBT 1.02 s, from 1.95)  | **99** · 96 · 100 · 100   |
| `/playground/`        | 55 · 97 · 100 · 100 (LCP 4.9 s, TBT 2.1 s)       | 88 · 97 · 100 · 100       |

The landing page reached 100 for accessibility on both form factors and the docs page's blocking
time halved. Three things were still under 4.5:1, each a token: Shiki's `github-dark` comment
grey, Fumadocs' default, at 3.98:1 on the site's dark surface — the dark theme is now
`github-dark-default`, whose comment grey clears 6:1 on the same surface, declared once in
`src/site/code-themes.ts` for the MDX pipeline (`source.config.ts`) and the demos' code
component alike; the purple `--info` badge colour at 4.2:1 on its own tint in dark, now
`#bf7fd9` (6.4:1); and `--fg-faint` at 4.37:1 on the accent tint a selected row carries, now
`#808b9a` (4.73:1). Against the local export with those three, Lighthouse's accessibility
category is 100 on `/`, `/docs/packages/opc/`, `/playground/` and `/demos/render-svg/`,
and on the deploy that carried them (run 34809496384) the live site scores 100 for accessibility,
best practices and SEO on all four. One docs-desktop run scored a layout shift of 0.103
from the web font swapping in, where the baseline's three runs had 0.001; that is the font's
`swap`, not this change, and it is open below. The playground's mobile score is the deck: the
100-slide default opens on the main thread, and that is 12.1's.

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
- **The editor is direct manipulation over five gestures**, not the plan's three edits behind a
  panel, and a shape that inherits a property is given a local one rather than refused. Above.
- **Graphic frames are drawn by the editor's stage**, as labelled boxes, because render-svg draws
  nothing for them and the frame-edit tier needs something to select.

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
5. **The web font swaps in.** Fontsource's faces load with `font-display: swap`, and one
   Lighthouse run measured a 0.103 layout shift from it on a docs page; a `size-adjust`ed
   fallback face would hold the lines still until IBM Plex arrives.
6. **The text box approximates the line model.** Its line height is 1.2 × the size, which is
   PowerPoint's single spacing for the common faces and not the renderer's measured lines; a
   caret in the drawn text itself is phase 6's.
7. **The `pull_request` build has no Chromium cache.** `npx playwright install --with-deps
chromium` runs on every build; a cache keyed on the Playwright version would remove it.
