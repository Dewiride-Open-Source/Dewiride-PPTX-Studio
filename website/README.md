# website

The public site: documentation for every published `@pptx-studio/*` package and a live demo of
each, at
[dewiride-open-source.github.io/Dewiride-PPTX-Studio](https://dewiride-open-source.github.io/Dewiride-PPTX-Studio/).

```sh
cd website
npm install
npm run dev      # http://localhost:3000
```

Node ≥ 24.11. Use **npm here, not pnpm**: this directory is deliberately outside
`pnpm-workspace.yaml`, so `npm install` fetches the real tarballs from the public registry instead
of linking the workspace. That exclusion is the whole point - a site wired to the local packages
would prove nothing about what was published, and it is also the third consumer of the tarballs
beside the release candidate gate and the daily canary.

## What is here

| Directory        | Purpose                                                                                       |
| ---------------- | --------------------------------------------------------------------------------------------- |
| `src/app/`       | the routes: the landing page, `/docs`, `/demos/<package>`, `/playground`, the search index    |
| `src/demos/`     | one directory per package, `demo.tsx` the entry; the same component embeds in the docs        |
| `src/deck/`      | the one deck every demo reads, the parsed document, and the Worker for census/validate/export |
| `src/design/`    | the primitives every demo is built from; the tokens are in `src/app/globals.css`              |
| `src/docs/`      | the Fumadocs wiring and the site's own MDX components                                         |
| `src/reference/` | the API reference, read from `src/reference/generated/` at build                              |
| `src/status/`    | `plan.json`, written by `tools/repo/plan-status.ts` from `docs/plan/phases.json`              |
| `content/docs/`  | the documentation, in MDX                                                                     |
| `snippets/`      | the docs' code, type-checked by `npm run typecheck` against the installed packages            |
| `prerender/`     | what the build renders ahead of time: the CLI over the sample decks, and the API reference    |
| `deploy/`        | what a deploy proves before it uploads: the export served under the base path, in Chromium    |
| `public/decks/`  | byte-identical copies of corpus decks, with `manifest.json` and `pnpm corpus` holding them    |

## A static export

`npm run build` writes a static site to `out/` (`output: 'export'`), under the path in `BASE_PATH`
when one is set - GitHub Pages serves a project site under the repository's name. `prebuild` runs
first: `prerender/decks.mjs` renders every sample deck through `@pptx-studio/cli` into
`public/rendered/`, and `prerender/reference.mjs` reads each installed package's `dist/index.d.ts`
into `src/reference/generated/`. Both are generated, never committed.

`node deploy/check-export.mjs` serves `out/` the way Pages will and walks it in Chromium: the
Worker, a fetched deck, a mounted slide, the pre-rendered CLI output, the search index and the
first-load script budgets all have to hold under the base path before anything is uploaded.

**A machine with no fonts substitutes everything.** `renderDeck` leaves `systemFonts` at its default
of true, so it indexes the build machine's real faces; a runner that ships none draws every typeface
in a fallback. `render.json` carries the report either way - that is `FaceUse`, and the CLI page
shows it.

## Scripts

```sh
npm run dev           # the app, unprefixed
npm run build         # prebuild, then the static export into out/
npm run typecheck     # tsc --noEmit, snippets included
npm run lint          # eslint, with the site's own config
npm run smoke         # every sample deck through six packages, in Node
npm run check-export  # serve out/ and walk it in Chromium
```

## The decks

Eleven, in [`public/decks/`](./public/decks/), copied byte for byte from the project's corpus. All
are **CC0-1.0 and self-authored**; `public/decks/manifest.json` records each one's digest and
`pnpm corpus` refuses a copy that differs from its original. The `b*` decks were authored in real
PowerPoint; the `a*` decks are generated probes. `a46-hundred-slides.pptx` is the default, because
it is the only one that makes the strip and the zoom credible.
