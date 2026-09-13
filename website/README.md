# website

A Next.js app built on the **published** `@pptx-studio/*` packages — one tool per package.

```sh
cd website
npm install
npm run dev      # http://localhost:3000
```

Node ≥ 24.11. Use **npm here, not pnpm**: this directory is deliberately outside
`pnpm-workspace.yaml`, so `npm install` fetches the real tarballs from the public registry instead
of linking the workspace. That exclusion is the whole point — a sample wired to the local packages
would prove nothing about what was published.

## Why it exists

Eleven packages went to npm at 0.1.0 and nothing had ever installed them. `apps/studio` is
`private: true` with every dependency on `workspace:^`, and `pnpm pkg:qa` (`publint` + `attw`)
inspects the _shape_ of a tarball without ever installing or executing one. So the release path had
the same defect it was built to prevent: code nobody had run.

This app closes that. `npm run smoke` renders every shipped deck through the installed packages and
asserts the slide count, the SVG shape and a clean round trip. It runs twice in CI: the release gate
installs the tarballs this commit would publish into a copy of this directory, and the daily canary
installs `latest` from the registry with no lockfile. The ranges in `package.json` are the caret of
each workspace version, written by `pnpm version-packages` and refused by the gate when they are
not, so `npm install` here gets what the last release proved.

## The tools

| Route          | Package      | What it shows                                                              |
| -------------- | ------------ | -------------------------------------------------------------------------- |
| `/slides`      | `render-svg` | The deck drawn, a shape selected, the 2.11 debug overlay — and three edits |
| `/inheritance` | `model`      | Where each resolved property came from, and what `undefined` means         |
| `/package`     | `opc`        | Parts, content types, and each part's own relationships                    |
| `/markup`      | `xml`        | The tree with its byte offsets, and byte-identical re-emission, live       |
| `/census`      | `census`     | Parts, features and namespaces, counted in a Web Worker                    |
| `/validate`    | `validate`   | All 29 rules, with part URI and XPath on anything that fired               |
| `/roundtrip`   | `writer`     | Export, the preservation check, and a file you can open in PowerPoint      |
| `/geometry`    | `geometry`   | The 187 presets, their guides and their draggable adjust handles           |
| `/paint`       | `paint`      | Colour transforms in document order, the two-stop ramp, 54 pattern tiles   |
| `/text`        | `text`       | Measurement, line breaking, the autofit ladder, 41 autonumber schemes      |
| `/thumbnails`  | `cli`        | Slides rendered in Node when the site was built, with no LibreOffice       |

Every page runs entirely in the tab. `/thumbnails` shows what `prerender/decks.mjs` made of the
sample decks when the site was built, because the CLI is the one Node package; a deck you open is
not sent anywhere.

## Editing is three XML edits, and not a pretence

`/slides` can move a shape, recolour it and retype its first run. Each is an `XmlEdit` that hands
back its own inverse, which is the undo stack — so undo restores the _document_, including the
parts nothing here has ever parsed.

It is **not** Phase 5's command bus. There is no coalescing, no invalidation set, no snapping, no
resize solver and no multi-select. Two refusals are deliberate and worth reading when you hit them:

- Moving a shape with no `a:xfrm` of its own is refused, because writing one is a Change Layout
  decision rather than a drag. `/inheritance` shows why that shape has no geometry.
- Recolouring a shape with no `a:solidFill` of its own is refused for the same reason — the fill is
  coming down the placeholder chain or out of the theme's style matrix.

## A static export

`npm run build` writes a static site to `out/` (`output: 'export'`), under the path in `BASE_PATH`
when one is set — GitHub Pages serves a project site under the repository's name. `prebuild` runs
the CLI over `public/decks/` first, so `public/rendered/` is generated and never committed.
`node deploy/check-export.mjs` serves `out/` the way Pages will and walks it in Chromium: the
Worker, a fetched deck and the pre-rendered slides all have to answer under the base path before
anything is uploaded.

**A machine with no fonts substitutes everything.** `renderDeck` leaves `systemFonts` at its default
of true, so it indexes the build machine's real faces; a runner that ships none draws every typeface
in a fallback. `render.json` carries the report either way — that is `FaceUse`, and the page shows
it.

## The decks

Six, in [`public/decks/`](./public/decks/), copied from the project's corpus. All are **CC0-1.0 and
self-authored**. The `b*` ones were authored in real PowerPoint; the `a*` ones are generated
feature probes. `b02-layouts.pptx` contains the Blank layout, whose slide really is empty — an empty
`<svg>` is the correct output for it.

## Scripts

```sh
npm run dev         # the app
npm run build       # the static export, into out/ (prebuild renders the decks first)
npm run typecheck   # tsc --noEmit
npm run smoke       # render every shipped deck through the installed packages
```
