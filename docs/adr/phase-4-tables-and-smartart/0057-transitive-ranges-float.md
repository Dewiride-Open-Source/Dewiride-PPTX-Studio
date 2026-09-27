# 0057 — Transitive ranges float

Date: 2026-09-27
Status: **accepted** — the website pins `mdast-util-to-markdown` at 2.1.2 until fumadocs keeps a
handler's `attention` on its wrapper. The pin is the whole difference: under the same lockfile-less
install the build passes with it and fails 26 times without it. 7/7 mutants killed.

**Maintenance after 4.1 · supersedes the claim in
[0055](../phase-3-text/0055-the-website-installs-what-it-documents.md) § The lockfile that what varies
between runs is the twelve packages and their transitive `fflate`.** The canary went red on
2026-09-27 (run 36317786259, issue #57) with nobody having touched the repository. It was the first
of three consumers to find out. `website.yml` and the candidate gate would have followed on their
next run, and the release requires the candidate gate.

## What broke

`mdast-util-to-markdown` 2.1.3 was published at 09:31Z. The canary installed it at 12:04Z. `next
build` then failed on the 14 pages under `website/content/docs` that contain bold or italic text,
and passed on the 16 that contain neither:

```
RangeError: Maximum call stack size exceeded
  at strong (mdast-util-to-markdown/lib/handle/strong.js:29:23)
  at fumadocs-core/dist/mdx-plugins/stringifier.js:65:21
  at phrasing (mdast-util-to-markdown/lib/util/container-phrasing.js:559:27)
  at containerPhrasing (mdast-util-to-markdown/lib/util/container-phrasing.js:237:20)
  at strong (…)
```

2.1.3 moved attention into `containerPhrasing`. `strong` and `emphasis` now call
`state.containerPhrasing({ type: 'root', children: [node] })`. `containerPhrasing` reads
`state.handle.handlers[child.type].attention`, and when it finds one it serialises the attention
itself instead of calling the handler.

fumadocs-core's `remark-structure` builds the search index through `defaultStringifier`. Its
`_custom` handler replaces every entry of `state.handlers` with `modHandler(handler)`
(`packages/core/src/mdx-plugins/stringifier.ts:116`, `:181`). The returned function carries neither
`attention` nor `peek`. So `containerPhrasing` finds no `attention` and calls the wrapper; the wrapper
calls `strong`; `strong` wraps the node in a root and calls `containerPhrasing` again. This repeats
until the stack overflows.

The wrapper is the same on fumadocs' `dev` and `main` and in 16.15.14, the newest release, which
predates 2.1.3. Reported upstream with the fix, copying `peek` and `attention` onto the wrapper:
[fuma-nama/fumadocs#3604](https://github.com/fuma-nama/fumadocs/issues/3604).

## What 0055 got wrong

0055 declined a lockfile for good reasons, and they still hold:

- the release rewrites the ranges before it publishes;
- the gate and the canary need a fresh resolution to mean anything.

It then reasoned that "every off-scope dependency in the manifest is an exact pin, so what varies
between runs is the twelve packages and their transitive `fflate`." That is false. An exact pin at the
top says nothing about the ranges beneath it: fumadocs-core 16.15.x declares `mdast-util-to-markdown:
^2.1.2`, and a patch release two levels down reached all three consumers within hours.

`website/package.json#overrides` is where a transitive pin lives. The canary rewrites only the
`@pptx-studio/*` entries of that file and keeps the rest; `npm install --no-package-lock` still reads
`overrides` from the root manifest; `website.yml` installs the manifest as written. The candidate
gate did not carry it (below).

## The pin, and the two alternatives it beat

- **Switch `remark-structure` off.** `remarkStructureOptions: false` removes the stringifier and the
  search index with it. `src/app/search.json/route.ts` builds `/search.json` through
  `createFromSource`, which throws on a page with no structured data, and `deploy/check-export.mjs`
  searches the dialog. Refused: it trades a feature for a green build.
- **A `stringify` hook in `source.config.ts`** that short-circuits strong and emphasis before
  `modHandler` calls the real handler. Refused as a shortcut:
  - it depends on fumadocs internals;
  - it drops the markers from the indexed text;
  - an empty `strong` still recurses.
- **`"overrides": { "mdast-util-to-markdown": "2.1.2" }`.** 2.1.2 and 2.1.3 are the only versions
  every dependent in the tree accepts (fumadocs-core `^2.1.2`, ten others `^2.0.0`). An override
  rather than a direct dependency, because the website does not import it and npm refuses an
  override that disagrees with a direct dependency.

**Removal trigger:** a fumadocs-core release whose `modHandler` keeps `attention` on the wrapper.
Nothing will signal it, because the canary obeys the pin, so it is carried in `docs/plan/phases.json`
and checked whenever Dependabot's `docs` group moves fumadocs.

## The candidate gate carries the consumer's overrides

`scratchManifest` replaced the website's manifest with one that held only the candidate pins, so
the gate would have installed 2.1.3 whatever the website said. It now takes the consumer whole
(`dependencies`, `scripts`, `overrides`) and emits the consumer's overrides ahead of the candidates.
It refuses four things rather than carrying them:

- an override inside `@pptx-studio/*`, which would redirect a candidate;
- a dependency inside the scope, which would displace a pin;
- an override that disagrees with a direct dependency, which npm would refuse at install time;
- a nested override, which the `Record<string, string>` shape cannot carry.

## The website's own tests

`npm test` in `website/` (15 tests, `node --test`, against `@pptx-studio/opc` and `model`) ran in no
workflow. It now runs in `website.yml` after lint, and in the candidate gate against the candidate
tarballs. `tools/release/gate.test.ts` asserts the first.

## The website's dependencies

| package                    | from    | to       | why                                                                                                                        |
| -------------------------- | ------- | -------- | -------------------------------------------------------------------------------------------------------------------------- |
| `next`                     | 16.3.5  | 16.3.6   | GHSA-vcvr-r3jv-pc5j / CVE-2026-94545, critical, `next/og`; the site imports no `next/og`, but 16.3.5 is in range           |
| `@next/eslint-plugin-next` | 16.3.5  | 16.3.6   | in step with `next`                                                                                                        |
| `fumadocs-core`, `-ui`     | 16.15.9 | 16.15.14 | supersedes Dependabot #55 (16.15.12); no source change needed                                                              |
| `fumadocs-mdx`             | 15.4.0  | 15.4.5   | as above                                                                                                                   |
| `lucide-react`             | 1.45.0  | 1.48.0   | fumadocs-ui ≥ 16.15.12 depends on `^1.47.0`; one copy, not two. `Trash2` is an alias of `Trash` and is imported as `Trash` |
| `eslint`                   | 10.9.1  | 10.11.0  | no rule or API this site uses changed                                                                                      |
| `typescript-eslint`        | 8.68.0  | 8.70.1   | `recommended-type-checked` is the same rule list                                                                           |
| `@types/node`              | 26.5.1  | 26.6.3   | additive                                                                                                                   |

`typescript` stays at 6.0.3: [0001](../phase-0-foundation/0001-toolchain.md)'s trigger has not
fired, since typescript-eslint 8.70.1 and its canary 8.70.2-alpha.10 still peer `<6.1.0`, and
`typescript@7.0.2` exports no compiler API for `prerender/reference.mjs` to call. Dependabot #42 is
closed on that ground. `playwright` stays at 1.62.1 until it moves with the root, because it moves
Chromium.

## Verification

- **Fresh install, as the three consumers do it** (`npm install --no-package-lock`):
  `mdast-util-to-markdown@2.1.2 overridden` at every edge, and `lucide-react@1.48.0` deduped under
  fumadocs-ui.
- **The website, locally:**
  - typecheck clean, lint clean;
  - `npm test` 15/15;
  - smoke all green;
  - build green;
  - `check-export` green, with script gzipped against budget: `/` 218 of 230 kB,
    `/docs/packages/opc/` 240 of 250, `/demos/render-dom/` 203 of 230, `/playground/` 203 of 230.
- **The pin, mutated away:** the same fresh install resolves 2.1.3, and `next build` exits 1 with
  "Turbopack build failed with 26 errors", 26 of them the stack overflow. Restored, it builds.
- **Tools:** 88/88 in `tools/release`. Mutants, 7/7 killed:
  - the consumer's overrides winning over the pins;
  - each of the four refusals removed;
  - the consumer's overrides dropped;
  - `website.yml` not running `npm test`.

## Open questions

1. The pin has no expiry of its own. It comes out by hand when fumadocs#3604 ships in a release.
2. This tree is still installed, unpinned, in the release job that holds `id-token: write`
   (`release.yml:123`). A transitive range that floats to something hostile rather than something
   broken would run beside the publish identity. It stays open until the release job stops installing
   the website.
