# 0057 — Transitive ranges float

Date: 2026-09-27
Status: **accepted** — fumadocs fixed the stringifier six minutes after the report, and the
website takes 16.15.15 rather than pinning around it. The candidate gate now carries the consumer's
`overrides`, which it had been dropping. 7/7 mutants killed.

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

## The fix is upstream's

Reported as [fuma-nama/fumadocs#3604](https://github.com/fuma-nama/fumadocs/issues/3604) at
15:03Z, with the cause and the one-line fix. The timeline:

- **15:09Z:** fumadocs fixed it in `7b83e88b`, `Object.assign(wrapped, handler)` in `modHandler`.
- **15:17Z:** they published fumadocs-core and fumadocs-ui 16.15.15.
- **15:15Z:** they closed the issue as completed, without a comment.

16.15.15 carries SLSA provenance from `fuma-nama/fumadocs/.github/workflows/release.yml` at
`1bc5374f`, the version commit directly after the fix.

The website moves to 16.15.15 and pins nothing. Until the release existed, the working answer was
`"overrides": { "mdast-util-to-markdown": "2.1.2" }`. Measured under the lockfile-less install all
three consumers use, the build passed with it and failed with 26 errors without it. Two other
answers were rejected:

- **Switch `remark-structure` off.** `remarkStructureOptions: false` removes the stringifier and the
  search index with it. `/search.json` is built through `createFromSource`, which throws on a page with
  no structured data, and `deploy/check-export.mjs` searches the dialog. Refused: it trades a feature
  for a green build.
- **A `stringify` hook in `source.config.ts`** that short-circuits strong and emphasis before
  `modHandler` calls the real handler. Refused as a shortcut:
  - it depends on fumadocs internals;
  - it drops the markers from the indexed text;
  - an empty `strong` still recurses.

## What 0055 got wrong

0055 declined a lockfile for good reasons, and they still hold:

- the release rewrites the ranges before it publishes;
- the gate and the canary need a fresh resolution to mean anything.

It then reasoned that "every off-scope dependency in the manifest is an exact pin, so what varies
between runs is the twelve packages and their transitive `fflate`." That is false. An exact pin at the
top says nothing about the ranges beneath it: fumadocs-core 16.15.x declares `mdast-util-to-markdown:
^2.1.2`, and a patch release two levels down reached all three consumers within hours.

`website/package.json#overrides` is where a transitive pin lives when one is needed:

- the canary rewrites only the `@pptx-studio/*` entries of that file and keeps the rest;
- `npm install --no-package-lock` still reads `overrides` from the root manifest;
- `website.yml` installs the manifest as written.

## The candidate gate carries the consumer's overrides

The candidate gate is the one consumer that did not honour `overrides`. `scratchManifest` replaced
the website's manifest with one holding only the candidate pins, so any pin the website declared
would have been dropped, and the gate would have built a tree the website does not.
`candidate.ts` claims that "the only thing that changes between this gate, the canary and the deploy
is where the packages came from", and dropped overrides break that claim.

`scratchManifest` now takes the consumer whole (`dependencies`, `scripts`, `overrides`) and emits
the consumer's overrides ahead of the candidates. It refuses four things rather than carrying them:

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
| `fumadocs-core`, `-ui`     | 16.15.9 | 16.15.15 | the stringifier fix; supersedes Dependabot #55 (16.15.12); no source change needed                                         |
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
  fumadocs-core 16.15.15 with `mdast-util-to-markdown` 2.1.3 at all 11 edges.
- **The website, locally:**
  - typecheck clean, lint clean;
  - `npm test` 15/15;
  - smoke all green;
  - build green, with no stack overflow;
  - `check-export` green, with script gzipped against budget: `/` 215 of 230 kB,
    `/docs/packages/opc/` 234 of 250, `/demos/render-dom/` 202 of 230, `/playground/` 201 of 230.
- **The same install on 16.15.14:** `next build` exits 1 with "Turbopack build failed with 26
  errors", all 26 the stack overflow. With the 2.1.2 override it built.
- **Tools:** 88/88 in `tools/release`. Mutants, 7/7 killed:
  - the consumer's overrides winning over the pins;
  - each of the four refusals removed;
  - the consumer's overrides dropped;
  - `website.yml` not running `npm test`.

## Open questions

1. This tree is still installed, unpinned, in the release job that holds `id-token: write`
   (`release.yml:123`). A transitive range that floats to something hostile rather than something
   broken would run beside the publish identity. It stays open until the release job stops installing
   the website.
