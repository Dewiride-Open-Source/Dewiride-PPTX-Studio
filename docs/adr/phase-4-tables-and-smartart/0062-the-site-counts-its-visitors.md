# 0062 — The site counts its visitors

Date: 2026-09-29
Status: **accepted** — every exported page carries the owner's page-view counter exactly once, the
browser walk loads it on every page it visits without counting a visit, and the site's two
privacy sentences say what is now true. One mutant, killed by both new assertions.

**Maintenance between 4.1 and 4.2, amending
[0055](../phase-3-text/0055-the-website-installs-what-it-documents.md)'s walk.** The owner asked
for the Dewiride analytics tag on the public site:

```text
<script defer src="https://analytics.dewiride.com/dw.js" data-site="01a0eb67-…"></script>
<noscript><img src="https://analytics.dewiride.com/collect/pixel.gif?site=01a0eb67-…" …></noscript>
```

## Where it goes

`website/src/app/layout.tsx`, the root layout, is the one component every route renders — the
landing page, the docs, the demos, the playground and the 404. The script goes in `<head>` as
written and the pixel is the first child of `<body>`. React renders both into the static HTML as
the snippet spells them (`defer=""`, `referrerPolicy`, `style="position:absolute"`), and the walk
below saw no console error from hydrating them, so `next/script` was not needed.

The tag is rendered only when `NODE_ENV` is `production`. `next build` is always production; `next
dev` is not, so a page opened while developing the site is never counted.

## What the site said, and says now

Two sentences promised the opposite and would have been false the moment this deployed:

- the footer, on every page: "No cookies, no analytics, no third-party requests";
- the FAQ's "Is my file uploaded?": "the site sets no cookies and loads nothing from a third party".

The owner chose to name the counter and make no claim about cookies, because what `dw.js` stores
is not something this repository controls or measures. The footer now reads "Page views are counted
by analytics.dewiride.com. A deck you open here never leaves your browser." The FAQ says the same
and that the counter is the one thing loaded from another origin. Every sentence about the deck
itself — the picker, the playground, the landing page, the demos index — was already about the deck
only, and stays.

## The walk

ADR 0055's `check-export.mjs` fails on every failed request and every console error, and it runs
on every pull request, every push to `main` and every deploy. Left alone it would have counted a
visit from a GitHub runner on each run and failed whenever the analytics server was slow or down.
And its first-load budget read every `<script src>` from `out/`, so the external URL would have
thrown `ENOENT`.

- **The counter is answered by the walk.** `page.route` fulfils every request to
  `analytics.dewiride.com` with an empty script. No CI run counts a visit or depends on that
  server; every other failed request still fails the check.
- **Every page is checked for the tag.** Each `out/**/index.html` must contain the exact tag once
  with the owner's site id: 47 pages.
- **The browser asked for it.** The walk counts the requests it answered and fails at zero: 7, one
  per page it opens.
- **The budget is this site's bundle.** An absolute `src` is not the site's JavaScript and is not
  counted. The four budgeted pages load 215, 235, 202 and 202 kB gzipped, as before.

The candidate gate, the canary and the release build the site but do not walk it, so no browser
there loads the counter.

## Verification

- Website `typecheck`, `lint`, and a build under `BASE_PATH=/Dewiride-PPTX-Studio`; `check-export`
  passed every step above on Windows.
- **The mutant.** With the tag switched off in the layout and the site rebuilt, `check-export`
  failed twice: "47 page(s) without the analytics tag exactly once" and "no page asked for the
  page-view counter".
