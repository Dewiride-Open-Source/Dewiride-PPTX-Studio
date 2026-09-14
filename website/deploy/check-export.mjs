/**
 * Walk the static export in a real browser, served the way Pages serves it.
 *
 * ```
 * BASE_PATH=/Dewiride-PPTX-Studio node deploy/check-export.mjs
 * ```
 *
 * A build that succeeds proves nothing about a base path, a Worker chunk or a
 * fetched deck; only a page that runs does. Every failed request and every
 * console error fails the check, before anything is uploaded.
 */

import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { gzipSync } from 'node:zlib';

import { chromium } from 'playwright';

import { serveOut } from './serve-out.mjs';

const BASE_PATH = process.env['BASE_PATH'] ?? '';
const TIMEOUT = 30_000;
const OUT = join(import.meta.dirname, '..', 'out');

/** Gzipped kilobytes of first-load script per page, set just above what the build measured. */
const BUDGET_KB = {
  '/': 230,
  '/docs/packages/opc/': 250,
  '/demos/render-dom/': 230,
  '/playground/': 230,
};

/** The gzipped bytes of every module script a page loads; the nomodule polyfill never reaches a modern browser. */
function firstLoadKb(path) {
  const html = readFileSync(join(OUT, path, 'index.html'), 'utf8');
  const scripts = [...html.matchAll(/<script[^>]*src="([^"]+.js)"[^>]*>/g)]
    .filter((match) => !/noModule/.test(match[0]))
    .map((match) => match[1].replace(BASE_PATH, ''));
  let bytes = 0;
  for (const script of new Set(scripts)) bytes += gzipSync(readFileSync(join(OUT, script))).length;
  return bytes / 1024;
}

/** @type {string[]} */
const failures = [];
/** @param {string} what */
const fail = (what) => {
  failures.push(what);
  console.log(`FAIL ${what}`);
};
/** @param {string} what */
const ok = (what) => console.log(`ok   ${what}`);

// A build made on Windows writes the segment-cache prefetch files into nested
// directories, so their requests answer 404 (vercel/next.js#92339). The deploy
// builds on Linux, where the same answer is a real failure.
const SEGMENT_PREFETCH = /\/__next\.[^/]*__PAGE__\.txt/;
/** @param {string} url */
const knownWindowsDefect = (url) => process.platform === 'win32' && SEGMENT_PREFETCH.test(url);

const { origin, close } = await serveOut();
const browser = await chromium.launch();
const page = await browser.newPage();
page.on('console', (message) => {
  if (message.type() !== 'error') return;
  if (knownWindowsDefect(message.location().url)) return;
  fail(`console: ${message.text()}`);
});
page.on('pageerror', (error) => fail(`page error: ${error.message}`));
page.on('response', (response) => {
  if (response.status() < 400 || knownWindowsDefect(response.url())) return;
  fail(`${String(response.status())} ${response.url()}`);
});
page.on('requestfailed', (request) => {
  // A prefetch the browser cancelled on navigation is not a failed fetch.
  const why = request.failure()?.errorText ?? '';
  if (why === 'net::ERR_ABORTED') return;
  fail(`request failed: ${request.url()} ${why}`);
});

/** @param {string} path */
const at = (path) => `${origin}${BASE_PATH}${path}`;

try {
  await page.goto(at('/'), { waitUntil: 'networkidle', timeout: TIMEOUT });
  const title = await page.title();
  if (title.includes('PPTX Studio')) ok(`/ renders: "${title}"`);
  else fail(`/ title is "${title}"`);

  // The Worker: a chunk emitted by the bundler, fetched under the base path,
  // that loads the census package and answers.
  await page.goto(at('/demos/census/'), { waitUntil: 'networkidle', timeout: TIMEOUT });
  await page.getByText('namespaces', { exact: true }).waitFor({ timeout: TIMEOUT });
  ok('/demos/census/ reached the Worker and got a census back');

  // A deck fetched through publicUrl, drawn by the string renderer, and the
  // editor over it: a click gets a toolbar, a swatch changes the drawn fill.
  await page.goto(at('/demos/render-svg/'), { waitUntil: 'networkidle', timeout: TIMEOUT });
  const deckTitle = page.locator('.stage-surface [data-shape="11"]').first();
  await deckTitle.waitFor({ state: 'attached', timeout: TIMEOUT });
  ok('/demos/render-svg/ fetched a sample deck and drew a shape');
  const box = await deckTitle.boundingBox();
  if (box === null) throw new Error('the title shape has no box to click');
  await page.mouse.click(box.x + box.width / 2, box.y + box.height / 2);
  await page.getByRole('toolbar', { name: /Edit Deck title/ }).waitFor({ timeout: TIMEOUT });
  await page.getByRole('button', { name: 'Colour' }).click();
  await page.getByRole('button', { name: /Colour Deck title #E8453C/ }).click();
  await deckTitle
    .locator('path[fill="#E8453C"]')
    .first()
    .waitFor({ state: 'attached', timeout: TIMEOUT });
  ok('/demos/render-svg/ selected a shape and coloured it from the toolbar');

  // The live renderer: a slide mounted by render-dom, and the playground's bench over it.
  await page.goto(at('/demos/render-dom/'), { waitUntil: 'networkidle', timeout: TIMEOUT });
  await page.locator('[data-shape]').first().waitFor({ timeout: TIMEOUT });
  ok('/demos/render-dom/ mounted a live slide');
  await page.goto(at('/playground/'), { waitUntil: 'networkidle', timeout: TIMEOUT });
  await page.locator('[data-shape]').first().waitFor({ timeout: TIMEOUT });
  await page.getByRole('tab', { name: 'Census' }).click();
  await page.getByText('namespaces', { exact: true }).waitFor({ timeout: TIMEOUT });
  ok('/playground/ mounted the viewer and ran a census under a tab');

  // The CLI's output, rendered ahead of time.
  await page.goto(at('/demos/cli/'), { waitUntil: 'networkidle', timeout: TIMEOUT });
  await page.locator('img[src*="/rendered/"]').first().waitFor({ timeout: TIMEOUT });
  ok('/demos/cli/ shows a slide the CLI rendered at build time');

  // The docs shell, and the search index it exported.
  await page.goto(at('/docs/'), { waitUntil: 'networkidle', timeout: TIMEOUT });
  await page.locator('aside, nav').first().waitFor({ timeout: TIMEOUT });
  const index = await page.request.get(at('/search.json'));
  if (index.status() === 200) ok('/search.json is served');
  else fail(`/search.json answered ${String(index.status())}`);
  await page.keyboard.press('Control+k');
  await page.getByPlaceholder('Search').fill('started');
  await page.locator('[role="dialog"] button[aria-selected]').first().waitFor({ timeout: TIMEOUT });
  ok('/docs/ has a sidebar, and the search dialog finds a page');

  for (const [path, budget] of Object.entries(BUDGET_KB)) {
    const kb = firstLoadKb(path);
    if (kb <= budget)
      ok(`${path} loads ${kb.toFixed(0)} kB of script gzipped (budget ${String(budget)})`);
    else
      fail(
        `${path} loads ${kb.toFixed(0)} kB of script gzipped, over the ${String(budget)} kB budget`,
      );
  }

  const missing = await page.request.get(at('/no-such-page/'));
  if (missing.status() === 404 && (await missing.text()).includes('<html'))
    ok('an unknown path gets the 404 page');
  else fail(`an unknown path answered ${String(missing.status())}`);
} catch (cause) {
  fail(cause instanceof Error ? cause.message : String(cause));
} finally {
  await browser.close();
  close();
}

if (failures.length > 0) {
  console.log(`\n${String(failures.length)} failure(s)`);
  process.exitCode = 1;
} else {
  console.log('\nthe export runs');
}
