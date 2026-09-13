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

import { chromium } from 'playwright';

import { serveOut } from './serve-out.mjs';

const BASE_PATH = process.env['BASE_PATH'] ?? '';
const TIMEOUT = 30_000;

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
  await page.goto(at('/census/'), { waitUntil: 'networkidle', timeout: TIMEOUT });
  await page.getByText('namespaces', { exact: true }).waitFor({ timeout: TIMEOUT });
  ok('/census/ reached the Worker and got a census back');

  // A deck fetched through publicUrl, drawn by the string renderer.
  await page.goto(at('/slides/'), { waitUntil: 'networkidle', timeout: TIMEOUT });
  await page.locator('[data-shape]').first().waitFor({ timeout: TIMEOUT });
  ok('/slides/ fetched a sample deck and drew a shape');

  // The CLI's output, rendered ahead of time.
  await page.goto(at('/thumbnails/'), { waitUntil: 'networkidle', timeout: TIMEOUT });
  await page.locator('img[src*="/rendered/"]').first().waitFor({ timeout: TIMEOUT });
  ok('/thumbnails/ shows a slide the CLI rendered at build time');

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
