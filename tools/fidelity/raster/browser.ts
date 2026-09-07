/**
 * One Chromium, serving the working tree to itself.
 *
 * The page is fulfilled from `page.route` rather than from a server, so the
 * harness needs no port, no background process and nothing to clean up if it
 * throws - which is what lets CI run it as one command. The packages are loaded
 * as the built ESM through an import map, so what is rasterised is what an
 * `npm install` would get, not a bundle made for this test. ADR 0035.
 */

import { readFileSync } from 'node:fs';
import { join, normalize, sep } from 'node:path';

import { chromium, type Browser, type Page } from 'playwright';

import { REPO_ROOT } from '../../repo/root.ts';
import { FidelityError } from '../errors.ts';

/**
 * An origin that resolves nowhere, so a missed route cannot reach the network.
 *
 * `https`, because the raster digest is taken with `crypto.subtle` and that is
 * only defined in a secure context.
 */
const ORIGIN = 'https://fidelity.invalid';

/**
 * The only trees the page may read.
 *
 * `packages/opc/node_modules` is in there because `fflate` is the one bare
 * specifier the built ESM still carries and the import map has to point
 * somewhere real.
 */
const SERVED: readonly string[] = ['packages/', 'corpus/'];

const HTML_TYPE = 'text/html; charset=utf-8';

const TYPES: Readonly<Record<string, string>> = {
  '.js': 'text/javascript; charset=utf-8',
  '.mjs': 'text/javascript; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
  '.html': HTML_TYPE,
  '.png': 'image/png',
  '.pptx': 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
};

/**
 * Flags that pin what Skia does to text.
 *
 * Measured on win-x64 to change nothing: the raster digest is identical with
 * and without them, over four decks and five fresh browsers. They are kept as
 * insurance for the Linux runner, where FreeType hinting and LCD filtering are
 * real and version-dependent, and the ADR records that this is unverified
 * there rather than claiming they were shown to matter.
 */
export const PINNED_ARGS: readonly string[] = [
  '--force-color-profile=srgb',
  '--disable-lcd-text',
  '--disable-font-subpixel-positioning',
  '--font-render-hinting=none',
  '--disable-remote-fonts',
];

const IMPORT_MAP = {
  imports: {
    fflate: '/packages/opc/node_modules/fflate/esm/browser.js',
    '@pptx-studio/opc': '/packages/opc/dist/index.js',
    '@pptx-studio/xml': '/packages/xml/dist/index.js',
    '@pptx-studio/geometry': '/packages/geometry/dist/index.js',
    '@pptx-studio/paint': '/packages/paint/dist/index.js',
    '@pptx-studio/text': '/packages/text/dist/index.js',
    '@pptx-studio/model': '/packages/model/dist/index.js',
    '@pptx-studio/render-svg': '/packages/render-svg/dist/index.js',
  },
};

const PAGE = `<!doctype html><meta charset="utf-8"><title>fidelity</title>
<script type="importmap">${JSON.stringify(IMPORT_MAP)}</script>
<script type="module">
import * as opc from '@pptx-studio/opc';
import * as model from '@pptx-studio/model';
import * as rsvg from '@pptx-studio/render-svg';
import * as text from '@pptx-studio/text';
globalThis.pptx = { opc, model, rsvg, text };
</script>`;

export interface Harness {
  readonly page: Page;
  close(): Promise<void>;
}

/** Whether a URL path stays inside the trees this is willing to read. */
function resolveServed(pathname: string): string | null {
  const relative = pathname.replace(/^\/+/, '');
  if (!SERVED.some((root) => relative.startsWith(root))) return null;
  const full = normalize(join(REPO_ROOT, relative));
  if (!full.startsWith(REPO_ROOT + sep)) return null;
  return full;
}

export async function openHarness(): Promise<Harness> {
  const browser: Browser = await chromium.launch({ args: [...PINNED_ARGS] });
  const page = await browser.newPage({ deviceScaleFactor: 1 });

  await page.route(`${ORIGIN}/**`, async (route) => {
    const pathname = new URL(route.request().url()).pathname;
    if (pathname === '/harness.html') {
      await route.fulfill({ status: 200, contentType: HTML_TYPE, body: PAGE });
      return;
    }
    const file = resolveServed(pathname);
    if (file === null) {
      await route.fulfill({ status: 403, body: 'outside the served trees' });
      return;
    }
    const dot = pathname.lastIndexOf('.');
    const type = TYPES[pathname.slice(dot)] ?? 'application/octet-stream';
    try {
      await route.fulfill({ status: 200, contentType: type, body: readFileSync(file) });
    } catch {
      await route.fulfill({ status: 404, body: 'no such file' });
    }
  });

  await page.goto(`${ORIGIN}/harness.html`);
  await page.waitForFunction(() => 'pptx' in globalThis);

  return {
    page,
    close: async (): Promise<void> => {
      await browser.close();
    },
  };
}

/** The URL a repository-relative path is served at inside the page. */
export function servedUrl(relative: string): string {
  if (!SERVED.some((root) => relative.startsWith(root))) {
    throw new FidelityError('FID_ENV_UNKNOWN', `${relative} is not in a served tree`, relative);
  }
  return `/${relative}`;
}
