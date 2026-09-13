/**
 * Serve `out/` the way GitHub Pages will: under `BASE_PATH`, `dir/` from
 * `dir/index.html`, and `404.html` for anything else.
 *
 * ```
 * node deploy/serve-out.mjs            # prints the origin, serves until killed
 * ```
 */

import { createServer } from 'node:http';
import { existsSync, readFileSync, statSync } from 'node:fs';
import { extname, join, normalize } from 'node:path';

const OUT = join(import.meta.dirname, '..', 'out');
const BASE_PATH = process.env['BASE_PATH'] ?? '';

const TYPES = {
  '.html': 'text/html; charset=utf-8',
  '.js': 'text/javascript; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
  '.svg': 'image/svg+xml',
  '.png': 'image/png',
  '.ico': 'image/x-icon',
  '.txt': 'text/plain; charset=utf-8',
  '.woff2': 'font/woff2',
  '.woff': 'font/woff',
  '.xml': 'application/xml',
  '.pptx': 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
  '.pptm': 'application/vnd.ms-powerpoint.presentation.macroEnabled.12',
};

/** @param {string} pathname */
function fileFor(pathname) {
  if (BASE_PATH !== '' && !pathname.startsWith(`${BASE_PATH}/`) && pathname !== BASE_PATH) {
    return null;
  }
  const under = pathname.slice(BASE_PATH.length) || '/';
  const local = normalize(join(OUT, decodeURIComponent(under)));
  if (!local.startsWith(OUT)) return null;
  if (existsSync(local) && statSync(local).isDirectory()) {
    const index = join(local, 'index.html');
    return existsSync(index) ? index : null;
  }
  if (existsSync(local)) return local;
  const html = `${local}.html`;
  return existsSync(html) ? html : null;
}

export function serveOut() {
  const server = createServer((request, response) => {
    const pathname = new URL(request.url ?? '/', 'http://localhost').pathname;
    const file = fileFor(pathname);
    if (file === null) {
      response.writeHead(404, { 'content-type': TYPES['.html'] });
      response.end(readFileSync(join(OUT, '404.html')));
      return;
    }
    response.writeHead(200, { 'content-type': TYPES[extname(file)] ?? 'application/octet-stream' });
    response.end(readFileSync(file));
  });
  return new Promise((resolve) => {
    server.listen(0, '127.0.0.1', () => {
      const address = server.address();
      if (address === null || typeof address === 'string') throw new Error('no port');
      resolve({ origin: `http://127.0.0.1:${String(address.port)}`, close: () => server.close() });
    });
  });
}

if (process.argv[1] !== undefined && import.meta.filename === process.argv[1]) {
  const { origin } = await serveOut();
  console.log(`${origin}${BASE_PATH}/`);
}
