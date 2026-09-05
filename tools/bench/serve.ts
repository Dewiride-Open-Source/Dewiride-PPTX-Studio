#!/usr/bin/env node
/**
 * A static file server for `apps/studio`, cross-origin isolated.
 *
 * ```
 * node tools/bench/serve.ts               # http://127.0.0.1:5173
 * node tools/bench/serve.ts --port 8080 --decks C:/somewhere
 * ```
 *
 * Two headers are the reason this exists rather than any off-the-shelf server:
 *
 *   Cross-Origin-Opener-Policy: same-origin
 *   Cross-Origin-Embedder-Policy: require-corp
 *
 * Together they make the page **cross-origin isolated**, which is the only
 * state in which `performance.measureUserAgentSpecificMemory()` is available.
 * That call is the only way to find out what a 200 MB deck actually costs:
 * `performance.memory` counts the JS heap and would report a 200 MB
 * `ArrayBuffer` as costing nothing at all, since it lives outside it.
 *
 * Nothing here is hardened and nothing here should ever face a network. It
 * binds to the loopback address, serves a fixed set of roots, and rejects any
 * path that escapes them.
 */

import { createReadStream, statSync } from 'node:fs';
import { createServer } from 'node:http';
import { extname, join, normalize, resolve, sep } from 'node:path';

import { REPO_ROOT as ROOT } from '../repo/root.ts';

const TYPES: Readonly<Record<string, string>> = {
  '.html': 'text/html; charset=utf-8',
  '.js': 'text/javascript; charset=utf-8',
  '.map': 'application/json; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.pptx': 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
  '.pptm': 'application/vnd.ms-powerpoint.presentation.macroEnabled.12',
  '.png': 'image/png',
};

export interface ServeOptions {
  readonly port: number;
  /** Extra directory served under `/decks/`, for benchmark files kept outside the repo. */
  readonly decks: string | null;
}

function parse(argv: readonly string[]): ServeOptions {
  let port = 5173;
  let decks: string | null = null;
  for (let i = 0; i < argv.length; i++) {
    if (argv[i] === '--port') port = Number.parseInt(argv[++i] ?? '', 10);
    else if (argv[i] === '--decks') decks = resolve(argv[++i] ?? '');
    else throw new Error('unknown flag ' + String(argv[i]));
  }
  if (!Number.isFinite(port)) throw new Error('--port wants a number');
  return { port, decks };
}

/** Resolve a URL path inside `base`, or `null` if it tries to escape. */
function safeJoin(base: string, urlPath: string): string | null {
  const decoded = decodeURIComponent(urlPath);
  const candidate = normalize(join(base, decoded));
  return candidate === base || candidate.startsWith(base + sep) ? candidate : null;
}

export function startServer(options: ServeOptions): { url: string; close: () => Promise<void> } {
  const server = createServer((request, response) => {
    const url = new URL(request.url ?? '/', 'http://127.0.0.1');
    let path = url.pathname === '/' ? '/apps/studio/index.html' : url.pathname;

    let file: string | null;
    if (path.startsWith('/decks/') && options.decks !== null) {
      file = safeJoin(options.decks, path.slice('/decks'.length));
    } else {
      if (
        !path.startsWith('/apps/') &&
        !path.startsWith('/packages/') &&
        !path.startsWith('/corpus/')
      ) {
        path = '/apps/studio' + path;
      }
      file = safeJoin(ROOT, path);
    }

    // Every response carries the isolation headers, the deck included. A single
    // resource without `Cross-Origin-Resource-Policy` is enough to break
    // isolation for the whole page under `require-corp`.
    response.setHeader('Cross-Origin-Opener-Policy', 'same-origin');
    response.setHeader('Cross-Origin-Embedder-Policy', 'require-corp');
    response.setHeader('Cross-Origin-Resource-Policy', 'same-origin');
    response.setHeader('Cache-Control', 'no-store');

    if (file === null) {
      response.writeHead(403).end('outside the served roots');
      return;
    }
    let size: number;
    try {
      const stat = statSync(file);
      if (!stat.isFile()) throw new Error('not a file');
      size = stat.size;
    } catch {
      response.writeHead(404).end('not found: ' + path);
      return;
    }

    response.writeHead(200, {
      'Content-Type': TYPES[extname(file).toLowerCase()] ?? 'application/octet-stream',
      'Content-Length': String(size),
    });
    createReadStream(file).pipe(response);
  });

  server.listen(options.port, '127.0.0.1');
  return {
    url: 'http://127.0.0.1:' + String(options.port),
    close: () =>
      new Promise<void>((done, fail) => {
        server.close((error) => (error === undefined ? done() : fail(error)));
      }),
  };
}

// Run directly rather than imported: start and stay up.
if (
  process.argv[1] !== undefined &&
  import.meta.url.endsWith(process.argv[1].replace(/\\/g, '/'))
) {
  const options = parse(process.argv.slice(2));
  const { url } = startServer(options);
  console.log('serving ' + ROOT + ' at ' + url);
  console.log('  the explorer   ' + url + '/');
  if (options.decks !== null) console.log('  decks          ' + url + '/decks/<name>.pptx');
  console.log('  cross-origin isolated, so worker memory is measurable');
}
