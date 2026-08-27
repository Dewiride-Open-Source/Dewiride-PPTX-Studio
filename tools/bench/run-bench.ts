#!/usr/bin/env node
/**
 * Sub-phase 0.8's verification: parse a 200 MB / 300-slide deck in a browser
 * tab and record the timing.
 *
 * ```
 * pnpm build
 * node tools/bench/make-deck.ts media-200mb <dir>/media-200mb.pptx
 * node tools/bench/run-bench.ts --decks <dir> --repeat 5 --out corpus/bench/results.json
 * ```
 *
 * It drives the real page. `globalThis.pptxStudio.censusFromUrl` is the same
 * function a dropped file goes through, so the numbers are the ones a user
 * would get and not the ones a purpose-built harness can arrange. The tab is a
 * real Chromium - the one `pnpm test` already uses - and nothing is installed.
 *
 * The page is served cross-origin isolated, which is what makes
 * `performance.measureUserAgentSpecificMemory()` available inside the worker.
 * Without it the memory column is empty and the run is still valid; the timing
 * does not depend on it.
 */

import { readdirSync, writeFileSync } from 'node:fs';
import { basename, resolve } from 'node:path';
import { chromium } from 'playwright';
import { startServer } from './serve.ts';

interface RunTiming {
  readonly index: number;
  readonly totalMs: number;
  readonly openMs: number;
  readonly inflateMs: number;
  readonly scanMs: number;
  readonly xmlBytesScanned: number;
}

interface StageTiming {
  readonly xmlParts: number;
  readonly xmlBytes: number;
  readonly tokens: number;
  readonly inflateMs: number;
  readonly decodeMs: number;
  readonly tokenizeMs: number;
}

interface BrowserResult {
  readonly name: string;
  readonly bytes: number;
  readonly runs: readonly RunTiming[];
  readonly stages: StageTiming | null;
  readonly memory: {
    beforeBytes: number | null;
    afterBytes: number | null;
    available: boolean;
    note: string | null;
  };
  readonly environment: Record<string, unknown>;
  readonly dispatchMs: number;
  readonly wallMs: number;
  readonly census: {
    readonly archive: {
      readonly parts: number;
      readonly entries: number;
      readonly declaredInflatedBytes: number;
    };
    readonly presentation: { readonly slides: number } | null;
    readonly features: readonly { readonly key: string; readonly count: number }[];
    readonly problems: readonly { readonly severity: string; readonly code: string }[];
  };
}

/** The automation surface `apps/studio/src/main.ts` puts on the page. */
interface StudioWindow {
  pptxStudio: {
    censusFromUrl: (url: string, repeat?: number) => Promise<BrowserResult>;
  };
}

interface Options {
  readonly decks: string;
  readonly repeat: number;
  readonly port: number;
  readonly out: string | null;
  readonly only: string | null;
  readonly headed: boolean;
  /** Save a picture of the rendered explorer after the last deck. */
  readonly screenshot: string | null;
}

function parse(argv: readonly string[]): Options {
  let decks: string | null = null;
  let repeat = 3;
  let port = 5199;
  let out: string | null = null;
  let only: string | null = null;
  let headed = false;
  let screenshot: string | null = null;
  for (let i = 0; i < argv.length; i++) {
    const flag = argv[i];
    if (flag === '--decks') decks = resolve(argv[++i] ?? '');
    else if (flag === '--repeat') repeat = Number.parseInt(argv[++i] ?? '', 10);
    else if (flag === '--port') port = Number.parseInt(argv[++i] ?? '', 10);
    else if (flag === '--out') out = resolve(argv[++i] ?? '');
    else if (flag === '--only') only = argv[++i] ?? null;
    else if (flag === '--headed') headed = true;
    else if (flag === '--screenshot') screenshot = resolve(argv[++i] ?? '');
    else throw new Error('unknown flag ' + String(flag));
  }
  if (decks === null) {
    throw new Error(
      'usage: run-bench.ts --decks <dir-of-pptx> [--repeat 3] [--out results.json] ' +
        '[--only name] [--screenshot page.png] [--headed]',
    );
  }
  if (!Number.isFinite(repeat) || repeat < 1) throw new Error('--repeat wants a positive integer');
  return { decks, repeat, port, out, only, headed, screenshot };
}

const options = parse(process.argv.slice(2));

const decks = readdirSync(options.decks)
  .filter((name) => name.toLowerCase().endsWith('.pptx'))
  .filter((name) => options.only === null || name.includes(options.only))
  .sort();
if (decks.length === 0) throw new Error('no .pptx files in ' + options.decks);

const server = startServer({ port: options.port, decks: options.decks });
const browser = await chromium.launch({ headless: !options.headed });
const results: BrowserResult[] = [];
let chromiumVersion: string;

try {
  const page = await browser.newPage();
  const failures: string[] = [];
  page.on('pageerror', (error) => failures.push(String(error)));
  page.on('console', (message) => {
    if (message.type() === 'error') failures.push(message.text());
  });

  await page.goto(server.url + '/', { waitUntil: 'load' });
  // The hook only exists once `main.ts` has run and the worker has said hello.
  await page.waitForFunction(
    () => (globalThis as unknown as Partial<StudioWindow>).pptxStudio !== undefined,
    undefined,
    { timeout: 30_000 },
  );

  // Probed on the main thread as well as in the worker, because the two
  // answers are different and only the pair is informative: the memory API is
  // documented for worker scopes, and this browser exposes it to the document
  // and not to the worker. One `null` on its own would look like a missing
  // header rather than a missing implementation.
  const host = await page.evaluate(() => ({
    isolated: (globalThis as unknown as { crossOriginIsolated: boolean }).crossOriginIsolated,
    measureMemory:
      (performance as unknown as { measureUserAgentSpecificMemory?: unknown })
        .measureUserAgentSpecificMemory !== undefined,
  }));
  chromiumVersion = browser.version();
  console.log(
    'chromium ' +
      chromiumVersion +
      ', cross-origin isolated: ' +
      String(host.isolated) +
      ', measureUserAgentSpecificMemory on the page: ' +
      String(host.measureMemory),
  );
  console.log('');

  for (const deck of decks) {
    process.stdout.write('  ' + deck.padEnd(24) + ' … ');
    // `page.evaluate` has no timeout of its own, which is what this needs: a
    // 200 MB deck at three repeats is minutes of honest work, and a harness
    // that gave up halfway would report a failure that is entirely its own.
    const result: BrowserResult = await page.evaluate(
      async ({ url, repeat }: { url: string; repeat: number }) =>
        (globalThis as unknown as StudioWindow).pptxStudio.censusFromUrl(url, repeat),
      { url: '/decks/' + deck, repeat: options.repeat },
    );
    results.push(result);

    const best = [...result.runs].sort((a, b) => a.totalMs - b.totalMs)[0];
    console.log(
      (result.bytes / 1024 / 1024).toFixed(0).padStart(4) +
        ' MiB  best ' +
        (best === undefined ? '?' : (best.totalMs / 1000).toFixed(2) + ' s'),
    );
  }

  if (options.screenshot !== null) {
    // The benchmark exercises the rendering path - `censusFromUrl` calls the
    // same `render` a dropped file does - but never looks at it. A picture is
    // the only thing that catches a report that computed correctly and drew
    // nothing.
    await page.setViewportSize({ width: 1200, height: 2000 });
    await page.screenshot({ path: options.screenshot });
    console.log('\nwrote ' + options.screenshot);
  }

  if (failures.length > 0) {
    console.error('\npage errors:');
    for (const failure of failures) console.error('  ' + failure);
    process.exitCode = 1;
  }
} finally {
  await browser.close();
  await server.close();
}

// ------------------------------------------------------------------- report

const mib = (bytes: number): string => (bytes / 1024 / 1024).toFixed(1);
const rate = (bytes: number, ms: number): string =>
  ms <= 0 ? '-' : (bytes / 1024 / 1024 / (ms / 1000)).toFixed(0) + ' MiB/s';

for (const result of results) {
  const best = [...result.runs].sort((a, b) => a.totalMs - b.totalMs)[0];
  const median = [...result.runs].sort((a, b) => a.totalMs - b.totalMs)[
    Math.floor(result.runs.length / 2)
  ];
  console.log('');
  console.log('='.repeat(72));
  console.log(basename(result.name));
  console.log('='.repeat(72));
  console.log(
    '  archive        ' +
      mib(result.bytes) +
      ' MiB, ' +
      String(result.census.archive.entries) +
      ' entries, ' +
      String(result.census.archive.parts) +
      ' parts, ' +
      String(result.census.presentation?.slides ?? 0) +
      ' slides',
  );
  console.log('  inflated       ' + mib(result.census.archive.declaredInflatedBytes) + ' MiB');
  if (best !== undefined && median !== undefined) {
    console.log(
      '  census         best ' +
        (best.totalMs / 1000).toFixed(3) +
        ' s, median ' +
        (median.totalMs / 1000).toFixed(3) +
        ' s over ' +
        String(result.runs.length) +
        ' runs',
    );
    console.log(
      '    open         ' + best.openMs.toFixed(0) + ' ms   (central directory + content types)',
    );
    console.log(
      '    inflate      ' +
        best.inflateMs.toFixed(0) +
        ' ms   ' +
        rate(best.xmlBytesScanned, best.inflateMs),
    );
    console.log(
      '    scan         ' +
        best.scanMs.toFixed(0) +
        ' ms   ' +
        rate(best.xmlBytesScanned, best.scanMs) +
        '   (tokenize + count)',
    );
  }
  const stages = result.stages;
  if (stages !== null) {
    console.log(
      '  stages         ' +
        mib(stages.xmlBytes) +
        ' MiB of XML, ' +
        String(stages.tokens) +
        ' tokens, ' +
        String(stages.xmlParts) +
        ' parts',
    );
    console.log(
      '    inflate      ' +
        stages.inflateMs.toFixed(0) +
        ' ms   ' +
        rate(stages.xmlBytes, stages.inflateMs),
    );
    console.log(
      '    decode       ' +
        stages.decodeMs.toFixed(0) +
        ' ms   ' +
        rate(stages.xmlBytes, stages.decodeMs),
    );
    console.log(
      '    tokenize     ' +
        stages.tokenizeMs.toFixed(0) +
        ' ms   ' +
        rate(stages.xmlBytes, stages.tokenizeMs) +
        '   ' +
        ((stages.tokenizeMs * 1e6) / Math.max(1, stages.tokens)).toFixed(0) +
        ' ns/token',
    );
  }
  console.log('  dispatch       ' + result.dispatchMs.toFixed(1) + ' ms   (transfer, not copy)');
  if (!result.memory.available && result.memory.note !== null) {
    console.log('  worker memory  unavailable: ' + result.memory.note);
  }
  if (result.memory.available) {
    console.log(
      '  worker memory  ' +
        (result.memory.beforeBytes === null ? '?' : mib(result.memory.beforeBytes)) +
        ' MiB before, ' +
        (result.memory.afterBytes === null ? '?' : mib(result.memory.afterBytes)) +
        ' MiB after',
    );
  }
  const errors = result.census.problems.filter((problem) => problem.severity === 'error');
  console.log(
    '  problems       ' +
      String(result.census.problems.length) +
      ', of which ' +
      String(errors.length) +
      ' errors',
  );
}

if (options.out !== null) {
  // A summary, not the census. The full report of a 300-slide deck is a
  // thousand-row part table and a namespace histogram, and none of that is what
  // a benchmark record is for. What belongs in the repository is the shape of
  // each deck, the timings, and enough about the machine to know what the
  // timings mean.
  const summary = {
    recordedAt: new Date().toISOString().slice(0, 10),
    note:
      'Timings from one machine on one day. A record, not a fixture: nothing ' +
      'asserts on them, because a slower laptop is not a regression.',
    chromium: chromiumVersion,
    decks: results.map((result) => ({
      name: result.name,
      bytes: result.bytes,
      entries: result.census.archive.entries,
      parts: result.census.archive.parts,
      slides: result.census.presentation?.slides ?? 0,
      inflatedBytes: result.census.archive.declaredInflatedBytes,
      census: {
        runs: result.runs.length,
        bestMs: Math.min(...result.runs.map((run) => run.totalMs)),
        medianMs: [...result.runs].sort((a, b) => a.totalMs - b.totalMs)[
          Math.floor(result.runs.length / 2)
        ]?.totalMs,
        openMs: result.runs[0]?.openMs,
        inflateMs: result.runs[0]?.inflateMs,
        scanMs: result.runs[0]?.scanMs,
      },
      stages: result.stages,
      dispatchMs: result.dispatchMs,
      environment: result.environment,
      memory: result.memory,
      problems: result.census.problems.length,
      features: Object.fromEntries(
        result.census.features.map((feature) => [feature.key, feature.count]),
      ),
    })),
  };
  writeFileSync(options.out, JSON.stringify(summary, null, 2) + '\n');
  console.log('\nwrote ' + options.out);
}
