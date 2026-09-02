import { mkdirSync, readFileSync } from 'node:fs';
import { spawnSync } from 'node:child_process';
import { basename, join, resolve } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { chromium, type Browser } from 'playwright';
import { startServer } from '../bench/serve.ts';

/**
 * Gate 1, end to end, through a real browser and then through real PowerPoint.
 *
 * ```
 * pnpm build
 * pnpm gate1                     # the kitchen-sink deck, both modes, + PowerPoint
 * pnpm gate1 --no-powerpoint     # everything except the last step
 * ```
 *
 * The gate reads:
 *
 * > a live demo that opens a deck with charts + SmartArt + animations + OLE +
 * > macros, re-saves it, and the downloaded file opens in real PowerPoint with
 * > zero repair prompt and zero visible difference.
 *
 * Three of those words decide the shape of this file.
 *
 * **"downloaded"**, so the bytes that reach PowerPoint have to come out of the
 * browser's own download machinery and not out of a `page.evaluate` that
 * returns an array. A `Blob`, an object URL and a synthetic click are what the
 * page actually does, and each of the three has its own way of being subtly
 * wrong - a stale revoke cancels the download, a missing `download` attribute
 * navigates instead of saving, and a `Uint8Array` view onto a larger buffer
 * writes the wrong bytes. None of those would fail a unit test of the export.
 *
 * **"re-saves"**, so it is run twice: once with no edit, which is what the gate
 * literally asks for, and once with the stamp, which is the only path in which
 * a part is serialised afresh rather than streamed. A demo that only ever
 * emits the bytes it read has not exercised the writer.
 *
 * **"real PowerPoint"**, so the last step is the COM oracle from 1.5 with
 * `OpenAndRepair` off. With it on - which is the documented default, and what
 * every obvious script gets - a file PowerPoint repairs opens successfully,
 * already repaired, and reports success. That is the one measurement this
 * project cannot afford to get wrong.
 *
 * Nothing here is installed and nothing outside the repository is read: the
 * Chromium is the one `pnpm test` already uses, the server binds to loopback
 * and serves three directories of this repository, and the only file PowerPoint
 * is pointed at is one this run just wrote into the scratch directory it was
 * given.
 */

const ROOT = resolve(fileURLToPath(import.meta.url), '../../..');
const ORACLE = resolve(ROOT, 'packages/cli/scripts/powerpoint-oracle.ps1');
const BUILT = resolve(ROOT, 'apps/studio/dist/main.js');

/** The five features the gate names, and the one deck that has all of them. */
export const GATE_DECK = 'corpus/decks/a43-kitchen-sink.pptm';

export type EditKind = 'none' | 'stamp';

/** What `powerpoint-oracle.ps1` said, kept apart from what we concluded. */
export type OpenVerdict = 'opened' | 'would-not-open' | 'unresolved' | 'not-run';

export interface GateRun {
  readonly edit: EditKind;
  /** Where the browser's download landed. */
  readonly file: string;
  readonly bytesIn: number;
  readonly bytesOut: number;
  /** Parts serialised afresh. Empty for `none`; one entry for `stamp`. */
  readonly rewritten: readonly string[];
  readonly streamed: number;
  /** Parts that differ from the source package, compared as documents. */
  readonly differences: readonly string[];
  readonly parts: number;
  readonly blocking: number;
  readonly rulesChecked: number;
  /** Did the file the browser produced survive being read back here? */
  readonly reopened: boolean;
  readonly powerPoint: OpenVerdict;
  readonly powerPointDetail: string;
  readonly ok: boolean;
}

export interface GateOptions {
  readonly deck: string;
  readonly out: string;
  readonly powerPoint: boolean;
  readonly headed: boolean;
  readonly port: number;
}

/** The subset of the page's automation hook this harness drives. */
interface StudioAutomation {
  exportFromUrl: (
    url: string,
    edit: EditKind,
  ) => Promise<{
    outcome: {
      rewritten: string[];
      streamed: number;
      bytesIn: number;
      bytesOut: number;
      comparison: { parts: number; differences: { part: string; detail: string }[] };
      report: { blocking: number; checked: string[] } | null;
    };
  }>;
}

/**
 * Ask PowerPoint, with repair off.
 *
 * `-AllowRepair` is deliberately not offered here. The interesting question at
 * this gate is not "can PowerPoint make something of it" - it can make
 * something of almost anything - but whether it opens the file as written, and
 * a flag that softened that would be a flag someone reached for on a bad day.
 */
function askPowerPoint(file: string): { verdict: OpenVerdict; detail: string } {
  const result = spawnSync(
    'powershell.exe',
    ['-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', ORACLE, '-File', file],
    { encoding: 'utf8', timeout: 120_000 },
  );
  const detail = ((result.stdout ?? '') + (result.stderr ?? '')).trim().split('\n').pop() ?? '';
  if (result.error !== undefined) return { verdict: 'unresolved', detail: result.error.message };
  if (result.status === 0) return { verdict: 'opened', detail };
  if (result.status === 1) return { verdict: 'would-not-open', detail };
  return { verdict: 'unresolved', detail };
}

/** Read the emitted file back with the writer, which is a second opinion. */
async function reopen(file: string): Promise<{ ok: boolean; parts: number }> {
  const writer = (await import(
    pathToFileURL(resolve(ROOT, 'packages/writer/dist/index.js')).href
  )) as {
    roundTripPackage: (bytes: Uint8Array) => {
      ok: boolean;
      comparison: { parts: { part: string }[] };
    };
  };
  const result = writer.roundTripPackage(new Uint8Array(readFileSync(file)));
  return { ok: result.ok, parts: result.comparison.parts.length };
}

export async function runGate1(options: GateOptions): Promise<GateRun[]> {
  try {
    readFileSync(BUILT);
  } catch {
    throw new Error('apps/studio/dist/main.js is missing - run `pnpm build` first');
  }
  mkdirSync(options.out, { recursive: true });

  const server = startServer({ port: options.port, decks: null });
  let browser: Browser | undefined;
  const runs: GateRun[] = [];

  try {
    browser = await chromium.launch({ headless: !options.headed });
    const context = await browser.newContext({ acceptDownloads: true });
    const page = await context.newPage();

    const failures: string[] = [];
    page.on('pageerror', (error) => failures.push('page error: ' + error.message));
    page.on('console', (message) => {
      if (message.type() === 'error') failures.push('console: ' + message.text());
    });

    await page.goto(server.url + '/', { waitUntil: 'load' });
    await page.waitForFunction(() => 'pptxStudio' in globalThis);

    for (const edit of ['none', 'stamp'] as const) {
      const download = page.waitForEvent('download');
      const result = await page.evaluate(
        ([url, kind]) =>
          (globalThis as unknown as { pptxStudio: StudioAutomation }).pptxStudio.exportFromUrl(
            url as string,
            kind as EditKind,
          ),
        ['/' + options.deck, edit],
      );

      const saved = await download;
      const file = join(options.out, saved.suggestedFilename().replace('.pptx-studio', '.' + edit));
      await saved.saveAs(file);

      const outcome = result.outcome;
      const reread = await reopen(file);
      const differences = outcome.comparison.differences.map((entry) => entry.part);
      const asked = options.powerPoint
        ? askPowerPoint(file)
        : { verdict: 'not-run' as OpenVerdict, detail: 'skipped with --no-powerpoint' };

      // A no-op export must differ nowhere and rewrite nothing; a stamped one
      // must differ in exactly one part and rewrite exactly that part. Both
      // halves matter: an export that rewrote a part without changing it would
      // pass the first check and fail the second, and it is the more worrying
      // of the two - it means the writer serialised something nobody edited.
      const expected = edit === 'none' ? 0 : 1;

      runs.push({
        edit,
        file,
        bytesIn: outcome.bytesIn,
        bytesOut: outcome.bytesOut,
        rewritten: outcome.rewritten,
        streamed: outcome.streamed,
        differences,
        parts: outcome.comparison.parts,
        blocking: outcome.report?.blocking ?? -1,
        rulesChecked: outcome.report?.checked.length ?? 0,
        reopened: reread.ok,
        powerPoint: asked.verdict,
        powerPointDetail: asked.detail,
        ok:
          outcome.report?.blocking === 0 &&
          reread.ok &&
          reread.parts === outcome.comparison.parts &&
          differences.length === expected &&
          outcome.rewritten.length === expected &&
          (asked.verdict === 'opened' || asked.verdict === 'not-run'),
      });
    }

    if (failures.length > 0) {
      throw new Error('the page reported: ' + failures.join('; '));
    }
    return runs;
  } finally {
    await browser?.close();
    await server.close();
    // The oracle attaches to a PowerPoint that was already running and never
    // quits that one, because it is the user's. It does start one when there is
    // none, and leaving that behind after a command that has finished is rude:
    // an invisible instance holding a file handle is exactly the thing somebody
    // hits an hour later and cannot explain.
    if (options.powerPoint) {
      spawnSync(
        'powershell.exe',
        [
          '-NoProfile',
          '-NonInteractive',
          '-ExecutionPolicy',
          'Bypass',
          '-File',
          ORACLE,
          '-QuitOnly',
        ],
        { encoding: 'utf8', timeout: 60_000 },
      );
    }
  }
}

export function formatGate1(deck: string, runs: readonly GateRun[]): string {
  const lines: string[] = ['gate1: ' + basename(deck)];
  for (const run of runs) {
    lines.push(
      '  ' +
        (run.ok ? 'PASS' : 'FAIL') +
        '  ' +
        run.edit.padEnd(6) +
        String(run.parts).padStart(4) +
        ' parts, ' +
        String(run.rewritten.length) +
        ' rewritten, ' +
        String(run.streamed) +
        ' streamed, ' +
        String(run.rulesChecked) +
        ' rules / ' +
        String(run.blocking) +
        ' blocking',
    );
    lines.push(
      '        ' +
        String(run.bytesIn) +
        ' bytes in, ' +
        String(run.bytesOut) +
        ' out; reopened ' +
        (run.reopened ? 'clean' : 'BROKEN') +
        '; PowerPoint: ' +
        run.powerPoint +
        (run.powerPointDetail === '' ? '' : ' (' + run.powerPointDetail + ')'),
    );
    lines.push(
      '        differs in: ' +
        (run.differences.length === 0 ? 'nothing' : run.differences.join(', ')),
    );
    lines.push('        ' + run.file);
  }
  return lines.join('\n');
}
