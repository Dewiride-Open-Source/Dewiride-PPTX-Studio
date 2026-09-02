import { spawnSync } from 'node:child_process';
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { readZip } from '@pptx-studio/opc';
import { firstDifference } from '@pptx-studio/xml';
import { validatePackage } from '@pptx-studio/validate';
import {
  bisectPackages,
  describeChange,
  flattenChanges,
  roundTripPackage,
  summarizeBisect,
  type BisectResult,
  type Change,
  type Oracle,
  type Verdict,
} from '@pptx-studio/writer';

/**
 * `pptx-studio bisect deck.pptx`
 *
 * Find the smallest change that makes a package fail.
 *
 * ## What it is for
 *
 * It is the debugger for this project, and the plan says so in as many words.
 * `validate` answers "does this break one of the twenty-nine rules"; `roundtrip`
 * answers "did we change anything"; neither answers the question you actually
 * have at two in the morning, which is *which* of the four hundred things that
 * changed is the one PowerPoint objects to.
 *
 * ## Two ways to call it
 *
 * With one deck, the broken package is **our own export of it**. That is the
 * workflow this exists for: we read a file, wrote it back, and PowerPoint now
 * wants to repair it - so the delta to search is exactly what the writer did.
 *
 * With two, they are the original and the broken package, in that order. Use
 * this when the broken one came from somewhere else: an edit made through the
 * library, a hand-mangled fixture, a file a user sent in.
 *
 * ## The oracles
 *
 * `--oracle validate` (the default) asks the twenty-nine rules. It needs no
 * PowerPoint, runs in milliseconds, and is what the tests use - but it can only
 * find failures we already know how to describe.
 *
 * `--oracle powerpoint` asks the real thing, on Windows, over COM. This is the
 * one that matters, because the interesting failures are the ones no rule
 * covers. It opens each candidate with `OpenAndRepair` switched **off**, which
 * is the only way to see a repair at all: PowerPoint's automation interface
 * repairs silently by default and reports success.
 *
 * `--oracle command --command "..."` runs anything else. `{}` in the command is
 * replaced with the path to the candidate package; a non-zero exit means the
 * candidate fails.
 */

export type OracleName = 'validate' | 'powerpoint' | 'command';

export interface BisectOptions {
  readonly oracle: OracleName;
  /** The shell command for `--oracle command`. `{}` becomes the candidate path. */
  readonly command: string | null;
  readonly maxRuns: number;
  /** Per-run ceiling in milliseconds, for the oracles that spawn something. */
  readonly timeout: number;
  /** Save the smallest failing package here. */
  readonly write: string | null;
  readonly json: boolean;
  readonly quiet: boolean;
  readonly out: string | null;
  /** Print a line per oracle run. Off by default; a bisection is not quick. */
  readonly progress: boolean;
}

export const BISECT_DEFAULTS: BisectOptions = {
  oracle: 'validate',
  command: null,
  maxRuns: 2000,
  timeout: 120_000,
  write: null,
  json: false,
  quiet: false,
  out: null,
  progress: false,
};

/** Where `powerpoint-oracle.ps1` lives, from `src/` and from `dist/` alike. */
export function oracleScriptPath(): string {
  return fileURLToPath(new URL('../scripts/powerpoint-oracle.ps1', import.meta.url));
}

/**
 * The twenty-nine rules, asked of one candidate.
 *
 * No baseline is supplied and none could be: a candidate is a package that
 * never existed until this run, and it has no history. That makes every fatal
 * finding count as ours, which is the safe direction here for the same reason
 * it is in `validate` - the question is "is this package broken", not "who
 * broke it".
 */
export function validateOracle(): Oracle {
  return (bytes) => {
    try {
      return validatePackage({ bytes }).ok ? 'passes' : 'fails';
    } catch {
      // A candidate that will not open at all is a failure of a different kind,
      // and one this reducer must not chase: it means the splice produced
      // something that is not a package, which is a bug here rather than a
      // finding about the deck.
      return 'unresolved';
    }
  };
}

interface SpawnOracleOptions {
  readonly timeout: number;
  readonly directory: string;
  readonly onSpawn?: (verdict: Verdict, detail: string) => void;
}

/** Write the candidate somewhere, run something over it, read the exit code. */
function spawningOracle(
  run: (path: string) => { status: number | null; signal: string | null; stderr: string },
  options: SpawnOracleOptions,
): Oracle {
  const path = join(options.directory, 'candidate.pptx');
  return (bytes) => {
    writeFileSync(path, bytes);
    const result = run(path);
    if (result.signal !== null || result.status === null) {
      // Killed, or timed out. Neither is an answer about the package.
      options.onSpawn?.('unresolved', result.signal ?? 'no exit status');
      return 'unresolved';
    }
    const verdict: Verdict =
      result.status === 0 ? 'passes' : result.status === 1 ? 'fails' : 'unresolved';
    options.onSpawn?.(verdict, 'exit ' + String(result.status));
    return verdict;
  };
}

export function powerPointOracle(options: SpawnOracleOptions): Oracle {
  const script = oracleScriptPath();
  return spawningOracle((path) => {
    const result = spawnSync(
      'powershell.exe',
      ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', script, '-File', path],
      { timeout: options.timeout, encoding: 'utf8', windowsHide: true },
    );
    return {
      status: result.status,
      signal: result.signal,
      stderr: result.stderr ?? '',
    };
  }, options);
}

export function commandOracle(command: string, options: SpawnOracleOptions): Oracle {
  return spawningOracle((path) => {
    const filled = command.includes('{}') ? command.replaceAll('{}', path) : command + ' ' + path;
    const result = spawnSync(filled, {
      shell: true,
      timeout: options.timeout,
      encoding: 'utf8',
      windowsHide: true,
    });
    // A shell that could not start the command reports 127, and a command
    // that does not exist is not a broken package. Only 0 and 1 are answers.
    return { status: result.status, signal: result.signal, stderr: result.stderr ?? '' };
  }, options);
}

/** Shut down a PowerPoint this run started, so a bisection leaves nothing behind. */
export function quitPowerPoint(timeout: number): void {
  spawnSync(
    'powershell.exe',
    ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', oracleScriptPath(), '-QuitOnly'],
    { timeout, encoding: 'utf8', windowsHide: true },
  );
}

export interface BisectStats {
  readonly original: string;
  readonly broken: string;
  readonly bytesOriginal: number;
  readonly bytesBroken: number;
  readonly entriesOriginal: number;
  readonly entriesBroken: number;
  readonly oracle: OracleName;
}

/**
 * A change's two sides, short enough to read and showing where they part.
 *
 * Truncating both from the front is the obvious thing and it is useless: two
 * versions of a 4 kB element that differ at character 3 000 come out as the
 * same sixty-eight characters twice, which reads as a bug in the bisector. So
 * the window is centred on the first character where they actually differ.
 */
function sides(change: Change, width = 66): string[] {
  const flatten = (text: string): string => text.replaceAll('\r', '').replaceAll('\n', '⏎');
  const was = change.was === null ? null : flatten(change.was);
  const text = change.text === null ? null : flatten(change.text);

  let from = 0;
  if (was !== null && text !== null && (was.length > width || text.length > width)) {
    const at = firstDifference(was, text);
    if (at > width / 2) from = Math.floor(at - width / 2);
  }

  const show = (value: string | null, mark: string): string => {
    if (value === null) return '    ' + mark + '  (absent)';
    if (value === '') return '    ' + mark + '  (nothing)';
    const head = from > 0 ? '…' : '';
    const body = value.slice(from, from + width);
    const tail = from + width < value.length ? '…' : '';
    return '    ' + mark + '  ' + head + body + tail;
  };
  return [show(was, '-'), show(text, '+')];
}

/**
 * Render a result.
 *
 * Exported for the same reason `formatRoundTrip` is: the branches worth reading
 * are the ones a green corpus cannot produce, and they need a test of their own.
 */
export function formatBisect(stats: BisectStats, result: BisectResult, quiet: boolean): string {
  const lines: string[] = [];

  if (!quiet) {
    lines.push(
      'bisect ' + stats.original,
      '   vs. ' + stats.broken,
      '',
      '  original  ' +
        String(stats.bytesOriginal) +
        ' bytes, ' +
        String(stats.entriesOriginal) +
        ' entries',
      '  broken    ' +
        String(stats.bytesBroken) +
        ' bytes, ' +
        String(stats.entriesBroken) +
        ' entries',
      '  delta     ' +
        String(flattenChanges(result.changes).length) +
        ' change(s) in ' +
        String(result.changes.length) +
        ' entry(s)',
      '  oracle    ' +
        stats.oracle +
        ', ' +
        String(result.runs) +
        ' run(s)' +
        (result.cached > 0 ? ', ' + String(result.cached) + ' from cache' : '') +
        (result.unresolved > 0 ? ', ' + String(result.unresolved) + ' unresolved' : ''),
      '',
    );
  }

  switch (result.outcome) {
    case 'identical':
      lines.push('  the two packages hold the same entries, byte for byte');
      break;
    case 'broken-passes':
      lines.push(
        '  the second package passes the oracle, so there is nothing to look for',
        '  (' +
          String(flattenChanges(result.changes).length) +
          ' change(s) between them, none fatal)',
      );
      break;
    case 'original-fails':
      lines.push(
        '  the *original* fails the oracle, so the cause is not in the difference',
        '  nothing was narrowed down; look at the original package first',
      );
      break;
    case 'localized':
      lines.push('  ' + summarizeBisect(result), '');
      for (const change of result.minimal) {
        lines.push('  ' + describeChange(change), ...sides(change), '');
      }
      break;
  }

  if (result.exhausted) {
    lines.push(
      '',
      '  stopped at the ceiling of ' + String(result.runs) + ' oracle run(s).',
      '  the answer above still fails, but it is not minimal - raise --max-runs.',
    );
  }
  if (result.truncated) {
    lines.push(
      '',
      '  the delta was too large to decompose completely, so some changes are',
      '  whole entries rather than elements - raise --max-changes.',
    );
  }

  return lines.join('\n') + '\n';
}

export interface BisectFileResult {
  readonly result: BisectResult;
  readonly output: string;
  readonly exitCode: number;
}

export function bisectFiles(
  originalPath: string,
  brokenPath: string | null,
  options: BisectOptions = BISECT_DEFAULTS,
  note: (text: string) => void = () => {},
): BisectFileResult {
  const original = new Uint8Array(readFileSync(originalPath));
  // One argument means "bisect what our own writer did to this deck", which is
  // the case this command was built for.
  const broken =
    brokenPath === null
      ? roundTripPackage(original).exported.bytes
      : new Uint8Array(readFileSync(brokenPath));

  const directory = mkdtempSync(join(tmpdir(), 'pptx-bisect-'));
  try {
    const spawnOptions: SpawnOracleOptions = { timeout: options.timeout, directory };
    const oracle: Oracle =
      options.oracle === 'validate'
        ? validateOracle()
        : options.oracle === 'powerpoint'
          ? powerPointOracle(spawnOptions)
          : commandOracle(options.command ?? '', spawnOptions);

    const result = bisectPackages(original, broken, {
      oracle,
      maxRuns: options.maxRuns,
      ...(options.progress
        ? {
            onRun: (run) => {
              note('  run ' + String(run.run).padStart(4) + '  ' + run.label + '\n');
            },
          }
        : {}),
    });

    const stats: BisectStats = {
      original: originalPath,
      broken: brokenPath ?? '(our own export of it)',
      bytesOriginal: original.length,
      bytesBroken: broken.length,
      entriesOriginal: readZip(original).entries.length,
      entriesBroken: readZip(broken).entries.length,
      oracle: options.oracle,
    };

    const output = options.json
      ? JSON.stringify(
          {
            original: originalPath,
            broken: stats.broken,
            oracle: options.oracle,
            outcome: result.outcome,
            runs: result.runs,
            cached: result.cached,
            unresolved: result.unresolved,
            exhausted: result.exhausted,
            truncated: result.truncated,
            delta: flattenChanges(result.changes).length,
            minimal: result.minimal.map((change) => ({
              entry: change.entry,
              kind: change.kind,
              where: change.where,
              was: change.was,
              text: change.text,
            })),
          },
          null,
          2,
        ) + '\n'
      : formatBisect(stats, result, options.quiet);

    if (options.write !== null) writeFileSync(options.write, result.bytes);

    return {
      result,
      output,
      // Non-zero whenever the deck is not clean, which is the three outcomes
      // that are not "nothing to report" - the same shape `validate` and
      // `roundtrip` use, so the three verbs can be chained in a script.
      exitCode: result.outcome === 'localized' || result.outcome === 'original-fails' ? 1 : 0,
    };
  } finally {
    rmSync(directory, { recursive: true, force: true });
    if (options.oracle === 'powerpoint') quitPowerPoint(options.timeout);
  }
}

export function runBisect(
  originalPath: string,
  brokenPath: string | null,
  options: BisectOptions,
  write: (text: string) => void,
): number {
  const result = bisectFiles(originalPath, brokenPath, options, write);
  if (options.out === null) write(result.output);
  else writeFileSync(options.out, result.output);
  return result.exitCode;
}
