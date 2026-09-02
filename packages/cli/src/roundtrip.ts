import { readFileSync, writeFileSync } from 'node:fs';
import { readZip } from '@pptx-studio/opc';
import {
  roundTripPackage,
  summarizeRoundTrip,
  type Difference,
  type RoundTripReport,
} from '@pptx-studio/writer';

/**
 * `pptx-studio roundtrip deck.pptx`
 *
 * Read a deck, write it back, and say whether it is still the same deck.
 *
 * ## What it is for
 *
 * It is the gate of Phase 1 with a command line on it. Sub-phase 1.6 runs it
 * over the corpus on every pull request, and the number it prints is the one
 * the README's badge carries - so this is a place where the output format is
 * part of the deliverable rather than decoration.
 *
 * It is also the first thing to reach for when a deck misbehaves. `validate`
 * answers "is this package well formed"; this answers "did we change it", which
 * is a different question and usually the one that matters, because the packages
 * that cause trouble are the ones that are well formed and subtly different.
 *
 * ## What "the same deck" means here
 *
 * Not the same bytes. The plan rules that out and `comparePackages` explains
 * why at length: entry order, deflate level, timestamps and attribute order all
 * differ legitimately between two archives holding one document. What is
 * compared is the canonical XML of every XML part, the relationship graph with
 * ids treated as opaque labels, and the SHA-256 of everything else.
 *
 * ## The exit code
 *
 * `0` when there is nothing to report. `1` when the two packages differ, and
 * `1` as well when the export was refused outright - the caller in CI wants one
 * bit, and both of those are "this deck did not round-trip". The two are still
 * distinguishable on stderr, because they need completely different work.
 */

export interface RoundTripOptions {
  /** Print the comparison as JSON instead of as text. */
  readonly json: boolean;
  /** Show only the differences, not the tally. */
  readonly quiet: boolean;
  /** Write the report to this path instead of stdout. */
  readonly out: string | null;
  /**
   * Also write the exported package here.
   *
   * The point of the flag is what happens next: open the file in PowerPoint.
   * Nothing in this repository can assert that a file opens clean, and until
   * 1.5 scripts it there is no substitute for doing it by hand.
   */
  readonly write: string | null;
}

export const ROUNDTRIP_DEFAULTS: RoundTripOptions = {
  json: false,
  quiet: false,
  out: null,
  write: null,
};

export interface RoundTripFileResult {
  readonly comparison: RoundTripReport;
  readonly output: string;
  readonly bytes: Uint8Array;
  readonly exitCode: number;
}

function describeDifference(difference: Difference): string {
  // Two lines, the second indented, because a `detail` may itself be three
  // lines when it carries an excerpt of the markup either side.
  const body = difference.detail
    .split('\n')
    .map((line) => '    ' + line.trimEnd())
    .join('\n');
  return '  ' + difference.kind.padEnd(13) + difference.part + '\n' + body;
}

/** What the header line reports, gathered once so the formatter is testable on its own. */
export interface RoundTripStats {
  readonly bytesIn: number;
  readonly bytesOut: number;
  readonly entriesIn: number;
  readonly entriesOut: number;
  readonly rewritten: number;
  readonly streamed: number;
}

/**
 * Render a comparison.
 *
 * Exported so that the branch that lists differences can be tested, which it
 * otherwise could not be: no deck in the corpus makes this writer produce one,
 * which is the point of the whole phase and leaves the most important output in
 * the file as the only output nothing exercises.
 */
export function formatRoundTrip(
  path: string,
  stats: RoundTripStats,
  comparison: RoundTripReport,
  quiet: boolean,
): string {
  const lines: string[] = [];
  if (!quiet) {
    lines.push(
      'roundtrip ' + path,
      '',
      '  read      ' + String(stats.bytesIn) + ' bytes, ' + String(stats.entriesIn) + ' entries',
      '  written   ' + String(stats.bytesOut) + ' bytes, ' + String(stats.entriesOut) + ' entries',
      '  export    ' +
        String(stats.rewritten) +
        ' part(s) re-serialized, ' +
        String(stats.streamed) +
        ' streamed',
      '  compared  ' + summarizeRoundTrip(comparison),
      '',
    );
  }

  if (comparison.ok) {
    lines.push('  no differences');
  } else {
    lines.push('  ' + String(comparison.differences.length) + ' difference(s)', '');
    for (const difference of comparison.differences) lines.push(describeDifference(difference), '');
  }

  if (!quiet && comparison.relabelled.length > 0) {
    // Never true for a package this writer produced, so when it is true it is
    // worth a line: it means the file being compared came from somewhere else.
    lines.push(
      '',
      '  ' + String(comparison.relabelled.length) + ' relationship id(s) renamed:',
      ...comparison.relabelled
        .slice(0, 10)
        .map((entry) => '    ' + entry.source + '  ' + entry.from + ' -> ' + entry.to),
    );
  }

  return lines.join('\n') + '\n';
}

export function roundTripFile(
  path: string,
  options: RoundTripOptions = ROUNDTRIP_DEFAULTS,
): RoundTripFileResult {
  const original = new Uint8Array(readFileSync(path));
  const result = roundTripPackage(original);
  const { comparison } = result;

  const stats: RoundTripStats = {
    bytesIn: original.length,
    bytesOut: result.exported.bytes.length,
    entriesIn: readZip(original).entries.length,
    entriesOut: readZip(result.exported.bytes).entries.length,
    rewritten: result.exported.rewritten.length,
    streamed: result.exported.streamed,
  };

  const output = options.json
    ? JSON.stringify(
        {
          file: path,
          ok: comparison.ok,
          bytesIn: stats.bytesIn,
          bytesOut: stats.bytesOut,
          rewritten: result.exported.rewritten,
          counts: comparison.counts,
          relabelled: comparison.relabelled,
          differences: comparison.differences,
          parts: comparison.parts,
        },
        null,
        2,
      ) + '\n'
    : formatRoundTrip(path, stats, comparison, options.quiet);

  return {
    comparison,
    output,
    bytes: result.exported.bytes,
    exitCode: comparison.ok ? 0 : 1,
  };
}

/** Run `roundTripFile` and put the results where the options say. */
export function runRoundTrip(
  path: string,
  options: RoundTripOptions,
  write: (text: string) => void,
): number {
  const result = roundTripFile(path, options);
  if (options.write !== null) writeFileSync(options.write, result.bytes);
  if (options.out === null) write(result.output);
  else writeFileSync(options.out, result.output);
  return result.exitCode;
}
