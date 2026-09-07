#!/usr/bin/env node
/**
 * Gate 2's verification, as one command.
 *
 * ```
 * pnpm build
 * pnpm gate2                                  # the committed oracle, no Office needed
 * pnpm gate2 --capture <dir> --out <dir>      # PowerPoint's own PNGs, if one was captured
 * ```
 *
 * Exit status: 0 the gate holds, 1 it does not, 2 the harness could not run.
 * The third is separate for the reason `check-roundtrip.ts` separates it - a run
 * that could not answer must never read as one that answered no.
 *
 * Unlike `pnpm gate1` this needs no PowerPoint: the reference is the oracle
 * sub-phase 3.9 committed, so the gate is reproducible from a clone. A capture
 * only raises the middle panel from an eight-pixel grid to the export it was
 * reduced from.
 */

import { resolve } from 'node:path';

import { REPO_ROOT as ROOT } from '../repo/root.ts';

import { runGate2, type GateOptions } from './gate2.ts';
import { writeGate2Report } from './report.ts';

const OUT = resolve(ROOT, 'fidelity/gate2');

function parse(argv: readonly string[]): GateOptions {
  let out = OUT;
  let capture: string | null = null;
  for (let i = 0; i < argv.length; i++) {
    const arg = argv[i];
    if (arg === '--out') out = resolve(argv[++i] ?? '');
    else if (arg === '--capture') capture = resolve(argv[++i] ?? '');
    else throw new Error('unknown argument ' + String(arg));
  }
  return { out, capture };
}

let options: GateOptions;
try {
  options = parse(process.argv.slice(2));
} catch (error) {
  process.stderr.write('gate2: ' + (error instanceof Error ? error.message : String(error)) + '\n');
  process.exit(2);
}

try {
  const run = await runGate2(options);
  const file = writeGate2Report(run, options.out);

  for (const entry of run.coverage) {
    const count = entry.slides.length;
    process.stdout.write(
      `  ${count === 0 ? 'x' : '.'} ${entry.label.padEnd(34)} ${String(count).padStart(3)} slide(s)\n`,
    );
  }
  for (const gap of run.notDrawn) process.stdout.write(`  x ${gap.key}: ${gap.reason}\n`);

  // Worst first, because a mean over fourteen slides hides the one that is
  // wrong and the whole point of the page is to show it.
  process.stdout.write('\n  slide                     meanBp  maxD  worst region\n');
  for (const slide of [...run.slides].sort((a, b) => a.meanBp - b.meanBp)) {
    const worst = slide.regions[0];
    const region =
      worst === undefined ? '-' : `${String(worst.cells)} cells @ maxD ${String(worst.maxD)}`;
    process.stdout.write(
      `  ${slide.key.padEnd(24)} ${String(slide.meanBp).padStart(6)} ` +
        `${String(slide.maxD).padStart(5)}  ${region}\n`,
    );
  }

  process.stdout.write(
    `gate2: ${String(run.slides.length)} slide(s) drawn beside PowerPoint from ` +
      `${String(run.scanned)} scanned, report at ${file}\n`,
  );
  process.exit(run.ok ? 0 : 1);
} catch (error) {
  process.stderr.write('gate2: ' + (error instanceof Error ? error.message : String(error)) + '\n');
  process.exit(2);
}
