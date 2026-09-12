#!/usr/bin/env node
/**
 * Gate 3's verification, as one command.
 *
 * ```
 * pnpm build
 * pnpm gate3                                   # the gate: a46 through the page, at every zoom, offline
 * pnpm gate3 --record                          # this machine's first zoom baseline
 * pnpm gate3 --record --why '<reason>' --expect <n>
 * pnpm gate3 --deck <path>                     # a measurement of any deck; scores and gates nothing
 * pnpm gate3 --out <dir> --port <n> --headed
 * ```
 *
 * Exit status: 0 the gate holds, 1 it does not, 2 the harness could not run - the third kept
 * apart for the reason `check-roundtrip.ts` keeps it apart. A `--deck` run exits 0 when it ran,
 * because it reached no verdict.
 */

import { resolve } from 'node:path';

import { parseRecordArgs } from '../fidelity/record.ts';
import { REPO_ROOT as ROOT } from '../repo/root.ts';

import { runGate3, type Gate3Options, type Gate3Run } from './gate3.ts';
import { writeGate3Report } from './report.ts';

const OUT = resolve(ROOT, 'fidelity/gate3');

const PORT = 5201;

function parse(argv: readonly string[]): Gate3Options {
  let out = OUT;
  let port = PORT;
  let headed = false;
  let deck: string | null = null;
  const rest: string[] = [];
  for (let i = 0; i < argv.length; i++) {
    const arg = argv[i];
    if (arg === '--out') out = resolve(argv[++i] ?? '');
    else if (arg === '--port') port = Number.parseInt(argv[++i] ?? '', 10);
    else if (arg === '--headed') headed = true;
    else if (arg === '--deck') deck = resolve(argv[++i] ?? '');
    else rest.push(String(arg));
  }
  if (!Number.isInteger(port) || port <= 0) throw new Error('--port wants a positive integer');
  const record = parseRecordArgs(rest);
  if (deck !== null && record.record)
    throw new Error('--deck is a measurement; it records nothing');
  return { out, port, headed, record, deck };
}

function pct(zoom: number): string {
  return `${String(zoom * 100).padStart(5)} %`;
}

function printRun(run: Gate3Run, file: string): void {
  const write = (line: string): void => {
    process.stdout.write(`${line}\n`);
  };
  if (run.mode === 'measurement') {
    write('  gate3: a measurement, not a gate - nothing here is scored or recorded');
  }
  write(
    `  ${String(run.slides)} slide(s) in ${run.deckPath}; parsed in ${run.timings.parseMs.toFixed(0)} ms, ` +
      `first slide at ${run.timings.firstStageMs.toFixed(0)} ms, strip at ${run.timings.stripMs.toFixed(0)} ms ` +
      `(${run.timings.stripCpuMs.toFixed(0)} ms on the thread)` +
      (run.timings.heapBytes === null
        ? ''
        : `, heap ${(run.timings.heapBytes / 1048576).toFixed(0)} MB`),
  );
  write('  zoom      width  drawn  meanBp   mount ms  shot ms   worst');
  for (const column of run.zooms) {
    const drawn = column.slides.length;
    const mean = (values: readonly number[]): string =>
      values.length === 0 ? '-' : (values.reduce((a, b) => a + b, 0) / values.length).toFixed(1);
    const worst = [...column.slides].sort((a, b) => a.meanBp - b.meanBp)[0];
    write(
      `  ${pct(column.zoom)} ${String(column.width).padStart(6)} ${String(drawn).padStart(6)} ` +
        `${run.mode === 'gate' ? String(column.meanBp).padStart(7) : '      -'} ` +
        `${mean(column.slides.map((s) => s.mountMs)).padStart(9)} ${mean(column.slides.map((s) => s.screenshotMs)).padStart(8)}   ` +
        (worst === undefined || run.mode !== 'gate' ? '-' : `${worst.key} ${String(worst.meanBp)}`),
    );
  }
  for (const gap of run.notDrawn) write(`  x ${gap.key}: ${gap.reason}`);
  for (const brk of run.breaks)
    write(`  x ${brk.key} differs at ${pct(brk.zoom)} beyond what the stroke rule owns`);
  for (const error of run.pageErrors) write(`  x page error: ${error}`);
  write(
    `  requests: ${String(run.requests.total)} (${String(run.requests.static)} static, ` +
      `${String(run.requests.deck)} deck, ${String(run.requests.other.length)} other, ` +
      `${String(run.requests.afterOffline.length)} after going offline)`,
  );
  for (const url of run.requests.other) write(`  x request the gate does not recognise: ${url}`);
  for (const url of run.requests.afterOffline) write(`  x request after going offline: ${url}`);
  if (run.baseline !== null) {
    if (run.baseline.recorded !== null) {
      const moved = [...run.baseline.changed, ...run.baseline.unrecorded, ...run.baseline.vanished];
      write(
        `  recorded ${run.baseline.recorded}` +
          (moved.length === 0 ? '' : `, moving ${String(moved.length)}`),
      );
    } else {
      for (const key of run.baseline.changed)
        write(`  x ${key} rasterises differently than recorded`);
      for (const key of run.baseline.vanished) write(`  x ${key} is recorded and was not drawn`);
    }
  }
  write(`  report at ${file}`);
}

let options: Gate3Options;
try {
  options = parse(process.argv.slice(2));
} catch (error) {
  process.stderr.write(`gate3: ${error instanceof Error ? error.message : String(error)}\n`);
  process.exit(2);
}

try {
  const run = await runGate3(options);
  const file = writeGate3Report(run, options.out);
  printRun(run, file);
  if (run.mode === 'measurement') {
    process.stdout.write('gate3: measured, not gated\n');
    process.exit(0);
  }
  process.stdout.write(`gate3: ${run.ok ? 'holds' : 'does not hold'}\n`);
  process.exit(run.ok ? 0 : 1);
} catch (error) {
  process.stderr.write(`gate3: ${error instanceof Error ? error.message : String(error)}\n`);
  process.exit(2);
}
