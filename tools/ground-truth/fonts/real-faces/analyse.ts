/**
 * Experiment T14, step 2 - read the measurement, score it, and fail on it.
 *
 * ```
 * node tools/ground-truth/fonts/real-faces/analyse.ts <work-dir> [--summary <p>]
 * ```
 *
 * Nothing here reads a font or opens a browser: it scores the artifact
 * `tools/ground-truth/fonts/real-faces/measure-in-browser.ts` wrote, so a
 * maintainer can re-score a red run anywhere. ADR 0042.
 */

import { appendFileSync, readFileSync, writeFileSync } from 'node:fs';
import { join, resolve } from 'node:path';

import { reportMarkdown, score, type Run } from './score.ts';

const USAGE = 'usage: analyse.ts <work-dir> [--summary <path>]';

const args = process.argv.slice(2);
const work = args[0];
if (work === undefined || work.startsWith('--')) throw new Error(USAGE);

const summaryAt = args.indexOf('--summary');
const summary = summaryAt < 0 ? undefined : args[summaryAt + 1];
if (summaryAt >= 0 && summary === undefined) throw new Error(`${USAGE}\n--summary needs a path`);

const workDir = resolve(work);
const run = JSON.parse(readFileSync(join(workDir, 'real-faces.json'), 'utf8')) as Run;

const verdict = score(run);
const report = reportMarkdown(run, verdict);

writeFileSync(join(workDir, 'verdict.json'), `${JSON.stringify(verdict, null, 2)}\n`);
if (summary !== undefined) appendFileSync(summary, report);
console.log(report);

if (verdict.failures.length > 0) {
  throw new Error(
    `${String(verdict.failures.length)} gate(s) failed: ` +
      [...new Set(verdict.failures.map((failure) => failure.kind))].join(', '),
  );
}
