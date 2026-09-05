#!/usr/bin/env node
/**
 * Round-trip every committed deck, and hold the badge to the answer.
 *
 * ```
 * pnpm roundtrip           # measure, and fail if the committed badge disagrees
 * pnpm roundtrip --write   # measure, and write the badge
 * ```
 *
 * Sub-phase 1.6. Two jobs in one command, because they must not be able to
 * disagree: it measures the round trip, and it checks that the number on the
 * README is the number it just measured.
 *
 * Exit status: 0 clean, 1 a deck failed or the badge is stale, 2 the check
 * could not run. The third is separate for the same reason `check-corpus.ts`
 * separates it - a missing build is a different incident from a broken deck,
 * and a run that cannot answer must not be read as one that answered no.
 *
 * Needs `pnpm build`, because nothing links `@pptx-studio/*` into `tools/` and
 * the specifier does not resolve under plain `node`. `write-census-keys.ts`
 * reaches into `dist` the same way and for the same reason.
 */

import { mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { relative, resolve } from 'node:path';
import { pathToFileURL } from 'node:url';

import { REPO_ROOT as ROOT } from '../../repo/root.ts';
import {
  badgeText,
  roundTripCorpus,
  summaryFor,
  type CorpusRoundTrip,
  type RoundTrip,
} from './report.ts';

const DIST = resolve(ROOT, 'packages/writer/dist/index.js');
const BADGE = resolve(ROOT, '.github/badges/roundtrip.json');

const write = process.argv.includes('--write');

function fail(message: string, code: number): never {
  process.stderr.write('roundtrip: ' + message + '\n');
  process.exit(code);
}

let roundTripPackage: RoundTrip;
try {
  const writer = (await import(pathToFileURL(DIST).href)) as { roundTripPackage: RoundTrip };
  roundTripPackage = writer.roundTripPackage;
} catch (error) {
  fail(
    'could not load ' +
      relative(ROOT, DIST).split('\\').join('/') +
      ' - run `pnpm build` first.\n  ' +
      (error instanceof Error ? error.message : String(error)),
    2,
  );
}

let report: CorpusRoundTrip;
try {
  report = roundTripCorpus(ROOT, roundTripPackage);
} catch (error) {
  fail(
    'the corpus could not be read: ' + (error instanceof Error ? error.message : String(error)),
    2,
  );
}

// The job summary goes to GitHub when there is a GitHub to send it to, and to
// nowhere otherwise. It is not printed to stdout: a local `pnpm check` wants
// one line, and a table of fifty-two decks buried between the lint output and
// the test output is a table nobody reads.
const summaryPath = process.env['GITHUB_STEP_SUMMARY'];
if (summaryPath !== undefined && summaryPath !== '') {
  try {
    writeFileSync(summaryPath, summaryFor(report), { flag: 'a' });
  } catch {
    // A summary that could not be written is not a reason to fail the gate.
  }
}

const text = badgeText(report);
const committed = (() => {
  try {
    return readFileSync(BADGE, 'utf8');
  } catch {
    return null;
  }
})();

if (write) {
  mkdirSync(resolve(ROOT, '.github/badges'), { recursive: true });
  writeFileSync(BADGE, text);
}

const headline =
  String(report.passed) +
  '/' +
  String(report.total) +
  ' deck(s), ' +
  String(report.parts) +
  ' part(s) (' +
  String(report.xml) +
  ' xml, ' +
  String(report.relationships) +
  ' rels, ' +
  String(report.binary) +
  ' binary)';

if (!report.ok) {
  for (const deck of report.decks.filter((entry) => !entry.ok)) {
    process.stderr.write('  ' + deck.id + ': ' + deck.differences.join(', ') + '\n');
  }
  fail(headline + ' - see the differences above', 1);
}

if (!write && committed !== text) {
  process.stderr.write(
    'roundtrip: ' +
      headline +
      '\n' +
      (committed === null
        ? '  .github/badges/roundtrip.json does not exist.\n'
        : '  .github/badges/roundtrip.json says something else:\n' +
          committed
            .trimEnd()
            .split('\n')
            .map((line) => '    ' + line)
            .join('\n') +
          '\n') +
      '  the badge is a claim about this repository, so it is checked rather than\n' +
      '  published from a run. Regenerate it with `pnpm roundtrip --write`.\n',
  );
  process.exit(1);
}

process.stdout.write(
  'roundtrip: ' + headline + (write && committed !== text ? ' (badge updated)' : '') + '\n',
);
