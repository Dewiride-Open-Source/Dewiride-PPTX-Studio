#!/usr/bin/env node
/**
 * Gate 1's verification, as one command.
 *
 * ```
 * pnpm build
 * pnpm gate1                        # a43-kitchen-sink, both modes, and PowerPoint
 * pnpm gate1 --no-powerpoint        # everything a machine without Office can do
 * pnpm gate1 --deck corpus/decks/a26-ole.pptx --headed
 * ```
 *
 * Exit status: 0 every run passed, 1 one did not, 2 the harness could not run.
 * The third is separate for the reason `check-corpus.ts` and `check-roundtrip.ts`
 * separate it - a missing build is a different incident from a broken package,
 * and a run that could not answer must never read as one that answered no.
 *
 * Not in CI, and it cannot be: no hosted runner has PowerPoint on it. GitHub's
 * Windows images ship Visual Studio's Office *development* workload and not one
 * Office application, so the last step of this command exists only on a machine
 * with the real thing. `--no-powerpoint` is what a runner could do, and what it
 * would prove is the browser half - which `apps/studio/src/export.test.ts`
 * already asserts in Chromium on every `pnpm test`.
 */

import { resolve } from 'node:path';

import { REPO_ROOT as ROOT } from '../repo/root.ts';
import { formatGate1, runGate1, GATE_DECK, type GateOptions } from './gate1.ts';

const SCRATCH = resolve(ROOT, 'node_modules/.gate1');

function parse(argv: readonly string[]): GateOptions {
  let deck = GATE_DECK;
  let out = SCRATCH;
  let powerPoint = true;
  let headed = false;
  let port = 5199;

  for (let i = 0; i < argv.length; i++) {
    const arg = argv[i];
    if (arg === '--deck') deck = (argv[++i] ?? '').replace(/\\/g, '/');
    else if (arg === '--out') out = resolve(argv[++i] ?? '');
    else if (arg === '--no-powerpoint') powerPoint = false;
    else if (arg === '--headed') headed = true;
    else if (arg === '--port') port = Number.parseInt(argv[++i] ?? '', 10);
    else throw new Error('unknown argument ' + String(arg));
  }
  if (!Number.isFinite(port)) throw new Error('--port wants a number');
  if (deck.startsWith('/') || deck.includes('..')) {
    throw new Error('--deck is a path inside this repository, relative to its root');
  }
  return { deck, out, powerPoint, headed, port };
}

let options: GateOptions;
try {
  options = parse(process.argv.slice(2));
} catch (error) {
  process.stderr.write('gate1: ' + (error instanceof Error ? error.message : String(error)) + '\n');
  process.exit(2);
}

try {
  const runs = await runGate1(options);
  process.stdout.write(formatGate1(options.deck, runs) + '\n');
  process.exit(runs.every((run) => run.ok) ? 0 : 1);
} catch (error) {
  process.stderr.write('gate1: ' + (error instanceof Error ? error.message : String(error)) + '\n');
  process.exit(2);
}
