/**
 * When a baseline may be rewritten, and when the run must stop instead.
 *
 * The gate has no tolerance to loosen, so `--record` is the only way to make a
 * red run green: it is the escape hatch, and these are the locks on it. ADR
 * 0035.
 */

import { FidelityError } from './errors.ts';

export interface RecordArgs {
  readonly record: boolean;
  /** Why the recorded digests are being rewritten, in the author's words. */
  readonly why: string | null;
  /** How many recorded digests the author says this rewrite touches. */
  readonly expect: number | null;
  /** Writing a platform's first baseline, which is the one record CI may do. */
  readonly bootstrap: boolean;
}

export function parseRecordArgs(argv: readonly string[]): RecordArgs {
  let record = false;
  let why: string | null = null;
  let expect: number | null = null;
  let bootstrap = false;
  for (let at = 0; at < argv.length; at++) {
    const arg = argv[at];
    if (arg === '--record') record = true;
    else if (arg === '--bootstrap') bootstrap = true;
    else if (arg === '--why') why = argv[++at] ?? null;
    else if (arg === '--expect') expect = Number(argv[++at] ?? Number.NaN);
    else throw new FidelityError('FID_RECORD_UNJUSTIFIED', `unknown argument ${String(arg)}`, arg);
  }
  return { record, why, expect, bootstrap };
}

/**
 * Whether this machine may rewrite a baseline at all.
 *
 * A first record has nothing to overwrite and needs no case made for it; every
 * later one is a claim that the renderer is right and the record is stale, and
 * that claim has to be written down by the person making it.
 */
export function assertMayRecord(
  args: RecordArgs,
  environment: { readonly inCi: boolean; readonly hasBaseline: boolean },
): void {
  if (!args.record) return;
  if (args.bootstrap && environment.hasBaseline) {
    throw new FidelityError(
      'FID_RECORD_UNJUSTIFIED',
      '--bootstrap writes a platform its first baseline, and this platform has one',
    );
  }
  // A runner may write a platform's first baseline and nothing else. That file
  // gates nothing until a person commits it, so it cannot turn a gate green;
  // rewriting one that exists could, and the check above has already refused it.
  if (environment.inCi && !args.bootstrap) {
    throw new FidelityError(
      'FID_RECORD_IN_CI',
      'CI may only write a platform its first baseline, and then only with --bootstrap',
    );
  }
  if (!environment.hasBaseline) return;
  if (args.why === null || args.why.trim() === '') {
    throw new FidelityError(
      'FID_RECORD_UNJUSTIFIED',
      'rewriting a recorded baseline needs --why "<what changed and why it is right>"',
    );
  }
  if (args.expect === null || !Number.isInteger(args.expect) || args.expect < 0) {
    throw new FidelityError(
      'FID_RECORD_UNJUSTIFIED',
      'rewriting a recorded baseline needs --expect <how many slides it changes>',
    );
  }
}

/**
 * That the rewrite touches the number of slides its author said it would.
 *
 * Knowing the count in advance is the difference between accepting a change and
 * noticing one: a re-record that moves more slides than the author expected is
 * a second, unexamined change riding along with the first.
 */
export function assertExpectedCount(args: RecordArgs, touched: number): void {
  if (!args.record || args.expect === null) return;
  if (args.expect !== touched) {
    throw new FidelityError(
      'FID_RECORD_UNJUSTIFIED',
      `--expect ${String(args.expect)} slide(s) but this rewrites ${String(touched)}`,
      String(touched),
    );
  }
}

/** Slides the run and the baseline do not agree exist, both ways round. */
export interface BaselineGaps {
  /** Rendered here and absent from the baseline, so nothing gates them. */
  readonly unrecorded: readonly string[];
  /** In the baseline and not rendered here, so the renderer lost a slide. */
  readonly vanished: readonly string[];
}

export function baselineGaps(recorded: Iterable<string>, rendered: Iterable<string>): BaselineGaps {
  const was = new Set(recorded);
  const now = new Set(rendered);
  return {
    unrecorded: [...now].filter((key) => !was.has(key)).sort(),
    vanished: [...was].filter((key) => !now.has(key)).sort(),
  };
}

/**
 * A slide with no recorded digest is ungated, and silence about that is a lie.
 *
 * The gate reports PASS by comparing against what it has; without this it would
 * report PASS over slides it has never seen, which is the one failure a gate
 * with no tolerance cannot afford.
 */
export function assertBaselineCovers(gaps: BaselineGaps): void {
  if (gaps.unrecorded.length === 0) return;
  throw new FidelityError(
    'FID_BASELINE_INCOMPLETE',
    `${String(gaps.unrecorded.length)} slide(s) have no recorded digest and are not gated: ` +
      `${gaps.unrecorded.join(', ')}`,
    gaps.unrecorded[0],
  );
}
