import type { Location } from './location.js';
import { ruleById, type RuleCategory, type RuleId, type Severity } from '../rules/rules.js';

/**
 * What validation produces.
 *
 * ## Origin, and why it is computed rather than guessed
 *
 * The hardest question this package has to answer is not "is this file valid".
 * It is **"did we break it"**, and those are different questions with different
 * consequences. A fatal finding in a deck the user imported ten seconds ago is
 * information; the same finding in a deck they just edited is a bug in us, and
 * handing over the bytes would give them a repair prompt with no explanation.
 *
 * Refusing both would be the strict-looking choice and it is the wrong one. It
 * would mean that a deck with one pre-existing defect - a dangling image
 * relationship, say, which PowerPoint tolerates and which is common in files
 * that have been through three other tools - could be opened in this editor and
 * never saved again. The editor would be refusing to give the user back their
 * own file over a problem it did not cause and cannot fix. `PartStore.write`
 * already made exactly this call for dangling relationships, and this is the
 * same call generalised to all thirty-one rules.
 *
 * So `origin` is not a per-rule judgement call. It is computed by running the
 * rules a second time against the package **as it was opened** and differencing
 * the two reports: a finding that is in both is `inherited`, one that is only in
 * the new report is `introduced`. That is exact, it needs no rule to reason
 * about history, and it cannot drift from what the rules actually do.
 *
 * The second pass only happens when the first found something fatal, so a clean
 * export - the overwhelmingly common case - pays nothing for it.
 */

export type Origin =
  /** The finding is in the package we are about to write and was not in the original. */
  | 'introduced'
  /** The finding was already there when the package was opened. */
  | 'inherited'
  /** No baseline was supplied, so the question was not asked. */
  | 'unknown';

export interface Finding {
  readonly rule: RuleId;
  readonly severity: Severity;
  readonly category: RuleCategory;
  readonly where: Location;
  /**
   * What is wrong, in one sentence, naming the values involved.
   *
   * Written to be read on its own: a message that says "invalid id" and leaves
   * the reader to go and look is a message that costs more than it saves.
   */
  readonly message: string;
  readonly origin: Origin;
}

/** A rule that did not run, and why. Never silent - see `Report.skipped`. */
export interface SkippedRule {
  readonly rule: RuleId;
  readonly why: string;
}

/** Something that could not be read. Not a rule violation; a gap in coverage. */
export interface ReadProblem {
  readonly part: string;
  readonly message: string;
}

export interface Report {
  /** Sorted: fatal before warning, then by part, then by offset. */
  readonly findings: readonly Finding[];
  /** Rules that ran. */
  readonly checked: readonly RuleId[];
  /**
   * Rules that did not run, and why.
   *
   * Present so that a report can never quietly mean less than it looks like it
   * means. Six rules need the package as it was opened; asked to validate a
   * package with no baseline they report nothing, and "reported nothing" and
   * "found nothing" are the same shape on a screen and opposite in meaning.
   */
  readonly skipped: readonly SkippedRule[];
  /** Parts that could not be read, and the rules that therefore did not see them. */
  readonly problems: readonly ReadProblem[];
  /** Fatal findings this session introduced. The number that refuses an export. */
  readonly blocking: number;
  /** True when nothing blocking was found. Not the same as "no findings". */
  readonly ok: boolean;
}

/** Fatal, then warning; then by part; then by offset; then by rule. */
function compareFindings(a: Finding, b: Finding): number {
  if (a.severity !== b.severity) return a.severity === 'fatal' ? -1 : 1;
  if (a.where.part !== b.where.part) return a.where.part.localeCompare(b.where.part);
  const offsetA = a.where.offset ?? -1;
  const offsetB = b.where.offset ?? -1;
  if (offsetA !== offsetB) return offsetA - offsetB;
  return a.rule.localeCompare(b.rule);
}

/**
 * The key two reports are differenced on.
 *
 * Includes the message, because the message carries the instance - the id that
 * collided, the guide that was not defined. Two findings of the same rule at
 * the same place with different values are different findings, and treating
 * them as one would let an introduced defect hide behind an inherited one.
 */
export function findingKey(finding: Finding): string {
  return [finding.rule, finding.where.part, finding.where.xpath ?? '', finding.message].join(
    '\u0000',
  );
}

export function buildReport(input: {
  readonly findings: readonly Finding[];
  readonly checked: readonly RuleId[];
  readonly skipped: readonly SkippedRule[];
  readonly problems: readonly ReadProblem[];
}): Report {
  const findings = [...input.findings].sort(compareFindings);
  const blocking = findings.filter(
    (finding) => finding.severity === 'fatal' && finding.origin !== 'inherited',
  ).length;
  return {
    findings,
    checked: input.checked,
    skipped: input.skipped,
    problems: input.problems,
    blocking,
    ok: blocking === 0,
  };
}

export function isReport(value: unknown): value is Report {
  return (
    typeof value === 'object' &&
    value !== null &&
    Array.isArray((value as Report).findings) &&
    typeof (value as Report).blocking === 'number'
  );
}

export interface FormatOptions {
  /** Include `inherited` findings. Default true. */
  readonly inherited?: boolean;
  /** Include warnings. Default true. */
  readonly warnings?: boolean;
  /** Append each rule's `why`, once per rule. Default false. */
  readonly explain?: boolean;
}

/**
 * A report as text.
 *
 * Grouped by part rather than by rule, because the question a reader has is
 * "what is wrong with this file", and a file is a set of parts. Grouping by
 * rule is the right shape for the corpus gate and the wrong one here.
 */
export function formatReport(report: Report, options: FormatOptions = {}): string {
  const showInherited = options.inherited ?? true;
  const showWarnings = options.warnings ?? true;

  const shown = report.findings.filter(
    (finding) =>
      (showInherited || finding.origin !== 'inherited') &&
      (showWarnings || finding.severity === 'fatal'),
  );

  const lines: string[] = [];
  if (shown.length === 0) {
    // 'nothing was found' and 'nothing survived the filter' are different
    // sentences, and a reader who asked to hide inherited findings has to be
    // able to tell which one they are looking at.
    lines.push(report.findings.length === 0 ? 'no findings' : 'no findings shown');
  } else {
    const byPart = new Map<string, Finding[]>();
    for (const finding of shown) {
      const bucket = byPart.get(finding.where.part);
      if (bucket) bucket.push(finding);
      else byPart.set(finding.where.part, [finding]);
    }
    for (const [part, findings] of byPart) {
      lines.push('');
      lines.push('  ' + (part === '/' ? '(package)' : part));
      for (const finding of findings) {
        const tag =
          finding.severity === 'fatal'
            ? finding.origin === 'inherited'
              ? 'fatal (already there)'
              : 'fatal'
            : 'warning';
        lines.push('    ' + finding.rule + '  ' + tag);
        if (finding.where.xpath !== null) lines.push('      ' + finding.where.xpath);
        lines.push('      ' + finding.message);
      }
    }
  }

  for (const problem of report.problems) {
    lines.push('');
    lines.push('  ' + problem.part + ' could not be read: ' + problem.message);
  }
  for (const skipped of report.skipped) {
    lines.push('');
    lines.push('  ' + skipped.rule + ' did not run: ' + skipped.why);
  }

  if (options.explain === true) {
    const seen = new Set<RuleId>();
    for (const finding of shown) {
      if (seen.has(finding.rule)) continue;
      seen.add(finding.rule);
      const rule = ruleById(finding.rule);
      if (rule === undefined) continue;
      lines.push('');
      lines.push('  ' + rule.id + ' - ' + rule.title);
      lines.push('    ' + rule.why);
    }
  }

  lines.push('');
  lines.push(
    'validate: ' +
      String(report.checked.length) +
      ' rule(s), ' +
      String(report.findings.length) +
      ' finding(s), ' +
      String(report.blocking) +
      ' blocking',
  );
  return lines.join('\n');
}
