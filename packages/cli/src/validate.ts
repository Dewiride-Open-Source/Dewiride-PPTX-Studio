import { readFileSync, writeFileSync } from 'node:fs';
import { formatReport, validatePackage, type Report } from '@pptx-studio/validate';

/**
 * `pptx-studio validate deck.pptx`
 *
 * The rules live in `@pptx-studio/validate`, which is a browser package with no
 * Node in it - the same split `inspect` has, for the same reason. What is left
 * here is reading a path, writing a stream, and choosing an exit code.
 *
 * ## What this command can and cannot ask
 *
 * Six of the twenty-nine rules compare a package against the package **as it
 * was opened**, and a file dropped on the command line has no such history.
 * They are skipped, and the report says so by name rather than counting them as
 * passes - `--quiet` still prints the skipped list, because a report that
 * silently means less than it looks like is the failure mode this package
 * exists to avoid.
 *
 * That makes `validate` a diagnostic rather than the export gate. The gate is
 * `assertValid`, called from the writer in sub-phase 1.3 with both packages in
 * hand; this is what you run against a file somebody sent you.
 *
 * ## The exit code
 *
 * `1` when anything fatal was found, `0` otherwise; a warning never fails the
 * command. With no baseline every fatal is treated as ours, which is the safe
 * direction when the question "did we break this" cannot be asked at all.
 */

export interface ValidateOptions {
  /** Print the report as JSON instead of as text. */
  readonly json: boolean;
  /** Append each rule's rationale, once per rule that fired. */
  readonly explain: boolean;
  /** Hide warnings; show only what would refuse an export. */
  readonly quiet: boolean;
  /** Write to this path instead of stdout. */
  readonly out: string | null;
}

export const VALIDATE_DEFAULTS: ValidateOptions = {
  json: false,
  explain: false,
  quiet: false,
  out: null,
};

export interface ValidateResult {
  readonly report: Report;
  readonly output: string;
  readonly exitCode: number;
}

export function validateFile(
  path: string,
  options: ValidateOptions = VALIDATE_DEFAULTS,
): ValidateResult {
  const bytes = new Uint8Array(readFileSync(path));
  const report = validatePackage({ bytes });

  const output = options.json
    ? JSON.stringify(report, null, 2) + '\n'
    : formatReport(report, { explain: options.explain, warnings: !options.quiet }) + '\n';

  const fatal = report.findings.filter((finding) => finding.severity === 'fatal').length;
  return { report, output, exitCode: fatal > 0 ? 1 : 0 };
}

/** Run `validateFile` and put the result where the options say. */
export function runValidate(
  path: string,
  options: ValidateOptions,
  write: (text: string) => void,
): number {
  const result = validateFile(path, options);
  if (options.out === null) write(result.output);
  else writeFileSync(options.out, result.output);
  return result.exitCode;
}
