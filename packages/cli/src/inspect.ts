import { readFileSync, writeFileSync } from 'node:fs';
import { censusPackage, formatCensus, type PackageCensus } from '@pptx-studio/census';

/**
 * `pptx-studio inspect deck.pptx`
 *
 * Everything this command knows about `.pptx` lives in `@pptx-studio/census`,
 * which is a browser package with no Node in it. That split is deliberate and
 * it is the reason the census exists as a package at all: the browser explorer
 * in `apps/studio` computes exactly the same object in a Web Worker, so a
 * question answered here and a question answered in a tab cannot drift.
 *
 * What is left for this file is the two things a browser cannot do - read a
 * path off the filesystem, and write to a stream - plus argument handling.
 */

export interface InspectOptions {
  readonly json: boolean;
  readonly parts: boolean;
  readonly namespaces: boolean;
  readonly top: number;
  /** Write to this path instead of stdout. */
  readonly out: string | null;
}

export const INSPECT_DEFAULTS: InspectOptions = {
  json: false,
  parts: false,
  namespaces: false,
  top: 15,
  out: null,
};

export interface InspectResult {
  readonly census: PackageCensus;
  readonly output: string;
  /**
   * Process exit code.
   *
   * `1` when the census found something it classes as an error, `0` otherwise -
   * a warning or a note never fails the command. That distinction is what makes
   * `inspect` usable in a script: "did this deck load, and is anything in it
   * structurally broken" is a yes/no question, while "is anything in it
   * unusual" is not.
   */
  readonly exitCode: number;
}

/** Run a census over one file and render it. Does no I/O of its own beyond the read. */
export function inspectFile(
  path: string,
  options: InspectOptions = INSPECT_DEFAULTS,
): InspectResult {
  const bytes = new Uint8Array(readFileSync(path));
  const census = censusPackage(bytes);

  const output = options.json
    ? JSON.stringify(census, null, 2) + '\n'
    : formatCensus(census, {
        parts: options.parts,
        namespaces: options.namespaces,
        top: options.top,
      });

  const errors = census.problems.filter((problem) => problem.severity === 'error').length;
  return { census, output, exitCode: errors > 0 ? 1 : 0 };
}

/** Run `inspectFile` and put the result where the options say. */
export function runInspect(
  path: string,
  options: InspectOptions,
  write: (text: string) => void,
): number {
  const result = inspectFile(path, options);
  if (options.out === null) write(result.output);
  else writeFileSync(options.out, result.output);
  return result.exitCode;
}
