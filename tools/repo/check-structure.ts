/**
 * `pnpm structure` - no directory grows into an unnavigable pile.
 *
 * ```
 * node tools/repo/check-structure.ts
 * ```
 */

import { execFileSync } from 'node:child_process';

import { REPO_ROOT as ROOT } from './root.ts';
import { checkStructure, MAX_SOURCE_FILES } from './structure.ts';

/** Untracked files count: a directory is hard to read before its files are committed. */
function paths(): string[] {
  const list = (args: readonly string[]): string[] =>
    execFileSync('git', args, { cwd: ROOT, encoding: 'utf8' })
      .split('\n')
      .filter((line) => line.length > 0);
  return [...list(['ls-files']), ...list(['ls-files', '--others', '--exclude-standard'])];
}

const problems = checkStructure(paths());

if (problems.length > 0) {
  for (const problem of problems) {
    console.error(
      `${problem.directory} holds ${String(problem.count)} source files, more than the ${String(MAX_SOURCE_FILES)} one directory may hold:`,
    );
    for (const file of problem.files) console.error(`    ${file}`);
    console.error('  group them by what they do - see hard rule 4 in CLAUDE.md\n');
  }
  console.error(`${String(problems.length)} directory/directories over the limit`);
  process.exitCode = 1;
} else {
  console.log(`structure: no directory holds more than ${String(MAX_SOURCE_FILES)} source files`);
}
