/**
 * Every published tarball carries the licence and the notice.
 *
 * ```
 * node tools/repo/check-legal.ts [--fix]
 * ```
 *
 * Apache-2.0 4(a) says a copy of the licence travels with the work and 4(d) says
 * the NOTICE text does too. A monorepo makes that easy to miss, because the two
 * files sit at the root and the thing that ships is a package directory. npm
 * includes a `LICENSE` at a package root whether or not `files` names it; it
 * does not do the same for `NOTICE`, which is exactly the one carrying Apache
 * POI's attribution for the preset geometry.
 */

import { readdirSync, readFileSync, writeFileSync } from 'node:fs';

import { PACKAGES, SCOPE } from './layering/layers.ts';
import { repoPath } from './root.ts';

const CARRIED = ['LICENSE', 'NOTICE'] as const;

interface Manifest {
  readonly name?: string;
  readonly private?: boolean;
  readonly files?: readonly string[];
}

interface Problem {
  readonly directory: string;
  readonly says: string;
}

export function checkLegal(fix: boolean): readonly Problem[] {
  const root = Object.fromEntries(
    CARRIED.map((name) => [name, readFileSync(repoPath(name), 'utf8')]),
  );
  const problems: Problem[] = [];

  for (const name of readdirSync(repoPath('packages'))) {
    if (!(name in PACKAGES)) continue;
    const directory = `packages/${name}`;
    let manifest: Manifest;
    try {
      manifest = JSON.parse(readFileSync(repoPath(directory, 'package.json'), 'utf8')) as Manifest;
    } catch {
      problems.push({ directory, says: 'has no readable package.json' });
      continue;
    }
    if (manifest.private === true) continue;
    if (manifest.name !== `${SCOPE}/${name}`) {
      problems.push({ directory, says: `is named ${String(manifest.name)}` });
      continue;
    }

    for (const carried of CARRIED) {
      const at = repoPath(directory, carried);
      let mine: string | null;
      try {
        mine = readFileSync(at, 'utf8');
      } catch {
        mine = null;
      }
      if (mine !== root[carried]) {
        if (fix) writeFileSync(at, root[carried] ?? '');
        else {
          problems.push({
            directory,
            says: mine === null ? `has no ${carried}` : `has a ${carried} that is not the root one`,
          });
        }
      }
      if (!(manifest.files ?? []).includes(carried)) {
        problems.push({ directory, says: `does not list ${carried} in "files"` });
      }
    }
  }
  return problems;
}

const problems = checkLegal(process.argv.includes('--fix'));
if (problems.length > 0) {
  console.error('legal: a published package must carry the licence and the notice\n');
  for (const problem of problems) console.error(`  ${problem.directory} ${problem.says}`);
  console.error('\n  node tools/repo/check-legal.ts --fix  copies the root files into place');
  process.exitCode = 1;
} else {
  console.log(`legal: every publishable package carries ${CARRIED.join(' and ')}`);
}
