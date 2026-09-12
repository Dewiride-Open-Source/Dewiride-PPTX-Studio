/**
 * Point the example at the versions this tree holds, as a release would publish them.
 *
 * ```
 * node tools/release/candidate/example.ts
 * ```
 *
 * A caret below 1.0 never reaches the next minor, and the example's source is
 * written against the tree. Run by `pnpm version-packages` after `changeset
 * version`; the candidate gate refuses a manifest this did not write. ADR 0051.
 */

import { readFileSync, readdirSync, writeFileSync } from 'node:fs';

import { repoPath } from '../../repo/root.ts';
import { exampleRanges, type Versioned } from './consumer.ts';

const SCOPE = '@pptx-studio/';
const MANIFEST = repoPath('examples/nextjs-studio/package.json');

interface Manifest {
  readonly name: string;
  readonly version: string;
  readonly private?: boolean;
}

const published: Versioned[] = readdirSync(repoPath('packages'))
  .map(
    (dir) => JSON.parse(readFileSync(repoPath(`packages/${dir}/package.json`), 'utf8')) as Manifest,
  )
  .filter((manifest) => manifest.private !== true)
  .map(({ name, version }) => ({ name, version }));

const example = JSON.parse(readFileSync(MANIFEST, 'utf8')) as {
  dependencies: Record<string, string>;
} & Record<string, unknown>;
const dependencies = exampleRanges(example.dependencies, published, SCOPE);

const changed = Object.entries(dependencies).filter(
  ([name, range]) => example.dependencies[name] !== range,
);
writeFileSync(MANIFEST, `${JSON.stringify({ ...example, dependencies }, null, 2)}\n`);
for (const [name, range] of changed) {
  console.log(`${name}: ${String(example.dependencies[name])} -> ${range}`);
}
console.log(`${String(changed.length)} range(s) rewritten in examples/nextjs-studio/package.json`);
