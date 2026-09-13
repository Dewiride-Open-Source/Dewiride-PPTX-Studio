/**
 * Point the website at the versions this tree holds, as a release would publish them.
 *
 * ```
 * node tools/release/candidate/website.ts
 * ```
 *
 * A caret below 1.0 never reaches the next minor, and the website's source is
 * written against the tree. Run by `pnpm version-packages` after `changeset
 * version`; the candidate gate refuses a manifest this did not write. ADR 0051.
 */

import { readFileSync, readdirSync, writeFileSync } from 'node:fs';

import { repoPath } from '../../repo/root.ts';
import { websiteRanges, type Versioned } from './consumer.ts';

const SCOPE = '@pptx-studio/';
const MANIFEST = repoPath('website/package.json');

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

const website = JSON.parse(readFileSync(MANIFEST, 'utf8')) as {
  dependencies: Record<string, string>;
} & Record<string, unknown>;
const dependencies = websiteRanges(website.dependencies, published, SCOPE);

const changed = Object.entries(dependencies).filter(
  ([name, range]) => website.dependencies[name] !== range,
);
writeFileSync(MANIFEST, `${JSON.stringify({ ...website, dependencies }, null, 2)}\n`);
for (const [name, range] of changed) {
  console.log(`${name}: ${String(website.dependencies[name])} -> ${range}`);
}
console.log(`${String(changed.length)} range(s) rewritten in website/package.json`);
