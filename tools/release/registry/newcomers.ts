/**
 * Refuse a release that would create a package npm does not have yet.
 *
 * ```
 * node tools/release/newcomers.ts
 * ```
 *
 * Runs before `changeset publish`, which publishes one package at a time and
 * cannot be rolled back. ADR 0046.
 */

import { readFileSync, readdirSync } from 'node:fs';

import { repoPath } from '../repo/root.ts';
import { bootstrapInstructions, newcomers, publishableIn } from './bootstrap.ts';
import { escapedName } from './oidc.ts';

const REGISTRY = 'https://registry.npmjs.org';

interface Manifest {
  readonly name: string;
  readonly version: string;
  readonly private?: boolean;
}

const manifests: Manifest[] = readdirSync(repoPath('packages')).map(
  (dir) => JSON.parse(readFileSync(repoPath(`packages/${dir}/package.json`), 'utf8')) as Manifest,
);

const { ignore } = JSON.parse(readFileSync(repoPath('.changeset/config.json'), 'utf8')) as {
  ignore?: string[];
};

const packages = publishableIn(manifests, ignore ?? []);

/**
 * Whether npm holds the name, asked of its dist-tags rather than its packument.
 *
 * The packument of a freshly created package 404s for a while after the publish
 * that created it, which is the one moment this check is asked about; the
 * dist-tags endpoint answers as soon as a version exists. Measured on
 * `@pptx-studio/render-dom`: packument 404 and dist-tags 200, at the same second.
 */
async function exists(name: string): Promise<boolean> {
  const at = `${REGISTRY}/-/package/${escapedName(name)}/dist-tags`;
  const response = await fetch(at, { cache: 'no-store' });
  if (response.status === 404) return false;
  if (response.ok) return true;
  throw new Error(`${name}: the registry answered ${String(response.status)}, so this is unknown`);
}

const published = new Set<string>();
for (const entry of packages) {
  if (await exists(entry.name)) published.add(entry.name);
}

const missing = newcomers(packages, published);
console.log(
  `${String(packages.length)} package(s) would publish, ${String(missing.length)} of them new`,
);
if (missing.length > 0) {
  console.error(bootstrapInstructions(missing));
  process.exitCode = 1;
}
