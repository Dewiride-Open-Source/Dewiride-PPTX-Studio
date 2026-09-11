/**
 * Which published packages carry a provenance attestation, and which do not.
 *
 * ```
 * node tools/release/attestations.ts [--require-all]
 * ```
 *
 * Provenance is a claim about the published artifact, so only the registry can
 * answer it - and nothing here asked until ten packages had been described as
 * attested for a month without being. Runs in the canary, which is where
 * questions about the world rather than the commit belong. ADR 0047.
 */

import { readFileSync, readdirSync } from 'node:fs';

import { repoPath } from '../repo/root.ts';

const REGISTRY = 'https://registry.npmjs.org';

interface Manifest {
  readonly name: string;
  readonly private?: boolean;
}

/** `dist.attestations` is where npm records provenance for a published version. */
interface Version {
  readonly dist?: { readonly attestations?: { readonly provenance?: unknown } };
}

const names = readdirSync(repoPath('packages'))
  .map(
    (dir) => JSON.parse(readFileSync(repoPath(`packages/${dir}/package.json`), 'utf8')) as Manifest,
  )
  .filter((manifest) => manifest.private !== true)
  .map((manifest) => manifest.name)
  .sort();

const attested: string[] = [];
const bare: string[] = [];
const absent: string[] = [];

for (const name of names) {
  const encoded = name.replace('/', '%2f');
  const tags = await fetch(`${REGISTRY}/-/package/${encoded}/dist-tags`, { cache: 'no-store' });
  if (tags.status === 404) {
    absent.push(name);
    continue;
  }
  if (!tags.ok) throw new Error(`${name}: the registry answered ${String(tags.status)}`);
  const latest = ((await tags.json()) as { latest?: string }).latest;
  if (latest === undefined) throw new Error(`${name}: no latest tag`);

  const at = await fetch(`${REGISTRY}/${encoded}/${latest}`, { cache: 'no-store' });
  if (!at.ok) throw new Error(`${name}@${latest}: the registry answered ${String(at.status)}`);
  const version = (await at.json()) as Version;
  const line = `${name}@${latest}`;
  if (version.dist?.attestations === undefined) bare.push(line);
  else attested.push(line);
}

for (const line of attested) console.log(`provenance  ${line}`);
for (const line of bare) console.log(`none        ${line}`);
for (const name of absent) console.log(`absent      ${name}`);
console.log(
  `\n${String(attested.length)} attested, ${String(bare.length)} without, ` +
    `${String(absent.length)} not published`,
);

// A release publishes over OIDC, which attests automatically, so anything
// without one was published some other way and a consumer cannot verify it.
if (process.argv.includes('--require-all') && bare.length > 0) {
  console.error(
    `\n${String(bare.length)} published package(s) carry no provenance, so ` +
      '`npm audit signatures` cannot verify them. A patch release over OIDC attests them.',
  );
  process.exitCode = 1;
}
