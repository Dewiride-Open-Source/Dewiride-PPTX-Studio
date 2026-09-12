/**
 * What a consumer can make of every version this scope has published.
 *
 * ```
 * node tools/release/registry/attestations.ts
 * ```
 *
 * Provenance is a claim about the published artifact, so only the registry can
 * answer it. Runs in the canary, where questions about the world belong, and
 * fails on a version a consumer can neither verify nor be warned off. ADR 0051.
 */

import { readFileSync, readdirSync } from 'node:fs';

import { repoPath } from '../../repo/root.ts';
import { escapedName } from './oidc.ts';
import {
  standings,
  unverifiable,
  unverifiableMessage,
  type Packument,
  type VersionStanding,
} from './provenance.ts';

const REGISTRY = 'https://registry.npmjs.org';

interface Manifest {
  readonly name: string;
  readonly private?: boolean;
}

const names = readdirSync(repoPath('packages'))
  .map(
    (dir) => JSON.parse(readFileSync(repoPath(`packages/${dir}/package.json`), 'utf8')) as Manifest,
  )
  .filter((manifest) => manifest.private !== true)
  .map((manifest) => manifest.name)
  .sort();

const all: VersionStanding[] = [];
const absent: string[] = [];

for (const name of names) {
  // The whole packument: it lags only the minutes after a name is created, which
  // is `newcomers.ts`'s moment and never a daily monitor's.
  const response = await fetch(`${REGISTRY}/${escapedName(name)}`, { cache: 'no-store' });
  if (response.status === 404) {
    absent.push(name);
    continue;
  }
  if (!response.ok) throw new Error(`${name}: the registry answered ${String(response.status)}`);
  all.push(...standings(name, (await response.json()) as Packument));
}

const width = Math.max(...all.map((entry) => entry.standing.length));
for (const entry of all) {
  console.log(
    `${entry.standing.padEnd(width)}  ${entry.name}@${entry.version}${entry.latest ? '  latest' : ''}`,
  );
}
for (const name of absent) console.log(`absent      ${name}`);
const count = (standing: string): number => all.filter((e) => e.standing === standing).length;
console.log(
  `\n${String(count('provenance'))} with provenance, ${String(count('deprecated'))} deprecated, ` +
    `${String(count('neither'))} neither, ${String(absent.length)} not published`,
);

const failures = unverifiable(all);
if (failures.length > 0) {
  console.error(`\n${unverifiableMessage(failures)}`);
  process.exitCode = 1;
}
