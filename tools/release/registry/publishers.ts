/**
 * Ask npm whether this workflow may publish each package, before it publishes any.
 *
 * ```
 * node tools/release/registry/publishers.ts
 * ```
 *
 * Runs in the release before `changeset publish`, which goes one package at a
 * time and cannot be rolled back. Only the workflow a publisher names can ask
 * this, so nothing else can ask it on that workflow's behalf. ADR 0049.
 */

import { readFileSync, readdirSync } from 'node:fs';

import { repoPath } from '../../repo/root.ts';
import { publishableIn } from './bootstrap.ts';
import {
  type Attempt,
  audienceFor,
  escapedName,
  exchangeUrl,
  idTokenUrl,
  missingPublisherInstructions,
  pending,
  refusals,
} from './oidc.ts';

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

const onNpm = new Set<string>();
for (const entry of packages) {
  const at = `${REGISTRY}/${escapedName(entry.name)}/${entry.version}`;
  const response = await fetch(at, { cache: 'no-store' });
  if (response.ok) onNpm.add(`${entry.name}@${entry.version}`);
  else if (response.status !== 404) {
    throw new Error(`${entry.name}: the registry answered ${String(response.status)}`);
  }
}

const requestUrl = process.env['ACTIONS_ID_TOKEN_REQUEST_URL'];
const requestToken = process.env['ACTIONS_ID_TOKEN_REQUEST_TOKEN'];
if (requestUrl === undefined || requestToken === undefined) {
  throw new Error('this runs in GitHub Actions, in a job with `permissions: id-token: write`');
}

const issued = await fetch(idTokenUrl(requestUrl, audienceFor(REGISTRY)), {
  headers: { authorization: `Bearer ${requestToken}` },
});
if (!issued.ok) throw new Error(`GitHub refused an ID token: ${String(issued.status)}`);
const idToken = ((await issued.json()) as { value?: string }).value;
if (idToken === undefined) throw new Error('GitHub returned no ID token');

/** npm's refusal carries a message; its success carries a token, which is never read. */
async function attempt(name: string): Promise<Attempt> {
  const response = await fetch(exchangeUrl(REGISTRY, name), {
    method: 'POST',
    headers: { accept: 'application/json', authorization: `Bearer ${idToken}` },
  });
  if (response.ok) return { name, allowed: true, detail: 'may publish from this workflow' };
  const status = String(response.status);
  // Only an answer that means "not you" is a refusal; anything else is unknown,
  // and reporting unknown as refused would block a release npm would accept.
  if (response.status !== 401 && response.status !== 403 && response.status !== 404) {
    throw new Error(`${name}: the exchange answered ${status}, so this is unknown`);
  }
  const said = ((await response.json().catch(() => ({}))) as { message?: string }).message;
  return { name, allowed: false, detail: `${status}, ${said ?? 'no message'}` };
}

const attempts: Attempt[] = [];
for (const entry of packages) attempts.push(await attempt(entry.name));

const willPublish = new Set(pending(packages, onNpm).map((entry) => entry.name));
for (const { name, allowed, detail } of attempts) {
  const mark = allowed ? 'publisher  ' : 'none       ';
  console.log(`${mark}${name}${willPublish.has(name) ? ' (this release)' : ''} - ${detail}`);
}
const allowed = attempts.filter((a) => a.allowed).length;
console.log(
  `\n${String(allowed)} of ${String(attempts.length)} can publish over OIDC, ` +
    `${String(willPublish.size)} pending in this release`,
);

const blocked = refusals(attempts, willPublish);
if (blocked.length > 0) {
  console.error(`\n${missingPublisherInstructions(blocked)}`);
  process.exitCode = 1;
}
