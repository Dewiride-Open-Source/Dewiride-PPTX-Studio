/**
 * Whether this workflow may publish each package, asked before anything is.
 *
 * npm trades a GitHub ID token for a publish token one package at a time, so
 * the right to publish is a per-package fact and `changeset publish` would
 * discover a missing one halfway through. ADR 0048.
 */

import type { Publishable } from './bootstrap.ts';

/** npm derives the audience from the registry host. */
export function audienceFor(registry: string): string {
  return `npm:${new URL(registry).hostname}`;
}

/** A scoped name as npm escapes it inside a registry path. */
export function escapedName(name: string): string {
  return name.replaceAll('/', '%2f');
}

/** Where an ID token is traded for a token that may publish this one package. */
export function exchangeUrl(registry: string, name: string): string {
  return `${registry}/-/npm/v1/oidc/token/exchange/package/${escapedName(name)}`;
}

/** GitHub's issuer already carries a query string, so the audience appends. */
export function idTokenUrl(requestUrl: string, audience: string): string {
  return `${requestUrl}&audience=${encodeURIComponent(audience)}`;
}

/** What the registry answered for one package. */
export interface Attempt {
  readonly name: string;
  readonly allowed: boolean;
  /** What npm said, for an operator who has to act on a refusal. */
  readonly detail: string;
}

/**
 * The packages a release would send, which is those npm does not already hold.
 *
 * `changeset publish` skips a version the registry already has, so this is the
 * same set it would try, and after `changeset version` it is every bump.
 */
export function pending(
  packages: readonly Publishable[],
  onNpm: ReadonlySet<string>,
): readonly Publishable[] {
  return packages.filter((entry) => !onNpm.has(`${entry.name}@${entry.version}`));
}

/** The attempts that failed and were required to succeed. */
export function refusals(
  attempts: readonly Attempt[],
  required: ReadonlySet<string>,
): readonly Attempt[] {
  return attempts.filter((attempt) => !attempt.allowed && required.has(attempt.name));
}

/**
 * What to do about them, as the message the release fails with.
 *
 * Written out in full because the operator reading it is mid-release, and a
 * trusted publisher is attached on the website rather than from the CLI.
 */
export function missingPublisherInstructions(attempts: readonly Attempt[]): string {
  const list = attempts.map((a) => `  - ${a.name}: ${a.detail}`).join('\n');
  return [
    `${String(attempts.length)} package(s) this release would publish cannot mint a`,
    'publish token from this workflow, so npm would have refused them one at a',
    'time and left the release half shipped. Nothing was published.',
    '',
    list,
    '',
    'A 404 here reads "package not found" and does not mean that: npm answers it',
    'whenever it cannot match this workflow run to a publisher on the package.',
    '',
    'To attach one, per package:',
    '  1. npmjs.com -> the package -> Settings -> Trusted publishing.',
    '  2. GitHub Actions, this repository, workflow release.yml, no environment.',
    '  3. Dispatch the release again. The changesets are still in the tree.',
  ].join('\n');
}
