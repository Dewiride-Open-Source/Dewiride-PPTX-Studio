/**
 * What a consumer can make of each published version: verify it, or be warned
 * off it, or neither.
 *
 * Provenance is per version and cannot be added later, so `npm deprecate` is
 * the one thing that reaches a consumer who pins one without it. ADR 0051.
 */

/** One published version, as the packument writes it. */
export interface PublishedVersion {
  readonly dist?: { readonly attestations?: unknown };
  /** The deprecation message; an empty one is how `npm deprecate` lifts it. */
  readonly deprecated?: string;
}

export interface Packument {
  readonly 'dist-tags'?: Readonly<Record<string, string>>;
  readonly versions?: Readonly<Record<string, PublishedVersion>>;
}

export type Standing = 'provenance' | 'deprecated' | 'neither';

export interface VersionStanding {
  readonly name: string;
  readonly version: string;
  readonly standing: Standing;
  readonly latest: boolean;
}

/** Provenance first: a deprecated version a consumer can still verify is verifiable. */
export function standingOf(version: PublishedVersion): Standing {
  if (version.dist?.attestations !== undefined) return 'provenance';
  if (typeof version.deprecated === 'string' && version.deprecated !== '') return 'deprecated';
  return 'neither';
}

/** Every version the registry holds for one package, `latest` read from the tag. */
export function standings(name: string, packument: Packument): readonly VersionStanding[] {
  const latest = packument['dist-tags']?.['latest'];
  if (latest === undefined) throw new Error(`${name}: no latest tag`);
  return Object.entries(packument.versions ?? {}).map(([version, published]) => ({
    name,
    version,
    standing: standingOf(published),
    latest: version === latest,
  }));
}

/** A version nobody can verify and nobody was warned off, or a `latest` that is not attested. */
export function unverifiable(all: readonly VersionStanding[]): readonly VersionStanding[] {
  return all.filter(
    (entry) => entry.standing === 'neither' || (entry.latest && entry.standing !== 'provenance'),
  );
}

export function unverifiableMessage(failures: readonly VersionStanding[]): string {
  const lines = failures.map(
    (entry) =>
      `  ${entry.name}@${entry.version}${entry.latest ? ' (latest)' : ''}: ${entry.standing}`,
  );
  return (
    `${String(failures.length)} published version(s) cannot be verified by \`npm audit signatures\` ` +
    'and carry no deprecation saying so:\n' +
    `${lines.join('\n')}\n` +
    'A version a release has replaced is deprecated by hand - `npm deprecate <name>@<version> ' +
    '"<why>"`, one OTP per package. A latest without provenance is republished over OIDC.'
  );
}
