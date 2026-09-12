/**
 * What the canary makes of each published version, decided over packuments
 * shaped like the registry's and never over the network. ADR 0051.
 */

import { describe, expect, it } from 'vitest';

import {
  standingOf,
  standings,
  unverifiable,
  unverifiableMessage,
  type Packument,
  type VersionStanding,
} from './provenance.ts';

const ATTESTED = {
  dist: { attestations: { url: 'https://registry.npmjs.org/-/npm/v1/attestations/x' } },
};
const DEPRECATED = { deprecated: 'Published without a provenance attestation. Use 0.1.1.' };
const BARE = { dist: {} };

const packument = (latest: string, versions: NonNullable<Packument['versions']>): Packument => ({
  'dist-tags': { latest },
  versions,
});

describe('one version', () => {
  it('has provenance when the registry recorded an attestation', () => {
    expect(standingOf(ATTESTED)).toBe('provenance');
  });

  it('is deprecated when a message was left for whoever installs it', () => {
    expect(standingOf(DEPRECATED)).toBe('deprecated');
  });

  it('is neither when nobody can verify it and nobody was warned', () => {
    expect(standingOf(BARE)).toBe('neither');
    expect(standingOf({})).toBe('neither');
  });

  it('reads an attested and deprecated version as verifiable', () => {
    expect(standingOf({ ...ATTESTED, ...DEPRECATED })).toBe('provenance');
  });

  it('reads an emptied deprecation as no deprecation, which is how npm lifts one', () => {
    expect(standingOf({ deprecated: '' })).toBe('neither');
  });
});

describe('a package', () => {
  it('marks latest from the tag, not from the order the versions are listed in', () => {
    const all = standings(
      '@pptx-studio/xml',
      packument('0.1.1', { '0.2.0-next.1': ATTESTED, '0.1.0': BARE, '0.1.1': ATTESTED }),
    );
    expect(all.filter((entry) => entry.latest).map((entry) => entry.version)).toEqual(['0.1.1']);
    expect(all.map((entry) => entry.standing)).toEqual(['provenance', 'neither', 'provenance']);
  });

  it('refuses a packument with no latest tag rather than guessing one', () => {
    expect(() => standings('@pptx-studio/xml', { versions: { '0.1.0': BARE } })).toThrowError(
      /no latest tag/,
    );
  });

  it('is empty for a packument that lists no versions', () => {
    expect(standings('@pptx-studio/xml', { 'dist-tags': { latest: '0.1.0' } })).toEqual([]);
  });
});

describe('what fails the canary', () => {
  const entry = (
    version: string,
    standing: VersionStanding['standing'],
    latest = false,
  ): VersionStanding => ({ name: '@pptx-studio/xml', version, standing, latest });

  it('is nothing when every version is attested', () => {
    expect(
      unverifiable([entry('0.1.0', 'provenance'), entry('0.1.1', 'provenance', true)]),
    ).toEqual([]);
  });

  it('is nothing when an older version was deprecated instead', () => {
    expect(
      unverifiable([entry('0.1.0', 'deprecated'), entry('0.1.1', 'provenance', true)]),
    ).toEqual([]);
  });

  it('is an older version that was neither, named with its version', () => {
    const failures = unverifiable([entry('0.1.0', 'neither'), entry('0.1.1', 'provenance', true)]);
    expect(failures).toEqual([entry('0.1.0', 'neither')]);
    expect(unverifiableMessage(failures)).toContain('@pptx-studio/xml@0.1.0');
    expect(unverifiableMessage(failures)).toContain('npm deprecate');
  });

  it('is a latest that is only deprecated, because a consumer following a caret lands on it', () => {
    const failures = unverifiable([
      entry('0.1.0', 'provenance'),
      entry('0.1.1', 'deprecated', true),
    ]);
    expect(failures.map((f) => f.version)).toEqual(['0.1.1']);
    expect(unverifiableMessage(failures)).toContain('(latest)');
  });

  it('is a latest that is neither', () => {
    expect(unverifiable([entry('0.1.0', 'neither', true)])).toHaveLength(1);
  });
});
