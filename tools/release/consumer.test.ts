/**
 * The release candidate gate, shown refusing each way an install can cheat.
 *
 * The failure this exists to catch is silent: a package nobody bumped keeps its
 * published version, its siblings' `^0.1.0` is satisfied from the registry, and
 * the suite passes against the previous release. ADR 0046.
 */

import { describe, expect, it } from 'vitest';

import {
  fileSpec,
  packedFrom,
  provenanceFailures,
  scratchManifest,
  type Lockfile,
  type LockEntry,
  type Packed,
} from './consumer.ts';

const PACKED: readonly Packed[] = [
  { name: '@pptx-studio/xml', version: '0.1.0', filename: '/tmp/tgz/pptx-studio-xml-0.1.0.tgz' },
  { name: '@pptx-studio/cli', version: '0.2.0', filename: '/tmp/tgz/pptx-studio-cli-0.2.0.tgz' },
];

const INTEGRITY: Record<string, string> = {
  '/tmp/tgz/pptx-studio-xml-0.1.0.tgz': 'sha512-XML',
  '/tmp/tgz/pptx-studio-cli-0.2.0.tgz': 'sha512-CLI',
};
const integrityOf = (filename: string): string => INTEGRITY[filename] ?? 'sha512-UNKNOWN';

function lockOf(packages: Record<string, LockEntry>): Lockfile {
  return { packages };
}

/** What a clean run produces: both packages, from their own tarballs. */
const GREEN = lockOf({
  '': { name: 'pptx-studio-release-candidate' },
  'node_modules/@pptx-studio/xml': {
    name: '@pptx-studio/xml',
    resolved: 'file:/tmp/tgz/pptx-studio-xml-0.1.0.tgz',
    integrity: 'sha512-XML',
  },
  'node_modules/@pptx-studio/cli': {
    name: '@pptx-studio/cli',
    resolved: 'file:/tmp/tgz/pptx-studio-cli-0.2.0.tgz',
    integrity: 'sha512-CLI',
  },
});

describe('reading what pnpm packed', () => {
  /** The real shape: an array, and every entry carries a nested `files` list. */
  const REPORT = JSON.stringify(
    [
      {
        name: '@pptx-studio/opc',
        version: '0.1.0',
        filename: '/tmp/tgz/pptx-studio-opc-0.1.0.tgz',
        files: [{ path: 'dist/index.js' }, { path: 'package.json' }],
      },
      {
        name: '@pptx-studio/xml',
        version: '0.1.0',
        filename: '/tmp/tgz/pptx-studio-xml-0.1.0.tgz',
        files: [{ path: 'README.md' }],
      },
    ],
    null,
    2,
  );

  it('reads every tarball out of a nested report', () => {
    expect(packedFrom(REPORT).map((entry) => entry.name)).toEqual([
      '@pptx-studio/opc',
      '@pptx-studio/xml',
    ]);
  });

  it('takes the filename from the report rather than rebuilding it', () => {
    // The default flattens the scope, so `@pptx-studio/opc` is not a path.
    expect(packedFrom(REPORT)[0]?.filename).toBe('/tmp/tgz/pptx-studio-opc-0.1.0.tgz');
  });

  it('reads several documents when the report is not one array', () => {
    const concatenated = `{"name":"a","version":"1.0.0","filename":"/tmp/a.tgz","files":[{"path":"x"}]}
{"name":"b","version":"1.0.0","filename":"/tmp/b.tgz","files":[]}`;
    expect(packedFrom(concatenated).map((entry) => entry.name)).toEqual(['a', 'b']);
  });

  it('is not confused by a brace inside a string', () => {
    const tricky = '[{"name":"a","version":"1.0.0","filename":"/tmp/a}b\\".tgz","files":[]}]';
    expect(packedFrom(tricky)[0]?.filename).toBe('/tmp/a}b".tgz');
  });

  it('ignores anything that is not a packed tarball', () => {
    expect(packedFrom('{"lifecycle":"prepack"} not json at all [1,2,3]')).toEqual([]);
  });
});

describe('the scratch manifest', () => {
  it('pins every candidate as both a dependency and an override', () => {
    const manifest = scratchManifest(PACKED, {});
    expect(manifest.dependencies['@pptx-studio/xml']).toBe(
      'file:/tmp/tgz/pptx-studio-xml-0.1.0.tgz',
    );
    // npm refuses an override whose spec differs from the direct dependency, and
    // without the override the edges between packages resolve from the registry.
    expect(manifest.overrides).toEqual(manifest.dependencies);
  });

  it('carries the consumer own dependencies through without pinning them', () => {
    const manifest = scratchManifest(PACKED, { next: '16.3.4' });
    expect(manifest.dependencies['next']).toBe('16.3.4');
    expect(manifest.overrides['next']).toBeUndefined();
  });

  it('writes a file: specifier with forward slashes, on any platform', () => {
    expect(fileSpec('C:\\work\\tgz\\a.tgz')).toBe('file:C:/work/tgz/a.tgz');
  });

  it('orders the packages by name, so the manifest does not churn', () => {
    const names = Object.keys(scratchManifest(PACKED, {}).overrides);
    expect(names).toEqual(['@pptx-studio/cli', '@pptx-studio/xml']);
  });
});

describe('provenance', () => {
  it('passes when every package came from the tarball this run packed', () => {
    expect(provenanceFailures(GREEN, PACKED, '@pptx-studio/', integrityOf)).toEqual([]);
  });

  it('fails the silent case: a package resolved from the registry instead', () => {
    const fallen = lockOf({
      ...GREEN.packages,
      'node_modules/@pptx-studio/xml': {
        name: '@pptx-studio/xml',
        resolved: 'https://registry.npmjs.org/@pptx-studio/xml/-/xml-0.1.0.tgz',
        integrity: 'sha512-PUBLISHED',
      },
    });
    const failures = provenanceFailures(fallen, PACKED, '@pptx-studio/', integrityOf);
    expect(failures).toHaveLength(1);
    expect(failures[0]).toContain('registry.npmjs.org');
  });

  it('fails a symlink, which is the workspace link farm leaking in', () => {
    const linked = lockOf({
      ...GREEN.packages,
      'node_modules/@pptx-studio/cli': { name: '@pptx-studio/cli', link: true },
    });
    expect(provenanceFailures(linked, PACKED, '@pptx-studio/', integrityOf)[0]).toContain(
      'a symlink',
    );
  });

  it('fails a tarball whose bytes are not the ones packed', () => {
    const stale = lockOf({
      ...GREEN.packages,
      'node_modules/@pptx-studio/cli': {
        name: '@pptx-studio/cli',
        resolved: 'file:/tmp/tgz/pptx-studio-cli-0.2.0.tgz',
        integrity: 'sha512-FROM-A-PREVIOUS-RUN',
      },
    });
    expect(provenanceFailures(stale, PACKED, '@pptx-studio/', integrityOf)[0]).toContain(
      'is not the tarball',
    );
  });

  it('fails a package that was packed and never reached the tree', () => {
    const missing = lockOf({
      '': { name: 'pptx-studio-release-candidate' },
      'node_modules/@pptx-studio/xml': {
        name: '@pptx-studio/xml',
        resolved: 'file:/tmp/tgz/pptx-studio-xml-0.1.0.tgz',
        integrity: 'sha512-XML',
      },
    });
    expect(provenanceFailures(missing, PACKED, '@pptx-studio/', integrityOf)).toEqual([
      '@pptx-studio/cli was packed but never installed',
    ]);
  });

  it('reads the name out of the path when the entry does not carry one', () => {
    const unnamed = lockOf({
      'node_modules/@pptx-studio/xml': {
        resolved: 'https://registry.npmjs.org/@pptx-studio/xml/-/xml-0.1.0.tgz',
      },
      'node_modules/@pptx-studio/cli': {
        resolved: 'file:/tmp/tgz/pptx-studio-cli-0.2.0.tgz',
        integrity: 'sha512-CLI',
      },
    });
    const failures = provenanceFailures(unnamed, PACKED, '@pptx-studio/', integrityOf);
    expect(failures).toHaveLength(1);
    expect(failures[0]).toContain('registry.npmjs.org');
  });

  it('ignores everything outside the scope, which is where real dependencies live', () => {
    const withDeps = lockOf({
      ...GREEN.packages,
      'node_modules/fflate': {
        name: 'fflate',
        resolved: 'https://registry.npmjs.org/fflate/-/fflate-0.8.3.tgz',
        integrity: 'sha512-FFLATE',
      },
    });
    expect(provenanceFailures(withDeps, PACKED, '@pptx-studio/', integrityOf)).toEqual([]);
  });
});
