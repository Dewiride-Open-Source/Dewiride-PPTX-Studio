/**
 * The release candidate gate, shown refusing each way an install can cheat.
 *
 * The failure this exists to catch is silent: a package nobody bumped keeps its
 * published version, its siblings' `^0.1.0` is satisfied from the registry, and
 * the suite passes against the previous release. ADR 0046.
 */

import { spawnSync } from 'node:child_process';
import { mkdirSync, mkdtempSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';

import { describe, expect, it } from 'vitest';

import { repoPath } from '../../repo/root.ts';
import {
  websiteRanges,
  fileSpec,
  packedFrom,
  provenanceFailures,
  readPackedRecord,
  scratchManifest,
  tamperedTarballs,
  staleRanges,
  type Consumer,
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

describe('both scripts refuse a tarball that moved after it was packed', () => {
  // Each refuses before it installs or asks the registry anything, so a
  // tampered directory is all either needs to show it. ADR 0059.
  function tamperedDirectory(): string {
    const dir = mkdtempSync(join(tmpdir(), 'release-'));
    const tarballs = join(dir, 'tarballs');
    mkdirSync(tarballs);
    writeFileSync(join(tarballs, 'pptx-studio-xml-0.1.0.tgz'), 'these are not the packed bytes');
    const record = [
      {
        name: '@pptx-studio/xml',
        version: '0.1.0',
        file: 'pptx-studio-xml-0.1.0.tgz',
        integrity: 'sha512-AAAA',
      },
    ];
    writeFileSync(join(tarballs, 'packed.json'), JSON.stringify(record));
    return dir;
  }

  it.each([
    [
      'candidate.ts --packed',
      (dir: string) => [repoPath('tools/release/candidate/candidate.ts'), dir, '--packed'],
    ],
    [
      'publish.ts',
      (dir: string) => [repoPath('tools/release/registry/publish.ts'), join(dir, 'tarballs')],
    ],
  ])(
    '%s',
    (_label, argv) => {
      const dir = tamperedDirectory();
      // No registry to reach, so a regression can fail this test but never publish from it.
      const env = { ...process.env, npm_config_registry: 'http://127.0.0.1:1/' };
      const result = spawnSync(process.execPath, argv(dir), { encoding: 'utf8', env });
      rmSync(dir, { recursive: true, force: true });
      expect(result.status).not.toBe(0);
      expect(result.stderr).toContain('pptx-studio-xml-0.1.0.tgz');
      expect(result.stderr).toMatch(/changed since they were packed|not the ones packed/);
    },
    30_000,
  );
});

describe('what the release hands from prove to publish', () => {
  const RECORD = [
    {
      name: '@pptx-studio/xml',
      version: '0.1.0',
      file: 'pptx-studio-xml-0.1.0.tgz',
      integrity: 'sha512-XML',
    },
    {
      name: '@pptx-studio/cli',
      version: '0.2.0',
      file: 'pptx-studio-cli-0.2.0.tgz',
      integrity: 'sha512-CLI',
    },
  ];
  const on = (file: string): string => INTEGRITY[`/tmp/tgz/${file}`] ?? 'sha512-UNKNOWN';

  it('reads every recorded tarball back', () => {
    expect(readPackedRecord(JSON.stringify(RECORD))).toEqual(RECORD);
  });

  it.each(['../evil.tgz', 'sub/dir.tgz', 'C:\\x.tgz', 'not-a-tarball.js', ''])(
    'refuses a file name that is not a bare tarball: %j',
    (file) => {
      const bad = JSON.stringify([{ ...RECORD[0], file }]);
      expect(() => readPackedRecord(bad)).toThrow(/not a tarball/);
    },
  );

  it('refuses an entry that lacks a field', () => {
    const { integrity: _, ...partial } = RECORD[0]!;
    expect(() => readPackedRecord(JSON.stringify([partial]))).toThrow(/lacks/);
  });

  it('passes tarballs whose bytes are the ones recorded', () => {
    expect(tamperedTarballs(RECORD, on)).toEqual([]);
  });

  it('names a tarball whose bytes changed after it was recorded', () => {
    const swapped = [{ ...RECORD[0]!, integrity: 'sha512-BEFORE' }, RECORD[1]!];
    expect(tamperedTarballs(swapped, on)).toEqual([
      'pptx-studio-xml-0.1.0.tgz: sha512-XML, recorded sha512-BEFORE',
    ]);
  });
});

describe('the scratch manifest', () => {
  const SCOPE = '@pptx-studio/';
  const consumer = (over: Partial<Consumer> = {}): Consumer => ({
    dependencies: {},
    scripts: {},
    overrides: {},
    ...over,
  });

  it('pins every candidate as both a dependency and an override', () => {
    const manifest = scratchManifest(PACKED, consumer(), SCOPE);
    expect(manifest.dependencies['@pptx-studio/xml']).toBe(
      'file:/tmp/tgz/pptx-studio-xml-0.1.0.tgz',
    );
    // npm refuses an override whose spec differs from the direct dependency, and
    // without the override the edges between packages resolve from the registry.
    expect(manifest.overrides).toEqual(manifest.dependencies);
  });

  it('carries the consumer own scripts, so the gate can run them', () => {
    // The gate runs `npm run smoke`, which only exists if the scripts survive
    // replacing the website's manifest.
    const scripts = { smoke: 'node smoke.mjs' };
    expect(scratchManifest(PACKED, consumer({ scripts }), SCOPE).scripts).toEqual(scripts);
  });

  it('carries the consumer own dependencies through without pinning them', () => {
    const manifest = scratchManifest(PACKED, consumer({ dependencies: { next: '16.3.4' } }), SCOPE);
    expect(manifest.dependencies['next']).toBe('16.3.4');
    expect(manifest.overrides['next']).toBeUndefined();
  });

  it('writes a file: specifier with forward slashes, on any platform', () => {
    expect(fileSpec('C:\\work\\tgz\\a.tgz')).toBe('file:C:/work/tgz/a.tgz');
  });

  it('orders the packages by name, so the manifest does not churn', () => {
    const names = Object.keys(scratchManifest(PACKED, consumer(), SCOPE).overrides);
    expect(names).toEqual(['@pptx-studio/cli', '@pptx-studio/xml']);
  });

  describe("the consumer's own overrides", () => {
    const pin = { overrides: { 'mdast-util-to-markdown': '2.1.2' } };

    it('carries a transitive pin as an override and never as a dependency', () => {
      // Without it the website resolves a transitive range fresh and the gate
      // builds a tree the website itself could not. ADR 0057.
      const manifest = scratchManifest(PACKED, consumer(pin), SCOPE);
      expect(manifest.overrides['mdast-util-to-markdown']).toBe('2.1.2');
      expect(manifest.dependencies['mdast-util-to-markdown']).toBeUndefined();
    });

    it('still pins every candidate, with the same spec as its dependency', () => {
      const manifest = scratchManifest(PACKED, consumer(pin), SCOPE);
      for (const entry of PACKED) {
        expect(manifest.overrides[entry.name], entry.name).toBe(fileSpec(entry.filename));
        expect(manifest.overrides[entry.name]).toBe(manifest.dependencies[entry.name]);
      }
    });

    it('refuses an override in scope, which would redirect a candidate', () => {
      const redirect = consumer({ overrides: { '@pptx-studio/xml': '0.0.1' } });
      expect(() => scratchManifest(PACKED, redirect, SCOPE)).toThrow(/only a candidate/);
    });

    it('refuses a dependency in scope, which would displace a candidate pin', () => {
      const own = consumer({ dependencies: { '@pptx-studio/cli': '^0.1.0' } });
      expect(() => scratchManifest(PACKED, own, SCOPE)).toThrow(/candidate pin belongs/);
    });

    it('refuses an override that disagrees with a direct dependency, as npm would', () => {
      const clash = consumer({ dependencies: { next: '16.3.6' }, overrides: { next: '16.3.5' } });
      expect(() => scratchManifest(PACKED, clash, SCOPE)).toThrow(/npm refuses that/);
    });

    it('carries an override that agrees with the direct dependency', () => {
      const agree = consumer({ dependencies: { next: '16.3.6' }, overrides: { next: '16.3.6' } });
      expect(scratchManifest(PACKED, agree, SCOPE).overrides['next']).toBe('16.3.6');
    });

    it('refuses a nested override rather than dropping it', () => {
      const nested = consumer({ overrides: { foo: { bar: '1.0.0' } } });
      expect(() => scratchManifest(PACKED, nested, SCOPE)).toThrow(/nested object/);
    });

    it('orders the consumer overrides by name, ahead of the candidates', () => {
      const two = consumer({ overrides: { zod: '4.0.0', 'mdast-util-to-markdown': '2.1.2' } });
      expect(Object.keys(scratchManifest(PACKED, two, SCOPE).overrides)).toEqual([
        'mdast-util-to-markdown',
        'zod',
        '@pptx-studio/cli',
        '@pptx-studio/xml',
      ]);
    });
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

describe('the website manifest', () => {
  const SCOPE = '@pptx-studio/';
  const written = { '@pptx-studio/xml': '^0.1.0', '@pptx-studio/cli': '^0.2.0', next: '16.3.4' };

  it('passes when every range is the caret of the version this run packed', () => {
    expect(staleRanges(written, PACKED, SCOPE)).toEqual([]);
  });

  it('refuses a caret the packed version has grown past, which can never resolve to it', () => {
    const stale = staleRanges({ ...written, '@pptx-studio/cli': '^0.1.0' }, PACKED, SCOPE);
    expect(stale).toHaveLength(1);
    expect(stale[0]).toContain('^0.1.0');
    expect(stale[0]).toContain('0.2.0');
  });

  it('asks for exactly what pack writes, so a caret one patch behind is stale too', () => {
    // `^0.1.0` would still resolve to 0.1.1; the manifest is written by the
    // release and a hand-edited range is the thing being caught.
    const behind = [{ ...PACKED[0]!, version: '0.1.1' }, PACKED[1]!];
    expect(staleRanges(written, behind, SCOPE)).toEqual([
      '@pptx-studio/xml: has ^0.1.0, packed 0.1.1',
    ]);
  });

  it('refuses a range on a package this run did not pack', () => {
    const stale = staleRanges({ ...written, '@pptx-studio/model': '^1.0.0' }, PACKED, SCOPE);
    expect(stale).toEqual(['@pptx-studio/model: ^1.0.0, not a package this run packed']);
  });

  it('ignores everything outside the scope', () => {
    expect(staleRanges({ ...written, next: '15.0.0' }, PACKED, SCOPE)).toEqual([]);
    expect(websiteRanges({ ...written, next: '15.0.0' }, PACKED, SCOPE)).toEqual({
      ...written,
      next: '15.0.0',
    });
  });

  it('writes the caret of each packed version and keeps the order it was given', () => {
    const behind = [{ ...PACKED[0]!, version: '0.1.1' }, PACKED[1]!];
    const ranges = websiteRanges(written, behind, SCOPE);
    expect(Object.keys(ranges)).toEqual(['@pptx-studio/xml', '@pptx-studio/cli', 'next']);
    expect(ranges['@pptx-studio/xml']).toBe('^0.1.1');
    expect(staleRanges(ranges, behind, SCOPE)).toEqual([]);
  });
});
