#!/usr/bin/env node
/**
 * Write `corpus/reject/`, and the manifest that claims it.
 *
 * ```
 * node tools/corpus/tiers/rejects/build-reject.ts --out corpus/reject --manifest
 * node tools/corpus/tiers/rejects/build-reject.ts --out corpus/reject --check
 * ```
 *
 * Same shape as `tools/corpus/tiers/a-generated/build-probes.ts` and for the same reasons.
 * `--out` has no default: `C008-orphan` fails on any file under `corpus/` that
 * no manifest claims, so a generator that wrote there by accident would break
 * the gate for whoever ran it next. `--check` rebuilds every fixture and
 * compares against what is on disk without writing - committed bytes and a
 * recipe that reproduces them are two different claims, and hashing the file
 * only checks one of them.
 *
 * The manifest is generated rather than hand-written, and everything in it
 * comes from a reviewed source: `description`, `rule` and `refusal` are
 * literals in `fixtures.ts`, which is code and is read in review; `sha256` and
 * `bytes` are measured from the file.
 */

import { createHash } from 'node:crypto';
import { mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { join, resolve } from 'node:path';

import { REPO_ROOT as ROOT } from '../../../repo/root.ts';
import { packageBytes } from './deck.ts';
import { outputName, REJECT_FIXTURES, type RejectFixture } from './fixtures.ts';

interface Options {
  readonly out: string;
  readonly manifest: boolean;
  readonly check: boolean;
}

function parseArgs(argv: readonly string[]): Options {
  let out = '';
  let manifest = false;
  let check = false;
  for (let i = 0; i < argv.length; i++) {
    const arg = argv[i];
    if (arg === '--out') out = argv[++i] ?? '';
    else if (arg === '--manifest') manifest = true;
    else if (arg === '--check') check = true;
    else throw new Error('unknown argument ' + String(arg));
  }
  if (out === '') {
    throw new Error('--out <dir> is required: this generator defaults nowhere. See the header.');
  }
  return { out, manifest, check };
}

function sha256(bytes: Uint8Array): string {
  return createHash('sha256').update(bytes).digest('hex');
}

export function buildFixture(fixture: RejectFixture): Uint8Array {
  return packageBytes(fixture.parts);
}

const options = parseArgs(process.argv.slice(2));
const outDir = resolve(ROOT, options.out);

const built = REJECT_FIXTURES.map((fixture) => ({
  fixture,
  bytes: buildFixture(fixture),
}));

if (options.check) {
  const wrong: string[] = [];
  for (const { fixture, bytes } of built) {
    const path = join(outDir, outputName(fixture));
    let onDisk: Uint8Array;
    try {
      onDisk = new Uint8Array(readFileSync(path));
    } catch {
      wrong.push(fixture.id + ': not on disk');
      continue;
    }
    if (sha256(onDisk) !== sha256(bytes)) {
      wrong.push(
        fixture.id +
          ': rebuilt bytes differ (' +
          String(onDisk.byteLength) +
          ' on disk, ' +
          String(bytes.byteLength) +
          ' rebuilt)',
      );
    }
  }
  if (wrong.length > 0) {
    console.error('corpus: ' + String(wrong.length) + ' reject fixture(s) do not rebuild');
    for (const line of wrong) console.error('  ' + line);
    process.exit(1);
  }
  console.log('corpus: ' + String(built.length) + ' reject fixture(s) rebuild byte-for-byte');
  process.exit(0);
}

mkdirSync(outDir, { recursive: true });
for (const { fixture, bytes } of built) {
  writeFileSync(join(outDir, outputName(fixture)), bytes);
  console.log(
    '  ' +
      fixture.id.padEnd(24) +
      String(bytes.byteLength).padStart(6) +
      ' bytes   ' +
      fixture.rule +
      '   ' +
      sha256(bytes).slice(0, 16),
  );
}

if (options.manifest) {
  const manifest = {
    manifestVersion: 1,
    collection: 'reject',
    generatedIn: '1.2',
    adr: 'docs/adr/phase-1-round-trip/0010-the-repair-firewall.md',
    generator: 'tools/corpus/tiers/rejects/build-reject.ts',
    entries: built.map(({ fixture, bytes }) => ({
      id: fixture.id,
      // Not `deck`. A package PowerPoint refuses is not evidence that any
      // feature works, and `C-COV` counts a deck's `features` as coverage.
      kind: 'fixture',
      storage: 'committed',
      license: 'CC0-1.0',
      source: 'synthetic',
      path: outputName(fixture),
      sha256: sha256(bytes),
      bytes: bytes.byteLength,
      description: fixture.description,
      rule: fixture.rule,
      refusal: fixture.refusal,
      tags: ['reject', fixture.rule],
      addedIn: '1.2',
      format: 'pptx',
    })),
  };
  const path = join(outDir, 'manifest.json');
  writeFileSync(path, JSON.stringify(manifest, null, 2) + '\n');
  console.log('\nwrote ' + path);
}
