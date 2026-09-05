#!/usr/bin/env node
/**
 * Write Tier C - the decks `packages/opc`'s own writer produces - and the
 * manifest that claims them.
 *
 * ```
 * pnpm build
 * node tools/corpus/tiers/c-written/build-written.ts --out corpus/written --manifest
 * node tools/corpus/tiers/c-written/build-written.ts --check --out corpus/written
 * npx prettier --write corpus/written/manifest.json
 * ```
 *
 * `--out` has no default, for the reason `build-probes.ts` gives: `C008-orphan`
 * fails on any file under `corpus/` that no manifest claims, so a generator that
 * wrote there by accident would break the gate for whoever ran it next.
 *
 * `--check` is `C-REGEN` for this tier. Unlike Tier B it applies at all, because
 * a no-op write is deterministic: `passthroughEntry` moves every already-
 * compressed stream across without inflating it, so the output does not depend
 * on which `fflate` is installed. `decks.test.ts` runs the same comparison under
 * Vitest and needs no build, so `--check` is the convenience and the test is the
 * gate.
 *
 * ## Why this imports `dist`
 *
 * The same reason `write-census-keys.ts` does, and at the same cost. Nothing
 * links workspace packages into `tools/`, and the resolution that lets a Vitest
 * file write `import { PartStore } from '@pptx-studio/opc'` comes from Vitest's
 * alias table rather than from Node. So this script needs `pnpm build` first and
 * says so, and the test - which does not - is where the claim is enforced.
 *
 * There is a second reason to prefer `dist` here rather than merely tolerate it:
 * the deck committed under `corpus/written` should be the output of the writer
 * this project *ships*, not of its sources. If `tsdown` ever changed the emitted
 * bytes, a fixture built from `src` would hide it and this one cannot.
 */

import { createHash } from 'node:crypto';
import { mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { join, resolve } from 'node:path';
import { pathToFileURL } from 'node:url';

import { REPO_ROOT as ROOT } from '../../../repo/root.ts';
import { WRITTEN_DECKS, type WrittenDeck } from './decks.ts';

const DIST = resolve(ROOT, 'packages/opc/dist/index.js');

interface Options {
  readonly out: string;
  readonly only: string | null;
  readonly manifest: boolean;
  readonly check: boolean;
}

function parseArgs(argv: readonly string[]): Options {
  let out = '';
  let only: string | null = null;
  let manifest = false;
  let check = false;
  for (let i = 0; i < argv.length; i++) {
    const arg = argv[i];
    if (arg === '--out') out = argv[++i] ?? '';
    else if (arg === '--id') only = argv[++i] ?? null;
    else if (arg === '--manifest') manifest = true;
    else if (arg === '--check') check = true;
    else throw new Error('unknown argument ' + String(arg));
  }
  if (out === '') {
    throw new Error('--out <dir> is required: this generator defaults nowhere. See the header.');
  }
  return { out, only, manifest, check };
}

const options = parseArgs(process.argv.slice(2));
const outDir = resolve(ROOT, options.out);

const decks =
  options.only === null ? WRITTEN_DECKS : WRITTEN_DECKS.filter((deck) => deck.id === options.only);
if (decks.length === 0) throw new Error('no deck with id ' + String(options.only));

const opc = (await import(pathToFileURL(DIST).href)) as {
  PartStore: { open(bytes: Uint8Array): { write(): Uint8Array } };
};

const sha256 = (bytes: Uint8Array): string => createHash('sha256').update(bytes).digest('hex');

interface Built {
  readonly deck: WrittenDeck;
  readonly bytes: Uint8Array;
  readonly sha256: string;
  readonly entries: number;
}

/** ZIP entry count, read from the end-of-central-directory record. */
function countEntries(bytes: Uint8Array): number {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  for (let at = bytes.length - 22; at >= 0; at--) {
    if (view.getUint32(at, true) === 0x06054b50) return view.getUint16(at + 10, true);
  }
  throw new Error('no end-of-central-directory record');
}

function build(deck: WrittenDeck): Built {
  const source = new Uint8Array(readFileSync(resolve(ROOT, deck.source)));
  const actual = sha256(source);
  if (actual !== deck.sourceSha256) {
    throw new Error(
      deck.id +
        ' is built from ' +
        deck.source +
        ', which has changed.\n  pinned  ' +
        deck.sourceSha256 +
        '\n  on disk ' +
        actual +
        '\nRe-pin it in decks.ts once you have looked at what moved, and expect this deck to ' +
        'change with it.',
    );
  }
  const bytes = opc.PartStore.open(source).write();
  return { deck, bytes, sha256: sha256(bytes), entries: countEntries(bytes) };
}

const built = decks.map(build);

if (options.check) {
  let bad = 0;
  for (const item of built) {
    const path = join(outDir, item.deck.id + '.pptx');
    let onDisk: Uint8Array;
    try {
      onDisk = new Uint8Array(readFileSync(path));
    } catch {
      console.error('  MISSING  ' + item.deck.id);
      bad += 1;
      continue;
    }
    const hash = sha256(onDisk);
    if (hash === item.sha256) continue;
    console.error(
      '  DIFFERS  ' + item.deck.id + '\n    on disk ' + hash + '\n    rebuilt ' + item.sha256,
    );
    bad += 1;
  }
  if (bad > 0) {
    console.error('\ncorpus: ' + String(bad) + ' deck(s) do not rebuild to their committed bytes');
    process.exit(1);
  }
  console.log('corpus: ' + String(built.length) + ' deck(s) rebuild byte-for-byte');
  process.exit(0);
}

mkdirSync(outDir, { recursive: true });
for (const item of built) {
  writeFileSync(join(outDir, item.deck.id + '.pptx'), item.bytes);
  console.log(
    '  ' +
      item.deck.id.padEnd(16) +
      String(item.bytes.byteLength).padStart(8) +
      ' bytes  ' +
      String(item.entries).padStart(3) +
      ' entries  ' +
      item.sha256.slice(0, 16),
  );
}

if (options.manifest) {
  if (options.only !== null) {
    throw new Error('--manifest writes every entry, so it cannot be combined with --id');
  }
  const manifest = {
    $comment:
      'Tier C of the sub-phase 1.1 corpus: decks written by the writer this project ships. Each ' +
      'one is another committed deck read through PartStore.open and written straight back out, ' +
      'so the parts belong to whoever authored them and the ZIP headers are ours. Generated by ' +
      'tools/corpus/tiers/c-written/build-written.ts --manifest; every description and features map is a ' +
      'literal in decks.ts beside it.',
    manifestVersion: 1,
    collection: 'written',
    generatedIn: '1.1',
    adr: 'docs/adr/phase-1-round-trip/0009-the-corpus.md',
    generator: 'tools/corpus/tiers/c-written/build-written.ts',
    // C018, and the whole reason the field has two halves. The container is
    // ours; the parts are Microsoft's, byte for byte, because a no-op write
    // passes every entry through. The xml value is spelled exactly as
    // corpus/authored spells it so that C-LEX counts one XML producer across
    // the two collections rather than two.
    serializers: {
      xml: 'Microsoft PowerPoint 16.0.20326',
      container: 'packages/opc',
    },
    targetCount: 1,
    entries: [...built]
      .sort((a, b) => a.deck.id.localeCompare(b.deck.id))
      .map((item) => ({
        id: item.deck.id,
        kind: 'deck',
        storage: 'committed',
        path: item.deck.id + '.pptx',
        license: 'CC0-1.0',
        source: 'self-authored',
        producer: 'packages/opc',
        sha256: item.sha256,
        bytes: item.bytes.byteLength,
        slides: item.deck.slides,
        zipEntries: item.entries,
        format: 'pptx',
        features: item.deck.features,
        description: item.deck.description,
        sourceNote:
          'Written from ' +
          item.deck.source +
          ' (sha256 ' +
          item.deck.sourceSha256 +
          '), a deck PowerPoint authored. The XML in every part is Microsoft-serialized and ' +
          'unmodified; only the ZIP container is ours, which is what makes this a distinct ' +
          'producer for C-LEX and pure redundancy for C-COV. Re-authoring Tier B changes this ' +
          'deck too, and the build refuses rather than writing silently when the source moves.',
        recipe: {
          tool: 'tools/corpus/tiers/c-written/build-written.ts',
          args: ['--out', 'corpus/written', '--id', item.deck.id],
        },
        addedIn: '1.1',
      })),
  };
  const path = join(outDir, 'manifest.json');
  writeFileSync(path, JSON.stringify(manifest, null, 2) + '\n');
  console.log('\nwrote ' + path);
}
