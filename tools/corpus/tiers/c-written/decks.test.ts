import { createHash } from 'node:crypto';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';

import { REPO_ROOT as ROOT } from '../../../repo/root.ts';
import { censusPackage } from '@pptx-studio/census';
import { PartStore } from '@pptx-studio/opc';
import { describe, expect, it } from 'vitest';
import { WRITTEN_DECKS, type WrittenDeck } from './decks.ts';
import { readZip } from '../../../ground-truth/lib/zip.ts';
import { readHeaders } from '../../lexical/inventory.ts';

const CORPUS = join(ROOT, 'corpus/written');

/**
 * `C-CENSUS` and `C-REGEN` for Tier C, and the claim the tier exists to make.
 *
 * Tier A's `C-REGEN` holds the generator to its own output. Tier B has none and
 * cannot: PowerPoint stamps `dcterms:created` on every save. Tier C is back in
 * the reproducible half, and for a reason worth stating rather than assuming -
 * a no-op write touches no part, so every already-compressed entry is passed
 * through as the DEFLATE stream it already was and nothing is recompressed.
 * That is what makes a deflating writer's output safe to pin a SHA-256 to, and
 * `rebuilds byte-for-byte` below is the assertion that it stays true.
 *
 * The rest of this file is about the diff against the source, because that diff
 * is the entire contribution of the tier. Its census adds nothing - the parts
 * are the same bytes - so a test that only checked features would pass while
 * saying nothing at all about the writer.
 *
 * These tests import `@pptx-studio/opc` by name and get `src` through Vitest's
 * alias table, so they need no build. `build-written.ts` imports `dist` and does
 * need one. The gate is here; the script is the convenience.
 */

function load(id: string): Uint8Array {
  return new Uint8Array(readFileSync(join(CORPUS, id + '.pptx')));
}

function sourceOf(deck: WrittenDeck): Uint8Array {
  return new Uint8Array(readFileSync(join(ROOT, deck.source)));
}

const sha256 = (bytes: Uint8Array): string => createHash('sha256').update(bytes).digest('hex');

/**
 * `version made by`, per entry, straight out of the central directory.
 *
 * `readZip` does not surface it and has no reason to - it is a field a reader
 * ignores. It is here because it is one of the three places our writer and
 * PowerPoint differ, and a divergence nothing reads is exactly the kind that
 * gets asserted in prose and never checked.
 *
 * The central-directory reader itself moved to `lexical.ts` when `C-LEX` needed
 * the same fields for every deck in the corpus. This is the one caller that
 * cares about a single field.
 */
const versionsMadeBy = (archive: Uint8Array): Set<number> =>
  new Set(readHeaders(archive).entries.map((header) => header.versionMadeBy));

describe.each(WRITTEN_DECKS.map((deck) => [deck.id, deck] as const))('%s', (id, deck) => {
  it('reports exactly the census its entry declares', () => {
    const census = censusPackage(load(id));
    const found: Record<string, number> = {};
    for (const feature of census.features) found[feature.key] = feature.count;

    expect(found).toEqual(deck.features);
  });

  it('is a package with nothing wrong with it', () => {
    const census = censusPackage(load(id));

    expect(census.problems).toEqual([]);
    expect(census.relationships.dangling).toEqual([]);
    expect(census.relationships.unreachableParts).toEqual([]);
  });

  it('is the deck the manifest claims it is', () => {
    const manifest = JSON.parse(readFileSync(join(CORPUS, 'manifest.json'), 'utf8')) as {
      entries: { id: string; sha256: string; bytes: number; slides: number; zipEntries: number }[];
    };
    const claimed = manifest.entries.find((row) => row.id === id);
    const bytes = load(id);

    expect(claimed, 'no manifest entry for ' + id).toBeDefined();
    expect(claimed?.sha256).toBe(sha256(bytes));
    expect(claimed?.bytes).toBe(bytes.byteLength);
    expect(claimed?.slides).toBe(deck.slides);
    expect(claimed?.zipEntries).toBe(readZip(bytes).length);
  });

  it('is built from the source it pins', () => {
    expect(sha256(sourceOf(deck))).toBe(deck.sourceSha256);
  });

  it('rebuilds byte-for-byte from that source', () => {
    expect(sha256(PartStore.open(sourceOf(deck)).write())).toBe(sha256(load(id)));
  });

  it('is a fixpoint: writing it again changes nothing', () => {
    expect(sha256(PartStore.open(load(id)).write())).toBe(sha256(load(id)));
  });

  it('preserves every part of its source, in order and byte-for-byte', () => {
    const before = readZip(sourceOf(deck));
    const after = readZip(load(id));

    expect(after.map((e) => e.name)).toEqual(before.map((e) => e.name));
    for (const [i, entry] of before.entries()) {
      expect(sha256(after[i]!.bytes), entry.name + ' changed').toBe(sha256(entry.bytes));
      expect(after[i]!.store, entry.name + ' changed compression method').toBe(entry.store);
    }
  });
});

/**
 * The three header fields our writer does not carry across, asserted in both
 * directions: present in the source, absent here.
 *
 * Every one is legal to drop and none is a bug, which is precisely why they
 * need a test rather than a paragraph. A reader ignores all three, so nothing
 * else in this repository would ever notice if the writer started emitting
 * them - or if PowerPoint stopped.
 */
describe('what the writer normalises away', () => {
  const deck = WRITTEN_DECKS[0]!;
  const source = (): Uint8Array => sourceOf(deck);
  const written = (): Uint8Array => load(deck.id);

  it('drops the 0xA220 growth hints, and that is the whole size difference', () => {
    const hints = readZip(source()).filter((e) => e.extra !== undefined);

    expect(hints.map((e) => e.extra!.length)).toEqual([520, 520, 264, 264, 264]);
    for (const entry of hints) {
      const id = new DataView(entry.extra!.buffer, entry.extra!.byteOffset).getUint16(0, true);
      expect(id).toBe(0xa220);
    }
    expect(readZip(written()).filter((e) => e.extra !== undefined)).toEqual([]);

    const dropped = hints.reduce((total, e) => total + e.extra!.length, 0);
    expect(source().byteLength - written().byteLength).toBe(dropped);
  });

  it('writes general-purpose flags of 0 where PowerPoint writes 0x0006', () => {
    const before = readZip(source());

    // 0x0006 is bits 1 and 2: for method 8 they are a compression-level hint
    // and mean nothing to a decompressor. PowerPoint sets them on every
    // deflated entry and on none of its stored ones.
    expect(new Set(before.filter((e) => e.store !== true).map((e) => e.flags))).toEqual(
      new Set([0x0006]),
    );
    expect(new Set(before.filter((e) => e.store === true).map((e) => e.flags))).toEqual(
      new Set([0x0000]),
    );
    expect(new Set(readZip(written()).map((e) => e.flags))).toEqual(new Set([0x0000]));
  });

  it('writes version-made-by 20 where PowerPoint writes 45', () => {
    expect(versionsMadeBy(source())).toEqual(new Set([45]));
    expect(versionsMadeBy(written())).toEqual(new Set([20]));
  });
});

describe('the collection', () => {
  const manifest = JSON.parse(readFileSync(join(CORPUS, 'manifest.json'), 'utf8')) as {
    collection: string;
    targetCount: number;
    entries: Record<string, unknown>[];
  };

  it('claims one entry per deck and no more', () => {
    expect(manifest.collection).toBe('written');
    expect(manifest.entries).toHaveLength(WRITTEN_DECKS.length);
    expect(manifest.targetCount).toBe(WRITTEN_DECKS.length);
  });

  it('names one producer, and it is ours', () => {
    for (const entry of manifest.entries) expect(entry['producer']).toBe('packages/opc');
  });

  it('carries both a recipe and the note saying what it was made from', () => {
    for (const entry of manifest.entries) {
      // The pairing is the point. Tier B has a `sourceNote` and no `recipe`
      // because its bytes cannot be reproduced; Tier A has a `recipe` and no
      // `sourceNote` because there is no source but its own module. Tier C is
      // the only one that needs both, and an entry missing either is making a
      // claim about its own provenance that is not true.
      expect(entry['recipe'], String(entry['id']) + ' must be reproducible').toBeDefined();
      expect(entry['sourceNote'], String(entry['id']) + ' must name its source').toBeDefined();
    }
  });

  it('uses the c-prefixed id shape', () => {
    for (const deck of WRITTEN_DECKS) expect(deck.id).toMatch(/^c\d{2}-[a-z0-9-]+$/);
  });
});
