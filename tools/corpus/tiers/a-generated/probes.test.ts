import { createHash } from 'node:crypto';
import { readFileSync } from 'node:fs';
import { join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { censusPackage } from '@pptx-studio/census';
import { describe, expect, it } from 'vitest';
import { buildProbePackage } from './package.ts';
import { PROBE_DECKS } from './decks/index.ts';
import { outputName } from './types.ts';
import { readZip } from '../../ground-truth/zip.ts';

const CORPUS = resolve(fileURLToPath(import.meta.url), '../../../../corpus/decks');

/**
 * `C-CENSUS` and `C-REGEN`, for the decks that exist.
 *
 * The interesting assertion is the first. Each probe deck states, as a literal
 * in its own module, the census it should produce; the census is an independent
 * token scanner in `packages/census` that knows nothing about the generator.
 * Making the two agree exactly - not "contains", not "at least" - is a real
 * check on both, and it is the reason the `features` maps are written by hand
 * rather than tallied at the emit site the way `tools/bench/deck.ts` does. A
 * map derived from the markup would restate the markup instead of checking it.
 *
 * Which means these tests fail whenever a deck's markup changes, on purpose.
 * The fix is to read the new map out of the failure and put it in the deck
 * module, having looked at what moved.
 */

function build(id: string): { bytes: Uint8Array; parts: ReturnType<typeof readZip> } {
  const deck = PROBE_DECKS.find((d) => d.id === id);
  if (deck === undefined) throw new Error('no such probe deck: ' + id);
  const { bytes } = buildProbePackage(deck.build());
  return { bytes, parts: readZip(bytes) };
}

describe.each(PROBE_DECKS.map((deck) => [deck.id, deck] as const))('%s', (id, deck) => {
  it('reports exactly the census its module declares', () => {
    const { bytes } = build(id);
    const census = censusPackage(bytes);
    const found: Record<string, number> = {};
    for (const feature of census.features) found[feature.key] = feature.count;

    expect(found).toEqual(deck.features);
  });

  it('is a package with nothing wrong with it', () => {
    const { bytes } = build(id);
    const census = censusPackage(bytes);

    expect(census.problems).toEqual([]);
    expect(census.relationships.dangling).toEqual([]);
    expect(census.relationships.unreachableParts).toEqual([]);
  });

  it('builds byte-for-byte identically every time', () => {
    // Not a tautology: a `Date.now()`, a `Math.random()` or an iteration over a
    // `Set` built from object identity would each break this and none would
    // show up anywhere else until `C-REGEN` failed on somebody else's machine.
    const first = createHash('sha256').update(build(id).bytes).digest('hex');
    const again = createHash('sha256').update(build(id).bytes).digest('hex');
    expect(again).toBe(first);
  });

  it('stores every entry, so its hash does not depend on which zlib built it', () => {
    // ADR 0009 committed bytes rather than recipes because a recipe's hash is
    // the hash of a build. Deflating would put zlib's version straight back
    // into `C-REGEN`. `a35-zip-shapes` is the one deck that opts out, and it
    // opts out because it is the deck *about* compression - which is the whole
    // reason the exception is where it is rather than anywhere else.
    const { bytes } = build(id);
    const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    let offset = 0;
    let checked = 0;
    let deflated = 0;
    while (view.getUint32(offset, true) === 0x04034b50) {
      const method = view.getUint16(offset + 8, true);
      if (id === 'a35-zip-shapes') {
        expect(method, 'compression method').toBeOneOf([0, 8]);
        if (method === 8) deflated += 1;
      } else {
        expect(method, 'compression method').toBe(0);
      }
      const compressed = view.getUint32(offset + 18, true);
      const nameLength = view.getUint16(offset + 26, true);
      const extraLength = view.getUint16(offset + 28, true);
      offset += 30 + nameLength + extraLength + compressed;
      checked += 1;
    }
    expect(checked).toBeGreaterThan(15);
    // And the exception has to be a real one: a deck that claimed to deflate
    // and stored everything would pass the branch above while testing nothing.
    if (id === 'a35-zip-shapes') expect(deflated).toBeGreaterThan(10);
  });

  it('has the archive shape its module declares', () => {
    // Only `a35-zip-shapes` sets any of this, and the point of asserting it
    // here rather than trusting the generator is that the general-purpose flag
    // and the extra field are invisible to every other test in the repository:
    // the census reads parts, `C-REGEN` reads one hash, and neither would
    // notice the writer silently dropping a flag it was asked to set.
    if (id !== 'a35-zip-shapes') return;
    const { bytes } = build(id);
    const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    const decoder = new TextDecoder();
    const shapes = new Map<string, { method: number; flags: number; extra: number }>();
    let offset = 0;
    while (view.getUint32(offset, true) === 0x04034b50) {
      const flags = view.getUint16(offset + 6, true);
      const method = view.getUint16(offset + 8, true);
      const compressed = view.getUint32(offset + 18, true);
      const nameLength = view.getUint16(offset + 26, true);
      const extraLength = view.getUint16(offset + 28, true);
      const name = decoder.decode(bytes.subarray(offset + 30, offset + 30 + nameLength));
      shapes.set(name, { method, flags, extra: extraLength });
      if (extraLength > 0) {
        // Microsoft's growth hint: id, length, signature, padding value.
        const at = offset + 30 + nameLength;
        expect(view.getUint16(at, true), name + ' extra field id').toBe(0xa220);
        expect(view.getUint16(at + 2, true)).toBe(extraLength - 4);
        expect(view.getUint16(at + 4, true), name + ' growth hint signature').toBe(0xa028);
        expect(view.getUint16(at + 6, true)).toBe(extraLength - 8);
      }
      offset += 30 + nameLength + extraLength + compressed;
    }

    expect(shapes.get('[Content_Types].xml')).toEqual({ method: 8, flags: 0x0006, extra: 520 });
    expect(shapes.get('_rels/.rels')?.flags).toBe(0x0006);
    expect(shapes.get('ppt/presentation.xml')?.extra).toBe(64);
    expect(shapes.get('ppt/slides/slide1.xml')?.method).toBe(0);
    expect(shapes.get('ppt/slides/slide2.xml')).toMatchObject({ method: 0, flags: 0x0006 });
    expect(shapes.get('ppt/slides/slide3.xml')?.flags).toBe(0x0800);
    expect(shapes.get('docProps/app.xml')).toMatchObject({ method: 0, extra: 128 });
    // Stored because DEFLATE made it bigger, not because the deck asked.
    expect(shapes.get('ppt/media/image2.png')?.method).toBe(0);
    expect(shapes.get('ppt/media/image1.png')?.method).toBe(8);
  });

  it('opens with [Content_Types].xml first and no directory entries', () => {
    const { parts } = build(id);
    expect(parts[0]?.name).toBe('[Content_Types].xml');
    expect(parts.some((part) => part.name.endsWith('/'))).toBe(false);
  });

  it('rebuilds to the bytes committed under corpus/decks', () => {
    // `C-REGEN`. Committed bytes and a recipe that reproduces them are two
    // different claims, and hashing the file only checks the first: the
    // manifest's `sha256` says "these bytes are what was reviewed", and this
    // says "and the generator still produces them". Without it a generator
    // change silently orphans every fixture from the code that made it.
    const onDisk = new Uint8Array(readFileSync(join(CORPUS, outputName(deck))));
    const rebuilt = build(id).bytes;
    expect(createHash('sha256').update(rebuilt).digest('hex')).toBe(
      createHash('sha256').update(onDisk).digest('hex'),
    );
  });

  it('is the deck the manifest claims it is', () => {
    const manifest = JSON.parse(readFileSync(join(CORPUS, 'manifest.json'), 'utf8')) as {
      entries: { id: string; sha256: string; bytes: number; features: Record<string, number> }[];
    };
    const claimed = manifest.entries.find((row) => row.id === id);
    const bytes = build(id).bytes;

    expect(claimed, 'no manifest entry for ' + id).toBeDefined();
    expect(claimed?.sha256).toBe(createHash('sha256').update(bytes).digest('hex'));
    expect(claimed?.bytes).toBe(bytes.byteLength);
    expect(claimed?.features).toEqual(deck.features);
  });
});

describe('the roster as a whole', () => {
  it('has unique, kebab-case, sorted ids', () => {
    const ids = PROBE_DECKS.map((deck) => deck.id);
    expect(new Set(ids).size).toBe(ids.length);
    expect([...ids].sort((a, b) => a.localeCompare(b))).toEqual(ids);
    for (const id of ids) expect(id).toMatch(/^a\d{2}-[a-z0-9-]+$/);
  });

  it('stays inside the corpus size caps', () => {
    // 512 KiB per committed file and 12 MiB total, both enforced by
    // `C012-size-cap`. Checking here as well means a deck that grows past the
    // cap fails in the test that built it rather than in the gate.
    let total = 0;
    for (const deck of PROBE_DECKS) {
      const { bytes } = buildProbePackage(deck.build());
      expect(bytes.byteLength, deck.id).toBeLessThan(512 * 1024);
      total += bytes.byteLength;
    }
    expect(total).toBeLessThan(12 * 1024 * 1024);
  });

  it('describes itself in the manifest voice', () => {
    for (const deck of PROBE_DECKS) {
      expect(deck.description.length, deck.id).toBeGreaterThan(80);
      expect(deck.title.startsWith('PPTX Studio corpus:'), deck.id).toBe(true);
    }
  });
});
