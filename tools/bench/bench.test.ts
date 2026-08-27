import { createHash } from 'node:crypto';
import { mkdtempSync, readFileSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { inflateSync } from 'node:zlib';
import { censusPackage } from '@pptx-studio/census';
import { afterAll, beforeAll, describe, expect, it } from 'vitest';
import { RECIPES, writeDeck, type DeckRecipe, type DeckSummary } from './deck.ts';
import { mulberry32, noisePng, sizeForBytes } from './png.ts';
import { ZipStream } from './zip-stream.ts';
import { readZip } from '../ground-truth/zip.ts';

/**
 * The benchmark deck generator, checked against the census.
 *
 * The interesting assertion is the last group. The generator counts what it put
 * into a deck and the census counts what it finds in one, and the two are
 * written independently - so making them agree is a real check on both. A
 * census verified only against its own fixtures verifies nothing.
 *
 * Everything here runs on the `small` recipe, which is two megabytes and takes
 * a fraction of a second. The 200 MB deck is built by hand and measured in a
 * browser; a test suite is the wrong place for either.
 */

const ROOT = resolve(fileURLToPath(import.meta.url), '../../..');
const PROBE_FONT = resolve(ROOT, 'corpus/ground-truth/fonts/probe-v2_2.eot');

let directory: string;

function build(recipe: DeckRecipe, name = 'deck.pptx'): { path: string; summary: DeckSummary } {
  const path = join(directory, name);
  const zip = new ZipStream(path);
  const summary = writeDeck(zip, recipe, recipe.embedFont ? PROBE_FONT : null);
  return { path, summary };
}

const SMALL: DeckRecipe = { ...(RECIPES['small'] as DeckRecipe) };

beforeAll(() => {
  directory = mkdtempSync(join(tmpdir(), 'pptx-studio-bench-'));
});
afterAll(() => {
  rmSync(directory, { recursive: true, force: true });
});

describe('the streaming archive writer', () => {
  it('writes a readable ZIP32 with the content-type stream first', () => {
    const { path } = build(SMALL, 'order.pptx');
    const entries = readZip(new Uint8Array(readFileSync(path)));

    expect(entries[0]?.name).toBe('[Content_Types].xml');
    expect(entries.some((entry) => entry.name.endsWith('/'))).toBe(false);
    expect(entries.length).toBeGreaterThan(100);
  });

  it('stores what the caller says to store, and deflates the rest', () => {
    const path = join(directory, 'methods.zip');
    const zip = new ZipStream(path);
    const noise = noisePng(64, 7);
    zip.add('stored.png', noise, { store: true });
    zip.add('text.xml', '<a>' + 'x'.repeat(5000) + '</a>');
    const { bytes, entries } = zip.close();

    expect(entries).toBe(2);
    const archive = new Uint8Array(readFileSync(path));
    expect(archive.byteLength).toBe(bytes);
    // The stored entry cannot be smaller than its payload; the deflated one must be.
    expect(archive.byteLength).toBeGreaterThan(noise.length);
    expect(archive.byteLength).toBeLessThan(noise.length + 5000);
  });
});

describe('the noise images', () => {
  it('is a valid PNG: signature, IHDR, IDAT, IEND', () => {
    const png = noisePng(32, 1);
    expect([...png.slice(0, 8)]).toEqual([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
    const text = new TextDecoder('latin1').decode(png);
    expect(text.indexOf('IHDR')).toBe(12);
    expect(text).toContain('IDAT');
    expect(text.lastIndexOf('IEND')).toBe(png.length - 8);
  });

  it('carries the pixels it says it does', () => {
    const size = 40;
    const png = noisePng(size, 2);
    const view = new DataView(png.buffer, png.byteOffset, png.byteLength);
    expect(view.getUint32(16, false)).toBe(size);
    expect(view.getUint32(20, false)).toBe(size);
    expect(png[24]).toBe(8); // bit depth
    expect(png[25]).toBe(2); // truecolour

    // The IDAT payload has to inflate back to one filter byte plus RGB per row.
    const start = png.indexOf(0x49, 33); // 'I' of IDAT, after IHDR
    const length = view.getUint32(start - 4, false);
    const raw = inflateSync(png.subarray(start + 4, start + 4 + length));
    expect(raw.length).toBe(size * (size * 3 + 1));
  });

  it('is incompressible, which is the whole point of it', () => {
    // A deck padded with zeroes deflates to nothing, and then the ZIP reader
    // never does the work a real 200 MB file makes it do.
    const png = noisePng(200, 3);
    const zip = new ZipStream(join(directory, 'compressible.zip'));
    zip.add('image.png', png);
    const { bytes } = zip.close();
    expect(bytes).toBeGreaterThan(png.length);
  });

  it('is deterministic, so a deck built twice hashes the same', () => {
    expect(noisePng(24, 99)).toEqual(noisePng(24, 99));
    expect(noisePng(24, 99)).not.toEqual(noisePng(24, 100));
    const random = mulberry32(5);
    const first = [random(), random(), random()];
    const again = mulberry32(5);
    expect([again(), again(), again()]).toEqual(first);
  });

  it('lands near the size it was asked for', () => {
    for (const target of [32 * 1024, 256 * 1024, 640 * 1024]) {
      const actual = noisePng(sizeForBytes(target), 4).length;
      expect(actual / target).toBeGreaterThan(0.9);
      expect(actual / target).toBeLessThan(1.15);
    }
  });
});

describe('the generated deck is a package', () => {
  it('rebuilds byte-identically from the same recipe', () => {
    const digest = (path: string): string =>
      createHash('sha256').update(readFileSync(path)).digest('hex');
    const first = build(SMALL, 'repeat-a.pptx');
    const second = build(SMALL, 'repeat-b.pptx');
    expect(digest(second.path)).toBe(digest(first.path));
  });

  it('gives every part a content type', () => {
    const { path } = build(SMALL, 'types.pptx');
    const entries = readZip(new Uint8Array(readFileSync(path)));
    const types = new TextDecoder().decode(
      entries.find((entry) => entry.name === '[Content_Types].xml')?.bytes ?? new Uint8Array(),
    );

    const defaults = new Set([...types.matchAll(/Extension="([^"]+)"/g)].map((m) => m[1]));
    const overrides = new Set([...types.matchAll(/PartName="([^"]+)"/g)].map((m) => m[1]));

    for (const entry of entries) {
      if (entry.name === '[Content_Types].xml') continue;
      const extension = entry.name.slice(entry.name.lastIndexOf('.') + 1);
      const covered = defaults.has(extension) || overrides.has('/' + entry.name);
      expect(covered, entry.name + ' has no Default and no Override').toBe(true);
    }
  });

  it('leaves no relationship pointing at a part that is not there', () => {
    // Dangling edges are what put a red X and an unexplained repair prompt in
    // front of a user, and a generator that emits one would poison every
    // measurement taken with its output.
    const { path } = build(SMALL, 'rels.pptx');
    const entries = readZip(new Uint8Array(readFileSync(path)));
    const names = new Set(entries.map((entry) => entry.name));
    const decoder = new TextDecoder();

    let checked = 0;
    for (const entry of entries) {
      if (!entry.name.includes('_rels/')) continue;
      // A relationship target resolves against the folder of its **source
      // part**, and the package-level `_rels/.rels` has no source part - its
      // targets resolve against the package root. Treating it like any other
      // relationship part puts every root target one folder too deep.
      const source = entry.name.replace(/_rels\/([^/]*)\.rels$/, '$1');
      const folder = source.slice(0, source.lastIndexOf('/') + 1);
      const xml = decoder.decode(entry.bytes);
      for (const match of xml.matchAll(/Target="([^"]+)"([^>]*)/g)) {
        const [, target = '', rest = ''] = match;
        if (rest.includes('External')) continue;
        const resolved = new URL(target, 'file:///' + folder).pathname.slice(1);
        expect(names.has(resolved), entry.name + ' -> ' + target).toBe(true);
        checked += 1;
      }
    }
    expect(checked).toBeGreaterThan(50);
  });
});

describe('the generator and the census agree about what is in the deck', () => {
  it('finds every feature the recipe put in, and no others', () => {
    const { path, summary } = build(SMALL, 'census.pptx');
    const census = censusPackage(new Uint8Array(readFileSync(path)));

    const found = new Map(census.features.map((feature) => [feature.key, feature.count]));
    for (const [key, expected] of Object.entries(summary.expected)) {
      expect(found.get(key) ?? 0, key).toBe(expected);
    }
  });

  it('agrees about the presentation itself', () => {
    const { path } = build(SMALL, 'structure.pptx');
    const census = censusPackage(new Uint8Array(readFileSync(path)));

    expect(census.presentation?.slides).toBe(SMALL.slides);
    expect(census.presentation?.notesSlides).toBe(SMALL.notes ? SMALL.slides : 0);
    expect(census.presentation?.slideMasters).toBe(1);
    expect(census.presentation?.sections).toBe(SMALL.sections);
    expect(census.presentation?.embedTrueTypeFonts).toBe(true);
    expect(census.presentation?.embeddedFonts[0]?.typeface).toBe('ProbeAlpha');
    expect(census.presentation?.slideWidthEmu).toBe(12192000);
  });

  it('reports a deck with nothing wrong with it', () => {
    const { path } = build(SMALL, 'clean.pptx');
    const census = censusPackage(new Uint8Array(readFileSync(path)));

    expect(census.problems).toEqual([]);
    expect(census.relationships.dangling).toEqual([]);
    expect(census.relationships.unreachableParts).toEqual([]);
  });

  it('sees the extension namespaces the deck actually uses', () => {
    const { path } = build(SMALL, 'namespaces.pptx');
    const census = censusPackage(new Uint8Array(readFileSync(path)));
    const byUri = new Map(census.namespaces.map((namespace) => [namespace.uri, namespace]));

    const a14 = byUri.get('http://schemas.microsoft.com/office/drawing/2010/main');
    expect(a14?.required, 'a14 is named by an mc:Choice/@Requires').toBe(true);
    expect(a14?.extension).toBe(true);

    const p14 = byUri.get('http://schemas.microsoft.com/office/powerpoint/2010/main');
    expect(p14?.count, 'p14:sectionLst and p14:section').toBeGreaterThan(0);

    const dsp = byUri.get('http://schemas.microsoft.com/office/drawing/2008/diagram');
    expect(dsp?.count, 'the SmartArt drawing fallback').toBeGreaterThan(0);
  });

  it('scales with the recipe rather than with a fixture', () => {
    const bigger: DeckRecipe = { ...SMALL, slides: 6, tableEvery: 2, chartEvery: 3 };
    const { path, summary } = build(bigger, 'scaled.pptx');
    const census = censusPackage(new Uint8Array(readFileSync(path)));
    const found = new Map(census.features.map((feature) => [feature.key, feature.count]));

    expect(summary.expected['table']).toBe(3);
    expect(summary.expected['chart']).toBe(2);
    expect(found.get('table')).toBe(3);
    expect(found.get('chart')).toBe(2);
    expect(census.presentation?.slides).toBe(6);
  });
});
