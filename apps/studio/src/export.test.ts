import { PartStore } from '@pptx-studio/opc';
import { childElements, parseXml, textContent } from '@pptx-studio/xml';
import { beforeAll, describe, expect, it } from 'vitest';
import { exportDeck, stampLastModifiedBy, STAMP } from './export.js';

/**
 * The export path, in a real browser, on a real deck.
 *
 * This is the half of Gate 1 that no package can test. `packages/writer` runs
 * its own suite in Chromium and has since 0.1, so "the writer works in a
 * browser" is not the open question - the open question is whether the *path*
 * does: a deck fetched over HTTP, opened, checked, written, and handed back as
 * bytes something can be downloaded from.
 *
 * The decks are fetched rather than built, and that is the point of doing it
 * here. `packages/validate/src/testing/deck.ts` says why the core packages
 * cannot: browser mode has no filesystem, so the committed corpus is not
 * reachable from a package test and the writer's own suite works on synthetic
 * packages instead. Vitest serves the workspace root, so from here the corpus
 * is one `fetch` away - and `a43-kitchen-sink` is the whole of what the gate
 * asks about in one file.
 */

const KITCHEN_SINK = '/corpus/decks/a43-kitchen-sink.pptm';
const MINIMAL = '/corpus/decks/a01-minimal.pptx';

async function load(path: string): Promise<Uint8Array> {
  const response = await fetch(path);
  if (!response.ok) throw new Error('could not fetch ' + path + ': ' + String(response.status));
  return new Uint8Array(await response.arrayBuffer());
}

let kitchenSink: Uint8Array;
let minimal: Uint8Array;

beforeAll(async () => {
  [kitchenSink, minimal] = await Promise.all([load(KITCHEN_SINK), load(MINIMAL)]);
});

describe('a deck with all five features, handed back unchanged', () => {
  it('rewrites nothing, and streams every part it read', () => {
    const { outcome } = exportDeck(kitchenSink);

    // The claim the architecture rests on, in the one form that can be false:
    // a part serialised on a no-op export is a part somebody's writer decided
    // it understood.
    expect(outcome.rewritten).toEqual([]);
    expect(outcome.streamed).toBe(55);
    expect(outcome.edited).toBeNull();
  });

  it('compares identical as a document, part by part', () => {
    const { outcome } = exportDeck(kitchenSink);

    expect(outcome.comparison.differences).toEqual([]);
    expect(outcome.comparison.parts).toBe(55);
    expect(outcome.comparison.same).toBe(55);
    // 35 as canonical XML, 17 as relationship graphs, and 3 by SHA-256: the
    // embedded workbook, the EMF preview and `ppt/vbaProject.bin`.
    expect(outcome.comparison.xml).toBe(35);
    expect(outcome.comparison.relationships).toBe(17);
    expect(outcome.comparison.binary).toBe(3);
  });

  it('runs the preservation check rather than skipping it', () => {
    // Both are `null`-shaped on a screen and opposite in meaning, which is why
    // `skipped` carries a reason instead of a boolean.
    const { outcome } = exportDeck(kitchenSink);
    expect(outcome.preservation.skipped).toBeNull();
    expect(outcome.preservation.checked).toBe(56);
    expect(outcome.preservation.rewritten).toBe(0);
  });

  it('passes all twenty-nine rules with nothing blocking', () => {
    const { outcome } = exportDeck(kitchenSink);
    expect(outcome.report?.ok).toBe(true);
    expect(outcome.report?.blocking).toBe(0);
    expect(outcome.report?.checked).toHaveLength(29);
    // Not "no findings": six rules need a baseline and this path supplies one,
    // so a skipped rule here would mean the baseline went missing.
    expect(outcome.report?.skipped).toEqual([]);
  });

  it('keeps the parts the gate is actually about', () => {
    const store = PartStore.open(exportDeck(kitchenSink).bytes);
    for (const part of [
      '/ppt/charts/chart1.xml',
      '/ppt/diagrams/data1.xml',
      '/ppt/diagrams/drawing1.xml',
      '/ppt/embeddings/oleObject1.xlsx',
      '/ppt/drawings/vmlDrawing1.vml',
      '/ppt/vbaProject.bin',
    ]) {
      expect(store.has(part), part).toBe(true);
    }
    // The macro-enabled content type, which is the whole of what makes a
    // package a `.pptm`.
    expect(store.contentTypeOf('/ppt/presentation.xml')).toBe(
      'application/vnd.ms-powerpoint.presentation.macroEnabled.main+xml',
    );
  });

  it('survives being exported twice', () => {
    // Idempotence, which a round trip does not imply on its own: a writer that
    // normalised something once and then normalised it differently would pass
    // one pass and fail this.
    const once = exportDeck(kitchenSink).bytes;
    const twice = exportDeck(once).bytes;
    expect([...twice]).toEqual([...once]);
  });
});

describe('the stamp, which is the only edit this page can make', () => {
  it('rewrites exactly one part and streams the rest', () => {
    const { outcome } = exportDeck(kitchenSink, 'stamp');

    expect(outcome.edited).toBe('/docProps/core.xml');
    expect(outcome.rewritten).toEqual(['/docProps/core.xml']);
    expect(outcome.streamed).toBe(54);
    // The preservation check deliberately does not compare the part this
    // session wrote, which is why the comparison below exists.
    expect(outcome.preservation.rewritten).toBe(1);
    expect(outcome.preservation.checked).toBe(55);
  });

  it('changes that part and nothing else', () => {
    const { outcome } = exportDeck(kitchenSink, 'stamp');

    expect(outcome.comparison.differences).toHaveLength(1);
    expect(outcome.comparison.differences[0]?.part).toBe('/docProps/core.xml');
    expect(outcome.comparison.same).toBe(outcome.comparison.parts - 1);
  });

  it('writes the value into cp:lastModifiedBy, and leaves its siblings alone', () => {
    const before = PartStore.open(kitchenSink);
    const after = PartStore.open(exportDeck(kitchenSink, 'stamp').bytes);
    const read = (store: PartStore, local: string): string | undefined => {
      const root = parseXml(store.read('/docProps/core.xml')).root;
      const child = childElements(root).find((node) => node.local === local);
      return child === undefined ? undefined : textContent(child);
    };

    expect(read(before, 'lastModifiedBy')).not.toBe(STAMP);
    expect(read(after, 'lastModifiedBy')).toBe(STAMP);
    expect(read(after, 'title')).toBe(read(before, 'title'));
    expect(read(after, 'creator')).toBe(read(before, 'creator'));
    expect(read(after, 'modified')).toBe(read(before, 'modified'));
  });

  it('leaves the rest of the part byte-for-byte, because only one node was dirty', () => {
    // 0.5's per-node fidelity, seen from the outside. The serializer re-emits
    // clean nodes by slicing the source, so the only thing that may differ in
    // this part is the text inside one element.
    const before = new TextDecoder().decode(PartStore.open(kitchenSink).read('/docProps/core.xml'));
    const after = new TextDecoder().decode(
      PartStore.open(exportDeck(kitchenSink, 'stamp').bytes).read('/docProps/core.xml'),
    );
    const strip = (text: string): string =>
      text.replace(/<cp:lastModifiedBy>[^<]*<\/cp:lastModifiedBy>/, '');
    expect(strip(after)).toBe(strip(before));
  });

  it('is a no-op the second time, because the value is already there', () => {
    const once = exportDeck(kitchenSink, 'stamp').bytes;
    const twice = exportDeck(once, 'stamp').bytes;
    // The part is still rewritten - the edit is applied and the tree serialised
    // - but the bytes it produces are the ones already in the archive.
    expect([...twice]).toEqual([...once]);
  });

  it('reports that it found nothing rather than inventing an element', () => {
    const store = PartStore.open(minimal);
    store.removePart('/docProps/core.xml');
    expect(stampLastModifiedBy(store, STAMP)).toBeNull();
  });
});

describe('the smallest deck, so the numbers are checkable by hand', () => {
  it('round-trips with the same shape as the big one', () => {
    const { outcome } = exportDeck(minimal);
    expect(outcome.rewritten).toEqual([]);
    expect(outcome.comparison.differences).toEqual([]);
    expect(outcome.report?.ok).toBe(true);
    expect(outcome.bytesOut).toBeGreaterThan(0);
  });

  it('keeps every part name it read, and adds none', () => {
    const before = PartStore.open(minimal).partNames.sort();
    const after = PartStore.open(exportDeck(minimal).bytes).partNames.sort();
    expect(after).toEqual(before);
  });
});
