import { describe, expect, it } from 'vitest';

import { mediaFromStore, type MediaStore } from './media.js';

/** A store of one relationship per part, remembering which part was asked. */
function storeOf(
  rels: Readonly<Record<string, Readonly<Record<string, string>>>>,
  types: Readonly<Record<string, string>>,
): MediaStore & { asked: string[] } {
  const asked: string[] = [];
  return {
    asked,
    relationships(part) {
      asked.push(part);
      return { targetOf: (id) => rels[part]?.[id] };
    },
    contentTypeOf: (part) => types[part],
    read: (part) => new TextEncoder().encode(part),
  };
}

describe('mediaFromStore', () => {
  it('resolves an rId against the part it was written in, not the slide', () => {
    const store = storeOf(
      { '/ppt/slideLayouts/slideLayout1.xml': { rId2: '/ppt/media/image1.png' } },
      { '/ppt/media/image1.png': 'image/png' },
    );
    const image = mediaFromStore(store)('rId2', '/ppt/slideLayouts/slideLayout1.xml');
    expect(image?.contentType).toBe('image/png');
    expect(new TextDecoder().decode(image?.bytes)).toBe('/ppt/media/image1.png');
    expect(store.asked).toEqual(['/ppt/slideLayouts/slideLayout1.xml']);
    // The same rId on the slide names nothing, and the resolver says so rather than guessing.
    expect(mediaFromStore(store)('rId2', '/ppt/slides/slide1.xml')).toBeUndefined();
  });

  it('paints nothing for a target the package does not type', () => {
    const store = storeOf({ '/ppt/slides/slide1.xml': { rId3: '/ppt/media/odd.blob' } }, {});
    expect(mediaFromStore(store)('rId3', '/ppt/slides/slide1.xml')).toBeUndefined();
  });
});
