import type { MediaResolver } from './blip.js';

/** The three questions a picture asks of a package, as `PartStore` answers them. */
export interface MediaStore {
  relationships(part: string): { targetOf(id: string): string | undefined };
  contentTypeOf(part: string): string | undefined;
  read(part: string): Uint8Array;
}

/** An rId means the rels of the part the fill was written in; an untyped or missing target paints nothing. */
export function mediaFromStore(store: MediaStore): MediaResolver {
  return (embed, part) => {
    const target = store.relationships(part).targetOf(embed);
    if (target === undefined) return undefined;
    const contentType = store.contentTypeOf(target);
    if (contentType === undefined) return undefined;
    return { bytes: store.read(target), contentType };
  };
}
