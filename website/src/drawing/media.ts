import type { PartStore } from '@pptx-studio/opc';
import type { MediaResolver } from '@pptx-studio/render-svg';

/**
 * How a picture reaches its bytes.
 *
 * The renderer resolves no relationships of its own: an `r:embed` means the
 * rels of the part the fill was written in, and only the caller knows which
 * part that was. This is the same four lines the CLI uses.
 */
export function mediaFrom(store: PartStore): MediaResolver {
  return (embed: string, part: string) => {
    const target = store.relationships(part).targetOf(embed);
    if (target === undefined) return undefined;
    const contentType = store.contentTypeOf(target);
    if (contentType === undefined) return undefined;
    return { bytes: store.read(target), contentType };
  };
}
