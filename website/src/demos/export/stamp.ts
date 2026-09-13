import type { PartStore } from '@pptx-studio/opc';
import {
  applyEdit,
  descendantElements,
  namespaceOf,
  newText,
  parseXml,
  serializeXml,
} from '@pptx-studio/xml';

import type { ReplacedPart } from '@/deck/worker/protocol';

const CORE_PART = '/docProps/core.xml';
const CORE_NS = 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties';

/**
 * The smallest edit a package can carry: `cp:lastModifiedBy` in the core properties.
 *
 * It changes nothing on any slide, but it is the only path in which a part is
 * serialised afresh - and an export that can only emit the bytes it read has not
 * exercised the writer. Returns null when the deck has no such element.
 */
export function stampLastModifiedBy(store: PartStore, by: string): ReplacedPart | null {
  if (!store.has(CORE_PART)) return null;
  const document = parseXml(store.read(CORE_PART));
  const element = [...descendantElements(document.root)].find(
    (one) => one.local === 'lastModifiedBy' && namespaceOf(one) === CORE_NS,
  );
  if (element === undefined) return null;
  const text = element.children.find((child) => child.type === 'text');
  if (text === undefined) {
    applyEdit({ kind: 'insertChild', parent: element, index: 0, node: newText(by) });
  } else {
    applyEdit({ kind: 'setValue', node: text, value: by });
  }
  const bytes = serializeXml(document);
  return {
    part: CORE_PART,
    bytes: bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer,
  };
}
