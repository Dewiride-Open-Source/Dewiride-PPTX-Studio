import { describe, expect, it } from 'vitest';
import {
  CONTENT_TYPES_PART,
  FONT_DATA_CONTENT_TYPE,
  FONT_DATA_EXTENSION,
  OPC_NS,
  REL_TYPE,
  ROOT_RELS_PART,
  zipEntryNameFor,
  type PartName,
} from './index.js';

describe('@pptx-studio/opc', () => {
  it('names the two package-level streams exactly, in their own namespaces', () => {
    // These two constants deliberately live in different namespaces, and the
    // asymmetry is load-bearing rather than sloppy. `[Content_Types].xml` is a
    // ZIP entry name because it is not a part at all - nothing types it and
    // nothing relates to it, so it has no part name. `_rels/.rels` *is* a part,
    // so it is named the way parts are named, with a leading slash. Conflating
    // the two is how a lookup silently misses.
    expect(CONTENT_TYPES_PART).toBe('[Content_Types].xml');
    expect(ROOT_RELS_PART).toBe('/_rels/.rels');
    expect(zipEntryNameFor(ROOT_RELS_PART as PartName)).toBe('_rels/.rels');
  });

  it('distinguishes the package relationship namespace from the officeDocument one', () => {
    // Easy to conflate: the package part `_rels/.rels` is in the *package*
    // namespace, while `r:id` attributes inside part XML are in the
    // officeDocument namespace. Mixing them produces a file PowerPoint repairs.
    expect(OPC_NS.relationships).toBe(
      'http://schemas.openxmlformats.org/package/2006/relationships',
    );
    expect(REL_TYPE.slideLayout).toBe(
      'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout',
    );
  });

  it('pins the embedded-font Default entry', () => {
    expect(FONT_DATA_EXTENSION).toBe('fntdata');
    expect(FONT_DATA_CONTENT_TYPE).toBe('application/x-fontdata');
  });
});
