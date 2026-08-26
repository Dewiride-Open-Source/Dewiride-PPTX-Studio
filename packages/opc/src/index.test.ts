import { describe, expect, it } from 'vitest';
import {
  CONTENT_TYPES_PART,
  FONT_DATA_CONTENT_TYPE,
  FONT_DATA_EXTENSION,
  OPC_NS,
  REL_TYPE,
  ROOT_RELS_PART,
} from './index.js';

describe('@pptx-studio/opc', () => {
  it('names the two package-level parts exactly', () => {
    expect(CONTENT_TYPES_PART).toBe('[Content_Types].xml');
    expect(ROOT_RELS_PART).toBe('_rels/.rels');
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
