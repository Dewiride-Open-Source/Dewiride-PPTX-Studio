import { describe, expect, it } from 'vitest';
import { CONTENT_TYPE, FONT_DATA_CONTENT_TYPE } from '../constants.js';
import { ContentTypes, isXmlContentType } from './content-types.js';
import { isOpcError } from '../errors.js';
import { contentTypesXml } from '../testing/build-package.js';

const encoder = new TextEncoder();
const decoder = new TextDecoder();

function parse(xml: string): ContentTypes {
  return ContentTypes.parse(encoder.encode(xml));
}

function codeOf(fn: () => unknown): string {
  try {
    fn();
  } catch (error) {
    if (isOpcError(error)) return error.code;
    return 'NOT_AN_OPC_ERROR: ' + (error instanceof Error ? error.constructor.name : typeof error);
  }
  return 'DID_NOT_THROW';
}

const OFFICE = contentTypesXml(
  '<Default Extension="rels" ContentType="' + CONTENT_TYPE.relationships + '"/>',
  '<Default Extension="xml" ContentType="application/xml"/>',
  '<Default Extension="png" ContentType="image/png"/>',
  '<Override PartName="/ppt/presentation.xml" ContentType="' + CONTENT_TYPE.presentation + '"/>',
  '<Override PartName="/ppt/slides/slide1.xml" ContentType="' + CONTENT_TYPE.slide + '"/>',
);

describe('resolving a part to a content type', () => {
  it('prefers an Override to a Default, and says which it used', () => {
    const ct = parse(OFFICE);
    expect(ct.resolve('/ppt/slides/slide1.xml')).toEqual({
      contentType: CONTENT_TYPE.slide,
      origin: 'override',
    });
    expect(ct.resolve('/ppt/tableStyles.xml')).toEqual({
      contentType: 'application/xml',
      origin: 'default',
    });
  });

  it('types the root relationship part through the rels Default', () => {
    // `/_rels/.rels` has no Override in any package we have measured, so its
    // extension has to resolve to `rels` for the package to be writable at all.
    expect(parse(OFFICE).for('/_rels/.rels')).toBe(CONTENT_TYPE.relationships);
    expect(parse(OFFICE).for('/ppt/slides/_rels/slide1.xml.rels')).toBe(CONTENT_TYPE.relationships);
  });

  it('matches extensions case-insensitively', () => {
    const ct = parse(OFFICE);
    expect(ct.for('/ppt/media/image1.PNG')).toBe('image/png');
    expect(ct.for('/ppt/media/image2.Png')).toBe('image/png');
  });

  it('matches Override part names case-insensitively', () => {
    expect(parse(OFFICE).for('/PPT/Slides/Slide1.XML')).toBe(CONTENT_TYPE.slide);
  });

  it('returns undefined for a part nothing types, and throws only when asked to', () => {
    const ct = parse(contentTypesXml('<Default Extension="xml" ContentType="application/xml"/>'));
    expect(ct.for('/ppt/media/image1.png')).toBeUndefined();
    expect(codeOf(() => ct.require('/ppt/media/image1.png'))).toBe('ERR_MISSING_CONTENT_TYPE');
  });

  it('has nothing to fall back on for a part with no extension', () => {
    const ct = parse(OFFICE);
    expect(ct.for('/ppt/embeddings/oleObject')).toBeUndefined();
  });

  it('does not treat a Default for an unused extension as an error', () => {
    // Office ships these constantly - a template carries `Default xlsx` whether
    // or not any chart in it has an embedded workbook.
    const ct = parse(
      contentTypesXml('<Default Extension="xlsx" ContentType="' + CONTENT_TYPE.spreadsheet + '"/>'),
    );
    expect(ct.defaults).toHaveLength(1);
  });
});

describe('a stream that disagrees with itself', () => {
  it('refuses two Defaults for one extension with different types', () => {
    expect(
      codeOf(() =>
        parse(
          contentTypesXml(
            '<Default Extension="xml" ContentType="application/xml"/>',
            '<Default Extension="XML" ContentType="text/xml"/>',
          ),
        ),
      ),
    ).toBe('ERR_CONTENT_TYPE_CONFLICT');
  });

  it('tolerates a duplicate that agrees, but marks itself dirty so it is not re-emitted', () => {
    // PowerPoint refuses two `<Default>`s for one extension with 0x80CB8000
    // even when the two elements are byte-identical — verified by duplicating
    // the `png` Default in a real deck. We read such a package rather than
    // failing, but leaving `dirty` false would pass the original stream
    // through untouched and hand back a file that does not open.
    const ct = parse(
      contentTypesXml(
        '<Default Extension="xml" ContentType="application/xml"/>',
        '<Default Extension="xml" ContentType="Application/XML"/>',
      ),
    );
    expect(ct.for('/a.xml')).toBe('Application/XML');
    expect(ct.dirty).toBe(true);
    expect(ct.defaults).toHaveLength(1);
    expect(decoder.decode(ct.serialize()).match(/<Default /g)).toHaveLength(1);
  });

  it('marks itself dirty for a duplicate Override that agrees', () => {
    const ct = parse(
      contentTypesXml(
        '<Override PartName="/a.xml" ContentType="application/xml"/>',
        '<Override PartName="/A.XML" ContentType="application/xml"/>',
      ),
    );
    expect(ct.dirty).toBe(true);
    expect(ct.overrides).toHaveLength(1);
  });

  it('stays clean for an ordinary stream, so it passes through untouched', () => {
    expect(parse(OFFICE).dirty).toBe(false);
  });

  it('refuses two Overrides for one part with different types', () => {
    expect(
      codeOf(() =>
        parse(
          contentTypesXml(
            '<Override PartName="/a.xml" ContentType="application/xml"/>',
            '<Override PartName="/A.XML" ContentType="text/xml"/>',
          ),
        ),
      ),
    ).toBe('ERR_CONTENT_TYPE_CONFLICT');
  });

  it('refuses a content type that is not a media type', () => {
    expect(
      codeOf(() => parse(contentTypesXml('<Default Extension="xml" ContentType="nonsense"/>'))),
    ).toBe('ERR_INVALID_CONTENT_TYPE');
  });

  it('refuses entries missing a required attribute, and an empty extension', () => {
    expect(codeOf(() => parse(contentTypesXml('<Default Extension="xml"/>')))).toBe(
      'ERR_MALFORMED_XML',
    );
    expect(codeOf(() => parse(contentTypesXml('<Override ContentType="application/xml"/>')))).toBe(
      'ERR_MALFORMED_XML',
    );
    expect(
      codeOf(() => parse(contentTypesXml('<Default Extension="" ContentType="application/xml"/>'))),
    ).toBe('ERR_MALFORMED_XML');
  });

  it('refuses an Override whose PartName is not a part name', () => {
    expect(
      codeOf(() =>
        parse(contentTypesXml('<Override PartName="ppt/x.xml" ContentType="application/xml"/>')),
      ),
    ).toBe('ERR_INVALID_PART_NAME');
    expect(
      codeOf(() =>
        parse(contentTypesXml('<Override PartName="/../x.xml" ContentType="application/xml"/>')),
      ),
    ).toBe('ERR_INVALID_PART_NAME');
  });

  it('refuses a root element that is not Types', () => {
    expect(codeOf(() => parse('<Relationships/>'))).toBe('ERR_MALFORMED_XML');
  });

  it('ignores an element it does not know, since one cannot mistype a part', () => {
    const ct = parse(
      contentTypesXml(
        '<Default Extension="xml" ContentType="application/xml"/>',
        '<FutureThing Extension="xml"/>',
      ),
    );
    expect(ct.for('/a.xml')).toBe('application/xml');
  });
});

describe('changing the map', () => {
  it('starts clean when parsed and turns dirty on the first change', () => {
    const ct = parse(OFFICE);
    expect(ct.dirty).toBe(false);
    ct.setOverride('/ppt/slides/slide2.xml', CONTENT_TYPE.slide);
    expect(ct.dirty).toBe(true);
  });

  it('ensureFor does nothing when a Default already covers the part', () => {
    const ct = parse(OFFICE);
    ct.ensureFor('/ppt/media/image9.png', 'image/png');
    expect(ct.dirty).toBe(false);
    expect(ct.overrides).toHaveLength(2);
  });

  it('ensureFor adds an Override rather than claiming an extension for everyone', () => {
    const ct = parse(OFFICE);
    ct.ensureFor('/ppt/media/thing.png', 'image/jpeg');
    expect(ct.dirty).toBe(true);
    expect(ct.defaults.map((d) => d.extension)).toEqual(['rels', 'xml', 'png']);
    expect(ct.for('/ppt/media/thing.png')).toBe('image/jpeg');
    expect(ct.for('/ppt/media/other.png')).toBe('image/png');
  });

  it('removes an Override and falls back to the Default underneath it', () => {
    const ct = parse(OFFICE);
    expect(ct.removeOverride('/ppt/slides/slide1.xml')).toBe(true);
    expect(ct.for('/ppt/slides/slide1.xml')).toBe('application/xml');
    expect(ct.removeOverride('/ppt/slides/slide1.xml')).toBe(false);
  });

  it('keeps the remaining Overrides addressable after a removal', () => {
    const ct = parse(
      contentTypesXml(
        '<Override PartName="/a.xml" ContentType="application/a"/>',
        '<Override PartName="/b.xml" ContentType="application/b"/>',
        '<Override PartName="/c.xml" ContentType="application/c"/>',
      ),
    );
    ct.removeOverride('/a.xml');
    expect(ct.for('/b.xml')).toBe('application/b');
    expect(ct.for('/c.xml')).toBe('application/c');
    ct.removeOverride('/b.xml');
    expect(ct.for('/c.xml')).toBe('application/c');
    expect(ct.overrides).toHaveLength(1);
  });

  it('refuses a Default for something that is not an extension', () => {
    const ct = ContentTypes.empty();
    expect(codeOf(() => ct.setDefault('a.png', 'image/png'))).toBe('ERR_INVALID_CONTENT_TYPE');
    expect(codeOf(() => ct.setDefault('', 'image/png'))).toBe('ERR_INVALID_CONTENT_TYPE');
  });
});

describe('writing the stream back', () => {
  it('reproduces Office’s layout exactly: declaration, Defaults, Overrides, no whitespace', () => {
    expect(decoder.decode(parse(OFFICE).serialize())).toBe(OFFICE);
  });

  it('round-trips through itself', () => {
    const once = parse(OFFICE).serialize();
    expect(decoder.decode(ContentTypes.parse(once).serialize())).toBe(decoder.decode(once));
  });

  it('escapes an attribute rather than emitting something unparseable', () => {
    const ct = ContentTypes.empty();
    ct.setOverride('/a&b.xml', 'application/xml');
    const back = ContentTypes.parse(ct.serialize());
    expect(back.for('/a&b.xml')).toBe('application/xml');
  });

  it('writes an empty map as an empty Types element', () => {
    const xml = decoder.decode(ContentTypes.empty().serialize());
    expect(xml).toContain('<Types xmlns=');
    expect(xml).toContain('</Types>');
    expect(ContentTypes.parse(new TextEncoder().encode(xml)).defaults).toHaveLength(0);
  });
});

describe('which parts hold XML', () => {
  it('answers for the +xml suffix, which covers almost every part', () => {
    expect(isXmlContentType(CONTENT_TYPE.slide)).toBe(true);
    expect(isXmlContentType(CONTENT_TYPE.theme)).toBe(true);
    expect(isXmlContentType(CONTENT_TYPE.relationships)).toBe(true);
  });

  it('answers for the two media types that are XML without saying so', () => {
    // `vmlDrawing` is the one that bites: every OLE object in every deck has
    // one, it carries the on-slide preview and the `@spid` that resolves to it,
    // and its content type has no `+xml` anywhere in it.
    expect(isXmlContentType('application/vnd.openxmlformats-officedocument.vmlDrawing')).toBe(true);
    expect(isXmlContentType('application/inkml+xml')).toBe(true);
    expect(isXmlContentType('application/xml')).toBe(true);
    expect(isXmlContentType('text/xml')).toBe(true);
  });

  it('is case-insensitive, because RFC 2045 makes the media type so', () => {
    // The spelling Office writes has a capital D. A lookup table that copied
    // that spelling and lower-cased its input would match nothing at all, which
    // is exactly the bug this test was written after.
    expect(isXmlContentType('application/vnd.openxmlformats-OFFICEdocument.VMLDRAWING')).toBe(true);
    expect(isXmlContentType('APPLICATION/XML')).toBe(true);
  });

  it('ignores parameters', () => {
    expect(isXmlContentType('application/xml; charset=utf-8')).toBe(true);
  });

  it('says no to media, and to nothing at all', () => {
    expect(isXmlContentType(CONTENT_TYPE.png)).toBe(false);
    expect(isXmlContentType(FONT_DATA_CONTENT_TYPE)).toBe(false);
    expect(isXmlContentType('application/vnd.ms-office.vbaProject')).toBe(false);
    expect(isXmlContentType(undefined)).toBe(false);
  });
});
