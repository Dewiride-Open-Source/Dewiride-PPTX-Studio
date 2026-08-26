import { describe, expect, it } from 'vitest';
import { isOpcError } from './errors.js';
import { escapeAttribute, localName, readFlatXml, XML_DECLARATION } from './flat-xml.js';

const encoder = new TextEncoder();

function parse(xml: string) {
  return readFlatXml(encoder.encode(xml), 'test.xml');
}

/** The code of the `OpcError` thrown by `fn`, or a description of what it did instead. */
function codeOf(fn: () => unknown): string {
  try {
    fn();
  } catch (error) {
    if (isOpcError(error)) return error.code;
    return 'NOT_AN_OPC_ERROR: ' + (error instanceof Error ? error.constructor.name : typeof error);
  }
  return 'DID_NOT_THROW';
}

describe('reading the two package grammars', () => {
  it('reads a content-type stream shaped like the one Office writes', () => {
    const elements = parse(
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
        '<Default Extension="rels" ContentType="application/x-rels"/>' +
        '<Override PartName="/ppt/presentation.xml" ContentType="application/x-pml"/>' +
        '</Types>',
    );
    expect(elements.map((e) => e.name)).toEqual(['Types', 'Default', 'Override']);
    expect(elements.map((e) => e.depth)).toEqual([0, 1, 1]);
    expect(elements[1]!.attrs.get('Extension')).toBe('rels');
    expect(elements[2]!.attrs.get('PartName')).toBe('/ppt/presentation.xml');
  });

  it('accepts single quotes, odd whitespace and a prefixed root', () => {
    const elements = parse(
      "<ct:Types  xmlns:ct='urn:x' >\n  <ct:Default\n   Extension = 'xml'\tContentType='application/xml' />\n</ct:Types>",
    );
    expect(elements[0]!.name).toBe('Types');
    expect(elements[0]!.qname).toBe('ct:Types');
    expect(elements[1]!.attrs.get('Extension')).toBe('xml');
  });

  it('skips comments, processing instructions and CDATA', () => {
    const elements = parse(
      '<?xml version="1.0"?><!-- <Default Extension="lies"/> --><Types><?php ?>' +
        '<Default Extension="xml"/><![CDATA[<Default Extension="alsolies"/>]]></Types>',
    );
    expect(elements.map((e) => e.name)).toEqual(['Types', 'Default']);
  });

  it('decodes the five predefined entities and numeric references', () => {
    const elements = parse(
      '<Relationships><Relationship Target="a.aspx?x=1&amp;y=2&lt;3&gt;&quot;&apos;&#65;&#x42;"/></Relationships>',
    );
    expect(elements[1]!.attrs.get('Target')).toBe('a.aspx?x=1&y=2<3>"\'AB');
  });

  it('reports the element name and depth for a nested document', () => {
    const elements = parse('<a><b><c/></b><d/></a>');
    expect(elements.map((e) => `${e.name}@${String(e.depth)}`)).toEqual([
      'a@0',
      'b@1',
      'c@2',
      'd@1',
    ]);
  });
});

describe('input that is trying something', () => {
  it('refuses a DOCTYPE outright, with its own code', () => {
    expect(codeOf(() => parse('<!DOCTYPE Types><Types/>'))).toBe('ERR_DOCTYPE_FORBIDDEN');
  });

  it('refuses the billion-laughs shape before expanding anything', () => {
    const xml =
      '<!DOCTYPE Types [<!ENTITY a "aaaaaaaaaa"><!ENTITY b "&a;&a;&a;&a;&a;&a;&a;&a;&a;&a;">]>' +
      '<Types><Default Extension="&b;"/></Types>';
    expect(codeOf(() => parse(xml))).toBe('ERR_DOCTYPE_FORBIDDEN');
  });

  it('refuses an external entity reference even with no DOCTYPE to declare it', () => {
    expect(codeOf(() => parse('<Types><Default Extension="&xxe;"/></Types>'))).toBe(
      'ERR_MALFORMED_XML',
    );
  });

  it('refuses an entity declaration on its own', () => {
    expect(codeOf(() => parse('<!ENTITY x "y"><Types/>'))).toBe('ERR_DOCTYPE_FORBIDDEN');
  });

  it('refuses a character reference that is not a code point', () => {
    expect(codeOf(() => parse('<Types a="&#xD800;"/>'))).toBe('ERR_MALFORMED_XML');
    expect(codeOf(() => parse('<Types a="&#1114112;"/>'))).toBe('ERR_MALFORMED_XML');
  });

  it('refuses a second root element', () => {
    expect(codeOf(() => parse('<Types/><Types/>'))).toBe('ERR_MALFORMED_XML');
  });

  it('refuses mismatched, unclosed and stray tags', () => {
    expect(codeOf(() => parse('<a><b></a></b>'))).toBe('ERR_MALFORMED_XML');
    expect(codeOf(() => parse('<a><b></b>'))).toBe('ERR_MALFORMED_XML');
    expect(codeOf(() => parse('</a>'))).toBe('ERR_MALFORMED_XML');
  });

  it('refuses malformed attributes', () => {
    expect(codeOf(() => parse('<a b/>'))).toBe('ERR_MALFORMED_XML');
    expect(codeOf(() => parse('<a b=c/>'))).toBe('ERR_MALFORMED_XML');
    expect(codeOf(() => parse('<a b="c/>'))).toBe('ERR_MALFORMED_XML');
    expect(codeOf(() => parse('<a b="1" b="2"/>'))).toBe('ERR_MALFORMED_XML');
  });

  it('refuses an unterminated comment rather than swallowing the document', () => {
    expect(codeOf(() => parse('<Types><!-- forever'))).toBe('ERR_MALFORMED_XML');
  });

  it('refuses bytes that are not valid UTF-8', () => {
    expect(codeOf(() => readFlatXml(new Uint8Array([0x3c, 0x61, 0xff, 0x2f, 0x3e]), 't'))).toBe(
      'ERR_MALFORMED_XML',
    );
  });

  it('refuses an empty document', () => {
    expect(codeOf(() => parse(''))).toBe('ERR_MALFORMED_XML');
  });

  it('enforces its element, attribute and depth ceilings', () => {
    const many = '<a>' + '<b/>'.repeat(50) + '</a>';
    expect(
      codeOf(() =>
        readFlatXml(encoder.encode(many), 't', {
          maxElements: 10,
          maxAttributes: 64,
          maxDepth: 8,
        }),
      ),
    ).toBe('ERR_MALFORMED_XML');

    const deep = '<a>'.repeat(20) + '</a>'.repeat(20);
    expect(codeOf(() => parse(deep))).toBe('ERR_MALFORMED_XML');
  });
});

describe('encoding', () => {
  it('reads UTF-8 with a BOM, and UTF-16 in both byte orders', () => {
    const xml = '<Types a="x"/>';
    const utf8Bom = new Uint8Array([0xef, 0xbb, 0xbf, ...encoder.encode(xml)]);
    expect(readFlatXml(utf8Bom, 't')[0]!.name).toBe('Types');

    const le: number[] = [0xff, 0xfe];
    const be: number[] = [0xfe, 0xff];
    for (const ch of xml) {
      const code = ch.charCodeAt(0);
      le.push(code & 0xff, code >> 8);
      be.push(code >> 8, code & 0xff);
    }
    expect(readFlatXml(new Uint8Array(le), 't')[0]!.attrs.get('a')).toBe('x');
    expect(readFlatXml(new Uint8Array(be), 't')[0]!.attrs.get('a')).toBe('x');
  });

  it('escapes an attribute value so it cannot terminate its own tag', () => {
    expect(escapeAttribute('a&b<c>d"e')).toBe('a&amp;b&lt;c&gt;d&quot;e');
    const round = parse('<a v="' + escapeAttribute('x"/><evil y="') + '"/>');
    expect(round).toHaveLength(1);
    expect(round[0]!.attrs.get('v')).toBe('x"/><evil y="');
  });

  it('strips a prefix', () => {
    expect(localName('ct:Override')).toBe('Override');
    expect(localName('Override')).toBe('Override');
  });

  it('pins the XML declaration to the bytes PowerPoint writes', () => {
    // Transcribed from `[Content_Types].xml` and `_rels/.rels` inside files
    // saved by PowerPoint 365 and PowerPoint 14: the declaration, then 0d 0a,
    // then the whole document on one line. The line break is invisible in any
    // tool that shows the part as text, which is exactly why it is asserted on
    // the bytes.
    expect([...encoder.encode(XML_DECLARATION).slice(-2)]).toEqual([0x0d, 0x0a]);
    expect(XML_DECLARATION).toBe('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n');
    expect(encoder.encode(XML_DECLARATION)[0]).toBe(0x3c); // no BOM
  });
});
