import { describe, expect, it } from 'vitest';
import { canonicalXml, excerpt, firstDifference } from './canonical.js';
import { NS } from './namespaces.js';
import { parseXmlString } from './xnode.js';

/**
 * The two halves of a normal form, and both have to be tested.
 *
 * A canonical form that erases too much reports two different documents as the
 * same, which is the failure that matters: it is silent, and it turns the
 * round-trip gate green while the deck loses something. A form that erases too
 * little reports a difference where there is none, which is loud, annoying and
 * gets the test disabled.
 *
 * So each block below is a pair. `same` collects the differences the form is
 * *supposed* to absorb; `different` collects the ones it must not.
 */

const canon = (xml: string): string => canonicalXml(parseXmlString(xml));

function same(name: string, a: string, b: string): void {
  it(name, () => {
    expect(canon(a)).toBe(canon(b));
  });
}

function different(name: string, a: string, b: string): void {
  it(name, () => {
    expect(canon(a)).not.toBe(canon(b));
  });
}

describe('what two producers are free to disagree about', () => {
  same('attribute order', '<a:off x="1" y="2"/>', '<a:off y="2" x="1"/>');

  same('quote style', `<a:off x='1'/>`, '<a:off x="1"/>');

  same('the long and short spelling of an empty element', '<a:off></a:off>', '<a:off/>');

  same('whitespace before the slash of an empty element', '<a:off x="1" />', '<a:off x="1"/>');

  same(
    'a character reference and the character it names',
    '<a:t>a&#x26;b</a:t>',
    '<a:t>a&amp;b</a:t>',
  );

  same('CDATA and the text it holds', '<a:t><![CDATA[a<b]]></a:t>', '<a:t>a&lt;b</a:t>');

  same(
    'a tab written literally in an attribute and one written as a reference',
    // XML 1.0 3.3.3 turns a literal tab into a space during normalisation and
    // leaves a reference to one alone, so these two are *not* the same value.
    // The pair here is the one that is: a reference either side.
    '<a:t v="a&#x9;b"/>',
    '<a:t v="a&#9;b"/>',
  );

  same(
    'the encoding named in the declaration',
    '<?xml version="1.0" encoding="UTF-8"?><a:t/>',
    '<?xml version="1.0"?><a:t/>',
  );

  it('absorbs a byte order mark', () => {
    expect(canon('\ufeff<a:t/>')).toBe(canon('<a:t/>'));
  });
});

describe('what carries meaning and must survive', () => {
  different('an attribute value', '<a:off x="1"/>', '<a:off x="2"/>');

  different('an attribute nothing else changed about', '<a:off x="1"/>', '<a:off x="1" y="2"/>');

  different('the order of two children', '<p:sp><a/><b/></p:sp>', '<p:sp><b/><a/></p:sp>');

  different(
    'a namespace prefix, even when the URI is the same',
    '<x:t xmlns:x="urn:a">v</x:t>',
    '<y:t xmlns:y="urn:a">v</y:t>',
  );

  different(
    'a namespace declaration nothing in the document uses',
    // The case C14N gets wrong for OOXML. PowerPoint writes the declaration so
    // that an `mc:Ignorable` token resolves, and there is often no element in
    // that namespace anywhere in the part.
    '<p:sld xmlns:p="urn:p" xmlns:a14="urn:a14" mc:Ignorable="a14" xmlns:mc="urn:mc"/>',
    '<p:sld xmlns:p="urn:p" mc:Ignorable="a14" xmlns:mc="urn:mc"/>',
  );

  different(
    'whitespace between elements',
    '<p:sp><a/><b/></p:sp>',
    '<p:sp>\n  <a/>\n  <b/>\n</p:sp>',
  );

  different(
    'leading whitespace inside a text run',
    // The reason whitespace is not collapsed. `V029` exists because this
    // difference is invisible to a producer and visible on a slide.
    '<a:t> hello</a:t>',
    '<a:t>hello</a:t>',
  );

  different('a comment', '<a:t><!-- note -->v</a:t>', '<a:t>v</a:t>');

  different('the presence of an XML declaration', '<?xml version="1.0"?><a:t/>', '<a:t/>');

  different(
    'the standalone pseudo-attribute',
    '<?xml version="1.0" standalone="yes"?><a:t/>',
    '<?xml version="1.0"?><a:t/>',
  );
});

describe('the shape of the output', () => {
  it('sorts namespace declarations ahead of everything else', () => {
    const out = canon('<p:sld a="1" xmlns:p="urn:p" xmlns="urn:d" p:b="2"/>');
    expect(out).toBe('<p:sld xmlns="urn:d" xmlns:p="urn:p" a="1" p:b="2"/>');
  });

  it('sorts unprefixed attributes before namespaced ones', () => {
    // An unprefixed attribute is in no namespace - it does not inherit the
    // element's default - so its empty URI sorts first.
    const out = canon('<t xmlns:z="urn:z" z:b="2" a="1"/>');
    expect(out).toBe('<t xmlns:z="urn:z" a="1" z:b="2"/>');
  });

  it('sorts namespaced attributes by URI and not by prefix', () => {
    // `z:` is bound to the earlier URI, so it comes first despite the prefix
    // sorting later. A form that sorted on the prefix would call two documents
    // different for having renamed a prefix consistently - which they have, and
    // which is a difference, but it would be *reported at the wrong place*.
    const out = canon('<t xmlns:a="urn:b" xmlns:z="urn:a" a:one="1" z:two="2"/>');
    expect(out).toBe('<t xmlns:a="urn:b" xmlns:z="urn:a" z:two="2" a:one="1"/>');
  });

  it('escapes the same characters C14N does', () => {
    expect(canon('<t v="a&lt;b&gt;c&amp;d&#x9;e">x&lt;y&gt;z</t>')).toBe(
      '<t v="a&lt;b&gt;c&amp;d&#x9;e">x&lt;y&gt;z</t>',
    );
  });

  it('keeps a processing instruction', () => {
    expect(canon('<?xml version="1.0"?><?mso-application progid="PowerPoint"?><t/>')).toBe(
      '<?xml version="1.0"?><?mso-application progid="PowerPoint"?><t/>',
    );
  });

  it('walks a deeply nested part without recursing', () => {
    const depth = 400;
    const xml = '<r>' + '<n>'.repeat(depth) + 'x' + '</n>'.repeat(depth) + '</r>';
    expect(canon(xml)).toBe(xml);
  });
});

describe('rewriting relationship ids', () => {
  const R = NS.r;
  const doc = parseXmlString(
    '<p:pic xmlns:p="urn:p" xmlns:r="' + R + '"><a:blip r:embed="rId3" cstate="print"/></p:pic>',
  );

  it('rewrites only attributes in the relationships namespace', () => {
    const out = canonicalXml(doc, {
      rewriteAttributeValue: (value, namespaceUri) =>
        namespaceUri === R && value === 'rId3' ? 'rId9' : value,
    });
    expect(out).toContain('r:embed="rId9"');
    expect(out).toContain('cstate="print"');
  });

  it('is told the attribute is in no namespace when it is unprefixed', () => {
    const seen: [string, string][] = [];
    canonicalXml(doc, {
      rewriteAttributeValue: (value, namespaceUri, local) => {
        seen.push([local, namespaceUri]);
        return value;
      },
    });
    expect(seen).toContainEqual(['cstate', '']);
    expect(seen).toContainEqual(['embed', R]);
  });

  it('reports the xmlns namespace for a declaration', () => {
    const seen = new Set<string>();
    canonicalXml(doc, {
      rewriteAttributeValue: (value, namespaceUri) => {
        seen.add(namespaceUri);
        return value;
      },
    });
    expect(seen.has(NS.xmlns)).toBe(true);
  });
});

describe('locating a difference', () => {
  it('reports the first index at which two forms diverge', () => {
    expect(firstDifference('<a:off x="1"/>', '<a:off x="2"/>')).toBe(10);
  });

  it('reports -1 when they agree', () => {
    expect(firstDifference('same', 'same')).toBe(-1);
  });

  it('reports the length of the shorter when one is a prefix of the other', () => {
    expect(firstDifference('abc', 'abcd')).toBe(3);
  });

  it('excerpts a window around the difference', () => {
    const text = 'x'.repeat(200) + 'HERE' + 'y'.repeat(200);
    const window = excerpt(text, 200, 4);
    expect(window).toBe('...xxxxHERE...');
  });

  it('does not mark a window that reaches the end of the string', () => {
    expect(excerpt('abcdef', 3, 10)).toBe('abcdef');
  });
});
