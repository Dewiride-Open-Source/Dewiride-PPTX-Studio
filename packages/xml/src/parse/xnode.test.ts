import { describe, expect, it } from 'vitest';
import { isXmlError } from './errors.js';
import {
  attribute,
  attributeNamespaceOf,
  attributeValue,
  checkTreeCoverage,
  childElements,
  declaredNamespaces,
  descendantElements,
  firstChild,
  namespaceOf,
  namespaceScope,
  parseXml,
  parseXmlString,
  prefixMap,
  resolvePrefix,
  sourceOf,
  startTagOf,
  textContent,
  undeclaredPrefixes,
  xmlSpace,
  DEFAULT_PARSE_LIMITS,
  type XDocument,
  type XElement,
} from './xnode.js';

function codeOf(fn: () => unknown): string {
  try {
    fn();
  } catch (error) {
    if (isXmlError(error)) return error.code;
    return 'NOT_AN_XML_ERROR: ' + (error instanceof Error ? error.constructor.name : typeof error);
  }
  return 'DID_NOT_THROW';
}

/** Parse, and assert the tree accounts for every character of its source. */
function parse(source: string): XDocument {
  const document = parseXmlString(source);
  expect(checkTreeCoverage(document)).toEqual([]);
  return document;
}

const only = (document: XDocument, qname: string): XElement =>
  [...descendantElements(document.root)].find((e) => e.qname === qname)!;

describe('building a tree', () => {
  it('gives every node a span that slices back to its own source', () => {
    const source = '<p:sld><p:cSld><p:spTree/></p:cSld></p:sld>';
    const document = parse(source);
    expect(sourceOf(document, document.root)).toBe(source);
    expect(sourceOf(document, only(document, 'p:cSld'))).toBe('<p:cSld><p:spTree/></p:cSld>');
    expect(sourceOf(document, only(document, 'p:spTree'))).toBe('<p:spTree/>');
  });

  it('keeps the start tag addressable separately from the element', () => {
    // The distinction earns its keep in 0.6: editing a child invalidates the
    // element's span but not its start tag, so the tag is still a slice.
    const document = parse('<a:p a="1"><a:r/></a:p>');
    expect(startTagOf(document, document.root)).toBe('<a:p a="1">');
  });

  it('remembers whether an empty element was written short or long', () => {
    expect(parse('<a><b/></a>').root.children[0]).toMatchObject({ selfClosing: true });
    expect(parse('<a><b></b></a>').root.children[0]).toMatchObject({ selfClosing: false });
  });

  it('holds the prolog and the epilogue as nodes, so the source is reconstructible', () => {
    const source = '<?xml version="1.0"?>\r\n<!--x--><a/>\r\n';
    const document = parse(source);
    expect(document.children.map((c) => c.type)).toEqual([
      'declaration',
      'text',
      'comment',
      'element',
      'text',
    ]);
    expect(document.children.map((c) => source.slice(c.start, c.end)).join('')).toBe(source);
  });

  it('keeps a byte order mark off the tree and on the document', () => {
    const document = parse('﻿<a/>');
    expect(document.bom).toBe(true);
    expect(document.children[0]!.start).toBe(1);
  });

  it('exposes the declaration it read', () => {
    const document = parse('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><a/>');
    expect(document.declaration).toMatchObject({
      version: '1.0',
      encoding: 'UTF-8',
      standalone: 'yes',
    });
    expect(parse('<a/>').declaration).toBeUndefined();
  });

  it('parses from bytes as well as from a string', () => {
    const bytes = new TextEncoder().encode('<a:t>café</a:t>');
    expect(textContent(parseXml(bytes).root)).toBe('café');
  });
});

describe('structural well-formedness', () => {
  it('refuses a mismatched end tag', () => {
    expect(codeOf(() => parseXmlString('<a><b></a></b>'))).toBe('ERR_MISMATCHED_TAG');
    expect(codeOf(() => parseXmlString('<a></b>'))).toBe('ERR_MISMATCHED_TAG');
    expect(codeOf(() => parseXmlString('</a>'))).toBe('ERR_MISMATCHED_TAG');
  });

  it('refuses an element left open at the end', () => {
    expect(codeOf(() => parseXmlString('<a><b/>'))).toBe('ERR_UNCLOSED_TAG');
  });

  it('refuses a document with no root, or with two', () => {
    expect(codeOf(() => parseXmlString('<?xml version="1.0"?>'))).toBe('ERR_ROOT_ELEMENT');
    expect(codeOf(() => parseXmlString('<a/><b/>'))).toBe('ERR_ROOT_ELEMENT');
  });

  it('refuses text outside the root but allows whitespace', () => {
    expect(codeOf(() => parseXmlString('stray<a/>'))).toBe('ERR_MALFORMED_XML');
    expect(parse('  <a/>  ').root.qname).toBe('a');
  });
});

describe('nesting depth', () => {
  const nest = (depth: number): string =>
    '<r>' + '<a:ext>'.repeat(depth) + '</a:ext>'.repeat(depth) + '</r>';

  it('handles nesting far deeper than anything real, without recursing', () => {
    // PowerPoint opens a slide part nested 5000 elements deep, so a
    // recursive-descent builder would meet a well-formed file it cannot read -
    // and would announce it with a `RangeError`, the one thing this package
    // has promised never to throw. The deepest part in our corpus reaches 23.
    const document = parseXmlString(nest(5000), { ...DEFAULT_PARSE_LIMITS, maxDepth: 8192 });
    expect(checkTreeCoverage(document)).toEqual([]);
    let deepest = document.root;
    let depth = 0;
    while (deepest.children.length > 0) {
      deepest = deepest.children[0] as XElement;
      depth++;
    }
    expect(depth).toBe(5000);
  });

  it('stops at the configured limit with a typed error', () => {
    expect(codeOf(() => parseXmlString(nest(2000)))).toBe('ERR_LIMIT_EXCEEDED');
  });

  it('counts nesting, not element count', () => {
    const wide = '<r>' + '<a/>'.repeat(5000) + '</r>';
    expect(parse(wide).root.children).toHaveLength(5000);
  });
});

// ---------------------------------------------------------------------------
// Namespaces. Every case below is taken from a real file or was built and
// opened in PowerPoint, because the plausible shortcuts are all wrong.
// ---------------------------------------------------------------------------
describe('namespace resolution', () => {
  const SLIDE = '<p:sld xmlns:a="urn:a" xmlns:p="urn:p"><p:cSld><a:t>x</a:t></p:cSld></p:sld>';

  it('resolves a prefix by walking up the tree', () => {
    const document = parse(SLIDE);
    expect(namespaceOf(only(document, 'a:t'))).toBe('urn:a');
    expect(namespaceOf(document.root)).toBe('urn:p');
  });

  it('honours a binding made part-way down the document', () => {
    // 213 namespace declarations in our corpus sit on a non-root element.
    const document = parse('<p:x xmlns:p="urn:p"><p:e><zz:y xmlns:zz="urn:z"/></p:e></p:x>');
    expect(namespaceOf(only(document, 'zz:y'))).toBe('urn:z');
    expect(resolvePrefix(only(document, 'p:e'), 'zz')).toBeUndefined();
  });

  it('distinguishes an undeclared default namespace from an unbound prefix', () => {
    // `xmlns=""` is real: a shipped Microsoft template contains
    // `<p14:discardImageEditData xmlns="" xmlns:p14="..." val="0"/>`.
    const document = parse('<r xmlns="urn:d"><c xmlns=""><g/></c></r>');
    expect(namespaceOf(document.root)).toBe('urn:d');
    expect(namespaceOf(only(document, 'c'))).toBe('');
    expect(namespaceOf(only(document, 'g'))).toBe('');
    expect(resolvePrefix(document.root, 'nope')).toBeUndefined();
  });

  it('lets one prefix mean two different things in sibling subtrees', () => {
    // Verified in PowerPoint, and the reason a package-wide or even part-wide
    // prefix table is a bug. Across our corpus `p14` is bound to
    // .../2010/main in 46 parts and .../2007/7/12/main in 8 others.
    const document = parse(
      '<r xmlns:p="urn:p"><a><zz:x xmlns:zz="urn:one"/></a><b><zz:x xmlns:zz="urn:two"/></b></r>',
    );
    const [inA, inB] = [...descendantElements(document.root)].filter((e) => e.qname === 'zz:x');
    expect(namespaceOf(inA!)).toBe('urn:one');
    expect(namespaceOf(inB!)).toBe('urn:two');
    expect(prefixMap(document).get('zz')).toEqual(new Set(['urn:one', 'urn:two']));
  });

  it('puts an unprefixed attribute in no namespace, not the default one', () => {
    // Namespaces in XML §6.2. Getting this wrong makes every unprefixed
    // attribute in a slide part look like it is in the PresentationML namespace.
    const document = parse('<r xmlns="urn:d" xmlns:q="urn:q" plain="1" q:tagged="2"/>');
    expect(attributeNamespaceOf(document.root, attribute(document.root, 'plain')!)).toBe('');
    expect(attributeNamespaceOf(document.root, attribute(document.root, 'q:tagged')!)).toBe(
      'urn:q',
    );
  });

  it('binds the xml and xmlns prefixes without a declaration', () => {
    const document = parse('<r xml:space="preserve"/>');
    expect(resolvePrefix(document.root, 'xml')).toBe('http://www.w3.org/XML/1998/namespace');
    expect(resolvePrefix(document.root, 'xmlns')).toBe('http://www.w3.org/2000/xmlns/');
  });

  it('lists what one element declares, and what is in scope at it', () => {
    const document = parse('<r xmlns="urn:d" xmlns:a="urn:a"><c xmlns:b="urn:b"/></r>');
    expect([...declaredNamespaces(only(document, 'c'))]).toEqual([['b', 'urn:b']]);
    expect([...namespaceScope(only(document, 'c'))].sort()).toEqual([
      ['', 'urn:d'],
      ['a', 'urn:a'],
      ['b', 'urn:b'],
    ]);
  });

  it('reports a prefix nothing binds', () => {
    const document = parseXmlString('<r xmlns:a="urn:a"><zz:x/></r>');
    expect(undeclaredPrefixes(document)).toEqual([{ prefix: 'zz', at: 19 }]);
    expect(undeclaredPrefixes(parse('<r xmlns:a="urn:a"><a:x a:n="1"/></r>'))).toEqual([]);
  });
});

describe('xml:space', () => {
  it('inherits down the tree until something overrides it', () => {
    const document = parse('<r xml:space="preserve"><a><b xml:space="default"><c/></b></a></r>');
    expect(xmlSpace(document.root)).toBe('preserve');
    expect(xmlSpace(only(document, 'a'))).toBe('preserve');
    expect(xmlSpace(only(document, 'c'))).toBe('default');
  });

  it('defaults to "default" when nothing declares it', () => {
    // Which is every part of our corpus: there is no xml:space anywhere in it,
    // 120 <a:t> elements carry edge whitespace without one, and PowerPoint
    // round-trips such a part with the whitespace intact. So this is a reader,
    // never a writer.
    const document = parse('<a:t>trailing   </a:t>');
    expect(xmlSpace(document.root)).toBe('default');
    expect(textContent(document.root)).toBe('trailing   ');
  });
});

describe('reading a tree', () => {
  const DOC = parse(
    '<p:sp><p:nvSpPr n="1"/><p:txBody><a:p><a:r><a:t>Hello </a:t></a:r><a:r><a:t>world</a:t></a:r></a:p></p:txBody></p:sp>',
  );

  it('finds children by name and lists element children', () => {
    expect(firstChild(DOC.root, 'p:txBody')?.qname).toBe('p:txBody');
    expect(firstChild(DOC.root, 'nope')).toBeUndefined();
    expect(childElements(DOC.root).map((c) => c.qname)).toEqual(['p:nvSpPr', 'p:txBody']);
  });

  it('reads attributes by qualified name', () => {
    expect(attributeValue(only(DOC, 'p:nvSpPr'), 'n')).toBe('1');
    expect(attributeValue(DOC.root, 'missing')).toBeUndefined();
  });

  it('walks descendants in document order', () => {
    expect([...descendantElements(DOC.root)].map((e) => e.qname)).toEqual([
      'p:sp',
      'p:nvSpPr',
      'p:txBody',
      'a:p',
      'a:r',
      'a:t',
      'a:r',
      'a:t',
    ]);
  });

  it('concatenates text across runs, CDATA included', () => {
    expect(textContent(DOC.root)).toBe('Hello world');
    expect(textContent(parse('<a>x<![CDATA[ & ]]>y</a>').root)).toBe('x & y');
  });

  it('links every node to its parent', () => {
    const t = only(DOC, 'a:t');
    expect(t.parent?.qname).toBe('a:r');
    expect(DOC.root.parent).toBeUndefined();
  });
});

describe('coverage checking itself', () => {
  it('reports a gap when a span is moved', () => {
    const document = parseXmlString('<a><b/><c/></a>');
    document.root.children[1]!.start += 1;
    expect(checkTreeCoverage(document)[0]).toMatchObject({ kind: 'gap' });
  });

  it('reports children hanging off a self-closing element', () => {
    const document = parseXmlString('<a><b/><c/></a>');
    const [b, c] = document.root.children as XElement[];
    b!.children.push(c!);
    expect(checkTreeCoverage(document)[0]).toMatchObject({ kind: 'malformed' });
  });

  it('is clean for every shape in this file', () => {
    for (const source of [
      '<a/>',
      '﻿<?xml version="1.0"?><a/>',
      '<?xml version="1.0"?>\r\n<!--c--><a><?pi?><b></b><![CDATA[x]]>t</a>\r\n',
      '<a>\r\n  <b n="1" />\r\n</a>',
      '<a n=\'1\' m="2"/>',
    ]) {
      expect(checkTreeCoverage(parseXmlString(source))).toEqual([]);
    }
  });
});
