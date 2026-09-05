import { describe, expect, it } from 'vitest';
import { isXmlError } from '../errors.js';
import { escapeAttributeValue, escapeText } from '../mce/references.js';
import { encodeXmlSource } from '../parse/source.js';
import { serializeNode, serializeXml, serializeXmlString } from './serialize.js';
import { checkRoundTrip } from './roundtrip.js';
import type { XAttribute } from '../parse/tokenizer.js';
import {
  checkDirtyInvariant,
  checkTreeCoverage,
  DEFAULT_PARSE_LIMITS,
  descendantElements,
  markAttributeDirty,
  markDirty,
  parseXml,
  parseXmlString,
  type XDocument,
  type XElement,
  type XNode,
} from '../parse/xnode.js';

function codeOf(fn: () => unknown): string {
  try {
    fn();
  } catch (error) {
    if (isXmlError(error)) return error.code;
    return 'NOT_AN_XML_ERROR: ' + (error instanceof Error ? error.constructor.name : typeof error);
  }
  return 'DID_NOT_THROW';
}

/**
 * Mark every node and every attribute dirty, so nothing can take the slice path.
 *
 * A test instrument, and the most valuable one in this file. Without it the
 * round-trip gate only ever exercises `source.slice(start, end)`, which is a
 * property of the tokenizer rather than of the serializer.
 */
function forceRebuild(document: XDocument): XDocument {
  const stack: XNode[] = [...document.children];
  while (stack.length > 0) {
    const node = stack.pop()!;
    node.dirty = true;
    if (node.type !== 'element') continue;
    for (const attribute of node.attributes) attribute.dirty = true;
    for (const child of node.children) stack.push(child);
  }
  return document;
}

/** Where two strings first disagree, or -1. Better failure output than a diff. */
function firstDifference(a: string, b: string): number {
  const limit = Math.min(a.length, b.length);
  for (let i = 0; i < limit; i++) if (a.charCodeAt(i) !== b.charCodeAt(i)) return i;
  return a.length === b.length ? -1 : limit;
}

function element(qname: string, parent: XElement | undefined): XElement {
  const colon = qname.indexOf(':');
  return {
    type: 'element',
    qname,
    prefix: colon < 0 ? '' : qname.slice(0, colon),
    local: colon < 0 ? qname : qname.slice(colon + 1),
    attributes: [],
    children: [],
    selfClosing: true,
    start: 0,
    end: 0,
    openTagEnd: 0,
    closeTagStart: 0,
    nameEnd: 0,
    trailingSpaceStart: 0,
    hasNamespaceDeclarations: false,
    dirty: true,
    parent,
  };
}

const SLIDE =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
  'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">' +
  '<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
  '<p:grpSpPr><a:xfrm><a:off x="0" y="0" /><a:ext cx="0" cy="0" /></a:xfrm></p:grpSpPr>' +
  '<p:sp><p:nvSpPr><p:cNvPr id="4" name="Rectangle 3"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>' +
  '<p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr><p:spPr></p:spPr><p:txBody><a:bodyPr/>' +
  '<a:lstStyle><a:extLst/></a:lstStyle><a:p><a:r><a:rPr lang="en-US" dirty="0"/>' +
  '<a:t>Hello &amp; welcome — café 日本語</a:t></a:r></a:p></p:txBody></p:sp>' +
  '</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>';

/** Every lexical shape the tokenizer distinguishes, in one list. */
const DOCUMENTS: readonly string[] = [
  '<a/>',
  '<a />',
  '<a   />',
  '<a></a>',
  '<a>text</a>',
  '<a b="1"/>',
  "<a b='1'/>",
  '<a b="1" />',
  '<a b=\'1\' c="2" d = "3"/>',
  '<a\r\n  b = "1"\r\n  c = "2"\r\n/>',
  '<a><!-- c --><?pi d?><![CDATA[x]]>t</a>',
  '<a><?pi?></a>',
  '<r xmlns="urn:d"><c xmlns=""><g a=\'1\' b="2" /></c></r>',
  '<?xml version="1.0"?><a/>',
  "<?xml version='1.0' encoding='utf-8' standalone='no'?><a/>",
  '<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>\r\n<a>\r\n  <b/>\r\n</a>\r\n',
  '\ufeff<?xml version="1.0"?><a/>',
  '\ufeff<a>café 日本語 \u{1F600}</a>',
  '<a>Smith &amp; Sons &lt;tag&gt; &quot;q&quot; &apos;a&apos;</a>',
  '<a>&#72;&#x69; ]]&gt; end</a>',
  '<a b="x&#9;y" c="p&#xA;q" d="&quot;&amp;&lt;"/>',
  '<a>  <b/>  <!--x-->  </a>',
  SLIDE,
];

describe('a document nobody edited serializes to itself', () => {
  it('is byte-identical for every lexical shape', () => {
    for (const source of DOCUMENTS) {
      const document = parseXmlString(source);
      expect(checkTreeCoverage(document)).toEqual([]);
      const serialized = serializeXmlString(document);
      expect({ source, at: firstDifference(source, serialized) }).toEqual({ source, at: -1 });
    }
  });

  it('is byte-identical through the bytes, not just the string', () => {
    for (const source of DOCUMENTS) {
      const bytes = encodeXmlSource(source);
      expect(Array.from(serializeXml(parseXml(bytes)))).toEqual(Array.from(bytes));
    }
  });

  it('preserves a byte order mark, which is not a node', () => {
    const bytes = encodeXmlSource('\ufeff<a/>');
    expect(Array.from(bytes.slice(0, 3))).toEqual([0xef, 0xbb, 0xbf]);
    expect(Array.from(serializeXml(parseXml(bytes)))).toEqual(Array.from(bytes));
  });

  it('agrees with itself about every subtree', () => {
    const document = parseXmlString(SLIDE);
    for (const child of document.root.children) {
      expect(serializeNode(document, child)).toBe(document.source.slice(child.start, child.end));
    }
  });
});

describe('the rebuild path, which the byte gate cannot see', () => {
  // `return document.source` passes the gate above. Everything below forces
  // every node through the rebuild instead, which is the half sub-phase 0.6
  // runs on.

  it('reparses to the same tree for every lexical shape', () => {
    for (const source of DOCUMENTS) {
      expect({ source, differences: checkRoundTrip(forceRebuild(parseXmlString(source))) }).toEqual(
        {
          source,
          differences: [],
        },
      );
    }
  });

  it('is byte-identical too, wherever the original spelling was already minimal', () => {
    // Not a guarantee and deliberately not asserted as one - `&#62;` and `>`
    // are the same character data, and we write whichever is shorter. But a
    // real slide part, written by PowerPoint, comes back exactly.
    const document = forceRebuild(parseXmlString(SLIDE));
    expect(firstDifference(SLIDE, serializeXmlString(document))).toBe(-1);
  });

  it('keeps the quote character an attribute was written with', () => {
    const document = forceRebuild(parseXmlString('<a b=\'1\' c="2"/>'));
    expect(serializeXmlString(document)).toBe('<a b=\'1\' c="2"/>');
  });

  it('keeps the space before a slash, and the indentation between attributes', () => {
    // 70 822 of 98 777 self-closing tags in the corpus are written this way.
    expect(serializeXmlString(forceRebuild(parseXmlString('<a:off x="0" y="0" />')))).toBe(
      '<a:off x="0" y="0" />',
    );
    expect(serializeXmlString(forceRebuild(parseXmlString('<a\r\n  b="1"\r\n/>')))).toBe(
      '<a\r\n  b="1"\r\n/>',
    );
  });

  it('keeps the long empty form, and the short one', () => {
    expect(serializeXmlString(forceRebuild(parseXmlString('<a></a>')))).toBe('<a></a>');
    expect(serializeXmlString(forceRebuild(parseXmlString('<a/>')))).toBe('<a/>');
    expect(serializeXmlString(forceRebuild(parseXmlString('<a />')))).toBe('<a />');
  });

  it('re-escapes "]]>" the way it was written, because text may not contain it', () => {
    const document = forceRebuild(parseXmlString('<a>x ]]&gt; y</a>'));
    expect(document.root.children[0]).toMatchObject({ value: 'x ]]> y' });
    expect(serializeXmlString(document)).toBe('<a>x ]]&gt; y</a>');
  });

  it('escapes ">" even where nothing forces it, because the corpus says to', () => {
    // Only `]]>` actually requires it, so the narrow rule is to write a bare
    // `>` everywhere else - which is what this did first, reasoning that a
    // rebuilt node should keep its producer's spelling. The corpus disagreed:
    // across 22 461 text nodes and 170 019 attribute values there are **zero**
    // literal `>` and **five** `&gt;`. The narrow rule re-spelled all five and
    // preserved nothing, so the broad rule is the one that loses less.
    expect(serializeXmlString(forceRebuild(parseXmlString('<a>1 &gt; 0</a>')))).toBe(
      '<a>1 &gt; 0</a>',
    );
    expect(serializeXmlString(forceRebuild(parseXmlString('<a b="1 &gt; 0"/>')))).toBe(
      '<a b="1 &gt; 0"/>',
    );
    // The bare form is still read correctly - it is legal input, just not
    // output we produce.
    expect(parseXmlString('<a>1 > 0</a>').root.children[0]).toMatchObject({ value: '1 > 0' });
    expect(serializeXmlString(forceRebuild(parseXmlString('<a>1 > 0</a>')))).toBe(
      '<a>1 &gt; 0</a>',
    );
  });

  it('normalizes CRLF inside a text node it rebuilds, and only there', () => {
    // CRLF and LF are the same value after §2.11, so a rebuilt text node cannot
    // know which it came from. It costs nothing in practice: `dirty` never
    // travels downwards, so a text node is only rebuilt when someone edited
    // *it*, and an untouched sibling keeps its bytes.
    const source = '<a>\r\n  <b/>\r\n</a>';
    expect(serializeXmlString(parseXmlString(source))).toBe(source);
    expect(serializeXmlString(forceRebuild(parseXmlString(source)))).toBe('<a>\n  <b/>\n</a>');

    const document = parseXmlString(source);
    markDirty(document.root); // the element, not its text children
    expect(serializeXmlString(document)).toBe(source);
  });
});

describe('the four characters a naive serializer loses', () => {
  // Writing a value out as-is loses each of these, silently, because XML
  // transforms the input on the way back in. This is the two-tier normalization
  // trap running in the opposite direction.

  it('keeps a carriage return in text, which §2.11 would turn into a line feed', () => {
    const document = parseXmlString('<a>x</a>');
    const text = document.root.children[0]!;
    if (text.type !== 'text') throw new Error('expected a text node');
    text.value = 'line1\rline2';
    markDirty(text);

    expect(serializeXmlString(document)).toBe('<a>line1&#xD;line2</a>');
    const back = parseXmlString(serializeXmlString(document));
    expect(back.root.children[0]).toMatchObject({ value: 'line1\rline2' });
    expect(checkRoundTrip(document)).toEqual([]);
  });

  it('keeps a tab, line feed and carriage return in an attribute, which §3.3.3 would flatten', () => {
    for (const [value, written] of [
      ['a\tb', 'a&#x9;b'],
      ['a\nb', 'a&#xA;b'],
      ['a\rb', 'a&#xD;b'],
    ] as const) {
      const document = parseXmlString('<a n="x"/>');
      const attribute = document.root.attributes[0]!;
      attribute.value = value;
      markAttributeDirty(document.root, attribute);

      expect(serializeXmlString(document)).toBe('<a n="' + written + '"/>');
      const back = parseXmlString(serializeXmlString(document));
      expect(back.root.attributes[0]!.value).toBe(value);
      expect(checkRoundTrip(document)).toEqual([]);
    }
  });

  it('would have flattened them, had they been written literally', () => {
    // The counterfactual, so the test above is visibly load-bearing rather than
    // an assertion about a string constant.
    expect(parseXmlString('<a n="a\tb"/>').root.attributes[0]!.value).toBe('a b');
    expect(parseXmlString('<a>a\rb</a>').root.children[0]).toMatchObject({ value: 'a\nb' });
  });

  it('escapes only the delimiter in use', () => {
    expect(escapeAttributeValue('he said "hi" and \'bye\'', '"')).toBe(
      "he said &quot;hi&quot; and 'bye'",
    );
    expect(escapeAttributeValue('he said "hi" and \'bye\'', "'")).toBe(
      'he said "hi" and &apos;bye&apos;',
    );
  });

  it('round-trips any value through either escaper', () => {
    const values = [
      '',
      'plain',
      '&<>"\'',
      'a\tb\nc\rd',
      ']]>',
      ']]]>',
      '&amp;',
      'café 日本語 \u{1F600}',
      '  padded  ',
    ];
    for (const value of values) {
      const asText = parseXmlString('<a>' + escapeText(value) + '</a>');
      expect({
        value,
        back: asText.root.children[0]?.type === 'text' ? asText.root.children[0].value : '',
      }).toEqual({ value, back: value });
      for (const quote of ['"', "'"] as const) {
        const asAttribute = parseXmlString(
          '<a n=' + quote + escapeAttributeValue(value, quote) + quote + '/>',
        );
        expect({ value, quote, back: asAttribute.root.attributes[0]!.value }).toEqual({
          value,
          quote,
          back: value,
        });
      }
    }
  });
});

describe('an edit reaches exactly as far as it was made', () => {
  it('rewrites one attribute and nothing else', () => {
    const document = parseXmlString(SLIDE);
    const target = [...descendantElements(document.root)].find(
      (el) => el.qname === 'p:cNvPr' && el.attributes.some((a) => a.value === 'Rectangle 3'),
    )!;
    const name = target.attributes.find((a) => a.qname === 'name')!;

    name.value = 'Rectangle 4';
    markAttributeDirty(target, name);

    const serialized = serializeXmlString(document);
    expect(serialized).toBe(SLIDE.replace('Rectangle 3', 'Rectangle 4'));
    expect(checkRoundTrip(document)).toEqual([]);
    expect(checkDirtyInvariant(document)).toEqual([]);
  });

  it('marks the ancestors of an edited node and nothing else', () => {
    const document = parseXmlString('<r><a><b><c/></b></a><d><e/></d></r>');
    const a = document.root.children[0] as XElement;
    const b = a.children[0] as XElement;
    const c = b.children[0] as XElement;
    const d = document.root.children[1] as XElement;

    markDirty(c);

    expect([document.root.dirty, a.dirty, b.dirty, c.dirty]).toEqual([true, true, true, true]);
    expect([d.dirty, (d.children[0] as XElement).dirty]).toEqual([false, false]);
    // The untouched sibling still re-emits as a slice, which is the point.
    expect(serializeNode(document, d)).toBe('<d><e/></d>');
    expect(serializeXmlString(document)).toBe('<r><a><b><c/></b></a><d><e/></d></r>');
  });

  it('gives an element that was written empty a real body when a child arrives', () => {
    const document = parseXmlString('<p:sp><p:spPr/></p:sp>');
    const spPr = document.root.children[0] as XElement;
    const noFill = element('a:noFill', spPr);
    spPr.children.push(noFill);
    markDirty(spPr);

    expect(serializeXmlString(document)).toBe('<p:sp><p:spPr><a:noFill/></p:spPr></p:sp>');
    expect(checkRoundTrip(document)).toEqual([]);
  });

  it('drops an attribute without disturbing the spacing of the rest', () => {
    const document = parseXmlString('<a:off x="0" y="0" z="0" />');
    document.root.attributes.splice(0, 1);
    markDirty(document.root);
    expect(serializeXmlString(document)).toBe('<a:off y="0" z="0" />');
  });

  it('inserts an attribute with one space, and keeps the neighbours as they were', () => {
    const document = parseXmlString('<a:off\n  x="0"\n  y="0"\n/>');
    const inserted: XAttribute = {
      qname: 'z',
      prefix: '',
      local: 'z',
      start: 0,
      nameStart: 0,
      nameEnd: 0,
      valueStart: 0,
      valueEnd: 0,
      end: 0,
      quote: '"',
      value: '1',
      dirty: true,
    };
    document.root.attributes.push(inserted);
    markDirty(document.root);
    expect(serializeXmlString(document)).toBe('<a:off\n  x="0"\n  y="0" z="1"\n/>');
  });
});

describe('the serializer refuses what it cannot write', () => {
  function withValue(source: string, mutate: (node: XNode) => void): () => string {
    const document = parseXmlString(source);
    const node = document.root.children[0]!;
    mutate(node);
    markDirty(node);
    return () => serializeXmlString(document);
  }

  it('refuses a comment or CDATA section that would swallow its own terminator', () => {
    expect(
      codeOf(withValue('<a><!--c--></a>', (n) => ((n as { value: string }).value = 'a--b'))),
    ).toBe('ERR_MALFORMED_XML');
    expect(
      codeOf(withValue('<a><!--c--></a>', (n) => ((n as { value: string }).value = 'trailing-'))),
    ).toBe('ERR_MALFORMED_XML');
    expect(
      codeOf(withValue('<a><![CDATA[c]]></a>', (n) => ((n as { value: string }).value = 'a]]>b'))),
    ).toBe('ERR_MALFORMED_XML');
    expect(
      codeOf(withValue('<a><?pi d?></a>', (n) => ((n as { data: string }).data = ' a?>b'))),
    ).toBe('ERR_MALFORMED_XML');
  });

  it('refuses a carriage return where nothing can escape one', () => {
    // Text has `&#xD;`; a comment and a CDATA section have nothing, so a value
    // holding one would come back as a line feed.
    expect(
      codeOf(withValue('<a><!--c--></a>', (n) => ((n as { value: string }).value = 'a\rb'))),
    ).toBe('ERR_MALFORMED_XML');
    expect(
      codeOf(withValue('<a><![CDATA[c]]></a>', (n) => ((n as { value: string }).value = 'a\rb'))),
    ).toBe('ERR_MALFORMED_XML');
  });

  it('refuses a character XML does not permit, and an unpaired surrogate', () => {
    expect(
      codeOf(withValue('<a>t</a>', (n) => ((n as { value: string }).value = 'a\u000Bb'))),
    ).toBe('ERR_INVALID_CHARACTER');
    expect(
      codeOf(withValue('<a>t</a>', (n) => ((n as { value: string }).value = 'a\ud800b'))),
    ).toBe('ERR_MALFORMED_XML');
    const document = parseXmlString('<a n="v"/>');
    const attribute = document.root.attributes[0]!;
    attribute.value = 'a\ud800';
    markAttributeDirty(document.root, attribute);
    expect(codeOf(() => serializeXmlString(document))).toBe('ERR_MALFORMED_XML');
  });

  it('refuses a name that is not one', () => {
    for (const qname of ['', 'a b', '1a', 'a:b:c', ':a', 'a:', 'a<b']) {
      const document = parseXmlString('<a/>');
      document.root.qname = qname;
      markDirty(document.root);
      expect({ qname, code: codeOf(() => serializeXmlString(document)) }).toEqual({
        qname,
        code: 'ERR_MALFORMED_XML',
      });
    }
  });

  it('refuses a clean node whose span does not index the source', () => {
    // The shape of a synthesized node someone forgot to mark dirty. Slicing it
    // would silently emit whatever happened to be at those offsets.
    const document = parseXmlString('<a><b/></a>');
    const stray = element('c', document.root);
    stray.dirty = false;
    stray.start = 900;
    stray.end = 999;
    document.root.children.push(stray);
    markDirty(document.root);
    expect(codeOf(() => serializeXmlString(document))).toBe('ERR_MALFORMED_XML');
  });

  it('throws an XmlError and never a RangeError, however deep the tree', () => {
    const depth = 5000;
    const source = '<a>'.repeat(depth) + '</a>'.repeat(depth);
    const limits = { ...DEFAULT_PARSE_LIMITS, maxDepth: depth + 1 };
    const document = forceRebuild(parseXmlString(source, limits));
    // A recursive serializer answers this with a RangeError, which is the one
    // thing this package has promised never to throw.
    expect(serializeXmlString(document)).toBe(source);
  });
});

describe('the round-trip check has teeth', () => {
  it('catches a value that was changed without being marked dirty', () => {
    // The footgun the whole dirty contract rests on: the serializer would
    // re-emit the slice and discard the edit, with no error anywhere.
    const document = parseXmlString('<a>x</a>');
    (document.root.children[0] as { value: string }).value = 'y';

    const differences = checkRoundTrip(document);
    expect(differences).toHaveLength(1);
    expect(differences[0]!.kind).toBe('structure');
    expect(differences[0]!.detail).toContain('"y"');
    expect(serializeXmlString(document)).toBe('<a>x</a>');
  });

  it('catches a namespace binding that was changed without being marked dirty', () => {
    const document = parseXmlString('<r xmlns:a="urn:a"><a:c/></r>');
    document.root.attributes[0]!.value = 'urn:b';

    const kinds = checkRoundTrip(document).map((d) => d.kind);
    expect(kinds).toContain('structure');
    expect(kinds).toContain('prefix');
  });

  it('catches an attribute marked dirty without its element', () => {
    const document = parseXmlString('<a n="x"/>');
    const attribute = document.root.attributes[0]!;
    attribute.value = 'y';
    attribute.dirty = true; // and not the element

    expect(checkDirtyInvariant(document)).toHaveLength(1);
    expect(checkDirtyInvariant(document)[0]!.detail).toContain('markAttributeDirty');
    expect(checkRoundTrip(document)).not.toEqual([]);
  });

  it('catches a dirty node under a clean parent', () => {
    const document = parseXmlString('<r><a><b/></a></r>');
    const a = document.root.children[0] as XElement;
    (a.children[0] as XElement).dirty = true; // no propagation

    const gaps = checkDirtyInvariant(document);
    expect(gaps).toHaveLength(1);
    expect(gaps[0]!.detail).toContain('discard the edit');
  });

  it('passes, and reports nothing, on a document that is genuinely fine', () => {
    for (const source of DOCUMENTS) {
      expect({ source, differences: checkRoundTrip(parseXmlString(source)) }).toEqual({
        source,
        differences: [],
      });
      expect(checkDirtyInvariant(parseXmlString(source))).toEqual([]);
    }
  });
});
