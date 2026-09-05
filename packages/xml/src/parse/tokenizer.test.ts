import { describe, expect, it } from 'vitest';
import { isXmlError } from '../errors.js';
import {
  checkSpanCoverage,
  DEFAULT_TOKENIZER_LIMITS,
  splitQName,
  tokenize,
  XmlTokenizer,
  type XmlElementToken,
  type XmlToken,
} from './tokenizer.js';

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
 * Tokenize, then assert the one property that matters: the spans account for
 * every character. Every test in this file goes through here, so a span bug
 * anywhere fails everywhere rather than hiding until 0.5.
 */
function scan(source: string): XmlToken[] {
  const tokens = tokenize(source);
  expect(checkSpanCoverage(source, tokens)).toEqual([]);
  const bom = source.charCodeAt(0) === 0xfeff ? '\ufeff' : '';
  expect(bom + tokens.map((t) => source.slice(t.start, t.end)).join('')).toBe(source);
  return tokens;
}

const types = (source: string): string[] => scan(source).map((t) => t.type);
const first = (source: string): XmlElementToken =>
  scan(source).find(
    (t): t is XmlElementToken => t.type === 'startTag' || t.type === 'emptyElementTag',
  )!;

// ---------------------------------------------------------------------------
// Everything in this block was built into a real slide part, written back into
// a real template, and opened by the installed PowerPoint 365. All of it opens.
// A tokenizer that handles only what our corpus happens to contain would reject
// most of it, and every one of these is something a third-party writer emits.
// ---------------------------------------------------------------------------
describe('constructs PowerPoint accepts, so we must too', () => {
  it('reads a comment anywhere it is legal', () => {
    expect(types('<?xml version="1.0"?><!--before--><a/><!--after-->')).toEqual([
      'declaration',
      'comment',
      'emptyElementTag',
      'comment',
    ]);
    expect(types('<a><!--inside--></a>')).toEqual(['startTag', 'comment', 'endTag']);
  });

  it('allows a single hyphen inside a comment', () => {
    const [comment] = scan('<a><!-- a - b --></a>').slice(1);
    expect(comment).toMatchObject({ type: 'comment', value: ' a - b ' });
  });

  it('reads a processing instruction', () => {
    const tokens = scan('<?xml version="1.0"?><?probe target="x"?><a/>');
    expect(tokens[1]).toMatchObject({
      type: 'processingInstruction',
      target: 'probe',
      data: ' target="x"',
    });
  });

  it('reads a CDATA section, which PowerPoint reads as text and then destroys', () => {
    // Resaving through PowerPoint turns `<![CDATA[CDATA TEXT]]>` into plain
    // text. We keep the lexical form, which is the difference between us and it.
    const tokens = scan('<a:t><![CDATA[a < b & c]]></a:t>');
    expect(tokens[1]).toMatchObject({ type: 'cdata', value: 'a < b & c' });
  });

  it('reads single-quoted attributes and remembers which quote was used', () => {
    const tag = first('<a:off x=\'1\' y="2"/>');
    expect(tag.attributes.map((a) => [a.qname, a.value, a.quote])).toEqual([
      ['x', '1', "'"],
      ['y', '2', '"'],
    ]);
  });

  it('reads a document with no declaration at all', () => {
    expect(types('<a/>')).toEqual(['emptyElementTag']);
  });

  it('reads a declaration written with single quotes', () => {
    const [decl] = scan("<?xml version='1.0' encoding='UTF-8' standalone='yes'?><a/>");
    expect(decl).toMatchObject({
      type: 'declaration',
      version: '1.0',
      encoding: 'UTF-8',
      standalone: 'yes',
    });
  });

  it('distinguishes an empty element written long-form from one written short', () => {
    // 33 elements in the corpus, all in docProps, use `<x></x>`. PowerPoint's
    // own writer collapses them; we do not.
    expect(types('<a><b></b></a>')).toEqual(['startTag', 'startTag', 'endTag', 'endTag']);
    expect(types('<a><b/></a>')).toEqual(['startTag', 'emptyElementTag', 'endTag']);
  });

  it('records the whitespace before a self-closing slash', () => {
    // 70 822 of the 98 777 self-closing tags in the corpus are written this way.
    const tight = first('<a:off x="0"/>');
    const spaced = first('<a:off x="0"   />');
    expect(tight.trailingSpaceStart).toBe(tight.end - 2);
    expect(spaced.trailingSpaceStart).toBe(spaced.end - 5);
  });

  it('reads a tag broken across lines and spaced around the equals sign', () => {
    const tag = first('<a:off\r\n  x = "1"\r\n  y = "2"\r\n/>');
    expect(tag.attributes.map((a) => a.value)).toEqual(['1', '2']);
  });

  it('reads a pretty-printed document', () => {
    // Two shipped templates in our corpus are fully indented, so this is not a
    // hypothetical producer.
    const tokens = scan('<a>\r\n  <b/>\r\n</a>');
    expect(
      tokens
        .filter((t) => t.type === 'text')
        .every((t) => 'whitespaceOnly' in t && t.whitespaceOnly),
    ).toBe(true);
  });

  it('reads an unescaped ">" in text and in an attribute value', () => {
    // Legal in both positions, and both open in PowerPoint. Only "<" and "&"
    // are forbidden in text; only "<" is forbidden in an attribute value.
    expect(scan('<a:t>a > b</a:t>')[1]).toMatchObject({ value: 'a > b' });
    expect(first('<p:cNvPr name="a > b"/>').attributes[0]!.value).toBe('a > b');
  });

  it('reads an apostrophe inside a double-quoted value, and vice versa', () => {
    expect(first('<a n="it\'s"/>').attributes[0]!.value).toBe("it's");
    expect(first('<a n=\'say "hi"\'/>').attributes[0]!.value).toBe('say "hi"');
  });

  it('starts the token stream after a byte order mark', () => {
    const tokens = scan('\ufeff<?xml version="1.0"?><a/>');
    expect(tokens[0]!.start).toBe(1);
  });

  it('reads whitespace before and after the root element', () => {
    expect(types('<?xml version="1.0"?>\r\n\r\n  <a/>\r\n')).toEqual([
      'declaration',
      'text',
      'emptyElementTag',
      'text',
    ]);
  });

  it('reads CR-only and LF-only line endings', () => {
    expect(types('<?xml version="1.0"?>\r<a/>')).toEqual([
      'declaration',
      'text',
      'emptyElementTag',
    ]);
    expect(types('<?xml version="1.0"?>\n<a/>')).toEqual([
      'declaration',
      'text',
      'emptyElementTag',
    ]);
  });
});

// ---------------------------------------------------------------------------
// And the four that PowerPoint refuses with 0x80070570. Refusing these is not
// us being stricter than the format.
// ---------------------------------------------------------------------------
describe('constructs PowerPoint refuses, so we refuse them too', () => {
  it('refuses a DOCTYPE outright, with its own code', () => {
    // Not `ERR_MALFORMED_XML`: a DTD is a well-formed attack surface, and
    // rejecting the construct removes XXE and entity expansion entirely rather
    // than defending against them.
    expect(codeOf(() => tokenize('<!DOCTYPE a><a/>'))).toBe('ERR_DOCTYPE_FORBIDDEN');
    expect(codeOf(() => tokenize('<!DOCTYPE a [<!ENTITY x "y">]><a>&x;</a>'))).toBe(
      'ERR_DOCTYPE_FORBIDDEN',
    );
  });

  it('refuses a duplicate attribute', () => {
    expect(codeOf(() => tokenize('<a:off x="0" y="0" x="0"/>'))).toBe('ERR_DUPLICATE_ATTRIBUTE');
  });

  it('refuses a character outside the Char production', () => {
    expect(codeOf(() => tokenize('<a:t>a\u000bb</a:t>'))).toBe('ERR_MALFORMED_XML');
  });

  it('refuses an undeclared entity', () => {
    expect(codeOf(() => tokenize('<a:t>a&nbsp;b</a:t>'))).toBe('ERR_UNKNOWN_ENTITY');
  });
});

describe('well-formedness the specification requires', () => {
  it('refuses a literal "]]>" in text', () => {
    // XML 1.0 §2.4: a parser cannot tell it from the end of a CDATA section.
    // PowerPoint round-trips the escaped form `a]]&gt;b` happily.
    expect(codeOf(() => tokenize('<a>x]]>y</a>'))).toBe('ERR_MALFORMED_XML');
    expect(scan('<a>x]]&gt;y</a>')[1]).toMatchObject({ value: 'x]]>y' });
  });

  it('refuses "--" inside a comment, and a comment ending in a hyphen', () => {
    expect(codeOf(() => tokenize('<a><!-- a -- b --></a>'))).toBe('ERR_MALFORMED_XML');
    expect(codeOf(() => tokenize('<a><!-- a ---></a>'))).toBe('ERR_MALFORMED_XML');
  });

  it('refuses "xml" as a processing-instruction target anywhere but the start', () => {
    expect(codeOf(() => tokenize('<a><?xml version="1.0"?></a>'))).toBe('ERR_MALFORMED_XML');
    expect(codeOf(() => tokenize('<?XML x?><a/>'))).toBe('ERR_MALFORMED_XML');
  });

  it('refuses a "<" in an attribute value', () => {
    expect(codeOf(() => tokenize('<a n="x < y"/>'))).toBe('ERR_MALFORMED_XML');
  });

  it('refuses an unquoted attribute value', () => {
    expect(codeOf(() => tokenize('<a n=1/>'))).toBe('ERR_MALFORMED_XML');
  });

  it('refuses attributes that are not separated by whitespace', () => {
    expect(codeOf(() => tokenize('<a n="1"m="2"/>'))).toBe('ERR_MALFORMED_XML');
  });

  it('refuses an end tag with attributes', () => {
    expect(codeOf(() => tokenize('<a></a n="1">'))).toBe('ERR_MALFORMED_XML');
  });

  it('refuses unterminated constructs rather than scanning forever', () => {
    expect(codeOf(() => tokenize('<a'))).toBe('ERR_MALFORMED_XML');
    expect(codeOf(() => tokenize('<a n="1'))).toBe('ERR_MALFORMED_XML');
    expect(codeOf(() => tokenize('<!-- open'))).toBe('ERR_MALFORMED_XML');
    expect(codeOf(() => tokenize('<![CDATA[ open'))).toBe('ERR_MALFORMED_XML');
    expect(codeOf(() => tokenize('<?pi open'))).toBe('ERR_MALFORMED_XML');
  });

  it('refuses a qualified name with two colons or an empty half', () => {
    expect(codeOf(() => tokenize('<a:b:c/>'))).toBe('ERR_MALFORMED_XML');
    expect(codeOf(() => tokenize('<a:/>'))).toBe('ERR_MALFORMED_XML');
  });
});

describe('what a token carries', () => {
  it('splits a qualified name without resolving it', () => {
    // Resolution needs a scope, and a scope needs the tree. The tokenizer's job
    // ends at the colon.
    expect(splitQName('a:off')).toEqual({ prefix: 'a', local: 'off' });
    expect(splitQName('off')).toEqual({ prefix: '', local: 'off' });
  });

  it('gives an attribute every offset needed to rebuild it exactly', () => {
    const source = '<a:off  x="1" />';
    const tag = first(source);
    const attr = tag.attributes[0]!;
    expect(source.slice(attr.start, attr.end)).toBe('  x="1"');
    expect(source.slice(attr.nameStart, attr.nameEnd)).toBe('x');
    expect(source.slice(attr.valueStart, attr.valueEnd)).toBe('1');
    expect(attr.dirty).toBe(false);
  });

  it('keeps the source of a value separate from its normalized form', () => {
    const source = '<a n="x&#9;y"/>';
    const attr = first(source).attributes[0]!;
    expect(source.slice(attr.valueStart, attr.valueEnd)).toBe('x&#9;y');
    expect(attr.value).toBe('x\ty');
  });

  it('marks a whitespace-only text run', () => {
    const tokens = scan('<a>\r\n  <b/>text</a>');
    const texts = tokens.filter((t) => t.type === 'text');
    expect(texts.map((t) => ('whitespaceOnly' in t ? t.whitespaceOnly : null))).toEqual([
      true,
      false,
    ]);
  });
});

describe('the pull interface', () => {
  it('returns undefined at the end and can be iterated', () => {
    const tokenizer = new XmlTokenizer('<a/>');
    expect(tokenizer.next()?.type).toBe('emptyElementTag');
    expect(tokenizer.next()).toBeUndefined();
    expect([...new XmlTokenizer('<a><b/></a>')].map((t) => t.type)).toEqual([
      'startTag',
      'emptyElementTag',
      'endTag',
    ]);
  });

  it('reports where a byte order mark leaves the cursor', () => {
    expect(new XmlTokenizer('\ufeff<a/>').contentStart).toBe(1);
    expect(new XmlTokenizer('<a/>').contentStart).toBe(0);
  });
});

describe('limits', () => {
  it('refuses an element with more attributes than allowed', () => {
    const wide = '<a ' + Array.from({ length: 40 }, (_, i) => `a${i}="1"`).join(' ') + '/>';
    expect(tokenize(wide)).toHaveLength(1);
    expect(codeOf(() => tokenize(wide, { ...DEFAULT_TOKENIZER_LIMITS, maxAttributes: 8 }))).toBe(
      'ERR_LIMIT_EXCEEDED',
    );
  });

  it('refuses a document with more tokens than allowed', () => {
    const long = '<a>' + '<b/>'.repeat(100) + '</a>';
    expect(codeOf(() => tokenize(long, { ...DEFAULT_TOKENIZER_LIMITS, maxTokens: 10 }))).toBe(
      'ERR_LIMIT_EXCEEDED',
    );
  });
});

describe('coverage checking itself', () => {
  it('reports a gap when a span is wrong', () => {
    const source = '<a/><b/>';
    const tokens = tokenize(source);
    const broken = [tokens[0]!, { ...tokens[1]!, start: tokens[1]!.start + 1 }];
    expect(checkSpanCoverage(source, broken)).toEqual([{ kind: 'gap', at: 4, text: '<' }]);
  });

  it('reports an overlap', () => {
    const source = '<a/><b/>';
    const tokens = tokenize(source);
    const broken = [tokens[0]!, { ...tokens[1]!, start: tokens[1]!.start - 1 }];
    expect(checkSpanCoverage(source, broken)[0]).toMatchObject({ kind: 'overlap' });
  });

  it('reports a short read at the end', () => {
    const source = '<a/>  ';
    expect(checkSpanCoverage(source, tokenize(source).slice(0, 1))[0]).toMatchObject({
      kind: 'epilogue',
    });
  });
});
