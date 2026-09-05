import { describe, expect, it } from 'vitest';
import { isXmlError } from '../errors.js';
import {
  boundNamespaces,
  checkMarkupCompatibility,
  effectiveAttributes,
  effectiveChildren,
  ignorableNamespaces,
  isAlternateContent,
  mustUnderstandNamespaces,
  processContentNames,
  selectAlternateContent,
} from './mce.js';
import { NS } from './namespaces.js';
import { qualifiedKey } from '../edit/schema-order.js';
import { descendantElements, firstChild, parseXmlString } from '../parse/xnode.js';
import type { XElement } from '../parse/xnode.js';

const P = 'http://schemas.openxmlformats.org/presentationml/2006/main';
const A14 = 'http://schemas.microsoft.com/office/drawing/2010/main';
const P14 = 'http://schemas.microsoft.com/office/powerpoint/2010/main';
const OLD_P14 = 'http://schemas.microsoft.com/office/powerpoint/2007/7/12/main';

const ROOT = `xmlns:p="${P}" xmlns:mc="${NS.mc}"`;

function parse(source: string): XElement {
  return parseXmlString(source).root;
}

function only(root: XElement, qname: string): XElement {
  for (const element of descendantElements(root)) if (element.qname === qname) return element;
  throw new Error('no <' + qname + '>');
}

function codeOf(fn: () => unknown): string {
  try {
    fn();
  } catch (error) {
    if (isXmlError(error)) return error.code;
    return 'NOT_AN_XML_ERROR';
  }
  return 'DID_NOT_THROW';
}

function qnames(nodes: Iterable<{ type: string }>): string[] {
  return [...nodes].filter((n): n is XElement => n.type === 'element').map((n) => n.qname);
}

describe('mc:AlternateContent is a switch over namespaces', () => {
  const INK =
    `<p:sld ${ROOT}><mc:AlternateContent>` +
    `<mc:Choice xmlns:a14="${A14}" Requires="a14"><a14:ink/></mc:Choice>` +
    '<mc:Fallback><p:pic/></mc:Fallback>' +
    '</mc:AlternateContent></p:sld>';

  it('takes the Choice when its namespace is supported', () => {
    const alternate = only(parse(INK), 'mc:AlternateContent');
    expect(isAlternateContent(alternate)).toBe(true);
    expect(selectAlternateContent(alternate, new Set([A14]))?.qname).toBe('mc:Choice');
  });

  it('falls back when it is not', () => {
    const alternate = only(parse(INK), 'mc:AlternateContent');
    expect(selectAlternateContent(alternate, new Set([P]))?.qname).toBe('mc:Fallback');
  });

  it('selects nothing at all when there is no Fallback, which is legal', () => {
    const source =
      `<p:sld ${ROOT}><mc:AlternateContent>` +
      `<mc:Choice xmlns:a14="${A14}" Requires="a14"><a14:ink/></mc:Choice>` +
      '</mc:AlternateContent></p:sld>';
    const alternate = only(parse(source), 'mc:AlternateContent');
    expect(selectAlternateContent(alternate, new Set())).toBeUndefined();
    expect(qnames(effectiveChildren(parse(source), new Set()))).toEqual([]);
  });

  it('takes the first Choice that fits, not the best one', () => {
    const source =
      `<p:sld ${ROOT} xmlns:x="urn:x" xmlns:y="urn:y">` +
      '<mc:AlternateContent>' +
      '<mc:Choice Requires="x"><p:one/></mc:Choice>' +
      '<mc:Choice Requires="y"><p:two/></mc:Choice>' +
      '</mc:AlternateContent></p:sld>';
    const alternate = only(parse(source), 'mc:AlternateContent');
    const both = selectAlternateContent(alternate, new Set(['urn:x', 'urn:y']))!;
    expect(firstChild(both, 'p:one')).toBeDefined();
  });

  it('requires every namespace a Choice names, not any of them', () => {
    const source =
      `<p:sld ${ROOT} xmlns:x="urn:x" xmlns:y="urn:y">` +
      '<mc:AlternateContent>' +
      '<mc:Choice Requires="x  y"><p:both/></mc:Choice>' +
      '<mc:Fallback><p:neither/></mc:Fallback>' +
      '</mc:AlternateContent></p:sld>';
    const alternate = only(parse(source), 'mc:AlternateContent');
    expect(
      firstChild(selectAlternateContent(alternate, new Set(['urn:x']))!, 'p:neither'),
    ).toBeDefined();
    expect(
      firstChild(selectAlternateContent(alternate, new Set(['urn:x', 'urn:y']))!, 'p:both'),
    ).toBeDefined();
  });

  it('reads Requires as an unprefixed attribute, because that is how it is written', () => {
    // `mc:Requires` is not a thing. A walker that looked for one would find no
    // requirement anywhere and take the first branch of every switch.
    const source =
      `<p:sld ${ROOT} xmlns:x="urn:x">` +
      '<mc:AlternateContent>' +
      '<mc:Choice Requires="x"><p:modern/></mc:Choice>' +
      '<mc:Fallback><p:old/></mc:Fallback>' +
      '</mc:AlternateContent></p:sld>';
    const alternate = only(parse(source), 'mc:AlternateContent');
    expect(firstChild(selectAlternateContent(alternate, new Set())!, 'p:old')).toBeDefined();
  });
});

describe('a prefix means what it means where it was written', () => {
  /**
   * The corpus fact this whole package is built around: `p14` is bound to
   * `…/2010/main` in 46 parts and to `…/2007/7/12/main` in 8 others. Here both
   * bindings live in one document, in sibling subtrees, and the same
   * `Requires="p14"` selects differently in each.
   */
  const TWO_BINDINGS =
    `<p:sld ${ROOT}>` +
    `<p:one xmlns:p14="${P14}"><mc:AlternateContent>` +
    '<mc:Choice Requires="p14"><p:modern/></mc:Choice>' +
    '<mc:Fallback><p:old/></mc:Fallback></mc:AlternateContent></p:one>' +
    `<p:two xmlns:p14="${OLD_P14}"><mc:AlternateContent>` +
    '<mc:Choice Requires="p14"><p:modern/></mc:Choice>' +
    '<mc:Fallback><p:old/></mc:Fallback></mc:AlternateContent></p:two>' +
    '</p:sld>';

  it('resolves the same Requires to two different answers in one part', () => {
    const root = parse(TWO_BINDINGS);
    const supported = new Set([P14]);
    const first = only(firstChild(root, 'p:one')!, 'mc:AlternateContent');
    const second = only(firstChild(root, 'p:two')!, 'mc:AlternateContent');
    expect(firstChild(selectAlternateContent(first, supported)!, 'p:modern')).toBeDefined();
    expect(firstChild(selectAlternateContent(second, supported)!, 'p:old')).toBeDefined();
  });

  it('finds both bindings when asked what the document needs', () => {
    expect(boundNamespaces(parse(TWO_BINDINGS)).has(P14)).toBe(true);
    expect(boundNamespaces(parse(TWO_BINDINGS)).has(OLD_P14)).toBe(true);
  });
});

describe('a prefix that resolves to nothing is loud', () => {
  it('throws rather than treating the Choice as unsupported', () => {
    const source =
      `<p:sld ${ROOT}><mc:AlternateContent>` +
      '<mc:Choice Requires="gone"><p:a/></mc:Choice>' +
      '<mc:Fallback><p:b/></mc:Fallback></mc:AlternateContent></p:sld>';
    const alternate = only(parse(source), 'mc:AlternateContent');
    expect(codeOf(() => selectAlternateContent(alternate, new Set()))).toBe('ERR_INVALID_MCE');
  });

  it('rejects a malformed switch rather than picking something', () => {
    const cases: Record<string, string> = {
      'no Choice at all': '<mc:AlternateContent><mc:Fallback/></mc:AlternateContent>',
      'a Choice after the Fallback':
        '<mc:AlternateContent><mc:Fallback/><mc:Choice Requires="p"/></mc:AlternateContent>',
      'two Fallbacks':
        '<mc:AlternateContent><mc:Choice Requires="p"/><mc:Fallback/><mc:Fallback/></mc:AlternateContent>',
      'a foreign child':
        '<mc:AlternateContent><mc:Choice Requires="p"/><p:sp/></mc:AlternateContent>',
      'a Choice with no Requires':
        '<mc:AlternateContent><mc:Choice><p:a/></mc:Choice></mc:AlternateContent>',
      'a Choice requiring nothing':
        '<mc:AlternateContent><mc:Choice Requires="  "><p:a/></mc:Choice></mc:AlternateContent>',
    };
    for (const [what, body] of Object.entries(cases)) {
      const alternate = only(parse(`<p:sld ${ROOT}>${body}</p:sld>`), 'mc:AlternateContent');
      expect(
        codeOf(() => selectAlternateContent(alternate, new Set([P]))),
        what,
      ).toBe('ERR_INVALID_MCE');
    }
  });
});

describe('the compatibility-rule attributes', () => {
  it('accumulates mc:Ignorable down the tree and never unsets it', () => {
    const source =
      `<p:sld ${ROOT} xmlns:x="urn:x" mc:Ignorable="x">` +
      '<p:inner xmlns:y="urn:y" mc:Ignorable="y"><p:leaf/></p:inner></p:sld>';
    const root = parse(source);
    expect([...ignorableNamespaces(root)]).toEqual(['urn:x']);
    expect([...ignorableNamespaces(only(root, 'p:leaf'))].sort()).toEqual(['urn:x', 'urn:y']);
  });

  it('reads mc:ProcessContent as qualified names, not prefixes', () => {
    const source = `<p:sld ${ROOT} xmlns:x="urn:x" mc:Ignorable="x" mc:ProcessContent="x:wrap x:other"/>`;
    expect([...processContentNames(parse(source))].sort()).toEqual(
      [qualifiedKey('urn:x', 'other'), qualifiedKey('urn:x', 'wrap')].sort(),
    );
  });

  it('rejects an mc:ProcessContent entry that is not qualified', () => {
    const source = `<p:sld ${ROOT} xmlns:x="urn:x" mc:Ignorable="x" mc:ProcessContent="wrap"/>`;
    expect(codeOf(() => processContentNames(parse(source)))).toBe('ERR_INVALID_MCE');
  });

  it('ignores a bare Ignorable, which is in no namespace at all', () => {
    const source = `<p:sld ${ROOT} xmlns:x="urn:x" Ignorable="x"><x:thing/></p:sld>`;
    expect(ignorableNamespaces(parse(source)).size).toBe(0);
  });

  it('reports mc:MustUnderstand naming something we do not know', () => {
    const source = `<p:sld ${ROOT} xmlns:x="urn:x" mc:MustUnderstand="x"/>`;
    expect([...mustUnderstandNamespaces(parse(source))]).toEqual(['urn:x']);
    expect(checkMarkupCompatibility(parse(source), new Set([P]))).toHaveLength(1);
    expect(checkMarkupCompatibility(parse(source), new Set([P, 'urn:x']))).toHaveLength(0);
  });
});

describe('what a consumer actually sees', () => {
  it('drops an ignorable element it does not understand, and keeps it when it does', () => {
    const source = `<p:sld ${ROOT} xmlns:x="urn:x" mc:Ignorable="x"><p:a/><x:extra/><p:b/></p:sld>`;
    expect(qnames(effectiveChildren(parse(source), new Set([P])))).toEqual(['p:a', 'p:b']);
    expect(qnames(effectiveChildren(parse(source), new Set([P, 'urn:x'])))).toEqual([
      'p:a',
      'x:extra',
      'p:b',
    ]);
  });

  it('processes the content of an ignored element when mc:ProcessContent says to', () => {
    const source =
      `<p:sld ${ROOT} xmlns:x="urn:x" mc:Ignorable="x" mc:ProcessContent="x:wrap">` +
      '<x:wrap><p:kept/></x:wrap></p:sld>';
    expect(qnames(effectiveChildren(parse(source), new Set([P])))).toEqual(['p:kept']);
  });

  it('keeps an element nobody declared ignorable, rather than losing content', () => {
    // MCE says the producer should have annotated it. Deleting what a producer
    // forgot is the worse error for an editor whose thesis is preservation.
    const source = `<p:sld ${ROOT} xmlns:x="urn:x"><x:unannounced/></p:sld>`;
    expect(qnames(effectiveChildren(parse(source), new Set([P])))).toEqual(['x:unannounced']);
  });

  it('flattens a selected branch into the parent, nested switches included', () => {
    const source =
      `<p:sld ${ROOT} xmlns:x="urn:x">` +
      '<p:before/>' +
      '<mc:AlternateContent><mc:Choice Requires="x"><p:unreachable/></mc:Choice><mc:Fallback>' +
      '<p:one/>' +
      '<mc:AlternateContent><mc:Choice Requires="x"><p:skipped/></mc:Choice>' +
      '<mc:Fallback><p:two/></mc:Fallback></mc:AlternateContent>' +
      '</mc:Fallback></mc:AlternateContent>' +
      '<p:after/></p:sld>';
    expect(qnames(effectiveChildren(parse(source), new Set([P])))).toEqual([
      'p:before',
      'p:one',
      'p:two',
      'p:after',
    ]);
  });

  it('hides the MCE attributes and the ignorable ones, and keeps the rest', () => {
    const source = `<p:sld ${ROOT} xmlns:x="urn:x" mc:Ignorable="x" name="Slide 1" x:hint="drop" p:keep="yes"/>`;
    expect([...effectiveAttributes(parse(source), new Set([P]))].map((a) => a.qname)).toEqual([
      'name',
      'p:keep',
    ]);
  });

  it('never changes the tree it walked', () => {
    const source =
      `<p:sld ${ROOT} xmlns:x="urn:x" mc:Ignorable="x"><mc:AlternateContent>` +
      '<mc:Choice Requires="x"><x:a/></mc:Choice><mc:Fallback><p:b/></mc:Fallback>' +
      '</mc:AlternateContent></p:sld>';
    const document = parseXmlString(source);
    void [...effectiveChildren(document.root, new Set([P]))];
    void checkMarkupCompatibility(document.root, new Set([P]));
    expect(document.root.dirty).toBe(false);
    for (const element of descendantElements(document.root)) expect(element.dirty).toBe(false);
  });
});

describe('checkMarkupCompatibility reports rather than throwing', () => {
  it('collects every broken construct in the subtree', () => {
    const source =
      `<p:sld ${ROOT}>` +
      '<p:a><mc:AlternateContent><mc:Choice Requires="missing"/></mc:AlternateContent></p:a>' +
      '<p:b><mc:AlternateContent><mc:Fallback/></mc:AlternateContent></p:b>' +
      '</p:sld>';
    const problems = checkMarkupCompatibility(parse(source), new Set([P]));
    expect(problems).toHaveLength(2);
    expect(problems.every((p) => p.detail.length > 0)).toBe(true);
  });

  it('says nothing about a document that is simply fine', () => {
    const source =
      `<p:sld ${ROOT} xmlns:x="urn:x" mc:Ignorable="x">` +
      '<mc:AlternateContent><mc:Choice Requires="x"><x:a/></mc:Choice>' +
      '<mc:Fallback><p:b/></mc:Fallback></mc:AlternateContent></p:sld>';
    expect(checkMarkupCompatibility(parse(source), new Set([P]))).toEqual([]);
  });
});
