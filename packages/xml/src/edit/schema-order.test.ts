import { describe, expect, it } from 'vitest';
import { applyEdit, insertInOrder, newAttribute, newElement } from './edit.js';
import { isXmlError } from '../errors.js';
import {
  childRank,
  childRanks,
  elementKey,
  insertionIndex,
  outOfOrderChildren,
  qualifiedKey,
  SCHEMA_SOURCES,
} from './schema-order.js';
import { childElements, descendantElements, firstChild, parseXmlString } from '../parse/xnode.js';
import type { XElement } from '../parse/xnode.js';

const P = 'http://schemas.openxmlformats.org/presentationml/2006/main';
const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';

const DECLARATIONS = `xmlns:a="${A}" xmlns:p="${P}"`;

const SLIDE =
  `<p:sld ${DECLARATIONS}>` +
  '<p:cSld><p:spTree>' +
  '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
  '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>' +
  '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>' +
  '<p:sp><p:nvSpPr><p:cNvPr id="2" name="Title 1"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>' +
  '<p:spPr><a:xfrm><a:off x="838200" y="365125"/><a:ext cx="10515600" cy="1325563"/></a:xfrm></p:spPr>' +
  '<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>Hello</a:t></a:r>' +
  '<a:endParaRPr lang="en-US"/></a:p></p:txBody>' +
  '</p:sp>' +
  '</p:spTree></p:cSld>' +
  '<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>' +
  '</p:sld>';

function find(source: string, qname: string): XElement {
  const document = parseXmlString(source);
  for (const element of descendantElements(document.root)) {
    if (element.qname === qname) return element;
  }
  throw new Error('no <' + qname + '> in the fixture');
}

function names(element: XElement): string[] {
  return childElements(element).map((child) => child.qname);
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

describe('the table came from the schemas and says what the schemas say', () => {
  it('records the SHA-256 of every input it was generated from', () => {
    expect(SCHEMA_SOURCES.map((s) => s.file)).toEqual([
      'pml.xsd',
      'dml-main.xsd',
      'dml-picture.xsd',
      'dml-lockedCanvas.xsd',
    ]);
    for (const source of SCHEMA_SOURCES) expect(source.sha256).toMatch(/^[0-9a-f]{64}$/);
  });

  it('orders p:sp the way ECMA-376 does, with extLst last', () => {
    const ranks = childRanks(find(SLIDE, 'p:sp'))!;
    const order = [...ranks].sort((a, b) => a[1] - b[1]).map(([key]) => key);
    expect(order).toEqual(
      ['nvSpPr', 'spPr', 'style', 'txBody', 'extLst'].map((n) => qualifiedKey(P, n)),
    );
  });

  it('gives every shape kind in a spTree the same rank, because that rank is z-order', () => {
    const ranks = childRanks(find(SLIDE, 'p:spTree'))!;
    const shapes = ['sp', 'grpSp', 'graphicFrame', 'cxnSp', 'pic', 'contentPart'].map((n) =>
      ranks.get(qualifiedKey(P, n)),
    );
    expect(new Set(shapes).size).toBe(1);
    expect(shapes[0]).toBeGreaterThan(ranks.get(qualifiedKey(P, 'grpSpPr'))!);
    expect(shapes[0]).toBeLessThan(ranks.get(qualifiedKey(P, 'extLst'))!);
  });

  it('puts a:endParaRPr after every run in a:p, and a:pPr before them', () => {
    const ranks = childRanks(find(SLIDE, 'a:p'))!;
    expect(ranks.get(qualifiedKey(A, 'pPr'))).toBeLessThan(ranks.get(qualifiedKey(A, 'r'))!);
    // A repeating xsd:group, so a:r, a:br and a:fld interleave freely.
    expect(ranks.get(qualifiedKey(A, 'br'))).toBe(ranks.get(qualifiedKey(A, 'r')));
    expect(ranks.get(qualifiedKey(A, 'fld'))).toBe(ranks.get(qualifiedKey(A, 'r')));
    expect(ranks.get(qualifiedKey(A, 'endParaRPr'))).toBeGreaterThan(
      ranks.get(qualifiedKey(A, 'r'))!,
    );
  });

  it('orders a:spPr geometry before fill before line before effects', () => {
    const ranks = childRanks(find(SLIDE, 'p:spPr'))!;
    const rank = (local: string): number => ranks.get(qualifiedKey(A, local))!;
    expect(rank('xfrm')).toBeLessThan(rank('prstGeom'));
    expect(rank('prstGeom')).toBe(rank('custGeom'));
    expect(rank('prstGeom')).toBeLessThan(rank('solidFill'));
    expect(rank('solidFill')).toBe(rank('gradFill'));
    expect(rank('solidFill')).toBeLessThan(rank('ln'));
    expect(rank('ln')).toBeLessThan(rank('effectLst'));
    expect(rank('effectLst')).toBeLessThan(rank('extLst'));
  });
});

describe('keys are namespaces, and prefixes are only spelling', () => {
  it('answers the same for a document that spells PresentationML differently', () => {
    const renamed = SLIDE.replace(/\bp:/g, 'pres:').replace('xmlns:p=', 'xmlns:pres=');
    const ranks = childRanks(find(renamed, 'pres:sp'))!;
    expect([...ranks].sort((a, b) => a[1] - b[1]).map(([key]) => key)).toEqual(
      ['nvSpPr', 'spPr', 'style', 'txBody', 'extLst'].map((n) => qualifiedKey(P, n)),
    );
  });

  it('refuses an element whose prefix nothing binds', () => {
    // Well-formed XML; `undeclaredPrefixes` reports it, and nothing can rank it.
    const orphan = find('<x:root xmlns:x="urn:x"><p:child/></x:root>', 'p:child');
    expect(elementKey(orphan)).toBeUndefined();
    expect(childRanks(orphan)).toBeUndefined();
  });
});

describe('the three names the parent alone cannot place', () => {
  it('gives a:xfrm two children under p:spPr and four under p:grpSpPr', () => {
    const inShape = childRanks(find(SLIDE, 'p:spPr').children[0] as XElement)!;
    expect([...inShape.keys()].sort()).toEqual([qualifiedKey(A, 'ext'), qualifiedKey(A, 'off')]);

    const inGroup = childRanks(find(SLIDE, 'p:grpSpPr').children[0] as XElement)!;
    expect([...inGroup.keys()].sort()).toEqual(
      [
        qualifiedKey(A, 'chExt'),
        qualifiedKey(A, 'chOff'),
        qualifiedKey(A, 'ext'),
        qualifiedKey(A, 'off'),
      ].sort(),
    );
    expect(inGroup.get(qualifiedKey(A, 'chOff'))).toBeGreaterThan(
      inGroup.get(qualifiedKey(A, 'ext'))!,
    );
  });

  it('separates a:path in a gradient from a:path in a geometry', () => {
    const source =
      `<p:sld ${DECLARATIONS}><a:gradFill><a:path path="circle"/></a:gradFill>` +
      '<a:custGeom><a:pathLst><a:path w="1" h="1"/></a:pathLst></a:custGeom></p:sld>';
    const document = parseXmlString(source);
    const gradient = firstChild(firstChild(document.root, 'a:gradFill')!, 'a:path')!;
    const geometry = firstChild(
      firstChild(firstChild(document.root, 'a:custGeom')!, 'a:pathLst')!,
      'a:path',
    )!;
    expect([...childRanks(gradient)!.keys()]).toEqual([qualifiedKey(A, 'fillToRect')]);
    expect(childRanks(geometry)!.has(qualifiedKey(A, 'moveTo'))).toBe(true);
  });

  it('refuses to guess when the grandparent is missing', () => {
    // A detached a:xfrm has no grandparent, so neither content model applies.
    const document = parseXmlString(`<a:xfrm xmlns:a="${A}"/>`);
    expect(childRanks(document.root)).toBeUndefined();
  });
});

describe('opaque content is refused, not guessed at', () => {
  it('has no ranks for an extLst entry, whose schema is xsd:any', () => {
    const source = `<p:sld ${DECLARATIONS}><p:extLst><p:ext uri="{X}"/></p:extLst></p:sld>`;
    const ext = firstChild(firstChild(parseXmlString(source).root, 'p:extLst')!, 'p:ext')!;
    expect(childRanks(ext)).toBeUndefined();
  });

  it('has no ranks for a:graphicData, where a chart or a diagram lives', () => {
    const source = `<p:sld ${DECLARATIONS}><a:graphic><a:graphicData uri="{X}"/></a:graphic></p:sld>`;
    const data = firstChild(
      firstChild(parseXmlString(source).root, 'a:graphic')!,
      'a:graphicData',
    )!;
    expect(childRanks(data)).toBeUndefined();
  });

  it('has no ranks for an element in an extension namespace', () => {
    const source = `<p:sld ${DECLARATIONS} xmlns:p14="urn:p14"><p14:thing/></p:sld>`;
    expect(childRanks(firstChild(parseXmlString(source).root, 'p14:thing')!)).toBeUndefined();
  });
});

describe('inserting a child at every position in a sequence', () => {
  const CHILDREN = ['p:nvSpPr', 'p:spPr', 'p:style', 'p:txBody', 'p:extLst'];

  /**
   * The plan's stated verification for this sub-phase, taken literally: every
   * subset of a real sequence, crossed with every child missing from it. 80
   * insertions, and each one must land so that the result is still a
   * subsequence of the schema's order.
   */
  it('lands schema-correct for all 32 subsets of p:sp', () => {
    let insertions = 0;
    for (let mask = 0; mask < 1 << CHILDREN.length; mask++) {
      const present = CHILDREN.filter((_, i) => (mask & (1 << i)) !== 0);
      for (const missing of CHILDREN.filter((c) => !present.includes(c))) {
        const source =
          `<p:sld ${DECLARATIONS}><p:sp>` +
          present.map((c) => `<${c}/>`).join('') +
          '</p:sp></p:sld>';
        const sp = firstChild(parseXmlString(source).root, 'p:sp')!;
        applyEdit(insertInOrder(sp, newElement(missing)));
        const after = names(sp);
        expect(after).toContain(missing);
        expect(after).toEqual(CHILDREN.filter((c) => after.includes(c)));
        insertions++;
      }
    }
    expect(insertions).toBe(80);
  });

  it('puts a new shape on top of the z-order, before extLst', () => {
    const source =
      `<p:sld ${DECLARATIONS}><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>` +
      '<p:sp/><p:pic/><p:extLst/></p:spTree></p:sld>';
    const tree = firstChild(parseXmlString(source).root, 'p:spTree')!;
    applyEdit(insertInOrder(tree, newElement('p:sp')));
    expect(names(tree)).toEqual(['p:nvGrpSpPr', 'p:grpSpPr', 'p:sp', 'p:pic', 'p:sp', 'p:extLst']);
  });
});

describe('nothing already in the document ever moves', () => {
  it('steps over an mc:AlternateContent instead of ranking it', () => {
    const source =
      `<p:sld ${DECLARATIONS} xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">` +
      '<p:sp><p:nvSpPr/><mc:AlternateContent/><p:txBody/></p:sp></p:sld>';
    const sp = firstChild(parseXmlString(source).root, 'p:sp')!;
    applyEdit(insertInOrder(sp, newElement('p:spPr')));
    // spPr lands after nvSpPr, and the AlternateContent has not moved: it is
    // still the child that follows it.
    expect(names(sp)).toEqual(['p:nvSpPr', 'p:spPr', 'mc:AlternateContent', 'p:txBody']);
  });

  it('leaves an out-of-order document out of order rather than sorting it', () => {
    const source = `<p:sld ${DECLARATIONS}><p:sp><p:txBody/><p:nvSpPr/></p:sp></p:sld>`;
    const sp = firstChild(parseXmlString(source).root, 'p:sp')!;
    expect(outOfOrderChildren(sp)).toHaveLength(1);
    applyEdit(insertInOrder(sp, newElement('p:spPr')));
    expect(names(sp)).toEqual(['p:txBody', 'p:nvSpPr', 'p:spPr']);
  });

  it('skips text and comments when choosing a position', () => {
    const source = `<p:sld ${DECLARATIONS}>\n  <p:sp>\n    <p:nvSpPr/>\n    <!-- note -->\n    <p:txBody/>\n  </p:sp>\n</p:sld>`;
    const sp = firstChild(parseXmlString(source).root, 'p:sp')!;
    const index = insertionIndex(sp, qualifiedKey(P, 'spPr'))!;
    applyEdit({ kind: 'insertChild', parent: sp, index, node: newElement('p:spPr') });
    expect(names(sp)).toEqual(['p:nvSpPr', 'p:spPr', 'p:txBody']);
  });
});

describe('refusing is part of the contract', () => {
  it('refuses a child the parent does not admit', () => {
    const sp = find(SLIDE, 'p:sp');
    expect(childRank(sp, qualifiedKey(P, 'sldSz'))).toBeUndefined();
    expect(codeOf(() => insertInOrder(sp, newElement('p:sldSz')))).toBe('ERR_SCHEMA_ORDER');
  });

  it('refuses an insertion into opaque content', () => {
    const source = `<p:sld ${DECLARATIONS}><p:extLst><p:ext uri="{X}"/></p:extLst></p:sld>`;
    const ext = firstChild(firstChild(parseXmlString(source).root, 'p:extLst')!, 'p:ext')!;
    expect(codeOf(() => insertInOrder(ext, newElement('p:sp')))).toBe('ERR_SCHEMA_ORDER');
  });

  it('refuses a child whose own prefix is bound to nothing here', () => {
    const sp = find(SLIDE, 'p:sp');
    expect(codeOf(() => insertInOrder(sp, newElement('nope:spPr')))).toBe('ERR_SCHEMA_ORDER');
  });

  it('accepts a child that brings its own namespace declaration', () => {
    const source = `<p:sld ${DECLARATIONS}><p:sp><p:nvSpPr/></p:sp></p:sld>`;
    const sp = firstChild(parseXmlString(source).root, 'p:sp')!;
    // `q` is bound to nothing in the document. The element declares it itself,
    // and the binding it is about to bring with it is the one that counts.
    const child = newElement('q:spPr', [newAttribute('xmlns:q', P)]);
    applyEdit(insertInOrder(sp, child));
    expect(names(sp)).toEqual(['p:nvSpPr', 'q:spPr']);
  });
});
