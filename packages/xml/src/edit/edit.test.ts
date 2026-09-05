import { describe, expect, it } from 'vitest';
import {
  applyEdit,
  applyEdits,
  insertInOrder,
  newAttribute,
  newElement,
  newText,
  type XmlEdit,
} from './edit.js';
import { isXmlError } from '../errors.js';
import {
  extensions,
  findExtension,
  planAddExtension,
  planRemoveExtension,
} from '../mce/ext-lst.js';
import { checkRoundTrip } from '../emit/roundtrip.js';
import { serializeXmlString } from '../emit/serialize.js';
import {
  checkDirtyInvariant,
  childElements,
  descendantElements,
  firstChild,
  parseXmlString,
  type XDocument,
  type XElement,
  type XText,
} from '../parse/xnode.js';

const P = 'http://schemas.openxmlformats.org/presentationml/2006/main';
const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';

/**
 * Deliberately awkward, in every way sub-phase 0.5 measured to matter.
 *
 * A CRLF after the declaration - 656 of 2834 corpus parts have one, and it is
 * the single thing a rebuild cannot recover. A space before `/>`, which 70 822
 * of 98 777 self-closing tags in the corpus have. A single-quoted attribute. An
 * entity reference in a value that a rebuild would re-spell. Odd spacing
 * between attributes. An element written long-form that is empty.
 *
 * If undo restores this byte for byte, it restores anything.
 */
const SLIDE =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' +
  `<p:sld xmlns:a="${A}" xmlns:p="${P}">` +
  '<p:cSld><p:spTree>' +
  '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
  '<p:grpSpPr/>' +
  '<p:sp>' +
  '<p:nvSpPr><p:cNvPr  id=\'2\'   name="Title &amp; Text"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>' +
  '<p:spPr><a:xfrm><a:off x="838200" y="365125" /><a:ext cx="10515600" cy="1325563"/></a:xfrm></p:spPr>' +
  '<p:txBody><a:bodyPr></a:bodyPr><a:lstStyle/>' +
  '<a:p><a:r><a:rPr lang="en-US"/><a:t>Hello &#62; world</a:t></a:r>' +
  '<a:endParaRPr lang="en-US"/></a:p></p:txBody>' +
  '</p:sp>' +
  '</p:spTree></p:cSld>' +
  '</p:sld>';

/** The first `qname`, or the first one beneath the first `within`. */
function find(document: XDocument, qname: string, nested?: string): XElement {
  const from = nested === undefined ? document.root : find(document, qname);
  for (const element of descendantElements(from)) {
    if (element.qname === (nested ?? qname)) return element;
  }
  throw new Error('no <' + (nested ?? qname) + '>');
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

interface Trip {
  readonly document: XDocument;
  /** The serialization while the edits are applied. */
  readonly after: string;
  /** The serialization once they have been undone. */
  readonly restored: string;
}

/** Apply, serialize, undo, serialize. The shape of every test in this file. */
function trip(source: string, plan: (document: XDocument) => XmlEdit[]): Trip {
  const document = parseXmlString(source);
  const inverses = applyEdits(plan(document));
  const after = serializeXmlString(document);
  expect(checkRoundTrip(document), 'the edited document must still read back as itself').toEqual(
    [],
  );
  applyEdits(inverses);
  return { document, after, restored: serializeXmlString(document) };
}

/** The gate: undo restores the bytes, and leaves no flag set behind it. */
function expectExactUndo(trip_: Trip, source: string): void {
  expect(trip_.restored).toBe(source);
  // Bytes alone are not enough, and mutation testing is how that was found out.
  // An inverse that restores the flags but not the values serializes correctly -
  // the flags send every node back to its source slice, which hides the wrong
  // value behind the right bytes - and leaves a tree that reads back as
  // something the file does not say. `checkRoundTrip` compares the tree against
  // its own serialization, so it sees exactly that.
  expect(checkRoundTrip(trip_.document), 'the restored tree agrees with its own bytes').toEqual([]);
  expect(trip_.document.root.dirty, 'the root is clean again, so it re-emits as a slice').toBe(
    false,
  );
  expect(checkDirtyInvariant(trip_.document)).toEqual([]);
  for (const element of descendantElements(trip_.document.root)) {
    expect(element.dirty, '<' + element.qname + '> is clean again').toBe(false);
    for (const attribute of element.attributes) expect(attribute.dirty).toBe(false);
  }
}

describe("every edit's inverse restores byte identity", () => {
  it('changing an attribute value', () => {
    const t = trip(SLIDE, (d) => [
      { kind: 'setAttribute', element: find(d, 'a:off'), qname: 'x', value: '999' },
    ]);
    expect(t.after).toContain('<a:off x="999" y="365125" />');
    expectExactUndo(t, SLIDE);
  });

  it('changing an attribute whose neighbour is spelled with an entity reference', () => {
    // The element rebuilds its start tag, so `name` takes the slice path and
    // keeps `&amp;`; the single quotes on `id` are the attribute's own and
    // survive too. Undo then restores the whole tag from source.
    const t = trip(SLIDE, (d) => [
      { kind: 'setAttribute', element: find(d, 'p:sp', 'p:cNvPr'), qname: 'id', value: '7' },
    ]);
    // The rebuilt tag keeps the two spaces before `id` and its single quotes,
    // both recovered from the source, and `name` takes the slice path entirely.
    expect(t.after).toContain('<p:cNvPr  id=\'7\'   name="Title &amp; Text"/>');
    expectExactUndo(t, SLIDE);
  });

  it('adding an attribute that was not there', () => {
    const t = trip(SLIDE, (d) => [
      { kind: 'setAttribute', element: find(d, 'a:bodyPr'), qname: 'rot', value: '5400000' },
    ]);
    expect(t.after).toContain('<a:bodyPr rot="5400000"></a:bodyPr>');
    expectExactUndo(t, SLIDE);
  });

  it('removing an attribute, spacing and quote character included', () => {
    const t = trip(SLIDE, (d) => [
      { kind: 'removeAttribute', element: find(d, 'a:off'), qname: 'y' },
    ]);
    expect(t.after).toContain('<a:off x="838200" />');
    expectExactUndo(t, SLIDE);
  });

  it('inserting an element in schema order', () => {
    const t = trip(SLIDE, (d) => [insertInOrder(find(d, 'p:spPr'), newElement('a:prstGeom'))]);
    expect(t.after).toContain('</a:xfrm><a:prstGeom/></p:spPr>');
    expectExactUndo(t, SLIDE);
  });

  it('removing an element, and every byte inside it', () => {
    const t = trip(SLIDE, (d) => {
      const body = find(d, 'p:txBody');
      return [{ kind: 'removeChild', parent: body.parent!, node: body }];
    });
    expect(t.after).not.toContain('a:t>');
    expectExactUndo(t, SLIDE);
  });

  it('replacing the text of a run', () => {
    const t = trip(SLIDE, (d) => {
      const text = find(d, 'a:t').children[0] as XText;
      return [{ kind: 'setValue', node: text, value: 'Goodbye' }];
    });
    expect(t.after).toContain('<a:t>Goodbye</a:t>');
    expectExactUndo(t, SLIDE);
  });

  it('a whole gesture of six edits, undone in one go', () => {
    const t = trip(SLIDE, (d) => {
      const off = find(d, 'a:off');
      const ext = find(d, 'a:ext');
      const spPr = find(d, 'p:spPr');
      const body = find(d, 'a:bodyPr');
      return [
        { kind: 'setAttribute', element: off, qname: 'x', value: '1' },
        { kind: 'setAttribute', element: off, qname: 'y', value: '2' },
        { kind: 'setAttribute', element: ext, qname: 'cx', value: '3' },
        { kind: 'setAttribute', element: ext, qname: 'cy', value: '4' },
        insertInOrder(spPr, newElement('a:prstGeom', [newAttribute('prst', 'rect')])),
        { kind: 'setAttribute', element: body, qname: 'anchor', value: 'ctr' },
      ];
    });
    expect(t.after).toContain('<a:off x="1" y="2" />');
    expect(t.after).toContain('<a:prstGeom prst="rect"/>');
    expectExactUndo(t, SLIDE);
  });

  it('two edits to the same element, of which only the second is undone', () => {
    // The second edit finds its element already dirty, so it journals nothing
    // for it. If it journalled the element anyway, undoing the second would
    // mark the element clean - and a clean element re-emits its whole start tag
    // from source, silently discarding the first edit with no error anywhere.
    const document = parseXmlString(SLIDE);
    const off = find(document, 'a:off');
    applyEdit({ kind: 'setAttribute', element: off, qname: 'x', value: '111' });
    const second = applyEdit({ kind: 'setAttribute', element: off, qname: 'y', value: '222' });
    applyEdit(second);
    expect(off.dirty).toBe(true);
    expect(serializeXmlString(document)).toContain('<a:off x="111" y="365125" />');
    expect(checkRoundTrip(document)).toEqual([]);
  });

  it('an edit inside an edited subtree, undone in the reverse order', () => {
    // The second edit's ancestors are already dirty, so it journals almost
    // nothing. Undoing it must therefore leave the first edit's dirt alone.
    const document = parseXmlString(SLIDE);
    const off = find(document, 'a:off');
    const first = applyEdit({ kind: 'setAttribute', element: off, qname: 'x', value: '1' });
    const second = applyEdit({
      kind: 'setAttribute',
      element: find(document, 'a:ext'),
      qname: 'cx',
      value: '2',
    });
    applyEdit(second);
    expect(off.dirty, 'the first edit still stands').toBe(true);
    expect(serializeXmlString(document)).toContain('<a:off x="1" y="365125" />');
    applyEdit(first);
    expect(serializeXmlString(document)).toBe(SLIDE);
    expect(document.root.dirty).toBe(false);
  });
});

describe('an inverse is always appliable, guards included', () => {
  // No OOXML type has mixed content, but a mangled part can, and the 40 000-case
  // fuzz sweep found this: take an element out of a parent that already held
  // text, and putting it back looks like *creating* mixed content, so the guard
  // refused the undo and the document could not be returned to what it was. A
  // guard that can refuse an inverse is worse than no guard.
  const MIXED = '<r xmlns="urn:x">text<c/>more</r>';

  it('puts back an element it took out of content that already held text', () => {
    const t = trip(MIXED, (d) => [
      { kind: 'removeChild', parent: d.root, node: firstChild(d.root, 'c')! },
    ]);
    expect(t.after).toBe('<r xmlns="urn:x">textmore</r>');
    expectExactUndo(t, MIXED);
  });

  it('still refuses a forward insertion into the same place', () => {
    const document = parseXmlString(MIXED);
    expect(
      codeOf(() =>
        applyEdit({
          kind: 'insertChild',
          parent: document.root,
          index: 0,
          node: newElement('d'),
        }),
      ),
    ).toBe('ERR_INVALID_EDIT');
  });
});

describe('a batch of edits is all of them or none of them', () => {
  it('rolls back to the exact starting bytes when one of them is refused', () => {
    // The refusal is only knowable at apply time: `p:txBody` holds no text until
    // the third edit puts some there, and only then does inserting an element
    // become mixed content. Without a rollback the caller would be holding a
    // half-edited document and the inverses of nothing.
    const source = `<p:sld xmlns:a="${A}" xmlns:p="${P}"><p:sp><p:nvSpPr/><p:txBody/></p:sp></p:sld>`;
    const document = parseXmlString(source);
    const sp = firstChild(document.root, 'p:sp')!;
    const body = firstChild(sp, 'p:txBody')!;
    const code = codeOf(() =>
      applyEdits([
        insertInOrder(sp, newElement('p:spPr')),
        { kind: 'setAttribute', element: sp, qname: 'zz', value: '1' },
        { kind: 'insertChild', parent: body, index: 0, node: newText('words') },
        { kind: 'insertChild', parent: body, index: 1, node: newElement('a:p') },
      ]),
    );
    expect(code).toBe('ERR_INVALID_EDIT');
    expect(serializeXmlString(document)).toBe(source);
    expect(document.root.dirty).toBe(false);
    expect(checkRoundTrip(document)).toEqual([]);
    for (const element of descendantElements(document.root)) expect(element.dirty).toBe(false);
  });
});

describe('planning several edits against one parent', () => {
  it('removes the node it named, not the position that node used to be at', () => {
    // The insertion shifts every index after it. Removal is addressed by node,
    // so it is unaffected - which is the reason for the asymmetry between the
    // two ops.
    const source = `<p:sld xmlns:a="${A}" xmlns:p="${P}"><p:sp><p:nvSpPr/><p:spPr/><p:txBody/></p:sp></p:sld>`;
    const t = trip(source, (d) => {
      const sp = firstChild(d.root, 'p:sp')!;
      return [
        insertInOrder(sp, newElement('p:style')),
        { kind: 'removeChild', parent: sp, node: firstChild(sp, 'p:txBody')! },
      ];
    });
    expect(t.after).toContain('<p:nvSpPr/><p:spPr/><p:style/>');
    expect(t.after).not.toContain('txBody');
    expectExactUndo(t, source);
  });
});

describe('an edit reaches no further than it was made', () => {
  it('changes exactly one substring of the part', () => {
    const t = trip(SLIDE, (d) => [
      { kind: 'setAttribute', element: find(d, 'a:off'), qname: 'x', value: '999' },
    ]);
    expect(t.after).toBe(SLIDE.replace('x="838200"', 'x="999"'));
  });

  it('leaves the CRLF after the declaration alone', () => {
    const t = trip(SLIDE, (d) => [
      { kind: 'setAttribute', element: find(d, 'a:off'), qname: 'x', value: '999' },
    ]);
    // The prolog's text node is clean, so it is sliced. A rebuild would have
    // turned this into a bare LF, which is exactly what 0.5 measured on 656
    // corpus parts.
    expect(t.after).toContain('?>\r\n<p:sld');
  });

  it('leaves an untouched entity reference in a text node spelled as it was', () => {
    const t = trip(SLIDE, (d) => [
      { kind: 'setAttribute', element: find(d, 'a:off'), qname: 'x', value: '999' },
    ]);
    expect(t.after).toContain('<a:t>Hello &#62; world</a:t>');
  });
});

describe('removing an element can leave two text runs side by side', () => {
  const PRETTY = `<p:sld xmlns:a="${A}" xmlns:p="${P}">\n  <p:cSld/>\n  <p:clrMapOvr/>\n</p:sld>`;

  it('and the document is unchanged by that, because XML has no such distinction', () => {
    const t = trip(PRETTY, (d) => [
      { kind: 'removeChild', parent: d.root, node: firstChild(d.root, 'p:cSld')! },
    ]);
    // Two whitespace text nodes now sit next to each other. They serialize as
    // one run and reparse as one node, which `checkRoundTrip` treats as the
    // same document - asserted inside `trip`.
    expect(t.after).toBe(`<p:sld xmlns:a="${A}" xmlns:p="${P}">\n  \n  <p:clrMapOvr/>\n</p:sld>`);
    const remaining = t.document.root.children.filter((c) => c.type === 'text');
    expect(remaining).toHaveLength(3);
    expectExactUndo(t, PRETTY);
  });
});

describe('inserted markup is markup we built, so it is checked', () => {
  it('refuses to create mixed content', () => {
    const document = parseXmlString(SLIDE);
    const t = find(document, 'a:t');
    expect(
      codeOf(() =>
        applyEdit({ kind: 'insertChild', parent: t, index: 0, node: newElement('a:r') }),
      ),
    ).toBe('ERR_INVALID_EDIT');
  });

  it('refuses a second attribute of the same name', () => {
    const document = parseXmlString(SLIDE);
    const off = find(document, 'a:off');
    expect(
      codeOf(() =>
        applyEdit({
          kind: 'addAttribute',
          element: off,
          index: 0,
          attribute: newAttribute('x', '0'),
        }),
      ),
    ).toBe('ERR_INVALID_EDIT');
  });

  it('refuses to move a node that is still attached somewhere', () => {
    const document = parseXmlString(SLIDE);
    const body = find(document, 'p:txBody');
    expect(
      codeOf(() => applyEdit({ kind: 'insertChild', parent: document.root, index: 0, node: body })),
    ).toBe('ERR_INVALID_EDIT');
  });

  it('refuses an index that is not a position', () => {
    const document = parseXmlString(SLIDE);
    expect(
      codeOf(() =>
        applyEdit({ kind: 'insertChild', parent: document.root, index: 99, node: newText('x') }),
      ),
    ).toBe('ERR_INVALID_EDIT');
    expect(
      codeOf(() =>
        applyEdit({ kind: 'removeChild', parent: document.root, node: newText('stranger') }),
      ),
    ).toBe('ERR_INVALID_EDIT');
  });

  it('refuses to remove an attribute that is not there', () => {
    const document = parseXmlString(SLIDE);
    expect(
      codeOf(() =>
        applyEdit({ kind: 'removeAttribute', element: find(document, 'a:off'), qname: 'z' }),
      ),
    ).toBe('ERR_INVALID_EDIT');
  });

  it('refuses to make a declaration somebody’s child', () => {
    const document = parseXmlString(SLIDE);
    expect(
      codeOf(() =>
        applyEdit({
          kind: 'insertChild',
          parent: document.root,
          index: 0,
          node: document.declaration!,
        }),
      ),
    ).toBe('ERR_INVALID_EDIT');
  });

  it('keeps whitespace-only text from counting as mixed content', () => {
    const source = `<p:sld xmlns:p="${P}">\n  <p:cSld/>\n</p:sld>`;
    const t = trip(source, (d) => [insertInOrder(d.root, newElement('p:clrMapOvr'))]);
    expect(t.after).toContain('<p:cSld/><p:clrMapOvr/>');
    expectExactUndo(t, source);
  });
});

describe('a new element is born dirty and knows nothing about the source', () => {
  it('serializes from its own name and attributes', () => {
    const element = newElement('a:ln', [newAttribute('w', '9525'), newAttribute('cap', 'flat')]);
    expect(element.dirty).toBe(true);
    expect(element.selfClosing).toBe(true);
    const document = parseXmlString(`<p:sld xmlns:a="${A}" xmlns:p="${P}"><p:spPr/></p:sld>`);
    applyEdit(insertInOrder(firstChild(document.root, 'p:spPr')!, element));
    expect(serializeXmlString(document)).toContain('<a:ln w="9525" cap="flat"/>');
  });

  it('switches to the long form the moment it has a child', () => {
    const document = parseXmlString(`<p:sld xmlns:a="${A}" xmlns:p="${P}"><p:spPr/></p:sld>`);
    const ln = newElement('a:ln');
    applyEdit(insertInOrder(firstChild(document.root, 'p:spPr')!, ln));
    applyEdit(insertInOrder(ln, newElement('a:solidFill')));
    expect(serializeXmlString(document)).toContain('<a:ln><a:solidFill/></a:ln>');
  });
});

describe('extLst is an ordered list of opaque things', () => {
  const WITH_EXT =
    `<p:sld xmlns:a="${A}" xmlns:p="${P}"><p:cSld/><p:extLst>` +
    '<p:ext uri="{FIRST}"><p14:a xmlns:p14="urn:p14" val="1"/></p:ext>' +
    '<p:ext uri="{SECOND}"><p14:b xmlns:p14="urn:p14" val="2"/></p:ext>' +
    '</p:extLst></p:sld>';

  it('reads entries in order and finds them by uri', () => {
    const document = parseXmlString(WITH_EXT);
    expect(extensions(document.root).map((e) => e.attributes[0]!.value)).toEqual([
      '{FIRST}',
      '{SECOND}',
    ]);
    expect(findExtension(document.root, '{SECOND}')).toBeDefined();
    expect(findExtension(document.root, '{MISSING}')).toBeUndefined();
  });

  it('appends a new entry rather than sorting the list', () => {
    const t = trip(WITH_EXT, (d) => planAddExtension(d.root, '{AAA}').edits);
    expect(t.after).toContain('<p:ext uri="{SECOND}">');
    expect(t.after.indexOf('{AAA}')).toBeGreaterThan(t.after.indexOf('{SECOND}'));
    expectExactUndo(t, WITH_EXT);
  });

  it('creates the extLst in schema order when there is none', () => {
    const source = `<p:sld xmlns:a="${A}" xmlns:p="${P}"><p:cSld/><p:clrMapOvr/></p:sld>`;
    const t = trip(source, (d) => planAddExtension(d.root, '{NEW}').edits);
    // p:extLst is last in CT_Slide's sequence, after p:clrMapOvr.
    expect(t.after).toContain('<p:clrMapOvr/><p:extLst><p:ext uri="{NEW}"/></p:extLst>');
    expectExactUndo(t, source);
  });

  it('lets the caller fill the entry in before anything is applied', () => {
    const document = parseXmlString(`<p:sld xmlns:a="${A}" xmlns:p="${P}"><p:cSld/></p:sld>`);
    const { edits, ext } = planAddExtension(document.root, '{NEW}');
    applyEdits(edits);
    applyEdit({
      kind: 'insertChild',
      parent: ext,
      index: 0,
      node: newElement('p14:creationId', [
        newAttribute('xmlns:p14', 'urn:p14'),
        newAttribute('val', '9'),
      ]),
    });
    expect(serializeXmlString(document)).toContain(
      '<p:ext uri="{NEW}"><p14:creationId xmlns:p14="urn:p14" val="9"/></p:ext>',
    );
  });

  it('removes an entry and leaves every byte of its neighbours alone', () => {
    const t = trip(WITH_EXT, (d) => planRemoveExtension(d.root, '{FIRST}'));
    expect(t.after).not.toContain('{FIRST}');
    expect(t.after).toContain('<p:ext uri="{SECOND}"><p14:b xmlns:p14="urn:p14" val="2"/></p:ext>');
    expectExactUndo(t, WITH_EXT);
  });

  it('keeps an emptied extLst, because removing it was not what was asked for', () => {
    const source = `<p:sld xmlns:a="${A}" xmlns:p="${P}"><p:extLst><p:ext uri="{ONE}"/></p:extLst></p:sld>`;
    const t = trip(source, (d) => planRemoveExtension(d.root, '{ONE}'));
    expect(t.after).toContain('<p:extLst></p:extLst>');
    expectExactUndo(t, source);
  });

  it('refuses to add an extension that is already there', () => {
    const document = parseXmlString(WITH_EXT);
    expect(codeOf(() => planAddExtension(document.root, '{FIRST}'))).toBe('ERR_INVALID_EDIT');
  });

  it('refuses to name an extLst whose namespace nothing in scope binds', () => {
    const document = parseXmlString('<sld xmlns="' + P + '"><cSld/></sld>');
    // The default namespace is PresentationML, so this one does resolve - and
    // the created element is unprefixed, matching how the file is written.
    const { edits } = planAddExtension(document.root, '{NEW}');
    applyEdits(edits);
    expect(serializeXmlString(document)).toContain('<extLst><ext uri="{NEW}"/></extLst>');
  });
});

describe('the tree is left consistent, not merely correct-looking', () => {
  it('detaches a removed node so nothing propagates into a tree it left', () => {
    const document = parseXmlString(SLIDE);
    const body = find(document, 'p:txBody');
    const parent = body.parent!;
    applyEdit({ kind: 'removeChild', parent, node: body });
    expect(body.parent).toBeUndefined();
    expect(childElements(parent).map((c) => c.qname)).toEqual(['p:nvSpPr', 'p:spPr']);
  });

  it('keeps whitespaceOnly in step with a text node it edited', () => {
    const document = parseXmlString(`<p:sld xmlns:p="${P}">\n</p:sld>`);
    const text = document.root.children[0] as XText;
    expect(text.whitespaceOnly).toBe(true);
    applyEdit({ kind: 'setValue', node: text, value: 'x' });
    expect(text.whitespaceOnly).toBe(false);
    applyEdit({ kind: 'setValue', node: text, value: ' ' });
    expect(text.whitespaceOnly).toBe(true);
  });

  it('keeps hasNamespaceDeclarations in step with an added or removed xmlns', () => {
    const document = parseXmlString(`<p:sld xmlns:p="${P}"><p:cSld/></p:sld>`);
    const cSld = firstChild(document.root, 'p:cSld')!;
    expect(cSld.hasNamespaceDeclarations).toBe(false);
    const undo = applyEdit({
      kind: 'setAttribute',
      element: cSld,
      qname: 'xmlns:q',
      value: 'urn:q',
    });
    expect(cSld.hasNamespaceDeclarations).toBe(true);
    applyEdit(undo);
    expect(cSld.hasNamespaceDeclarations).toBe(false);
  });
});
