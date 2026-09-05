/**
 * Edits, each with its exact inverse.
 *
 * The plan's exit criterion for this sub-phase is a strong one: **every edit's
 * inverse restores byte identity.** Not "an equivalent document" - the same
 * bytes. That is a much sharper claim than it first looks, and meeting it
 * decides the shape of everything below.
 *
 * ## Why value-restoration is not enough
 *
 * `dirty` is not a hint. It is an invariant: *while `node.dirty` is false,
 * `source.slice(node.start, node.end)` **is** this node's serialization.* An
 * edit falsifies it for the node it touches and for that node's ancestors, and
 * `markDirty` propagates accordingly.
 *
 * Now undo. If the inverse restores the value but leaves the flags set, the
 * part still round-trips to the same *document* - but the touched subtree is
 * rebuilt rather than sliced, so `<a:off x="0" y="0" />` may come back as
 * `<a:off x="0" y="0"/>`, a `&#62;` may come back as `>`, and a CRLF between
 * two elements comes back as LF. Sub-phase 0.5 measured that last one on 656 of
 * 2834 real parts. Undo would silently rewrite a quarter of the corpus.
 *
 * So an edit records the flags it set, and its inverse clears exactly those.
 * That is legitimate rather than a trick, and the reason is worth stating
 * precisely: **clearing `dirty` is sound exactly when the invariant it asserts
 * has been restored**, and an exact inverse is the definition of having
 * restored it. Nothing here clears a flag it did not itself set.
 *
 * ## Consequences that shape the ops
 *
 *   - Removal keeps the removed **object**, and re-insertion puts that same
 *     object back. A structural copy would carry no source span and would have
 *     to be rebuilt. This is why `removeChild`'s inverse holds an `XNode` and
 *     not a description of one.
 *   - Removing an attribute does not mark the attribute dirty, only its
 *     element, so that restoring it restores its original spacing and quote
 *     character too.
 *   - Undo is LIFO, and these edits rely on it. Each records only what it
 *     itself changed, so unwinding in order is exact; unwinding out of order is
 *     not defined, and neither is any history stack that would ask for it.
 */

import { isAllWhitespace } from '../parse/chars.js';
import { XmlError } from '../errors.js';
import { elementKey, insertionIndex } from './schema-order.js';
import { splitQName, type XAttribute } from '../parse/tokenizer.js';
import {
  declaredNamespaces,
  markDirty,
  resolvePrefix,
  type XCData,
  type XComment,
  type XElement,
  type XNode,
  type XText,
} from '../parse/xnode.js';

function fail(
  code: 'ERR_INVALID_EDIT' | 'ERR_SCHEMA_ORDER',
  message: string,
  name?: string,
): never {
  throw new XmlError(code, message, { name });
}

// ------------------------------------------------------------ new detached nodes

/**
 * An attribute that came from nowhere.
 *
 * Born dirty, and it has to be: the serializer re-emits a clean attribute by
 * slicing `[start, end)` of the source, and this one's span is `[0, 0)`. The
 * zeroed offsets are not a placeholder to be filled in later - they are how
 * `leadingSpace` knows to fall back to a single space rather than recovering
 * indentation that was never there.
 */
export function newAttribute(qname: string, value: string, quote: '"' | "'" = '"'): XAttribute {
  const { prefix, local } = splitQName(qname);
  return {
    qname,
    prefix,
    local,
    start: 0,
    nameStart: 0,
    nameEnd: 0,
    valueStart: 0,
    valueEnd: 0,
    end: 0,
    quote,
    value,
    dirty: true,
  };
}

/** An element that came from nowhere. Born dirty, for the same reason. */
export function newElement(
  qname: string,
  attributes: readonly XAttribute[] = [],
  children: readonly XNode[] = [],
): XElement {
  const { prefix, local } = splitQName(qname);
  const element: XElement = {
    type: 'element',
    qname,
    prefix,
    local,
    attributes: [...attributes],
    children: [],
    // Written `<x/>` while it has no children and `<x></x>` once it has one -
    // the serializer derives the form rather than reading this flag, so no
    // caller has to remember to change it.
    selfClosing: true,
    start: 0,
    end: 0,
    openTagEnd: 0,
    closeTagStart: 0,
    nameEnd: 0,
    trailingSpaceStart: 0,
    hasNamespaceDeclarations: attributes.some((a) => a.qname === 'xmlns' || a.prefix === 'xmlns'),
    dirty: true,
    parent: undefined,
  };
  for (const child of children) {
    child.parent = element;
    element.children.push(child);
  }
  return element;
}

/** A text node that came from nowhere. */
export function newText(value: string): XText {
  return {
    type: 'text',
    value,
    whitespaceOnly: isAllWhitespace(value, 0, value.length),
    start: 0,
    end: 0,
    dirty: true,
    parent: undefined,
  };
}

// ------------------------------------------------------------------- the edits

/**
 * The flags an edit set, so that its inverse can unset them.
 *
 * Present only on an edit that *is* an inverse - one handed back by
 * {@link applyEdit}. Constructing an edit by hand leaves it absent, which
 * simply means "set the flags and leave them set", the correct behaviour for a
 * forward edit.
 */
export interface DirtyRestore {
  readonly nodes: readonly XNode[];
  readonly attributes: readonly XAttribute[];
}

interface EditBase {
  readonly restore?: DirtyRestore | undefined;
}

/** Change an attribute's value, or add it if the element does not have it. */
export interface SetAttributeEdit extends EditBase {
  readonly kind: 'setAttribute';
  readonly element: XElement;
  readonly qname: string;
  readonly value: string;
}

/**
 * Put an attribute object at an exact position in an element's attribute list.
 *
 * The low-level counterpart of `setAttribute`, and mainly the inverse of
 * `removeAttribute`: restoring the object rather than the name and value is
 * what brings back the original spacing and quote character.
 */
export interface AddAttributeEdit extends EditBase {
  readonly kind: 'addAttribute';
  readonly element: XElement;
  readonly index: number;
  readonly attribute: XAttribute;
}

export interface RemoveAttributeEdit extends EditBase {
  readonly kind: 'removeAttribute';
  readonly element: XElement;
  readonly qname: string;
}

/**
 * Put a node at an exact index among a parent's children.
 *
 * **Not the sanctioned way to add an OOXML element** - {@link insertInOrder} is,
 * and it is the only thing that knows where the schema puts a child. This exists
 * for the places where the schema has nothing to say: an `a:ext` inside an
 * `extLst`, a branch of an `mc:AlternateContent`, a text node.
 */
export interface InsertChildEdit extends EditBase {
  readonly kind: 'insertChild';
  readonly parent: XElement;
  readonly index: number;
  readonly node: XNode;
}

/**
 * Take a node out of its parent.
 *
 * Addressed by node rather than by index, and the asymmetry with
 * {@link InsertChildEdit} is deliberate. An index is a position *at the moment
 * it is applied*, so a command that plans several edits against one parent -
 * insert a `a:ln`, drop the old `a:effectLst` - would have its second index
 * silently shifted by its first. Removal has a stable address available and
 * uses it. Insertion does not: the position is the whole content of the edit.
 */
export interface RemoveChildEdit extends EditBase {
  readonly kind: 'removeChild';
  readonly parent: XElement;
  readonly node: XNode;
}

/** Replace the value of a text node, a CDATA section or a comment. */
export interface SetValueEdit extends EditBase {
  readonly kind: 'setValue';
  readonly node: XText | XCData | XComment;
  readonly value: string;
}

export type XmlEdit =
  | SetAttributeEdit
  | AddAttributeEdit
  | RemoveAttributeEdit
  | InsertChildEdit
  | RemoveChildEdit
  | SetValueEdit;

// -------------------------------------------------------------------- applying

interface Journal {
  readonly nodes: XNode[];
  readonly attributes: XAttribute[];
}

function attributeIndex(element: XElement, qname: string): number {
  return element.attributes.findIndex((a) => a.qname === qname);
}

/**
 * Refuse to put an element into content that carries real text.
 *
 * No OOXML type we edit has mixed content: `<a:t>` holds text and nothing else,
 * every other element holds elements and inter-element whitespace. An element
 * dropped into `<a:t>Hello</a:t>` produces markup that is well-formed, passes
 * every check in this package, and is refused by PowerPoint with no diagnostic.
 *
 * **Forward edits only**, and the exception is not a loophole - it is the one
 * rule this file cannot break. Every other precondition here is *structural*:
 * an index in range, a node not already attached, no attribute of that name.
 * An exact inverse satisfies all of them by construction, because it puts back
 * a state that existed. This one is *semantic*, about the shape of the
 * surrounding document, and an inverse can fail it: remove an element from a
 * parent that already held text - which a real file never does but a mangled one
 * can - and putting it back is "creating mixed content" by this test, so undo
 * would be refused and the document could not be returned to what it was.
 *
 * The 40 000-case fuzz sweep found that; it showed up as batches whose rollback
 * threw. A guard that can refuse an inverse is worse than no guard, because the
 * property it breaks is the one everything above this layer depends on.
 */
function assertNotMixed(parent: XElement, node: XNode): void {
  if (node.type !== 'element') return;
  for (const child of parent.children) {
    if (child.type === 'text' && !child.whitespaceOnly) {
      fail(
        'ERR_INVALID_EDIT',
        'cannot insert <' +
          node.qname +
          '> into <' +
          parent.qname +
          '>, which holds text: OOXML has no mixed content and PowerPoint refuses it',
        parent.qname,
      );
    }
  }
}

function setAttribute(edit: SetAttributeEdit, journal: Journal): XmlEdit {
  const index = attributeIndex(edit.element, edit.qname);
  if (index < 0) {
    const attribute = newAttribute(edit.qname, edit.value);
    edit.element.attributes.push(attribute);
    if (attribute.prefix === 'xmlns' || attribute.qname === 'xmlns') {
      edit.element.hasNamespaceDeclarations = true;
    }
    markDirty(edit.element, journal.nodes);
    return { kind: 'removeAttribute', element: edit.element, qname: edit.qname };
  }
  const attribute = edit.element.attributes[index]!;
  const previous = attribute.value;
  attribute.value = edit.value;
  if (!attribute.dirty) {
    attribute.dirty = true;
    journal.attributes.push(attribute);
  }
  markDirty(edit.element, journal.nodes);
  return { kind: 'setAttribute', element: edit.element, qname: edit.qname, value: previous };
}

function addAttribute(edit: AddAttributeEdit, journal: Journal): XmlEdit {
  if (edit.index < 0 || edit.index > edit.element.attributes.length) {
    fail('ERR_INVALID_EDIT', 'attribute index ' + String(edit.index) + ' is out of range');
  }
  if (attributeIndex(edit.element, edit.attribute.qname) >= 0) {
    // XML 1.0 §3.1, and verified against PowerPoint in sub-phase 0.4:
    // `<a:off x="0" y="0" x="0"/>` is refused with 0x80070570.
    fail(
      'ERR_INVALID_EDIT',
      '<' + edit.element.qname + '> already has an attribute named "' + edit.attribute.qname + '"',
      edit.attribute.qname,
    );
  }
  edit.element.attributes.splice(edit.index, 0, edit.attribute);
  if (edit.attribute.prefix === 'xmlns' || edit.attribute.qname === 'xmlns') {
    edit.element.hasNamespaceDeclarations = true;
  }
  // Deliberately does not touch `attribute.dirty`. A restored attribute keeps
  // whatever it was, and a freshly built one is already dirty.
  markDirty(edit.element, journal.nodes);
  return { kind: 'removeAttribute', element: edit.element, qname: edit.attribute.qname };
}

function removeAttribute(edit: RemoveAttributeEdit, journal: Journal): XmlEdit {
  const index = attributeIndex(edit.element, edit.qname);
  if (index < 0) {
    fail(
      'ERR_INVALID_EDIT',
      '<' + edit.element.qname + '> has no attribute named "' + edit.qname + '"',
      edit.qname,
    );
  }
  const attribute = edit.element.attributes[index]!;
  edit.element.attributes.splice(index, 1);
  if (attribute.prefix === 'xmlns' || attribute.qname === 'xmlns') {
    edit.element.hasNamespaceDeclarations = edit.element.attributes.some(
      (a) => a.qname === 'xmlns' || a.prefix === 'xmlns',
    );
  }
  markDirty(edit.element, journal.nodes);
  return { kind: 'addAttribute', element: edit.element, index, attribute };
}

function insertChild(edit: InsertChildEdit, journal: Journal, forward: boolean): XmlEdit {
  if (edit.index < 0 || edit.index > edit.parent.children.length) {
    fail('ERR_INVALID_EDIT', 'child index ' + String(edit.index) + ' is out of range');
  }
  if (edit.node.parent !== undefined) {
    fail(
      'ERR_INVALID_EDIT',
      'that node is already a child of <' +
        edit.node.parent.qname +
        '>; remove it before inserting it elsewhere',
    );
  }
  if (edit.node.type === 'declaration') {
    fail('ERR_INVALID_EDIT', 'an XML declaration is not a child of any element');
  }
  if (forward) assertNotMixed(edit.parent, edit.node);
  edit.parent.children.splice(edit.index, 0, edit.node);
  edit.node.parent = edit.parent;
  // Not `markDirty(edit.node)`: a node restored by an undo is clean and must
  // stay clean, and a synthesized one is already dirty.
  markDirty(edit.parent, journal.nodes);
  return { kind: 'removeChild', parent: edit.parent, node: edit.node };
}

function removeChild(edit: RemoveChildEdit, journal: Journal): XmlEdit {
  const index = edit.parent.children.indexOf(edit.node);
  if (index < 0) {
    fail(
      'ERR_INVALID_EDIT',
      'that node is not a child of <' + edit.parent.qname + '>',
      edit.parent.qname,
    );
  }
  edit.parent.children.splice(index, 1);
  edit.node.parent = undefined;
  markDirty(edit.parent, journal.nodes);
  return { kind: 'insertChild', parent: edit.parent, index, node: edit.node };
}

function setValue(edit: SetValueEdit, journal: Journal): XmlEdit {
  const previous = edit.node.value;
  edit.node.value = edit.value;
  if (edit.node.type === 'text')
    edit.node.whitespaceOnly = isAllWhitespace(edit.value, 0, edit.value.length);
  markDirty(edit.node, journal.nodes);
  return { kind: 'setValue', node: edit.node, value: previous };
}

/**
 * Perform one edit and return the edit that undoes it.
 *
 * The returned edit carries a {@link DirtyRestore} listing exactly the flags
 * this call set, so applying it puts the tree back to the state it was in -
 * including which nodes still describe their own source.
 *
 * Values are not checked here. A string that cannot be written as XML is caught
 * by the serializer, which owns that question and already refuses unpaired
 * surrogates, characters outside `Char`, and the terminators of the constructs
 * that have no escape. Duplicating the check would put two answers in the tree
 * to the same question.
 */
export function applyEdit(edit: XmlEdit): XmlEdit {
  const journal: Journal = { nodes: [], attributes: [] };
  let inverse: XmlEdit;
  switch (edit.kind) {
    case 'setAttribute':
      inverse = setAttribute(edit, journal);
      break;
    case 'addAttribute':
      inverse = addAttribute(edit, journal);
      break;
    case 'removeAttribute':
      inverse = removeAttribute(edit, journal);
      break;
    case 'insertChild':
      inverse = insertChild(edit, journal, edit.restore === undefined);
      break;
    case 'removeChild':
      inverse = removeChild(edit, journal);
      break;
    case 'setValue':
      inverse = setValue(edit, journal);
      break;
  }

  // Last, and only now: this edit is somebody's inverse, and the flags it was
  // told to clear describe state that has just been restored. Clearing them
  // before the mutation would let the mutation's own propagation set them again.
  if (edit.restore !== undefined) {
    for (const node of edit.restore.nodes) node.dirty = false;
    for (const attribute of edit.restore.attributes) attribute.dirty = false;
  }

  return { ...inverse, restore: { nodes: journal.nodes, attributes: journal.attributes } };
}

/**
 * Perform a list of edits and return the list that undoes it, already reversed.
 *
 * So `applyEdits(applyEdits(edits))` is the identity - which is the property the
 * corpus gate measures, over every part it can reach.
 *
 * **All of them, or none of them.** If any edit is refused, the ones that
 * already landed are undone before the error is rethrown, and the document is
 * left exactly as it was found - dirty flags included, since the rollback is
 * made of the same exact inverses everything else here is.
 *
 * That is not a nicety. Some refusals are only knowable at apply time: whether
 * an element would create mixed content depends on the state the earlier edits
 * in the same batch left behind. Without the rollback, a caller planning a
 * gesture as eight edits and having the fifth refused would be holding a
 * half-applied document and no way back, because the inverses of the first four
 * are only returned on success. The 40 000-case fuzz sweep found 66 such cases;
 * that is how this came to be here.
 *
 * A failure during the rollback is not caught. It cannot happen without a defect
 * in this file, and swallowing it to re-raise the original would hide the worse
 * of the two problems.
 */
export function applyEdits(edits: readonly XmlEdit[]): XmlEdit[] {
  const inverses: XmlEdit[] = [];
  try {
    for (const edit of edits) inverses.push(applyEdit(edit));
  } catch (error) {
    for (let i = inverses.length - 1; i >= 0; i--) applyEdit(inverses[i]!);
    throw error;
  }
  inverses.reverse();
  return inverses;
}

// ------------------------------------------------------------ ordered insertion

/**
 * The namespace a detached element will be in once it is inside `parent`.
 *
 * A node that is not yet in a tree has no namespace scope, so asking it
 * directly answers `undefined` for every prefixed name. Its own declarations
 * win if it carries any, and otherwise the scope it is about to join decides.
 */
function keyForChild(parent: XElement, child: XElement): string | undefined {
  const own = declaredNamespaces(child).get(child.prefix);
  const namespace = own ?? resolvePrefix(parent, child.prefix);
  return namespace === undefined ? undefined : '{' + namespace + '}' + child.local;
}

/**
 * Plan the insertion of an OOXML element at the position the schema gives it.
 *
 * **This is the only sanctioned way to add an OOXML child.** Not because
 * `insertChild` is broken, but because the position is not a caller's to
 * choose: OOXML complex types are `xsd:sequence`, and Microsoft's own Open XML
 * SDK once shipped a build that swapped two elements in `slideMaster1.xml`,
 * after which PowerPoint refused the file.
 *
 * Returns the edit rather than performing it, so a command can plan a whole
 * gesture, decide it is legal, and apply it as one unit - which is the shape
 * sub-phase 5.1's command bus needs.
 *
 * Throws `ERR_SCHEMA_ORDER` when the schema does not place the child there. The
 * three reasons, all of which are real:
 *
 *   - the parent's content is `xsd:any` - an `extLst` entry or a
 *     `graphicData` payload - and the plan forbids rebuilding either;
 *   - the parent is in an extension namespace the standard does not describe;
 *   - the child is simply not a child that parent admits.
 *
 * In each case appending anyway would produce a file PowerPoint may reject with
 * no diagnostic. A caller that genuinely means to append into opaque content
 * says so with {@link InsertChildEdit}.
 */
export function insertInOrder(parent: XElement, child: XElement): InsertChildEdit {
  const childKey = keyForChild(parent, child);
  if (childKey === undefined) {
    fail(
      'ERR_SCHEMA_ORDER',
      'the prefix of <' + child.qname + '> is bound to nothing inside <' + parent.qname + '>',
      child.qname,
    );
  }
  const index = insertionIndex(parent, childKey);
  if (index === undefined) {
    const parentKey = elementKey(parent);
    fail(
      'ERR_SCHEMA_ORDER',
      'the schema gives <' +
        child.qname +
        '> no place inside <' +
        parent.qname +
        '>' +
        (parentKey === undefined ? ' (whose namespace prefix is unbound)' : ''),
      child.qname,
    );
  }
  return { kind: 'insertChild', parent, index, node: child };
}
