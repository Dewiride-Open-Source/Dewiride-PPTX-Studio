/**
 * `extLst` - an ordered list of opaque things, keyed by GUID.
 *
 * ```xml
 * <p:extLst>
 *   <p:ext uri="{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}">
 *     <p14:creationId xmlns:p14="…/powerpoint/2010/main" val="1234567890"/>
 *   </p:ext>
 * </p:extLst>
 * ```
 *
 * The schema for the inside of an `ext` is `xsd:any`. That is not a gap in the
 * standard - it is the extension point, and what sits there was written by a
 * version of Office that may postdate anything we know about. The plan's rule
 * is therefore flat: **never rebuild an `extLst` from a typed model.** Read it
 * as a list, address an entry by its `uri`, add and remove whole entries, and
 * leave every byte inside an entry alone.
 *
 * Three consequences that this module exists to enforce:
 *
 *   - **Order is preserved, never sorted.** A new entry is appended. GUIDs have
 *     no meaningful order, so any sort we invented would rewrite files for
 *     nothing, and rewriting an extension list is one of the reliable ways to
 *     make PowerPoint offer to repair a deck.
 *   - **`uri` is the key**, not position. It is `use="required"` on both
 *     `CT_Extension` and `CT_OfficeArtExtension`, and it is unprefixed - an
 *     attribute in no namespace.
 *   - **`extLst` is last in its parent's sequence**, in every complex type that
 *     has one. That falls out of the generated order table rather than being
 *     asserted here, which is the point of having generated it.
 */

import { XmlError } from '../errors.js';
import { insertInOrder, newAttribute, newElement, type XmlEdit } from '../edit/edit.js';
import { childRanks } from '../edit/schema-order.js';
import { attributeValue, childElements, prefixFor, type XElement } from '../parse/xnode.js';

/** The local name of both `p:extLst` and `a:extLst`, and of their entries. */
const EXT_LST = 'extLst';
const EXT = 'ext';

/** The `extLst` child of an element, whichever namespace the schema puts it in. */
export function extensionList(parent: XElement): XElement | undefined {
  for (const child of childElements(parent)) {
    if (child.local === EXT_LST) return child;
  }
  return undefined;
}

/** The entries of an element's `extLst`, in document order. Empty if it has none. */
export function extensions(parent: XElement): XElement[] {
  const list = extensionList(parent);
  if (list === undefined) return [];
  return childElements(list).filter((child) => child.local === EXT);
}

/**
 * The entry with this `uri`, or `undefined`.
 *
 * Compared exactly, braces and case included. PowerPoint writes these GUIDs in
 * upper case inside braces and matches them the same way; normalising here
 * would let two entries that Office considers distinct collide.
 */
export function findExtension(parent: XElement, uri: string): XElement | undefined {
  for (const ext of extensions(parent)) {
    if (attributeValue(ext, 'uri') === uri) return ext;
  }
  return undefined;
}

/**
 * The namespace the schema puts this parent's `extLst` in, if it admits one.
 *
 * Read out of the generated order table rather than assumed, because it is not
 * always the parent's own: `p:sp` takes a `p:extLst`, while `a:spPr` - which is
 * what `p:spPr` is typed as - takes an `a:extLst`.
 */
function extensionListNamespace(parent: XElement): string | undefined {
  const ranks = childRanks(parent);
  if (ranks === undefined) return undefined;
  for (const key of ranks.keys()) {
    if (key.endsWith('}' + EXT_LST)) return key.slice(1, key.indexOf('}'));
  }
  return undefined;
}

/**
 * Add an extension entry, creating the `extLst` if the element has none.
 *
 * Returns the edits to apply, and the entry element they will install, so that
 * a caller can fill the entry in before applying - the entry's content is
 * opaque to us, and building it is the caller's business.
 *
 * The `extLst` element, when it has to be created, is named with a prefix
 * already bound in scope. Declaring a namespace is an edit with its own
 * consequences for every descendant, and this function will not make that
 * decision quietly; if nothing in scope binds the namespace the schema calls
 * for, it refuses.
 */
export function planAddExtension(
  parent: XElement,
  uri: string,
): { edits: XmlEdit[]; ext: XElement } {
  const existing = findExtension(parent, uri);
  if (existing !== undefined) {
    throw new XmlError(
      'ERR_INVALID_EDIT',
      '<' + parent.qname + '> already carries the extension ' + uri,
      { name: uri },
    );
  }

  const edits: XmlEdit[] = [];
  let list = extensionList(parent);

  if (list === undefined) {
    const namespace = extensionListNamespace(parent);
    if (namespace === undefined) {
      throw new XmlError('ERR_SCHEMA_ORDER', 'the schema gives <' + parent.qname + '> no extLst', {
        name: parent.qname,
      });
    }
    const prefix = prefixFor(parent, namespace);
    if (prefix === undefined) {
      throw new XmlError(
        'ERR_SCHEMA_ORDER',
        'nothing in scope at <' +
          parent.qname +
          '> binds ' +
          namespace +
          ', so its extLst cannot be named',
        { name: parent.qname },
      );
    }
    list = newElement(prefix === '' ? EXT_LST : prefix + ':' + EXT_LST);
    edits.push(insertInOrder(parent, list));
  }

  // The entry is in the same namespace as the list that holds it, and takes the
  // list's own prefix so the two read as a pair in the file.
  const colon = list.qname.indexOf(':');
  const extQName = colon < 0 ? EXT : list.qname.slice(0, colon) + ':' + EXT;
  const ext = newElement(extQName, [newAttribute('uri', uri)]);
  // Appended, never sorted: the order an extension list arrived in is the order
  // it leaves in.
  edits.push({ kind: 'insertChild', parent: list, index: list.children.length, node: ext });
  return { edits, ext };
}

/**
 * Remove an extension entry by `uri`.
 *
 * An `extLst` left with no entries is kept. `<p:extLst/>` is schema-legal - the
 * entry group is `minOccurs="0"` - and real files contain them, so removing it
 * would be a rewrite we were not asked for, in the one part of the tree where
 * unrequested rewrites are most expensive.
 */
export function planRemoveExtension(parent: XElement, uri: string): XmlEdit[] {
  const list = extensionList(parent);
  if (list === undefined) return [];
  const ext = list.children.find(
    (child) =>
      child.type === 'element' && child.local === EXT && attributeValue(child, 'uri') === uri,
  );
  if (ext === undefined) return [];
  return [{ kind: 'removeChild', parent: list, node: ext }];
}
