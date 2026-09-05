/**
 * Where a new child goes.
 *
 * OOXML complex types are `xsd:sequence`. Microsoft's own Open XML SDK once
 * shipped a regression that swapped two elements in `slideMaster1.xml` and
 * PowerPoint refused the file, with no diagnostic beyond "found a problem". So
 * the position of an inserted element is not a matter of taste, and it is not
 * something a caller should be trusted to remember.
 *
 * ## A rank, not a list
 *
 * Every child name a parent admits carries an integer rank, and in a valid
 * document children appear in non-decreasing rank order. Names that share a
 * rank may appear in either order.
 *
 * That last part is not an approximation of the schema, it *is* the schema, and
 * it is load-bearing: `p:spTree` admits `sp`, `grpSp`, `graphicFrame`, `cxnSp`,
 * `pic` and `contentPart` under one repeating `xsd:choice`, and their freedom to
 * interleave is z-order. A strict list would make inserting a picture reorder
 * the slide.
 *
 * ## Keys are namespaces, never prefixes
 *
 * `p:` and `a:` are conventions, not rules; a producer may bind
 * PresentationML to any prefix it likes, and the corpus already contains one
 * prefix bound to two different URIs in different parts. So the table is keyed
 * `{namespace}local` and the prefix is resolved through the tree at lookup
 * time, the same way {@link namespaceOf} does everywhere else in this package.
 *
 * ## Three names need a grandparent
 *
 * The table is keyed by parent element name, because that is what a caller
 * holding an `XElement` has. The schema is keyed by complex type, and the two
 * are not the same map. Almost always the difference is invisible - `a:ext` is
 * `CT_PositiveSize2D` under `a:xfrm` and `CT_OfficeArtExtension` under
 * `a:extLst`, and only the second has element children at all - but three names
 * genuinely disagree:
 *
 * | name      | under          | children                       |
 * | --------- | -------------- | ------------------------------ |
 * | `a:xfrm`  | `p:spPr`       | `a:off`, `a:ext`               |
 * | `a:xfrm`  | `p:grpSpPr`    | plus `a:chOff`, `a:chExt`      |
 * | `a:path`  | `a:pathLst`    | `a:moveTo`, `a:lnTo`, …        |
 * | `a:path`  | `a:gradFill`   | `a:fillToRect`                 |
 * | `p:to`    | `p:animClr`    | a colour                       |
 * | `p:to`    | `p:set`        | an animation variant           |
 *
 * `a:xfrm` is the one that matters: every drag writes into it. So those get one
 * more level of key, and the codegen asserts that one more level is enough - it
 * reports any pair the grandparent still fails to separate, and there are none.
 */

import { CONTEXT_DATA, NAMESPACES, ORDER_DATA, SCHEMA_SOURCES } from './schema-order.gen.js';
import { namespaceOf, type XElement } from '../parse/xnode.js';

export { SCHEMA_SOURCES };

/**
 * The four namespaces the ordering table was generated from.
 *
 * Exported because "this element is not in the table" has two entirely
 * different meanings and only one of them is a problem. An `a:solidFill` the
 * table does not rank under `a:tblBg` is markup in a vocabulary we generated
 * from, so the table's silence is a statement. A `p14:honeycomb` under
 * `p:transition` is markup from a vocabulary we did not generate from, so the
 * table's silence means nothing at all - it is `mc:Ignorable` extension markup
 * doing exactly what it is for.
 *
 * `childRanks` deliberately collapses the two, because a caller trying to
 * *insert* a child cannot act on either. A caller trying to *judge* one can,
 * and `@pptx-studio/validate`'s `V012` is that caller.
 */
export const SCHEMA_NAMESPACES: readonly string[] = NAMESPACES;

/** True when the ordering table was generated from this element's namespace. */
export function isSchemaNamespace(namespace: string | undefined): boolean {
  return namespace !== undefined && NAMESPACES.includes(namespace);
}

/** The `{namespace}local` key used throughout this module. */
export function qualifiedKey(namespace: string, local: string): string {
  return '{' + namespace + '}' + local;
}

/**
 * The key of an element as it sits in its document.
 *
 * `undefined` when the element's prefix is bound to nothing, which is a
 * malformed document rather than an element in no namespace - the tokenizer
 * accepts it, `undeclaredPrefixes` reports it, and nothing here can order it.
 */
export function elementKey(element: XElement): string | undefined {
  const namespace = namespaceOf(element);
  return namespace === undefined ? undefined : qualifiedKey(namespace, element.local);
}

function expand(short: string): string {
  const namespace = NAMESPACES[Number(short.charAt(0))];
  if (namespace === undefined) throw new Error('schema-order.gen.ts is inconsistent');
  return qualifiedKey(namespace, short.slice(1));
}

/**
 * Decoded on first use, not at module load.
 *
 * The table is one string of 18 664 characters. A document that is only read
 * never inserts anything, so it never pays for this; and when it does, the cost
 * is one pass over that string rather than the parse of a large object literal
 * that a bundler would also have to keep intact.
 */
function decode(data: string, pair: boolean): Map<string, Map<string, number>> {
  const table = new Map<string, Map<string, number>>();
  if (data === '') return table;
  for (const entry of data.split(';')) {
    const equals = entry.indexOf('=');
    const ranks = new Map<string, number>();
    const groups = entry.slice(equals + 1).split(',');
    for (let rank = 0; rank < groups.length; rank++) {
      for (const name of groups[rank]!.split('|')) ranks.set(expand(name), rank);
    }
    const head = entry.slice(0, equals);
    if (!pair) {
      table.set(expand(head), ranks);
      continue;
    }
    const arrow = head.indexOf('>');
    table.set(expand(head.slice(0, arrow)) + '>' + expand(head.slice(arrow + 1)), ranks);
  }
  return table;
}

let byParent: Map<string, Map<string, number>> | undefined;
let byGrandparent: Map<string, Map<string, number>> | undefined;

/**
 * The ranks of every child name this parent admits, or `undefined`.
 *
 * `undefined` covers four different situations, and all four mean the same
 * thing to a caller - *we cannot place a child here* - so they are deliberately
 * not distinguished:
 *
 *   - the element is not in PresentationML or DrawingML at all (an extension
 *     namespace, or `mc:AlternateContent`);
 *   - its content model is `xsd:any`, which is `a:extLst`'s `a:ext`,
 *     `p:extLst`'s `p:ext`, and `a:graphicData` - the three places the plan
 *     forbids us to rebuild;
 *   - it has no element children in the schema;
 *   - it is one of the three ambiguous names and its grandparent does not
 *     resolve which content model applies.
 */
export function childRanks(parent: XElement): ReadonlyMap<string, number> | undefined {
  byParent ??= decode(ORDER_DATA, false);
  byGrandparent ??= decode(CONTEXT_DATA, true);
  const parentKey = elementKey(parent);
  if (parentKey === undefined) return undefined;
  const grandparent = parent.parent;
  if (grandparent !== undefined) {
    const grandparentKey = elementKey(grandparent);
    if (grandparentKey !== undefined) {
      const contextual = byGrandparent.get(grandparentKey + '>' + parentKey);
      if (contextual !== undefined) return contextual;
    }
  }
  return byParent.get(parentKey);
}

/** The rank of one child name under this parent, or `undefined` if it has none there. */
export function childRank(parent: XElement, childKey: string): number | undefined {
  return childRanks(parent)?.get(childKey);
}

/**
 * The index in `parent.children` at which a child of this name belongs.
 *
 * `undefined` when the schema does not place it - see {@link childRanks}.
 *
 * Two properties matter more than the arithmetic:
 *
 * **Nothing already present ever moves.** The result is an insertion point, and
 * the existing children keep both their order and, being untouched, their
 * `dirty` flags and their source spans. An `mc:AlternateContent` branch or a
 * `p14:` extension element sitting among the ranked children is skipped rather
 * than ranked, so it stays exactly where its producer put it. That is the plan's
 * rule - *elements not in the table pin to current position rather than
 * reorder* - and it is why this returns an index instead of a sorted array.
 *
 * **Equal rank means last, not first.** A new `p:pic` lands after every existing
 * shape in the `p:spTree`, which is the top of the z-order, which is where a
 * newly drawn shape goes.
 *
 * The scan does not stop early at the first child of higher rank. In a valid
 * document it could; in one whose children are already out of order it would
 * pick a position based on a prefix of the evidence, and running to the end is
 * both simpler to reason about and more forgiving.
 */
export function insertionIndex(parent: XElement, childKey: string): number | undefined {
  const ranks = childRanks(parent);
  if (ranks === undefined) return undefined;
  const rank = ranks.get(childKey);
  if (rank === undefined) return undefined;

  let index = 0;
  for (let i = 0; i < parent.children.length; i++) {
    const child = parent.children[i]!;
    if (child.type !== 'element') continue;
    const key = elementKey(child);
    if (key === undefined) continue;
    const other = ranks.get(key);
    if (other === undefined) continue;
    if (other <= rank) index = i + 1;
  }
  return index;
}

/**
 * Children that sit out of schema order, as adjacent pairs.
 *
 * Not used by insertion - which never reorders - but by the tests and by the
 * corpus gate, where the claim being checked is that 194 148 elements written
 * by PowerPoint agree with a table generated from the standard. Anything the
 * table cannot rank is skipped, so extension markup is not reported as
 * disorder.
 */
export function outOfOrderChildren(parent: XElement): { at: number; detail: string }[] {
  const ranks = childRanks(parent);
  if (ranks === undefined) return [];
  const problems: { at: number; detail: string }[] = [];
  let previousRank = -1;
  let previousName = '';
  for (const child of parent.children) {
    if (child.type !== 'element') continue;
    const key = elementKey(child);
    if (key === undefined) continue;
    const rank = ranks.get(key);
    if (rank === undefined) continue;
    if (rank < previousRank) {
      problems.push({
        at: child.start,
        detail:
          '<' + child.qname + '> follows <' + previousName + '> inside <' + parent.qname + '>',
      });
    }
    previousRank = rank;
    previousName = child.qname;
  }
  return problems;
}
