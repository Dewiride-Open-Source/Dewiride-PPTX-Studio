import { isRelationshipPartName } from '@pptx-studio/opc';
import {
  attributeValue,
  childElements,
  childRanks,
  descendantElements,
  elementKey,
  isSchemaNamespace,
  namespaceOf,
  outOfOrderChildren,
  type XElement,
} from '@pptx-studio/xml';
import type { Context } from './context.js';
import { elementLocation, xpathOf } from './location.js';

/**
 * `V010` … `V012`: where children are allowed to be.
 *
 * OOXML complex types are `xsd:sequence` almost everywhere, and this is the
 * category with the best-attested consequence in the whole appendix: Microsoft's
 * own Open XML SDK shipped a regression that swapped two elements in
 * `slideMaster1.xml`, changed nothing else, and PowerPoint refused the file.
 *
 * The ranks come from `@pptx-studio/xml`'s generated table, which is built from
 * the ECMA-376 Transitional schemas and checked against 194 148 elements
 * PowerPoint wrote. That second number is what makes the rule safe to make
 * fatal: a table generated from a standard and never compared with reality
 * would be a very confident way to reject files that work.
 *
 * The three rules divide the question by what the table can and cannot see:
 *
 * - `V010` - the child is in the model but in the wrong place.
 * - `V011` - `extLst` last, which has to hold in the places the table cannot
 *   rank at all.
 * - `V012` - the child is not in the model.
 *
 * `V010` skips anything it cannot rank, which is correct for ordering and is
 * exactly the blind spot `V012` exists to cover.
 */

/** Children in schema-sequence order. */
export function v010SchemaOrder(ctx: Context): void {
  forEachElement(ctx, (part, element) => {
    for (const problem of outOfOrderChildren(element)) {
      ctx.add(
        'V010',
        { part, xpath: xpathOf(element), offset: problem.at },
        problem.detail +
          '. OOXML complex types are xsd:sequence and PowerPoint enforces the order literally.',
      );
    }
  });
}

/**
 * `extLst` is last, and every `a:ext` names its extension.
 *
 * There is exactly one type in the whole generated table where `extLst` is not
 * last: `p:control`, whose model is `extLst` then `pic`. That is not an
 * exception worth handling gracefully - `V025` refuses `p:control` outright,
 * because PowerPoint refuses it in all eight forms we tried - but it is skipped
 * here rather than reported twice.
 */
export function v011ExtLstLast(ctx: Context): void {
  forEachElement(ctx, (part, element) => {
    if (element.local !== 'extLst') return;
    const parent = element.parent;
    if (parent === undefined || parent.local === 'control') return;

    const siblings = childElements(parent);
    const last = siblings[siblings.length - 1];
    if (last !== element) {
      ctx.add(
        'V011',
        elementLocation(part, element),
        '<' +
          element.qname +
          '> is followed by <' +
          (last?.qname ?? '?') +
          '> inside <' +
          parent.qname +
          '>. An extension list is last in every type that has one.',
      );
    }

    for (const ext of childElements(element)) {
      if (ext.local !== 'ext') continue;
      const uri = attributeValue(ext, 'uri');
      if (uri !== undefined && uri !== '') continue;
      ctx.add(
        'V011',
        elementLocation(part, ext),
        '<' +
          ext.qname +
          '> has no @uri. The uri is the only identity an extension has, and the contract for ' +
          'extension markup is that it is carried through untouched *keyed by that uri* - so one ' +
          'without it cannot be preserved, only guessed at.',
      );
    }
  });
}

/**
 * No child in a parent whose content model has no place for it.
 *
 * Deliberately conservative in four ways, because a false positive here refuses
 * a file somebody wanted:
 *
 * - only parents the generated table knows are checked, and `childRanks`
 *   returns `undefined` for every open (`xsd:any`) type, so `a:ext`,
 *   `p:ext` and `a:graphicData` - the three places the plan forbids us to
 *   rebuild - are never looked at;
 * - **only children in the four namespaces the table was generated from.**
 *   This is the one the corpus taught us. `childRanks` collapses "the schema
 *   forbids this here" and "we have never read this schema" into the same
 *   `undefined`, which is right for a caller trying to insert a child and
 *   catastrophic for one trying to judge an existing one: without the
 *   namespace test, the rule fired on forty `mc:AlternateContent` elements,
 *   on `a14:m`, on `p14:honeycomb` and on `a37-mce`'s deliberate `zz:` markup -
 *   which is `mc:Ignorable` extension markup doing exactly what it exists for;
 * - `mc:Choice` and `mc:Fallback` are unrankable parents, so their contents
 *   fall out for free.
 */
export function v012UnexpectedChild(ctx: Context): void {
  forEachElement(ctx, (part, element) => {
    const ranks = childRanks(element);
    if (ranks === undefined) return;
    for (const child of childElements(element)) {
      if (!isSchemaNamespace(namespaceOf(child))) continue;
      const key = elementKey(child);
      if (key === undefined || ranks.has(key)) continue;
      ctx.add(
        'V012',
        elementLocation(part, child),
        '<' +
          child.qname +
          '> is not a child <' +
          element.qname +
          '> admits. `a:ahXY` or `a:cxn` placed directly under `a:custGeom`, without its ' +
          '`a:ahLst`/`a:cxnLst` wrapper, is the case that made this a rule: a whole-package ' +
          'refusal that the ordering rule cannot see, because an unrankable child is skipped ' +
          'there rather than reported.',
      );
    }
  });
}

/**
 * Every element of every XML part, once.
 *
 * The three rules above each want the same walk. Sharing it means one traversal
 * of the package instead of three, and - more usefully - one place where the
 * decision "which parts does an order rule apply to" is made. Relationship
 * parts are excluded: `Relationships` is not in any schema we generate from, so
 * every one of its children would be unrankable and the rules would be checking
 * nothing while appearing to check everything.
 */
function forEachElement(ctx: Context, visit: (part: string, element: XElement) => void): void {
  for (const part of ctx.parts()) {
    if (isRelationshipPartName(part)) continue;
    const document = ctx.document(part);
    if (document === null) continue;
    for (const element of descendantElements(document.root)) visit(part, element);
  }
}
