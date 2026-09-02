import { CONTENT_TYPE, isRelationshipPartName } from '@pptx-studio/opc';
import {
  attribute,
  childElements,
  descendantElements,
  namespaceOf,
  NS,
  type XElement,
} from '@pptx-studio/xml';
import type { Context } from './context.js';
import { elementLocation } from './location.js';

/**
 * `V013` … `V017`: children and attributes that are not optional.
 *
 * Every one of these is a `minOccurs` the schema states plainly, and every one
 * is broken the same way: by an editor deleting the last of something. The last
 * paragraph goes and a `p:txBody` is left with no `a:p`; a shape is dragged out
 * of a group and the group's `p:grpSpPr` is dropped with it. That is why these
 * are separate rules from the ordering ones even though a missing child and a
 * misplaced child are both "the sequence is wrong": the messages have to say
 * *add this*, not *move this*, and they have to fire on the parent rather than
 * on a child that is not there to point at.
 */

/** The twelve attributes of `CT_ColorMapping`. All required, no defaults. */
const CLR_MAP_ATTRIBUTES = [
  'bg1',
  'tx1',
  'bg2',
  'tx2',
  'accent1',
  'accent2',
  'accent3',
  'accent4',
  'accent5',
  'accent6',
  'hlink',
  'folHlink',
] as const;

/** `p:presentation` carries `p:notesSz`. */
export function v013NotesSize(ctx: Context): void {
  for (const part of ctx.parts()) {
    const document = ctx.document(part);
    if (document === null) continue;
    const root = document.root;
    if (root.local !== 'presentation' || namespaceOf(root) !== NS.p) continue;

    if (childElements(root).some((child) => child.local === 'notesSz')) continue;
    ctx.add(
      'V013',
      elementLocation(part, root),
      'no <p:notesSz>. In CT_Presentation p:sldSz is [0..1] and p:notesSz is [1..1] - the ' +
        'asymmetry is the whole of this rule, and a generator that treats the two the same way ' +
        'writes a presentation part that is missing a required child.',
    );
  }
}

/** `p:clrMap` carries all twelve attributes. */
export function v014ColorMap(ctx: Context): void {
  forEachElement(ctx, (part, element) => {
    if (element.local !== 'clrMap' || namespaceOf(element) !== NS.p) return;
    const missing = CLR_MAP_ATTRIBUTES.filter((name) => attribute(element, name) === undefined);
    if (missing.length === 0) return;
    ctx.add(
      'V014',
      elementLocation(part, element),
      'the colour map is missing ' +
        missing.join(', ') +
        '. All twelve attributes of CT_ColorMapping are required and none has a default; a ' +
        'partial map is not a map with a gap, it is an invalid element - and it is the element ' +
        'that decides what bg1 and tx1 resolve to everywhere in the deck.',
    );
  });
}

/** `p:spTree` begins with `p:nvGrpSpPr` then `p:grpSpPr`. */
export function v015ShapeTreePrologue(ctx: Context): void {
  forEachElement(ctx, (part, element) => {
    if (element.local !== 'spTree' || namespaceOf(element) !== NS.p) return;
    const children = childElements(element);
    if (children[0]?.local !== 'nvGrpSpPr' || children[1]?.local !== 'grpSpPr') {
      ctx.add(
        'V015',
        elementLocation(part, element),
        'a shape tree begins with <nvGrpSpPr> then <grpSpPr>, in that order, before any shape. ' +
          'This one begins ' +
          (children.length === 0
            ? 'with nothing'
            : children
                .slice(0, 2)
                .map((child) => '<' + child.qname + '>')
                .join(' then ')) +
          '. A shape tree is a group shape, and CT_GroupShape requires both even when the slide ' +
          'is empty.',
      );
    }
  });
}

/** Every text body has `a:bodyPr` and at least one `a:p`. */
export function v016TextBody(ctx: Context): void {
  forEachElement(ctx, (part, element) => {
    if (element.local !== 'txBody') return;
    const children = childElements(element);
    if (!children.some((child) => child.local === 'bodyPr')) {
      ctx.add(
        'V016',
        elementLocation(part, element),
        '<' + element.qname + '> has no <a:bodyPr>. CT_TextBody requires it first.',
      );
    }
    if (!children.some((child) => child.local === 'p')) {
      ctx.add(
        'V016',
        elementLocation(part, element),
        '<' +
          element.qname +
          '> has no <a:p>. CT_TextBody is [1..unbounded] paragraphs, and an empty paragraph is ' +
          'not the same thing as no paragraph - a:endParaRPr is where an empty one gets its ' +
          'height. This is what deleting the last paragraph in an editor produces.',
      );
    }
  });
}

/** `p:graphicFrame` carries `p:xfrm` and `a:graphic`. */
export function v017GraphicFrame(ctx: Context): void {
  forEachElement(ctx, (part, element) => {
    if (element.local !== 'graphicFrame' || namespaceOf(element) !== NS.p) return;

    const xfrm = childElements(element).find((child) => child.local === 'xfrm');
    if (xfrm === undefined) {
      ctx.add(
        'V017',
        elementLocation(part, element),
        'no <p:xfrm>. Note the namespace: a graphic frame carries the PresentationML transform, ' +
          'not the a:xfrm every other shape uses. It has no placeholder-inherited geometry path ' +
          'either, which is why a rebound chart or table gets a copy of its layout transform ' +
          'rather than having its own deleted.',
      );
    } else if (namespaceOf(xfrm) !== NS.p) {
      ctx.add(
        'V017',
        elementLocation(part, xfrm),
        'the transform here is <' +
          xfrm.qname +
          '>, in the DrawingML namespace. CT_GraphicalObjectFrame takes p:xfrm; a:xfrm in its ' +
          'place is the same element name bound to the wrong schema.',
      );
    }

    // By local name and namespace, never by qname: `a:` is the conventional
    // prefix for DrawingML and not a guaranteed one, and a part that binds it
    // differently would silently stop being checked.
    const graphic = childElements(element).find(
      (child) => child.local === 'graphic' && namespaceOf(child) === NS.a,
    );
    if (graphic === undefined) {
      ctx.add(
        'V017',
        elementLocation(part, element),
        'no <a:graphic>. The frame is the wrapper; the graphic is the thing in it, and a frame ' +
          'without one names no table, chart, diagram or OLE object at all.',
      );
    }
  });
}

/**
 * Presentation-family parts only.
 *
 * A `.pptx` carries whole foreign documents inside it - an embedded workbook
 * behind every chart, a Word document behind some OLE objects - and their parts
 * are XML with elements called `txBody` and `graphicFrame` in namespaces these
 * rules do not govern. Checking them would be applying PresentationML's
 * cardinalities to SpreadsheetML's schema, which is why every rule above tests
 * the namespace and not only the local name, and why the walk itself stops at
 * parts whose content type belongs to another format.
 */
const FOREIGN_CONTENT_TYPES = new Set<string>([
  CONTENT_TYPE.spreadsheet,
  'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml',
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml',
]);

function forEachElement(ctx: Context, visit: (part: string, element: XElement) => void): void {
  for (const part of ctx.parts()) {
    if (isRelationshipPartName(part)) continue;
    const contentType = ctx.contentType(part);
    if (contentType !== undefined && FOREIGN_CONTENT_TYPES.has(contentType)) continue;
    const document = ctx.document(part);
    if (document === null) continue;
    for (const element of descendantElements(document.root)) visit(part, element);
  }
}

/** Exported for `refusal-rules.ts`, which walks the same parts for the same reason. */
export { forEachElement, FOREIGN_CONTENT_TYPES };
