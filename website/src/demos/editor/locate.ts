/**
 * Finding the pieces of a shape the editor writes to.
 *
 * Everything is matched by namespace URI and local name, never by the prefix in
 * the file: `p:` and `a:` are conventions PowerPoint happens to use and a
 * conforming writer may spell them anything.
 */

import {
  childElements,
  descendantElements,
  namespaceOf,
  textContent,
  type XElement,
  type XText,
} from '@pptx-studio/xml';

export const NS_P = 'http://schemas.openxmlformats.org/presentationml/2006/main';
export const NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main';

export function is(element: XElement, uri: string, local: string): boolean {
  return element.local === local && namespaceOf(element) === uri;
}

export function child(parent: XElement, uri: string, local: string): XElement | undefined {
  return childElements(parent).find((one) => is(one, uri, local));
}

export function attributeOf(element: XElement, qname: string): string | undefined {
  return element.attributes.find((attr) => attr.qname === qname)?.value;
}

/**
 * The `p:sp` / `p:pic` / `p:grpSp` / `p:graphicFrame` / `p:cxnSp` whose
 * `p:cNvPr/@id` is this one: two hops up from the id, whichever kind it is.
 */
export function shapeElement(root: XElement, cNvPrId: number): XElement | undefined {
  for (const element of descendantElements(root)) {
    if (!is(element, NS_P, 'cNvPr')) continue;
    const id = attributeOf(element, 'id');
    if (id === undefined || Number(id) !== cNvPrId) continue;
    return element.parent?.parent;
  }
  return undefined;
}

/** `p:cNvPr/@name`, or the id when the name is empty. */
export function nameOf(shape: XElement): string {
  for (const element of descendantElements(shape)) {
    if (!is(element, NS_P, 'cNvPr')) continue;
    const name = attributeOf(element, 'name');
    return name === undefined || name === '' ? `shape #${attributeOf(element, 'id') ?? '?'}` : name;
  }
  return 'the shape';
}

/** `p:spPr` or `p:grpSpPr`, where geometry and paint live. A graphic frame has neither. */
export function shapeProperties(shape: XElement): XElement | undefined {
  return child(shape, NS_P, 'spPr') ?? child(shape, NS_P, 'grpSpPr');
}

/** The `a:xfrm` a shape declares, or a graphic frame's own `p:xfrm`. */
export function transformElement(shape: XElement): XElement | undefined {
  const properties = shapeProperties(shape);
  return properties === undefined ? child(shape, NS_P, 'xfrm') : child(properties, NS_A, 'xfrm');
}

/**
 * The scale between a shape's `a:off` and slide EMU: the product of every
 * enclosing group's `ext / chExt`, per axis, with a zero read as one.
 */
export function childScale(shape: XElement): { readonly sx: number; readonly sy: number } {
  let sx = 1;
  let sy = 1;
  for (let group = shape.parent; group !== undefined; group = group.parent) {
    if (!is(group, NS_P, 'grpSp')) continue;
    const xfrm = transformElement(group);
    const ext = xfrm === undefined ? undefined : child(xfrm, NS_A, 'ext');
    const chExt = xfrm === undefined ? undefined : child(xfrm, NS_A, 'chExt');
    if (ext === undefined || chExt === undefined) continue;
    const ratio = (name: string): number => {
      const outer = Number(attributeOf(ext, `c${name}`) ?? '0');
      const inner = Number(attributeOf(chExt, `c${name}`) ?? '0');
      return outer === 0 || inner === 0 ? 1 : outer / inner;
    };
    sx *= ratio('x');
    sy *= ratio('y');
  }
  return { sx, sy };
}

const FILL_KINDS: ReadonlySet<string> = new Set([
  'noFill',
  'solidFill',
  'gradFill',
  'blipFill',
  'pattFill',
  'grpFill',
]);

/** The fill an element declares, whichever kind, or undefined when it inherits one. */
export function fillElement(parent: XElement): XElement | undefined {
  return childElements(parent).find(
    (one) => namespaceOf(one) === NS_A && FILL_KINDS.has(one.local),
  );
}

const GRAPHIC_KINDS: Readonly<Record<string, string>> = {
  'http://schemas.openxmlformats.org/drawingml/2006/chart': 'chart',
  'http://schemas.openxmlformats.org/drawingml/2006/table': 'table',
  'http://schemas.openxmlformats.org/drawingml/2006/diagram': 'SmartArt',
  'http://schemas.openxmlformats.org/presentationml/2006/ole': 'OLE object',
};

/** What a graphic frame holds, from its `a:graphicData/@uri`. */
export function graphicKind(frame: XElement): string {
  for (const element of descendantElements(frame)) {
    if (!is(element, NS_A, 'graphicData')) continue;
    const uri = attributeOf(element, 'uri');
    return uri === undefined ? 'graphic' : (GRAPHIC_KINDS[uri] ?? 'graphic');
  }
  return 'graphic';
}

export function textBody(shape: XElement): XElement | undefined {
  return child(shape, NS_P, 'txBody');
}

export function paragraphElements(body: XElement): XElement[] {
  return childElements(body).filter((one) => is(one, NS_A, 'p'));
}

/** `a:r`, `a:br` and `a:fld`, in order: what a paragraph is made of besides its properties. */
export function contentElements(paragraph: XElement): XElement[] {
  return childElements(paragraph).filter(
    (one) =>
      namespaceOf(one) === NS_A && (one.local === 'r' || one.local === 'br' || one.local === 'fld'),
  );
}

/** The text node of an `a:t`, which may be an empty element. */
export function textNodeOf(t: XElement): XText | undefined {
  return t.children.find((one): one is XText => one.type === 'text');
}

/** A paragraph's text between its `a:br`s; a field contributes the text PowerPoint cached for it. */
export function segmentsOf(paragraph: XElement): string[] {
  const segments = [''];
  for (const one of contentElements(paragraph)) {
    if (one.local === 'br') {
      segments.push('');
      continue;
    }
    const t = child(one, NS_A, 't');
    segments[segments.length - 1] += t === undefined ? '' : textContent(t);
  }
  return segments;
}
