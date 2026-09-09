/**
 * Finding the pieces of a shape that the three demo edits write to.
 *
 * Everything is matched by namespace URI and local name, never by the prefix in
 * the file: `p:` and `a:` are conventions PowerPoint happens to use and a
 * conforming writer may spell them anything.
 */

import {
  childElements,
  descendantElements,
  namespaceOf,
  type XElement,
  type XText,
} from '@pptx-studio/xml';

export const NS_P = 'http://schemas.openxmlformats.org/presentationml/2006/main';
export const NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main';

function named(parent: XElement, uri: string, local: string): XElement | undefined {
  return childElements(parent).find((child) => child.local === local && namespaceOf(child) === uri);
}

/**
 * The `p:sp` / `p:pic` / `p:grpSp` whose `p:cNvPr/@id` is this one.
 *
 * `cNvPr` sits inside `p:nvSpPr` (or `nvPicPr`, or `nvGrpSpPr`), whose parent
 * is the shape - so it is two hops up from the id, whichever kind it is.
 */
export function shapeElement(root: XElement, cNvPrId: number): XElement | undefined {
  for (const element of descendantElements(root)) {
    if (element.local !== 'cNvPr' || namespaceOf(element) !== NS_P) continue;
    const id = element.attributes.find((attr) => attr.qname === 'id')?.value;
    if (id === undefined || Number(id) !== cNvPrId) continue;
    return element.parent?.parent;
  }
  return undefined;
}

/** `p:spPr`, where geometry and paint live on any of the shape kinds. */
export function shapeProperties(shape: XElement): XElement | undefined {
  return named(shape, NS_P, 'spPr') ?? named(shape, NS_P, 'grpSpPr');
}

/** `a:off` inside the shape's own `a:xfrm`, or undefined when it inherits one. */
export function offsetElement(shape: XElement): XElement | undefined {
  const properties = shapeProperties(shape);
  if (properties === undefined) return undefined;
  const xfrm = named(properties, NS_A, 'xfrm');
  return xfrm === undefined ? undefined : named(xfrm, NS_A, 'off');
}

export interface SolidFill {
  readonly fill: XElement;
  /** The colour element, either `a:srgbClr` or one of the indirect kinds. */
  readonly color: XElement;
  readonly index: number;
}

/** The shape's own `a:solidFill`, when it declares one rather than inheriting. */
export function solidFill(shape: XElement): SolidFill | undefined {
  const properties = shapeProperties(shape);
  if (properties === undefined) return undefined;
  const fill = named(properties, NS_A, 'solidFill');
  if (fill === undefined) return undefined;
  const color = childElements(fill)[0];
  if (color === undefined) return undefined;
  return { fill, color, index: fill.children.indexOf(color) };
}

export interface RunText {
  readonly element: XElement;
  readonly node: XText | undefined;
}

/** The first `a:t` in the shape's text body, and its text node if it has one. */
export function firstRunText(shape: XElement): RunText | undefined {
  const body = named(shape, NS_P, 'txBody');
  if (body === undefined) return undefined;
  for (const element of descendantElements(body)) {
    if (element.local !== 't' || namespaceOf(element) !== NS_A) continue;
    return {
      element,
      node: element.children.find((child): child is XText => child.type === 'text'),
    };
  }
  return undefined;
}
