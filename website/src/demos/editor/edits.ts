/**
 * The editor's four gestures as XML edits on the slide's own part, planned in
 * batches so a later batch can see what the earlier ones did.
 *
 * Each writes what PowerPoint writes for the same gesture: a dragged
 * placeholder gets an `a:xfrm` of its own, a recoloured shape an `a:solidFill`
 * in its `p:spPr`, and the layout keeps everything else. `insertInOrder` puts
 * every new element where the schema does. ADR 0055.
 */

import {
  applyEdits,
  childElements,
  insertInOrder,
  newAttribute,
  newElement,
  newText,
  prefixFor,
  type XElement,
  type XNode,
  type XmlEdit,
} from '@pptx-studio/xml';
import type { Frame } from '@pptx-studio/render-svg';

import {
  NS_A,
  NS_P,
  attributeOf,
  child,
  childScale,
  contentElements,
  fillElement,
  is,
  paragraphElements,
  segmentsOf,
  shapeProperties,
  textBody,
  textNodeOf,
  transformElement,
} from './locate.ts';

/** A gesture this shape cannot take, in words a visitor can act on. */
export class NotEditable extends Error {}

/** Batches of edits; each batch is applied before the next is planned. */
export type Plan = Generator<readonly XmlEdit[], void, undefined>;

const emu = (value: number): string => String(Math.round(value));

const qn = (prefix: string, local: string): string =>
  prefix === '' ? local : `${prefix}:${local}`;

function drawingPrefix(near: XElement): string {
  const prefix = prefixFor(near, NS_A);
  if (prefix === undefined) {
    throw new NotEditable(
      'This slide binds no prefix for DrawingML, so nothing can be added to it.',
    );
  }
  return prefix;
}

/** A copy that came from nowhere; comments and instructions inside formatting are not carried. */
function cloneOf(element: XElement, qname = element.qname): XElement {
  const children: XNode[] = [];
  for (const one of element.children) {
    if (one.type === 'element') children.push(cloneOf(one));
    else if (one.type === 'text' || one.type === 'cdata') children.push(newText(one.value));
  }
  return newElement(
    qname,
    element.attributes.map((attr) => newAttribute(attr.qname, attr.value, attr.quote)),
    children,
  );
}

/** Siblings inserted in order, at the position the schema gives the first of them. */
function appended(parent: XElement, nodes: readonly XElement[]): XmlEdit[] {
  const [head, ...rest] = nodes;
  if (head === undefined) return [];
  const first = insertInOrder(parent, head);
  return [
    first,
    ...rest.map((node, i) => ({
      kind: 'insertChild' as const,
      parent,
      index: first.index + i + 1,
      node,
    })),
  ];
}

function newTransform(a: string, frame: Frame, dx: number, dy: number): XElement {
  const attributes = [];
  if (frame.rot !== 0) attributes.push(newAttribute('rot', String(Math.round(frame.rot * 60000))));
  if (frame.flipH) attributes.push(newAttribute('flipH', '1'));
  if (frame.flipV) attributes.push(newAttribute('flipV', '1'));
  return newElement(qn(a, 'xfrm'), attributes, [
    newElement(qn(a, 'off'), [
      newAttribute('x', emu(frame.x + dx)),
      newAttribute('y', emu(frame.y + dy)),
    ]),
    newElement(qn(a, 'ext'), [
      newAttribute('cx', emu(frame.cx)),
      newAttribute('cy', emu(frame.cy)),
    ]),
  ]);
}

/**
 * Move by a slide-EMU delta. A shape that inherits its transform is given one
 * at the frame it resolved to, which is what PowerPoint writes when a
 * placeholder is dragged; a group's child moves in the group's own units.
 */
export function* planMove(shape: XElement, frame: Frame, dx: number, dy: number): Plan {
  const xfrm = transformElement(shape);
  if (xfrm === undefined) {
    const properties = shapeProperties(shape);
    if (properties === undefined)
      throw new NotEditable('This frame declares no transform to move.');
    yield [insertInOrder(properties, newTransform(drawingPrefix(properties), frame, dx, dy))];
    return;
  }
  const off = child(xfrm, NS_A, 'off');
  if (off === undefined) {
    const a = drawingPrefix(xfrm);
    yield [
      insertInOrder(
        xfrm,
        newElement(qn(a, 'off'), [
          newAttribute('x', emu(frame.x + dx)),
          newAttribute('y', emu(frame.y + dy)),
        ]),
      ),
    ];
    return;
  }
  const { sx, sy } = childScale(shape);
  const x = Number(attributeOf(off, 'x') ?? '0');
  const y = Number(attributeOf(off, 'y') ?? '0');
  yield [
    { kind: 'setAttribute', element: off, qname: 'x', value: emu(x + dx / sx) },
    { kind: 'setAttribute', element: off, qname: 'y', value: emu(y + dy / sy) },
  ];
}

/** The `a:alpha` of a colour being replaced, so a transparent shape stays transparent. */
function alphaOf(fill: XElement | undefined): XElement | undefined {
  const color = fill === undefined ? undefined : childElements(fill)[0];
  return color === undefined ? undefined : child(color, NS_A, 'alpha');
}

function replaceFill(a: string, parent: XElement, value: string): readonly XmlEdit[] {
  const existing = fillElement(parent);
  if (existing !== undefined && is(existing, NS_A, 'solidFill')) {
    const color = childElements(existing)[0];
    if (color !== undefined && is(color, NS_A, 'srgbClr')) {
      return [{ kind: 'setAttribute', element: color, qname: 'val', value }];
    }
  }
  const alpha = alphaOf(existing);
  const fill = newElement(
    qn(a, 'solidFill'),
    [],
    [
      newElement(
        qn(a, 'srgbClr'),
        [newAttribute('val', value)],
        alpha === undefined ? [] : [cloneOf(alpha)],
      ),
    ],
  );
  // Planned before the removal: its index is read from the tree as it stands, and
  // the removal is addressed by node.
  const insert = insertInOrder(parent, fill);
  return existing === undefined
    ? [insert]
    : [insert, { kind: 'removeChild', parent, node: existing }];
}

/**
 * Paint a shape one colour: its own `a:solidFill`, set or written in `p:spPr`
 * in place of whichever fill it declared. A connector's colour is its line's.
 */
export function* planRecolour(shape: XElement, hex: string): Plan {
  const value = hex.replace('#', '').toUpperCase();
  if (!/^[0-9A-F]{6}$/.test(value)) throw new NotEditable(`${hex} is not a six-digit hex colour.`);
  if (is(shape, NS_P, 'pic'))
    throw new NotEditable('A picture keeps its image; it has no fill to colour.');
  if (is(shape, NS_P, 'grpSp')) {
    throw new NotEditable('A group has no colour of its own; colour the shapes inside it.');
  }
  const properties = shapeProperties(shape);
  if (properties === undefined) {
    throw new NotEditable(
      'A chart, table or diagram is preserved as it is; its content is not edited here.',
    );
  }
  const a = drawingPrefix(properties);
  if (is(shape, NS_P, 'cxnSp')) {
    const line = child(properties, NS_A, 'ln');
    if (line === undefined) {
      const fill = newElement(
        qn(a, 'solidFill'),
        [],
        [newElement(qn(a, 'srgbClr'), [newAttribute('val', value)])],
      );
      yield [insertInOrder(properties, newElement(qn(a, 'ln'), [], [fill]))];
      return;
    }
    yield replaceFill(a, line, value);
    return;
  }
  yield replaceFill(a, properties, value);
}

const same = (left: readonly string[], right: readonly string[]): boolean =>
  left.length === right.length && left.every((one, i) => one === right[i]);

/** Run properties for a paragraph that has no run: its `a:endParaRPr`, renamed. */
function runPropsFrom(a: string, paragraph: XElement): XElement | undefined {
  const end = child(paragraph, NS_A, 'endParaRPr');
  return end === undefined ? undefined : cloneOf(end, qn(a, 'rPr'));
}

function newRun(a: string, rPr: XElement | undefined, text: string): XElement {
  return newElement(
    qn(a, 'r'),
    [],
    [...(rPr === undefined ? [] : [cloneOf(rPr)]), newElement(qn(a, 't'), [], [newText(text)])],
  );
}

/** Runs and soft breaks for one paragraph's segments, all in one run's properties. */
function runsFor(
  a: string,
  rPr: XElement | undefined,
  segments: readonly string[],
  from: number,
): XElement[] {
  const nodes: XElement[] = [];
  segments.forEach((segment, index) => {
    if (index < from) return;
    if (index > 0) nodes.push(newElement(qn(a, 'br'), [], rPr === undefined ? [] : [cloneOf(rPr)]));
    nodes.push(newRun(a, rPr, segment));
  });
  return nodes;
}

function newParagraph(a: string, template: XElement, segments: readonly string[]): XElement {
  const pPr = child(template, NS_A, 'pPr');
  const firstRun = contentElements(template).find((one) => one.local === 'r');
  const rPr = firstRun === undefined ? runPropsFrom(a, template) : child(firstRun, NS_A, 'rPr');
  const end = child(template, NS_A, 'endParaRPr');
  return newElement(
    qn(a, 'p'),
    [],
    [
      ...(pPr === undefined ? [] : [cloneOf(pPr)]),
      ...runsFor(a, rPr, segments, 0),
      ...(end === undefined ? [] : [cloneOf(end)]),
    ],
  );
}

interface Rewrite {
  readonly paragraph: XElement;
  readonly segments: readonly string[];
  /** The run kept, whose `a:t` took the first segment; undefined when there was none. */
  readonly kept: XElement | undefined;
}

/**
 * Replace a shape's text: one entry per paragraph, split at its soft breaks.
 *
 * A paragraph whose text is unchanged is not touched. A changed one keeps its
 * `a:pPr` and its first run's `a:rPr`; the rest of its runs collapse into that
 * one, as PowerPoint does when a selection spanning them is typed over. New
 * paragraphs copy the last one's properties; surplus ones go.
 */
export function* planRetext(shape: XElement, wanted: readonly (readonly string[])[]): Plan {
  const body = textBody(shape);
  if (body === undefined) throw new NotEditable('This shape holds no text.');
  const a = drawingPrefix(body);
  const paragraphs = paragraphElements(body);
  if (paragraphs.length === 0)
    throw new NotEditable('This text body has no paragraph to write into.');
  const lines = wanted.length === 0 ? [['']] : wanted;

  const first: XmlEdit[] = [];
  const rewrites: Rewrite[] = [];
  paragraphs.forEach((paragraph, i) => {
    const segments = lines[i];
    if (segments === undefined) {
      first.push({ kind: 'removeChild', parent: body, node: paragraph });
      return;
    }
    if (same(segmentsOf(paragraph), segments)) return;
    const content = contentElements(paragraph);
    const kept = content.find((one) => one.local === 'r');
    for (const one of content) {
      if (one !== kept) first.push({ kind: 'removeChild', parent: paragraph, node: one });
    }
    if (kept !== undefined) {
      const t = child(kept, NS_A, 't');
      if (t === undefined) throw new NotEditable('A run without a:t cannot be retyped.');
      for (const one of t.children) {
        if (one.type === 'cdata') first.push({ kind: 'removeChild', parent: t, node: one });
      }
      const node = textNodeOf(t);
      const value = segments[0] ?? '';
      first.push(
        node === undefined
          ? { kind: 'insertChild', parent: t, index: 0, node: newText(value) }
          : { kind: 'setValue', node, value },
      );
    }
    rewrites.push({ paragraph, segments, kept });
  });
  yield first;

  const second: XmlEdit[] = [];
  for (const { paragraph, segments, kept } of rewrites) {
    const rPr = kept === undefined ? runPropsFrom(a, paragraph) : child(kept, NS_A, 'rPr');
    second.push(...appended(paragraph, runsFor(a, rPr, segments, kept === undefined ? 0 : 1)));
  }
  const template = paragraphs[Math.min(paragraphs.length, lines.length) - 1];
  if (template !== undefined) {
    second.push(
      ...appended(
        body,
        lines.slice(paragraphs.length).map((segments) => newParagraph(a, template, segments)),
      ),
    );
  }
  if (second.length > 0) yield second;
}

/** Take the shape out of its tree; the inverse puts it back at the same index. */
export function* planDelete(shape: XElement): Plan {
  const parent = shape.parent;
  if (parent === undefined) throw new NotEditable('This shape has no parent to be removed from.');
  yield [{ kind: 'removeChild', parent, node: shape }];
}

/**
 * Apply a plan batch by batch and return the inverses in the order that undoes
 * it. A refused batch rolls the earlier ones back before rethrowing, so a
 * refused gesture leaves the tree as it found it.
 */
export function applyAll(plan: Plan): XmlEdit[] {
  const inverses: XmlEdit[] = [];
  try {
    for (const batch of plan) {
      if (batch.length > 0) inverses.unshift(...applyEdits(batch));
    }
  } catch (cause) {
    applyEdits(inverses);
    throw cause;
  }
  return inverses;
}
