/**
 * Three edits, each planned against the slide's own XML and each handing back
 * its inverse.
 *
 * This is not Phase 5's command bus and does not pretend to be: there is no
 * coalescing, no invalidation set and no selection restore. What it does show is
 * the property those all rest on - an edit is an operation on the tree that
 * knows how to undo itself, so undo restores the document rather than our
 * understanding of it.
 */

import {
  applyEdit,
  newAttribute,
  newElement,
  newText,
  type XElement,
  type XmlEdit,
} from '@pptx-studio/xml';

import { firstRunText, offsetElement, solidFill } from './locate';

/** Why an edit could not be planned, in words a user can act on. */
export class NotEditable extends Error {}

const attributeOf = (element: XElement, qname: string): string | undefined =>
  element.attributes.find((attr) => attr.qname === qname)?.value;

/**
 * Move a shape by writing `a:off`.
 *
 * Refuses a shape that has no `a:xfrm` of its own, and the refusal is the
 * interesting half: a placeholder PowerPoint wrote has no geometry, because it
 * takes the layout's. Giving it one is what Change Layout has to undo later, so
 * it is a decision and not a drag.
 */
export function planMove(shape: XElement, dx: number, dy: number): readonly XmlEdit[] {
  const off = offsetElement(shape);
  if (off === undefined) {
    throw new NotEditable(
      'This shape declares no a:xfrm - it inherits its position from the layout. ' +
        'Writing one here is a Change Layout decision, not a drag.',
    );
  }
  const x = Number(attributeOf(off, 'x') ?? '0');
  const y = Number(attributeOf(off, 'y') ?? '0');
  return [
    { kind: 'setAttribute', element: off, qname: 'x', value: String(Math.round(x + dx)) },
    { kind: 'setAttribute', element: off, qname: 'y', value: String(Math.round(y + dy)) },
  ];
}

/**
 * Recolour a shape's own `a:solidFill`.
 *
 * A `a:schemeClr` is swapped for an `a:srgbClr`, which is exactly the "detach
 * from the theme" that sub-phase 7.10 does in bulk - and it is two edits rather
 * than one because the tree has no replace.
 */
export function planRecolour(shape: XElement, hex: string): readonly XmlEdit[] {
  const found = solidFill(shape);
  if (found === undefined) {
    throw new NotEditable(
      'This shape declares no a:solidFill - its fill is inherited, from the placeholder ' +
        'chain or the theme style matrix. The Inheritance tool says which.',
    );
  }
  const value = hex.replace('#', '').toUpperCase();
  if (found.color.local === 'srgbClr') {
    return [{ kind: 'setAttribute', element: found.color, qname: 'val', value }];
  }
  const replacement = newElement('a:srgbClr', [newAttribute('val', value)]);
  return [
    { kind: 'removeChild', parent: found.fill, node: found.color },
    { kind: 'insertChild', parent: found.fill, index: found.index, node: replacement },
  ];
}

/** Retype the first run. `a:t` may be an empty element, so a text node is made. */
export function planRetext(shape: XElement, value: string): readonly XmlEdit[] {
  const run = firstRunText(shape);
  if (run === undefined) {
    throw new NotEditable('This shape has no a:t to retype - it holds no text runs.');
  }
  return run.node === undefined
    ? [{ kind: 'insertChild', parent: run.element, index: 0, node: newText(value) }]
    : [{ kind: 'setValue', node: run.node, value }];
}

/** Apply in order and return the inverses in the order that undoes them. */
export function applyAll(edits: readonly XmlEdit[]): readonly XmlEdit[] {
  const inverses = edits.map((edit) => applyEdit(edit));
  return [...inverses].reverse();
}
