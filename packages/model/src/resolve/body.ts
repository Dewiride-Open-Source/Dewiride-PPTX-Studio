/**
 * `a:bodyPr`, resolved up the placeholder chain.
 *
 * Every attribute inherits on its own, from any level: a bare `<a:bodyPr/>` on a
 * slide is anchored where its master says, not at the schema default. ADR 0032.
 */

import { inheritanceChain } from './placeholder.js';
import type { Resolved, Shape, Sheet } from '../types.js';
import type { BodyProps, TextAnchor, TextWrap, VerticalText } from '../text.js';

const EMU_PER_POINT = 12700;

/**
 * The four insets in points, per attribute.
 *
 * Restated here rather than imported from `@pptx-studio/text`, which is a layer
 * below and which this package does not depend on - the same posture
 * `line-model.ts` takes towards `Spacing`.
 */
export interface ResolvedInsets {
  readonly left: number;
  readonly top: number;
  readonly right: number;
  readonly bottom: number;
}

/** Points. Asymmetric, and each edge defaults on its own. */
const DEFAULT_INSETS: ResolvedInsets = { left: 7.2, top: 3.6, right: 7.2, bottom: 3.6 };

/**
 * One property of the text frame, from the shape's own `a:bodyPr` or above it.
 *
 * `pick` returns `undefined` for "this level did not say", the same contract the
 * text cascade uses.
 */
export function resolveBody<T>(
  shape: Shape,
  sheet: Sheet,
  pick: (props: BodyProps) => T | undefined,
): Resolved<T> | undefined {
  for (const [index, link] of inheritanceChain(shape, sheet).entries()) {
    const props = link.shape.text?.bodyPr;
    if (props === undefined) continue;
    const value = pick(props);
    if (value === undefined) continue;
    return {
      value,
      origin: index === 0 ? 'shape' : link.origin,
      explicit: index === 0,
      sheet: link.sheet,
      shape: link.shape,
    };
  }
  return undefined;
}

export function resolveAnchor(shape: Shape, sheet: Sheet): Resolved<TextAnchor> | undefined {
  return resolveBody(shape, sheet, (p) => p.anchor);
}

export function resolveAnchorCtr(shape: Shape, sheet: Sheet): Resolved<boolean> | undefined {
  return resolveBody(shape, sheet, (p) => p.anchorCtr);
}

export function resolveVertical(shape: Shape, sheet: Sheet): Resolved<VerticalText> | undefined {
  return resolveBody(shape, sheet, (p) => p.vert);
}

export function resolveWrap(shape: Shape, sheet: Sheet): Resolved<TextWrap> | undefined {
  return resolveBody(shape, sheet, (p) => p.wrap);
}

export function resolveColumns(shape: Shape, sheet: Sheet): Resolved<number> | undefined {
  return resolveBody(shape, sheet, (p) => p.numCol);
}

/**
 * The four insets, in points, each falling back on its own.
 *
 * A frame stating only `lIns` still gets 3.6 above and below, so this resolves
 * four times rather than once.
 */
export function resolveInsets(shape: Shape, sheet: Sheet): ResolvedInsets {
  const at = (pick: (p: BodyProps) => number | undefined, fallback: number): number => {
    const found = resolveBody(shape, sheet, pick);
    return found === undefined ? fallback : found.value / EMU_PER_POINT;
  };
  return {
    left: at((p) => p.lIns, DEFAULT_INSETS.left),
    top: at((p) => p.tIns, DEFAULT_INSETS.top),
    right: at((p) => p.rIns, DEFAULT_INSETS.right),
    bottom: at((p) => p.bIns, DEFAULT_INSETS.bottom),
  };
}
