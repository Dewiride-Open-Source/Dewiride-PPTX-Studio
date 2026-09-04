/**
 * The style matrix, the 1000-offset, and `phClr`.
 *
 * A `p:style` holds four references into the theme. Each names an entry by
 * index and carries the colour that entry's `schemeClr val="phClr"` is to be
 * resolved with - so an entry such as
 * `<a:solidFill><a:schemeClr val="phClr"><a:lumMod val="60000"/></a:schemeClr></a:solidFill>`
 * is a *function* of a colour, and the reference is the call.
 *
 * Miss `phClr` and every themed shape renders black, which is the failure this
 * whole file exists to prevent.
 */

import {
  resolveColor,
  type Color,
  type Effect,
  type Fill,
  type Line,
  type Rgba,
} from '@pptx-studio/paint';

import { ModelError } from './errors.js';
import { colorContextOf, resolve, themeOf } from './resolve.js';
import type { Origin, Resolved, Shape, Sheet, StyleRef, Theme } from './types.js';

/**
 * The offset that separates the two style lists.
 *
 * `0` and `1000` mean *nothing at all*; `1..999` index the main list; `1001` and
 * up index the background list. Measured on both `a:fillRef` and `p:bgRef`
 * against a theme whose six entries paint six distinguishable colours - all six
 * agreed between the two elements, and `idx="0"` and `idx="1000"` both painted
 * nothing.
 */
export const STYLE_MATRIX_OFFSET = 1000;

/** Which of the two lists an index reaches, and where in it. */
export type StyleMatrixTarget =
  | { readonly list: 'none' }
  | { readonly list: 'main'; readonly at: number }
  | { readonly list: 'background'; readonly at: number };

/**
 * Where an index points, before anything is fetched.
 *
 * Out of range is not an error. A `fillRef idx="4"` against a three-entry list
 * paints the third entry, and a `bgRef idx="9999"` paints the last background
 * entry: PowerPoint clamps, opens the file without repairing it, and a renderer
 * that throws here refuses a deck PowerPoint shows.
 */
export function styleMatrixTarget(idx: number): StyleMatrixTarget {
  if (idx === 0 || idx === STYLE_MATRIX_OFFSET) return { list: 'none' };
  if (idx < STYLE_MATRIX_OFFSET) return { list: 'main', at: idx - 1 };
  return { list: 'background', at: idx - STYLE_MATRIX_OFFSET - 1 };
}

function clampInto<T>(entries: readonly T[], at: number, what: string, theme: Theme): T {
  if (entries.length === 0) {
    throw new ModelError('MODEL_STYLE_LIST_EMPTY', `${what} has no entries`, theme.partName, what);
  }
  return entries[Math.min(Math.max(at, 0), entries.length - 1)] as T;
}

/** The fill a `fillRef`/`bgRef` index names, or `null` for none. */
export function styleMatrixFill(ref: StyleRef, theme: Theme): Fill | null {
  const target = styleMatrixTarget(ref.idx);
  if (target.list === 'none') return null;
  const format = theme.format;
  if (format === null) {
    throw new ModelError('MODEL_NO_STYLE_MATRIX', 'the theme has no a:fmtScheme', theme.partName);
  }
  return target.list === 'main'
    ? clampInto(format.fillStyles, target.at, 'a:fillStyleLst', theme)
    : clampInto(format.bgFillStyles, target.at, 'a:bgFillStyleLst', theme);
}

/**
 * The line an `lnRef` index names, or `null` for none.
 *
 * `a:lnStyleLst` has no background counterpart, so an index above 1000 clamps
 * into the same list rather than reaching a second one.
 */
export function styleMatrixLine(ref: StyleRef, theme: Theme): Line | null {
  const target = styleMatrixTarget(ref.idx);
  if (target.list === 'none') return null;
  const format = theme.format;
  if (format === null) {
    throw new ModelError('MODEL_NO_STYLE_MATRIX', 'the theme has no a:fmtScheme', theme.partName);
  }
  return clampInto(format.lineStyles, target.at, 'a:lnStyleLst', theme);
}

/** The effects an `effectRef` index names. Empty for none. */
export function styleMatrixEffects(ref: StyleRef, theme: Theme): readonly Effect[] {
  const target = styleMatrixTarget(ref.idx);
  if (target.list === 'none') return [];
  const format = theme.format;
  if (format === null) {
    throw new ModelError('MODEL_NO_STYLE_MATRIX', 'the theme has no a:fmtScheme', theme.partName);
  }
  return clampInto(format.effectStyles, target.at, 'a:effectStyleLst', theme);
}

/**
 * The colour a style reference invokes its entry with.
 *
 * The reference's own colour child, resolved against the sheet - and it may
 * carry transforms of its own. PowerPoint's shape-style gallery writes
 * `<a:schemeClr val="accent1"><a:shade val="15000"/></a:schemeClr>` on seven of
 * its forty-two entries, so `phClr` is a fully transformed colour and not a
 * theme slot name.
 */
export function phClrOf(ref: StyleRef, sheet: Sheet): Rgba | null {
  if (ref.color === null) return null;
  return resolveColor(ref.color, colorContextOf(sheet));
}

/* -------------------------------------------------------------------------- */
/* the whole answer for one shape                                             */
/* -------------------------------------------------------------------------- */

/** What a shape paints, and where each half of the answer came from. */
export interface ResolvedAppearance {
  /** `null` means nothing is painted: no fill was declared and no style names one. */
  readonly fill: Fill | null;
  readonly fillOrigin: Origin | null;
  /** The colour `phClr` resolves to inside `fill`, when it came from the theme. */
  readonly fillPhClr: Rgba | null;
  readonly line: Line | null;
  readonly lineOrigin: Origin | null;
  readonly linePhClr: Rgba | null;
  readonly effects: readonly Effect[] | null;
}

function styleOf(shape: Shape, sheet: Sheet): Resolved<Shape['style']> | undefined {
  return resolve(shape, sheet, (s) => s.style);
}

/**
 * What a shape paints, following both routes and in the right order.
 *
 * An explicit `spPr` fill beats the `fillRef`, and an explicit `a:noFill` beats
 * it too - both measured. A shape with neither paints nothing rather than black.
 *
 * The two routes are searched independently: the `spPr` fill is looked for down
 * the whole inheritance chain first, and only if no level declared one does the
 * `p:style` - itself inherited down the same chain - get consulted.
 */
export function resolveAppearance(shape: Shape, sheet: Sheet): ResolvedAppearance {
  const theme = themeOf(sheet);
  const style = styleOf(shape, sheet);

  const directFill = resolve(shape, sheet, (s) => s.fill);
  const directLine = resolve(shape, sheet, (s) => s.line);
  const directEffects = resolve(shape, sheet, (s) => s.effects);

  let fill: Fill | null = directFill?.value ?? null;
  let fillOrigin: ResolvedAppearance['fillOrigin'] = directFill?.origin ?? null;
  let fillPhClr: Rgba | null = null;
  if (directFill === undefined && style !== undefined && theme !== null) {
    const ref = style.value?.fillRef;
    if (ref !== undefined) {
      fill = styleMatrixFill(ref, theme);
      fillOrigin = fill === null ? null : 'theme';
      fillPhClr = fill === null ? null : phClrOf(ref, sheet);
    }
  }

  let line: Line | null = directLine?.value ?? null;
  let lineOrigin: ResolvedAppearance['lineOrigin'] = directLine?.origin ?? null;
  let linePhClr: Rgba | null = null;
  if (directLine === undefined && style !== undefined && theme !== null) {
    const ref = style.value?.lnRef;
    if (ref !== undefined) {
      line = styleMatrixLine(ref, theme);
      lineOrigin = line === null ? null : 'theme';
      linePhClr = line === null ? null : phClrOf(ref, sheet);
    }
  }

  let effects: readonly Effect[] | null = directEffects?.value ?? null;
  if (directEffects === undefined && style !== undefined && theme !== null) {
    const ref = style.value?.effectRef;
    if (ref !== undefined) effects = styleMatrixEffects(ref, theme);
  }

  return { fill, fillOrigin, fillPhClr, line, lineOrigin, linePhClr, effects };
}

/**
 * A shape's fill as a single colour, when it is one.
 *
 * A convenience over `resolveAppearance` for the common case, and the shape the
 * ground-truth fixture is in: PowerPoint's object model reports one RGB per
 * shape, so this is what a measurement can be compared against.
 */
export function resolveSolidFill(shape: Shape, sheet: Sheet): Rgba | null {
  const appearance = resolveAppearance(shape, sheet);
  if (appearance.fill === null || appearance.fill.type !== 'solid') return null;
  const phClr = appearance.fillPhClr;
  return resolveColor(
    appearance.fill.color,
    colorContextOf(sheet, phClr === null ? undefined : phClr),
  );
}

/** Resolve any colour written on a sheet, with an optional `phClr` in scope. */
export function resolveOnSheet(color: Color, sheet: Sheet, phClr?: Rgba): Rgba {
  return resolveColor(color, colorContextOf(sheet, phClr));
}
