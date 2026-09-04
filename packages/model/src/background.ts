/**
 * What is behind everything on a slide.
 *
 * Two questions, and they are independent. *Which sheet's* background applies,
 * and *what* that background is.
 *
 * **Which sheet.** The nearest one that declares a `p:bg` at all: the slide,
 * else its layout, else the master. Measured on all three. An absent `p:bg` is
 * the only thing that inherits - a slide declaring `<p:bgPr><a:noFill/>` paints
 * nothing and does *not* fall through, which is the distinction between "I have
 * no opinion" and "my opinion is nothing", and the reason `Background` is a
 * value rather than a nullable fill.
 *
 * **What it is.** `p:bgPr` states a fill outright. `p:bgRef` names an entry of
 * the theme's style matrix and the colour to invoke its `phClr` with, using the
 * same 1000-offset as `a:fillRef` over the same two lists - measured on six
 * indices through both elements, agreeing on all six. PowerPoint's own master
 * writes `<p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef>`, which is the
 * first entry of `a:bgFillStyleLst` invoked with whatever `bg1` maps to.
 *
 * And `bg1` there is resolved through the colour map in force **on the sheet
 * that owns the shape**, not on the sheet that owns the background: a slide
 * whose `clrMapOvr` remaps `bg1` gets a different background out of the same
 * master. Measured.
 */

import { resolveColor, type Fill, type Rgba } from '@pptx-studio/paint';

import { colorContextOf, sheetChain, themeOf } from './resolve.js';
import { phClrOf, styleMatrixFill } from './style.js';
import type { Background, Sheet } from './types.js';

/** A background, with the sheet that declared it and how it was expressed. */
export interface ResolvedBackground {
  readonly fill: Fill;
  /** The sheet whose `p:bg` this is. */
  readonly sheet: Sheet;
  /** True when the sheet asked for it is the one that declared it. */
  readonly explicit: boolean;
  /** `bgPr` or `bgRef`, because they round-trip differently. */
  readonly form: 'bgPr' | 'bgRef';
  /** The colour `phClr` takes inside `fill`, for a `bgRef`. */
  readonly phClr: Rgba | null;
}

/** Which sheet's `p:bg` applies, without unpacking it. */
export function backgroundSheet(sheet: Sheet): { sheet: Sheet; background: Background } | null {
  for (const link of sheetChain(sheet)) {
    if (link.background !== undefined) return { sheet: link, background: link.background };
  }
  return null;
}

/**
 * The background of a slide, layout or master.
 *
 * `null` when no sheet in the chain declares one, which is a legal file: a deck
 * whose master has no `p:bg` paints nothing behind its slides.
 */
export function resolveBackground(sheet: Sheet): ResolvedBackground | null {
  const found = backgroundSheet(sheet);
  if (found === null) return null;

  if (found.background.kind === 'fill') {
    return {
      fill: found.background.fill,
      sheet: found.sheet,
      explicit: found.sheet === sheet,
      form: 'bgPr',
      phClr: null,
    };
  }

  const theme = themeOf(sheet);
  if (theme === null) return null;
  const fill = styleMatrixFill(found.background.ref, theme);
  if (fill === null) {
    // `idx="0"` and `idx="1000"` both paint nothing at all - not white, and not
    // the sheet above's background. Measured on both.
    return {
      fill: { type: 'none' },
      sheet: found.sheet,
      explicit: found.sheet === sheet,
      form: 'bgRef',
      phClr: null,
    };
  }
  return {
    fill,
    sheet: found.sheet,
    explicit: found.sheet === sheet,
    form: 'bgRef',
    // The map in force on the sheet being asked about, not on the sheet the
    // `bgRef` was written on.
    phClr: phClrOf(found.background.ref, sheet),
  };
}

/** The background as one colour, when it is a solid. `null` otherwise. */
export function resolveBackgroundColor(sheet: Sheet): Rgba | null {
  const background = resolveBackground(sheet);
  if (background === null || background.fill.type !== 'solid') return null;
  return resolveColor(
    background.fill.color,
    colorContextOf(sheet, background.phClr === null ? undefined : background.phClr),
  );
}
