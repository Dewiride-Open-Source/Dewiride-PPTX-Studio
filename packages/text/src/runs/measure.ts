/**
 * Measurement, over `OffscreenCanvas.measureText`, as the single source of truth.
 *
 * The plan's bet is that measuring and rendering with the same engine is what
 * makes autofit and line breaking correct - so nothing here consults a font
 * table, and nothing anywhere else in the package is allowed to measure by a
 * different route.
 *
 * ## How close this gets to PowerPoint
 *
 * T2 asked PowerPoint for the width of seven strings in fourteen faces at four
 * sizes, and asked Chromium for the same. Over 364 comparisons on installed
 * faces the disagreement is a **median of 0.15%**, p95 0.70%. The one large
 * outlier is Cambria at 12pt, where PowerPoint hints small text and Chromium
 * does not; the recorded maximum is 11.1%.
 *
 * That number is the honest ceiling on text fidelity and the plan promises to
 * track it rather than to claim it away. It is in
 * `corpus/ground-truth/text-metrics.json` under `browserAgreement`, and the
 * suite fails if it regresses.
 *
 * A caveat that cost a rewrite to notice: PowerPoint's own
 * `TextRange2.BoundWidth` is **wider than the string**, by exactly the face's
 * space advance, because the range includes the trailing paragraph mark.
 * Compared raw it reports a 5-20% disagreement that is entirely an artefact of
 * the reader. The fixture stores widths with the mark already subtracted.
 *
 * ## Units
 *
 * Everything is points, and the canvas is driven at one CSS pixel per point.
 * The layout space is a fixed 960x540pt with one `transform: scale(z)` for zoom
 * (bet 4), so there is no device-pixel ratio in this file on purpose: re-laying
 * out per zoom would change line breaks, because font metrics are not linear in
 * point size.
 */

import { TextError } from '../errors.js';

/**
 * The properties of a run that change its width.
 *
 * `sz`, `spc` and `kern` are all hundredths of a point, as `a:rPr` writes them.
 * They arrive resolved - the cascade is sub-phase 3.1's job - and `undefined`
 * means the cascade yielded nothing, which is not the same as zero.
 */
export interface RunFont {
  /** `a:latin/@typeface`, already resolved through any `+mj-lt`/`+mn-lt`. */
  readonly family: string;
  /** `a:rPr/@sz`, hundredths of a point. */
  readonly sz: number;
  readonly bold?: boolean | undefined;
  readonly italic?: boolean | undefined;
  /** `a:rPr/@spc`, hundredths of a point. May be negative. */
  readonly spc?: number | undefined;
  /** `a:rPr/@kern`: the minimum font size at which kerning happens. */
  readonly kern?: number | undefined;
}

/**
 * Whether kerning applies to a run.
 *
 * `@kern` is a **minimum font size**, not a flag - ECMA-376 Part 1,
 * `CT_TextCharacterProperties/@kern`: "Specifies the minimum font size at which
 * character kerning occurs for this text run." T2 measured all four readings
 * against Arial, which is the only one of the three probed faces with kern
 * pairs for the test string:
 *
 * | reading                                   | fits  |
 * |-------------------------------------------|-------|
 * | `kern > 0 && sz >= kern`                  | 12/12 |
 * | `kern > 0 && sz > kern`                   | 11/12 |
 * | non-zero means on, at any size            |  8/12 |
 * | present means on                          |  5/12 |
 *
 * The threshold is inclusive - a 12pt run with `kern="1200"` *is* kerned, which
 * is the single probe that separates the top two rows - and `kern="0"` means
 * never, which ECMA does not say.
 *
 * An absent `@kern` follows ECMA and kerns at every size. That path is rarer
 * than it looks: T2 measured an 8pt run in a file stating no `@kern` anywhere
 * and found it *unkerned*, because 3.1's tenth source - PowerPoint's built-in
 * `p:txStyles`, substituted when a master declares none - supplies `kern="1200"`.
 * So `undefined` here means the cascade genuinely produced nothing, and a caller
 * that has not run the cascade will get the wrong answer for the right reason.
 */
export function kerningEnabled(kern: number | undefined, sz: number): boolean {
  if (kern === undefined) return true;
  return kern > 0 && sz >= kern;
}

/**
 * A family name, quoted for a CSS font shorthand.
 *
 * Always quoted, never bare: an unquoted `Arial Black` is two identifiers, and
 * a family that happens to spell a CSS generic - a font really named `monospace`
 * is legal in OOXML - would silently become the generic.
 */
export function cssFamily(family: string): string {
  if (family.length === 0) {
    throw new TextError('TEXT_FONT_FAMILY', 'typeface name is empty', family);
  }
  // A quote or a backslash would break out of the quoted string and a control
  // character would break the shorthand. Spaces are fine and common - "Times
  // New Roman" is exactly why the name is quoted rather than escaped - so only
  // these are rejected, and rejected rather than stripped: a typeface name that
  // contains one is a corrupt file or an injection attempt, and neither should
  // be quietly repaired into something that parses.
  for (let i = 0; i < family.length; i += 1) {
    const code = family.charCodeAt(i);
    const isQuoteOrEscape = code === 0x22 || code === 0x5c;
    const isControl = code < 0x20 || code === 0x7f;
    if (isQuoteOrEscape || isControl) {
      throw new TextError(
        'TEXT_FONT_FAMILY',
        `typeface name ${JSON.stringify(family)} cannot be written into a CSS font shorthand`,
        family,
      );
    }
  }
  return `"${family}"`;
}

/**
 * A CSS `font` shorthand, in the order the property requires: style, weight,
 * size, family.
 *
 * The size is written in pixels because that is the unit the canvas takes, and
 * one pixel is one point here by construction - see the note on units above.
 */
export function cssFont(font: RunFont): string {
  const size = font.sz;
  if (!Number.isFinite(size) || size <= 0) {
    throw new TextError(
      'TEXT_SIZE',
      `font size ${String(size)} is not a positive size`,
      String(size),
    );
  }
  const style = font.italic === true ? 'italic ' : '';
  const weight = font.bold === true ? 'bold ' : '';
  return `${style}${weight}${String(size / 100)}px ${cssFamily(font.family)}`;
}

/**
 * `a:rPr/@spc` as a CSS `letterSpacing`.
 *
 * Measured twice over. PowerPoint's spacing is an **absolute length in points**,
 * not an em fraction: the same `@spc` widened a ten-character string by the same
 * number of points at 12pt and at 32pt. And it is applied after every character
 * *including the last* - scored 18/18 against 0/18 for the between-characters
 * reading. Chromium's `letterSpacing` does exactly the same, which is what makes
 * it usable directly rather than needing a correction.
 */
export function cssLetterSpacing(spc: number | undefined): string {
  if (spc === undefined) return '0px';
  if (!Number.isFinite(spc)) {
    throw new TextError('TEXT_SIZE', `a:rPr/@spc ${String(spc)} is not a length`, String(spc));
  }
  return `${String(spc / 100)}px`;
}

/** What a measurement returns. Points, in the fixed 960x540pt layout space. */
export interface Advance {
  /** The pen advance across the run: what a line's width is made of. */
  readonly width: number;
}

export interface TextMeasurer {
  /** The advance of `text` set in `font`, in points. */
  measure(text: string, font: RunFont): Advance;
}

/*
 * The canvas surface, declared rather than imported.
 *
 * This package is compiled with the DOM library available, but `OffscreenCanvas`
 * in a Worker and the DOM `CanvasRenderingContext2D` have drifted in lib.dom -
 * `letterSpacing` and `fontKerning` are on the 2d context at runtime in every
 * browser that matters and are not always in the checked-in types. Declaring the
 * exact surface used keeps `strict` honest without widening anything to `any`.
 */
export interface MeasuringContext {
  font: string;
  letterSpacing: string;
  fontKerning: string;
  measureText(text: string): {
    readonly width: number;
    readonly fontBoundingBoxAscent: number;
    readonly fontBoundingBoxDescent: number;
  };
}
interface MeasuringCanvas {
  getContext(id: '2d'): MeasuringContext | null;
}
type OffscreenCanvasCtor = new (width: number, height: number) => MeasuringCanvas;

/**
 * The one place in the package that acquires a measuring surface.
 *
 * The font guard and the baseline need a context too, and two routes to a
 * measurement are two measurements.
 */
export function createMeasuringContext(): MeasuringContext {
  const ctor = (globalThis as { OffscreenCanvas?: OffscreenCanvasCtor }).OffscreenCanvas;
  if (ctor === undefined) {
    throw new TextError(
      'TEXT_NO_CANVAS',
      'no OffscreenCanvas in this environment; the measurer is built for a Web Worker',
    );
  }
  const ctx = new ctor(1, 1).getContext('2d');
  if (ctx === null) {
    throw new TextError('TEXT_NO_CANVAS', 'OffscreenCanvas gave no 2d context');
  }
  return ctx;
}

/**
 * A measurer over a 1x1 `OffscreenCanvas`.
 *
 * The canvas is never drawn to, so its size is irrelevant and one pixel is
 * enough. One context is kept for the life of the measurer: setting `font` is
 * cheap, creating a context is not, and a slide asks for thousands of
 * measurements.
 */
export function createCanvasMeasurer(): TextMeasurer {
  const ctx = createMeasuringContext();

  return {
    measure(text: string, font: RunFont): Advance {
      ctx.font = cssFont(font);
      ctx.letterSpacing = cssLetterSpacing(font.spc);
      ctx.fontKerning = kerningEnabled(font.kern, font.sz) ? 'normal' : 'none';
      return { width: ctx.measureText(text).width };
    },
  };
}
