/**
 * Where the baseline sits inside a line box.
 *
 * 3.2 measured the box at 1.2 x the font size and stopped there; a renderer
 * cannot draw a glyph without knowing where in that box the pen goes. T8
 * measured it: the box is divided between ascent and descent in the typeface's
 * own proportion, 32 of 32 across eight faces and four sizes.
 * `corpus/ground-truth/text-rendering.json`, ADR 0034.
 */

import { TextError } from '../errors.js';

import { createMeasuringContext, cssFamily, type MeasuringContext } from '../runs/measure.js';

/**
 * A typeface's own box, as fractions of the em.
 *
 * `ascent + descent` is routinely more than one and sometimes more than 1.2 -
 * Verdana's is 1.215 - so this is not a subdivision of the em and the box it
 * divides is the line box.
 */
export interface FaceBox {
  readonly ascent: number;
  readonly descent: number;
  /**
   * The ideographic baseline below the alphabetic one, as a fraction of the em.
   *
   * Only vertical text reads it: an upright glyph's baseline sits `1 - this`
   * from the leading edge of its cell. ADR 0039.
   */
  readonly ideographic: number;
}

/**
 * The fraction of the line box that sits above the baseline.
 *
 * The rival every browser implements is CSS half-leading, which centres the
 * font box in the line box and puts the baseline an ascent below the top of it.
 * That scores 9 of 32 against this reading's 32, and is wrong by 1.4pt at 54pt
 * on Courier New - a fifth of a line.
 */
export function baselineShare(box: FaceBox): number {
  const total = box.ascent + box.descent;
  if (!Number.isFinite(total) || total <= 0 || box.ascent < 0 || box.descent < 0) {
    throw new TextError(
      'TEXT_FACE_METRICS',
      `a face box of ${String(box.ascent)} over ${String(box.descent)} is not a box`,
      String(total),
    );
  }
  return box.ascent / total;
}

/** How far below the top of the line box the baseline sits, in points. */
export function baselineDrop(lineHeightPt: number, box: FaceBox): number {
  if (!Number.isFinite(lineHeightPt) || lineHeightPt < 0) {
    throw new TextError(
      'TEXT_FACE_METRICS',
      `a line box of ${String(lineHeightPt)}pt is not a height`,
      String(lineHeightPt),
    );
  }
  return lineHeightPt * baselineShare(box);
}

/**
 * The size a face box is measured at.
 *
 * Chromium rounds the box to whole pixels, so the ratio read off a small size is
 * quantised: at 100px Arial reads 0.2250 where the answer is 0.2276, which is
 * half a point out at 54pt.
 */
export const FACE_BOX_PX = 1000;

/** U+4E00, the one character every CJK face has and no Latin face does. */
const IDEOGRAPH = '一';

/**
 * The ideographic baseline, below the alphabetic one, as a fraction of the em.
 *
 * Canvas reports a baseline under the alphabetic one as negative, and a face
 * with no ideographic metric of its own answers with its descent. Zero is the
 * reading for a face that answers with nothing usable, which puts an upright
 * glyph's alphabetic baseline on the trailing edge of its cell.
 */
function ideographicOffset(ctx: MeasuringContext): number {
  const offset = -ctx.measureText(IDEOGRAPH).ideographicBaseline / FACE_BOX_PX;
  return Number.isFinite(offset) && offset > 0 && offset < 1 ? offset : 0;
}

/** Measures a typeface's box on the same surface the measurer uses. */
export interface FaceBoxProbe {
  box(family: string): FaceBox;
}

/**
 * A probe over one `OffscreenCanvas`, caching each face.
 *
 * A slide asks for the same handful of families thousands of times and the
 * answer cannot change while the page is open.
 */
export function createFaceBoxProbe(context?: MeasuringContext): FaceBoxProbe {
  const ctx = context ?? createMeasuringContext();
  const cache = new Map<string, FaceBox>();
  return {
    box(family: string): FaceBox {
      const cached = cache.get(family);
      if (cached !== undefined) return cached;
      ctx.font = `${String(FACE_BOX_PX)}px ${cssFamily(family)}`;
      ctx.letterSpacing = '0px';
      ctx.fontKerning = 'normal';
      const metrics = ctx.measureText('Hxg');
      const box: FaceBox = {
        ascent: metrics.fontBoundingBoxAscent / FACE_BOX_PX,
        descent: metrics.fontBoundingBoxDescent / FACE_BOX_PX,
        ideographic: ideographicOffset(ctx),
      };
      baselineShare(box);
      cache.set(family, box);
      return box;
    },
  };
}
