/**
 * `TextMeasurer` and `FaceBoxProbe` over font tables instead of a canvas.
 *
 * This is a second measurement engine, and `@pptx-studio/text` says in as many
 * words that nothing inside it may measure by a second route. That rule holds:
 * this is outside it, in the one package where Node exists, and it is here
 * because the browser's engine is not available to a server. What makes it safe
 * is that T13 measured Chromium's arithmetic rather than guessing it, so the two
 * engines agree exactly on a face they both have. ADR 0042.
 */

import {
  kerningEnabled,
  type FaceBox,
  type FaceBoxProbe,
  type RunFont,
  type TextMeasurer,
} from '@pptx-studio/text';

import type { FontLibrary, Resolved } from './faces.js';
import { RenderError } from './errors.js';

/**
 * The fixed-point step Chromium reports an advance in.
 *
 * Quantising each glyph advance toward zero and each kern adjustment to nearest
 * fits all 432 of T13's widths, where summing exact floats fits 150 and is out
 * by up to 1.04e-4 px. ADR 0042.
 */
const FIXED = 65536;

/**
 * `hmtx.advanceWidth` is a `UFWORD`, so this is never negative and truncation
 * and flooring are one rule. ADR 0042.
 */
function quantiseAdvance(px: number): number {
  return Math.trunc(px * FIXED) / FIXED;
}

function quantiseKern(px: number): number {
  return Math.round(px * FIXED) / FIXED;
}

export interface FaceUse {
  readonly asked: string;
  readonly drawn: string;
  readonly file: string;
  readonly substituted: boolean;
}

export interface FontMeasurer {
  readonly measurer: TextMeasurer;
  readonly faceBox: FaceBoxProbe;
  /** Every typeface asked for, and what it was drawn in. */
  used(): readonly FaceUse[];
  /** Code points no face in the library could draw. */
  missing(): readonly number[];
}

/**
 * A measurer bound to one library.
 *
 * Both probes cache by family, because a slide asks for the same handful of
 * typefaces thousands of times and nothing can change between two asks.
 */
export function createFontMeasurer(library: FontLibrary): FontMeasurer {
  const resolved = new Map<string, Resolved>();
  const missing = new Set<number>();
  /** Which face draws a code point the asked-for face has no glyph for. */
  const fallbacks = new Map<string, Resolved | null>();

  const faceFor = (family: string, bold: boolean, italic: boolean): Resolved => {
    const cacheKey = `${family}\u0000${bold ? 'b' : ''}${italic ? 'i' : ''}`;
    const cached = resolved.get(cacheKey);
    if (cached !== undefined) return cached;
    const found = library.resolve(family, bold, italic);
    if (found === undefined) {
      // The library resolves any family it holds a face for, so this is the
      // empty library: there is nothing at all to draw with.
      const where =
        library.directories.length === 0
          ? 'no directory was searched'
          : `nothing was found in ${library.directories.join(', ')}`;
      throw new RenderError(
        'CLI_NO_FACE',
        `no face on this machine can stand in for ${JSON.stringify(family)}: ` +
          `the font library is empty, ${where}`,
        family,
      );
    }
    resolved.set(cacheKey, found);
    return found;
  };

  /**
   * The advance of one code point, in font units over its own em.
   *
   * A face that lacks the glyph does not draw a blank: the browser would fall
   * back per character, so the whole index is searched once per code point and
   * the answer cached. A code point nothing has is recorded and contributes the
   * face's own `.notdef` width, which is what is drawn.
   */
  const advanceOf = (face: Resolved, codePoint: number): { units: number; em: number } => {
    const own = face.face.advanceOf(codePoint);
    if (own !== undefined) return { units: own, em: face.face.metrics.unitsPerEm };

    const cacheKey = String(codePoint);
    let stand = fallbacks.get(cacheKey);
    if (stand === undefined) {
      stand = null;
      for (const entry of library.indexed) {
        if (entry.face.advanceOf(codePoint) !== undefined) {
          stand = { ...face, face: entry.face, drawn: entry.face.family, file: entry.file };
          break;
        }
      }
      fallbacks.set(cacheKey, stand);
    }
    if (stand === null) {
      missing.add(codePoint);
      return { units: 0, em: face.face.metrics.unitsPerEm };
    }
    const units = stand.face.advanceOf(codePoint) ?? 0;
    return { units, em: stand.face.metrics.unitsPerEm };
  };

  const measurer: TextMeasurer = {
    measure(text: string, font: RunFont) {
      if (!Number.isFinite(font.sz) || font.sz <= 0) {
        throw new RenderError(
          'CLI_FONT_SIZE',
          `font size ${String(font.sz)} is not a positive size`,
          font.family,
        );
      }
      const face = faceFor(font.family, font.bold === true, font.italic === true);
      // One CSS pixel is one point here, exactly as in the canvas measurer.
      const px = font.sz / 100;
      const kerns = kerningEnabled(font.kern, font.sz);
      const spacing = (font.spc ?? 0) / 100;

      const points = [...text];
      let width = 0;
      for (let i = 0; i < points.length; i++) {
        const codePoint = points[i]?.codePointAt(0) ?? 0;
        const { units, em } = advanceOf(face, codePoint);
        width += quantiseAdvance((units * px) / em);
        // `@spc` is an absolute length applied after every character, the last
        // included - T2 scored that 18/18 - and Chromium's `letterSpacing` does
        // the same, which is why the canvas measurer can pass it straight in.
        width += spacing;
        if (kerns && i + 1 < points.length) {
          const next = points[i + 1]?.codePointAt(0) ?? 0;
          const adjust = face.face.kernBetween(codePoint, next);
          if (adjust !== 0) {
            width += quantiseKern((adjust * px) / face.face.metrics.unitsPerEm);
          }
        }
      }
      return { width };
    },
  };

  const boxes = new Map<string, FaceBox>();
  const faceBox: FaceBoxProbe = {
    box(family: string): FaceBox {
      const cached = boxes.get(family);
      if (cached !== undefined) return cached;
      const { metrics } = faceFor(family, false, false).face;
      // Chromium answers the `BASE` `ideo` coordinate where the face has one and
      // its own box descent where it does not, 30 of 30 in T13.
      const ideographic = (metrics.ideographic ?? metrics.descent) / metrics.unitsPerEm;
      const box: FaceBox = {
        ascent: metrics.ascent / metrics.unitsPerEm,
        descent: metrics.descent / metrics.unitsPerEm,
        // Anything outside the em is unusable to `uprightPen`, and zero is what
        // the browser probe answers there. `packages/text/src/lines/baseline.ts`.
        ideographic: ideographic > 0 && ideographic < 1 ? ideographic : 0,
      };
      boxes.set(family, box);
      return box;
    },
  };

  return {
    measurer,
    faceBox,
    used(): readonly FaceUse[] {
      const seen = new Map<string, FaceUse>();
      for (const face of resolved.values()) {
        seen.set(face.asked, {
          asked: face.asked,
          drawn: face.drawn,
          file: face.file,
          substituted: face.substituted,
        });
      }
      return [...seen.values()].sort((a, b) => (a.asked < b.asked ? -1 : 1));
    },
    missing(): readonly number[] {
      return [...missing].sort((a, b) => a - b);
    },
  };
}
