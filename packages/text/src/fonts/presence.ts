/**
 * Whether the machine rendering this deck actually has the typeface it names.
 *
 * `document.fonts.check` cannot answer it: T7 asked it about 38 families and it
 * said yes to all 38, including 20 that do not exist. ADR 0033.
 */

import { TextError } from '../errors.js';
import { cssFamily, createMeasuringContext, type MeasuringContext } from '../runs/measure.js';

/**
 * The generics a family is stacked on so that absence becomes observable.
 *
 * All three, not one: Chromium resolves `monospace` to Consolas, `serif` to
 * Times New Roman and `sans-serif` to Arial, so any single anchor calls its own
 * resolution absent. One anchor scored 36-37 of 38 and all three scored 38.
 */
export const ANCHORS: readonly string[] = ['monospace', 'serif', 'sans-serif'];

/**
 * The string a fingerprint is taken over.
 *
 * Not three glyphs: no set of three or fewer single characters separates as many
 * families as the whole pool does, because Segoe UI and Leelawadee UI agree on
 * every character measured and differ only over a word.
 */
export const FINGERPRINT_TEXT = 'Hamburgefonstiv';

/**
 * The size a fingerprint is taken at.
 *
 * Fixed, because advances are **not** linear in size - hinting moved a small
 * advance by up to 11.1% from what the large one predicts - so two fingerprints
 * are comparable only at one size.
 */
export const FINGERPRINT_PX = 100;

/** A measurement over one CSS font stack. */
export interface FontProbe {
  /** The advance of `text` set at `px` in `stack`, which is a CSS family list. */
  advance(stack: string, px: number, text: string): number;
}

/** A probe over the same `OffscreenCanvas` route the measurer uses. */
export function createFontProbe(): FontProbe {
  const ctx: MeasuringContext = createMeasuringContext();
  return {
    advance(stack: string, px: number, text: string): number {
      if (!Number.isFinite(px) || px <= 0) {
        throw new TextError('TEXT_SIZE', `${String(px)} is not a positive size`, String(px));
      }
      ctx.font = `${String(px)}px ${stack}`;
      ctx.letterSpacing = '0px';
      ctx.fontKerning = 'normal';
      return ctx.measureText(text).width;
    },
  };
}

/** `"Family", monospace` - the family with one generic behind it. */
export function anchoredStack(family: string, anchor: string): string {
  return `${cssFamily(family)}, ${anchor}`;
}

/**
 * Whether this machine resolves `family` to something of its own.
 *
 * False means every anchor showed through, so the family contributed nothing.
 * True is the weaker claim it sounds like: a name the system maps to another
 * file - `Helvetica` to Arial here - answers true, and correctly, because the
 * text will lay out exactly as that file.
 */
export function isFontAvailable(family: string, probe: FontProbe): boolean {
  return ANCHORS.some(
    (anchor) =>
      probe.advance(anchoredStack(family, anchor), FINGERPRINT_PX, FINGERPRINT_TEXT) !==
      probe.advance(anchor, FINGERPRINT_PX, FINGERPRINT_TEXT),
  );
}

/**
 * The advance that identifies what `family` resolved to, at `FINGERPRINT_PX`.
 *
 * Two families with one fingerprint lay out identically. That is all a layout
 * engine needs, and all this can honestly claim: Arial and Helvetica share one,
 * because on this machine they are one file.
 */
export function fontFingerprint(family: string, probe: FontProbe): number {
  return probe.advance(cssFamily(family), FINGERPRINT_PX, FINGERPRINT_TEXT);
}

/** Whether two families lay out identically, which is what metric-compatible means. */
export function metricsAgree(a: string, b: string, probe: FontProbe): boolean {
  return fontFingerprint(a, probe) === fontFingerprint(b, probe);
}
