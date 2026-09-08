/**
 * What this machine will draw the corpus in, measured rather than configured.
 *
 * The plan's answer to the font trap was to vendor a metric-compatible clone
 * and alias to it in a `fonts.conf`. That pins a *cause*, and only the causes
 * anyone listed. This pins the *effect*: the advance every requested face
 * resolves to, and the shape of the glyphs it draws. A run on a machine whose
 * fonts differ from the recorded ones stops; it does not score. ADR 0035.
 */

import type { Page } from 'playwright';

import { FidelityError } from '../errors.ts';

/*
 * `page.evaluate` runs in Chromium and `tools/` compiles with `types: ["node"]`
 * and no DOM library, so declare exactly the surface this file's body uses.
 */
interface OffscreenCtx2D {
  fillStyle: string;
  font: string;
  clearRect(x: number, y: number, w: number, h: number): void;
  fillText(text: string, x: number, y: number): void;
  getImageData(x: number, y: number, w: number, h: number): { data: Uint8ClampedArray };
}
declare const OffscreenCanvas: new (
  width: number,
  height: number,
) => { getContext(id: '2d'): OffscreenCtx2D | null };
declare const navigator: { userAgent: string };

/** One face, as this machine actually resolves it. */
export interface FaceProbe {
  readonly family: string;
  /** False means every generic showed through, so nothing of this face is here. */
  readonly available: boolean;
  /** 3.7's own fingerprint: the advance of `Hamburgefonstiv` at 100px. */
  readonly advance: number;
  /**
   * SHA-256 of the 8-bit alpha of a pangram drawn in this face.
   *
   * The advance alone cannot tell a face from a metric-compatible clone of it -
   * being indistinguishable by advance is what "metric-compatible" means - and
   * a clone draws different glyphs. This is what separates them.
   */
  readonly coverageSha256: string;
}

export interface Environment {
  /** `win32-x64`, `linux-x64`: which recorded lock a run is compared against. */
  readonly envId: string;
  readonly chromium: string;
  readonly faces: readonly FaceProbe[];
}

/** The families a set of rendered slides drew with, sorted and deduplicated. */
export function familiesDrawn(
  slides: readonly { families: readonly string[] }[],
): readonly string[] {
  const found = new Set<string>();
  for (const slide of slides) for (const family of slide.families) found.add(family);
  // Default sort: UTF-16 code unit order, which is the same on every machine.
  return [...found].sort();
}

/** Measure each family, in the page, through `@pptx-studio/text`'s own probe. */
export async function probeEnvironment(
  page: Page,
  families: readonly string[],
  envId: string,
): Promise<Environment> {
  const chromium = await page.evaluate(() => navigator.userAgent);
  const faces = await page.evaluate(async (names: readonly string[]) => {
    const text = (
      globalThis as unknown as {
        pptx: {
          text: {
            createFontProbe: () => { advance: (stack: string, px: number, text: string) => number };
            fontFingerprint: (family: string, probe: unknown) => number;
            isFontAvailable: (family: string, probe: unknown) => boolean;
            cssFamily: (family: string) => string;
          };
        };
      }
    ).pptx.text;
    const probe = text.createFontProbe();
    /** Every ASCII letter, so the hash sees the shapes and not only the widths. */
    const PANGRAM = 'Sphinx of black quartz, judge my vow. 0123456789';

    const out = [];
    for (const family of names) {
      const canvas = new OffscreenCanvas(1024, 128);
      const context = canvas.getContext('2d');
      if (context === null) throw new Error('no 2d context');
      context.clearRect(0, 0, 1024, 128);
      context.fillStyle = '#000000';
      context.font = `64px ${text.cssFamily(family)}`;
      context.fillText(PANGRAM, 0, 96);
      const pixels = context.getImageData(0, 0, 1024, 128).data;
      // The alpha plane only, and not thresholded: a hinting change can leave
      // every glyph's silhouette identical and move every antialiasing ramp.
      const alpha = new Uint8Array(pixels.length / 4);
      for (let i = 0, a = 0; i < pixels.length; i += 4, a++) alpha[a] = pixels[i + 3] ?? 0;
      const digest = await crypto.subtle.digest('SHA-256', alpha);
      out.push({
        family,
        available: text.isFontAvailable(family, probe),
        advance: text.fontFingerprint(family, probe),
        coverageSha256: [...new Uint8Array(digest)]
          .map((b) => b.toString(16).padStart(2, '0'))
          .join(''),
      });
    }
    return out;
  }, families);

  return { envId, chromium, faces };
}

/**
 * Compare a measured environment against the one a baseline was recorded in.
 *
 * Throws on the first disagreement rather than collecting them: a machine whose
 * fonts or browser have moved cannot produce a comparable number for *any*
 * slide, so carrying on would only produce a longer wrong answer.
 */
export function assertSameEnvironment(measured: Environment, recorded: Environment): void {
  if (measured.chromium !== recorded.chromium) {
    throw new FidelityError(
      'FID_BROWSER_CHANGED',
      `the baseline was recorded through ${recorded.chromium} and this is ${measured.chromium}`,
      measured.chromium,
    );
  }
  const byName = new Map(recorded.faces.map((face) => [face.family, face]));
  for (const face of measured.faces) {
    const was = byName.get(face.family);
    if (was === undefined) continue;
    if (was.advance !== face.advance || was.available !== face.available) {
      throw new FidelityError(
        'FID_FONT_METRICS_CHANGED',
        `${face.family} measured ${String(face.advance)} here and ` +
          `${String(was.advance)} when the baseline was recorded`,
        face.family,
      );
    }
    if (was.coverageSha256 !== face.coverageSha256) {
      throw new FidelityError(
        'FID_FONT_COVERAGE_CHANGED',
        `${face.family} has the same advances as the recorded face and draws different glyphs`,
        face.family,
      );
    }
  }
}
