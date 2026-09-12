/**
 * Coverage profiles over an exported bitmap, in points: ink is `1 - min(r,g,b)/255` over white.
 */

import type { Bitmap } from './bmp.ts';

/** Ink coverage over a white slide: `1 - min(r,g,b)/255`. */
export function ink(bmp: Bitmap, x: number, y: number): number {
  const [r, g, b] = bmp.pixel(x, y);
  return 1 - Math.min(r, g, b) / 255;
}

/** The pixel row or column nearest a point coordinate; off the slide is a mislaid probe, so it throws. */
export function pixelAt(points: number, scale: number, limit: number, what: string): number {
  const at = Math.floor(points * scale);
  if (at < 0 || at >= limit) {
    throw new Error(
      `${what}: ${String(points)}pt is pixel ${String(at)}, outside 0..${String(limit - 1)} - the probe is off the slide`,
    );
  }
  return at;
}

export interface Sampled {
  /** Coverage per sample, in order. */
  readonly cover: number[];
  /** The point coordinate of sample 0's centre. */
  readonly start: number;
  /** Points per sample. */
  readonly step: number;
}

export function sampleRow(
  bmp: Bitmap,
  scale: number,
  y: number,
  x0: number,
  x1: number,
  what: string,
): Sampled {
  const row = pixelAt(y, scale, bmp.height, what);
  const from = Math.max(0, Math.round(x0 * scale));
  const to = Math.min(bmp.width - 1, Math.round(x1 * scale));
  const cover: number[] = [];
  for (let x = from; x <= to; x++) cover.push(ink(bmp, x, row));
  return { cover, start: (from + 0.5) / scale, step: 1 / scale };
}

export function sampleCol(
  bmp: Bitmap,
  scale: number,
  x: number,
  y0: number,
  y1: number,
  what: string,
): Sampled {
  const col = pixelAt(x, scale, bmp.width, what);
  const from = Math.max(0, Math.round(y0 * scale));
  const to = Math.min(bmp.height - 1, Math.round(y1 * scale));
  const cover: number[] = [];
  for (let y = from; y <= to; y++) cover.push(ink(bmp, col, y));
  return { cover, start: (from + 0.5) / scale, step: 1 / scale };
}

/** Where coverage crosses one half, interpolated between the straddling samples to a tenth of a pixel. */
export function crossings(s: Sampled): { at: number; rising: boolean }[] {
  const out: { at: number; rising: boolean }[] = [];
  for (let i = 0; i + 1 < s.cover.length; i++) {
    const a = s.cover[i]!;
    const b = s.cover[i + 1]!;
    if (a < 0.5 && b >= 0.5) {
      out.push({ at: s.start + (i + (0.5 - a) / (b - a)) * s.step, rising: true });
    } else if (a >= 0.5 && b < 0.5) {
      out.push({ at: s.start + (i + (a - 0.5) / (a - b)) * s.step, rising: false });
    }
  }
  return out;
}

/** Alternating ink/gap lengths from a crossing list, in points. */
export function runsOf(cross: { at: number; rising: boolean }[]): { on: number[]; off: number[] } {
  const on: number[] = [];
  const off: number[] = [];
  for (let i = 0; i + 1 < cross.length; i++) {
    const a = cross[i]!;
    const b = cross[i + 1]!;
    if (a.rising && !b.rising) on.push(b.at - a.at);
    else if (!a.rising && b.rising) off.push(b.at - a.at);
  }
  return { on, off };
}
