/**
 * The heatmap, as geometry rather than as a picture.
 *
 * Two properties carry it and both are checkable without looking: a cell that
 * differs by `d` is drawn in the band the histogram counts it in, and a run of
 * equal cells is one rectangle rather than several. ADR 0038.
 */

import { describe, expect, it } from 'vitest';

import { bandOf, gridSvg, heatmapSvg, RAMP } from './heatmap.ts';
import { HISTOGRAM_EDGES } from './metric/score.ts';
import type { Grid } from './metric/reduce.ts';

function rects(svg: string): { x: number; width: number; fill: string }[] {
  return [
    ...svg.matchAll(/<rect x="(\d+)" y="\d+" width="(\d+)" height="\d+" fill="([^"]+)"\/>/g),
  ].map((found) => ({
    x: Number(found[1]),
    width: Number(found[2]),
    fill: found[3] ?? '',
  }));
}

describe('the difference bands', () => {
  it('puts a difference in the bucket the histogram counts it in', () => {
    // The same walk `scoreOf` does, so the picture and the number cannot
    // disagree about which side of an edge a cell falls on.
    for (const [difference, band] of [
      [0, 0],
      [1, 1],
      [2, 2],
      [3, 2],
      [4, 3],
      [127, 7],
      [128, 8],
      [255, 8],
    ] as const) {
      expect(bandOf(difference), `d=${String(difference)}`).toBe(band);
    }
  });

  it('has one colour per bucket, so no two differences share a band', () => {
    expect(RAMP).toHaveLength(HISTOGRAM_EDGES.length + 1);
    expect(new Set(RAMP).size).toBe(RAMP.length);
  });
});

describe('the heatmap', () => {
  it('draws nothing at all where the two sides agree', () => {
    const svg = heatmapSvg([0, 0, 0, 0], 2, 2, 8);
    expect(rects(svg)).toStrictEqual([]);
    expect(svg).toContain('viewBox="0 0 16 16"');
  });

  it('merges a run of equal cells into one rectangle', () => {
    const svg = heatmapSvg([200, 200, 200, 200], 4, 1, 8);
    const drawn = rects(svg);
    expect(drawn).toHaveLength(1);
    expect(drawn[0]?.x).toBe(0);
    expect(drawn[0]?.width).toBe(32);
    expect(drawn[0]?.fill).toBe(RAMP[8]);
  });

  it('splits a run where the band changes, and skips the agreeing cells', () => {
    const svg = heatmapSvg([0, 200, 200, 1], 4, 1, 8);
    const drawn = rects(svg);
    expect(drawn.map((rect) => [rect.x, rect.width])).toStrictEqual([
      [8, 16],
      [24, 8],
    ]);
    expect(drawn[0]?.fill).toBe(RAMP[8]);
    expect(drawn[1]?.fill).toBe(RAMP[1]);
  });
});

describe('a grid, drawn back', () => {
  function flat(luma: number, cb: number, cr: number): Grid {
    return {
      width: 2,
      height: 2,
      cell: 8,
      luma: [luma, luma, luma, luma],
      chromaWidth: 1,
      chromaHeight: 1,
      cb: [cb],
      cr: [cr],
    };
  }

  it('takes white back to white and black back to black', () => {
    expect(rects(gridSvg(flat(255, 128, 128)))[0]?.fill).toBe('#ffffff');
    expect(rects(gridSvg(flat(0, 128, 128)))[0]?.fill).toBe('#000000');
  });

  it('reads chroma at half resolution, so one value covers four cells', () => {
    // Cr well above neutral is red; a grid that read chroma per luma cell would
    // run off the end of a one-entry plane and draw the fallback instead.
    const drawn = rects(gridSvg(flat(128, 128, 255)));
    expect(drawn).toHaveLength(2);
    for (const rect of drawn) {
      expect(rect.width).toBe(16);
      expect(rect.fill.slice(1, 3)).toBe('ff');
    }
  });
});
