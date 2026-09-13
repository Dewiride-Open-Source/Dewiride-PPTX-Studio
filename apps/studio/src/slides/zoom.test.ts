import { describe, expect, it } from 'vitest';

import {
  deviceRatioAt,
  fitZoom,
  stageSizeAt,
  THUMB_ZOOM,
  unitAt,
  watchDevicePixelRatio,
  ZOOMS,
  type Screen,
} from './zoom.js';

const WIDE = { cx: 12192000, cy: 6858000 };
const FOUR_THREE = { cx: 9144000, cy: 6858000 };

describe('zoom', () => {
  it('draws a 13.333-inch slide 960 wide at 100 %, which is one CSS pixel per point', () => {
    expect(stageSizeAt(1, WIDE)).toEqual({ width: 960, height: 540 });
    expect(stageSizeAt(4, WIDE)).toEqual({ width: 3840, height: 2160 });
    expect(stageSizeAt(0.25, WIDE)).toEqual({ width: 240, height: 135 });
    // The strip's thumbnail is the smallest export the oracle holds, 120 by 68.
    expect(stageSizeAt(THUMB_ZOOM, WIDE)).toEqual({ width: 120, height: 68 });
  });

  it('follows the slide aspect rather than assuming 16:9', () => {
    expect(stageSizeAt(1, FOUR_THREE)).toEqual({ width: 720, height: 540 });
    expect(unitAt(720, FOUR_THREE)).toBe(12700);
  });

  it('keeps the overlay one pixel: EMU per pixel is the slide over the stage', () => {
    expect(unitAt(960, WIDE)).toBe(12700);
    expect(unitAt(3840, WIDE)).toBe(3175);
  });

  it('fits the slide to a viewport and never below a twentieth', () => {
    expect(fitZoom(960, WIDE)).toBe(1);
    expect(fitZoom(480, WIDE)).toBe(0.5);
    expect(fitZoom(10, WIDE)).toBe(0.05);
  });

  it('offers a quarter, one, two and four', () => {
    expect(ZOOMS).toEqual([0.25, 1, 2, 4]);
  });

  it("rounds strokes at the display's ratio, or at the floor when the zoom is under it", () => {
    expect(deviceRatioAt(1, 1)).toBe(1);
    expect(deviceRatioAt(THUMB_ZOOM, 2)).toBe(2);
    // A thumbnail on a display at a third: an eighth of a third is under a twentieth, so the floor.
    expect(deviceRatioAt(THUMB_ZOOM, 1 / 3)).toBe(0.4);
    expect(deviceRatioAt(0.25, 0.1)).toBe(0.2);
  });

  it('watches the ratio the display has now, and the next one after it changes', () => {
    const queries: { query: string; fire: () => void }[] = [];
    let ratio = 1;
    const screen: Screen = {
      get devicePixelRatio() {
        return ratio;
      },
      matchMedia: (query) => ({
        addEventListener: (_type, listener) => {
          queries.push({ query, fire: listener });
        },
      }),
    };
    let changes = 0;
    const stop = watchDevicePixelRatio(screen, () => {
      changes += 1;
    });
    expect(queries.map((q) => q.query)).toEqual(['(resolution: 1dppx)']);
    ratio = 2;
    queries[0]!.fire();
    expect(changes).toBe(1);
    // The next query is for the ratio the display has now, or a second change would go unseen.
    expect(queries.map((q) => q.query)).toEqual(['(resolution: 1dppx)', '(resolution: 2dppx)']);
    stop();
    queries[1]!.fire();
    expect(changes).toBe(1);
    expect(queries).toHaveLength(2);
  });

  it('refuses a zoom or a slide with no size', () => {
    expect(() => stageSizeAt(0, WIDE)).toThrow(RangeError);
    expect(() => stageSizeAt(1, { cx: 0, cy: 0 })).toThrow(RangeError);
  });
});
