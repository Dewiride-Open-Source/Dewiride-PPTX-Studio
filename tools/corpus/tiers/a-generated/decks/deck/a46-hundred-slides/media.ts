/**
 * The two pictures the deck carries, drawn procedurally so the package owns them outright.
 *
 * Two parts shared by every picture slide's relationships: thirteen slides reach two files, which
 * is the shape of a real deck's media and the one no three-slide probe has.
 */

import { png } from '../../../assets/png.ts';
import { REL, type ProbePart, type ProbeRel } from '../../../markup/chassis.ts';

export const IMAGE_WIDTH = 96;
export const IMAGE_HEIGHT = 54;

function rgb(r: number, g: number, b: number): number {
  return (r << 16) | (g << 8) | b;
}

/** A sky darkening upward over two hills, the kind of stock photo a review deck opens with. */
export function skylinePng(): Uint8Array {
  return png(IMAGE_WIDTH, IMAGE_HEIGHT, (x, y) => {
    const hillA = 34 + 10 * Math.sin(x / 9);
    const hillB = 40 + 8 * Math.cos(x / 6 + 1);
    if (y > hillA && y > hillB) return rgb(46, 92, 58);
    if (y > Math.min(hillA, hillB)) return rgb(74, 128, 84);
    const t = y / IMAGE_HEIGHT;
    return rgb(Math.round(120 + 100 * t), Math.round(160 + 70 * t), Math.round(220 + 30 * t));
  });
}

/** Five bars on a pale ground, the picture a slide pastes when the chart is a screenshot. */
export function barsPng(): Uint8Array {
  const heights = [30, 44, 26, 50, 38];
  return png(IMAGE_WIDTH, IMAGE_HEIGHT, (x, y) => {
    const bar = Math.floor(x / 19);
    const within = x % 19;
    const top = IMAGE_HEIGHT - (heights[bar] ?? 0);
    if (within >= 3 && within < 17 && y >= top && y < IMAGE_HEIGHT - 2) {
      return bar % 2 === 0 ? rgb(41, 128, 185) : rgb(230, 126, 34);
    }
    return y >= IMAGE_HEIGHT - 2 ? rgb(90, 90, 90) : rgb(245, 245, 240);
  });
}

export const MEDIA_PARTS: readonly ProbePart[] = [
  {
    name: 'ppt/media/image1.png',
    bytes: skylinePng(),
    contentType: { kind: 'default', extension: 'png', type: 'image/png' },
  },
  { name: 'ppt/media/image2.png', bytes: barsPng() },
];

export const SKYLINE_RID = 'rId2';
export const BARS_RID = 'rId3';

/** Both pictures, related from any slide that draws one. */
export const MEDIA_RELS: readonly ProbeRel[] = [
  { id: SKYLINE_RID, type: REL + 'image', target: '../media/image1.png' },
  { id: BARS_RID, type: REL + 'image', target: '../media/image2.png' },
];
