/**
 * Zoom, as the page means it: CSS pixels per point, so 100 % draws a 13.333-inch slide 960 wide -
 * the width the fidelity harness scores at, and PowerPoint's own 75 %.
 */

import type { SlideSize } from '@pptx-studio/render-svg';

const EMU_PER_POINT = 12700;

/** The zooms the stage offers, besides fitting the viewport. */
export const ZOOMS: readonly number[] = [0.25, 1, 2, 4];

/** The strip's thumbnails, 120 pixels wide: the smallest export the oracle holds. */
export const THUMB_ZOOM = 0.125;

export type ZoomChoice = number | 'fit';

/** The stage's CSS size at a zoom; the height follows the slide's own aspect. */
export function stageSizeAt(zoom: number, size: SlideSize): { width: number; height: number } {
  if (!(zoom > 0) || !(size.cx > 0) || !(size.cy > 0)) {
    throw new RangeError(
      `zoom ${String(zoom)} over a ${String(size.cx)} by ${String(size.cy)} EMU slide`,
    );
  }
  const width = Math.round((zoom * size.cx) / EMU_PER_POINT);
  return { width, height: Math.round((width * size.cy) / size.cx) };
}

/** The zoom that fits the slide's width into `viewportWidthPx`, never under a twentieth. */
export function fitZoom(viewportWidthPx: number, size: SlideSize): number {
  return Math.max(0.05, viewportWidthPx / (size.cx / EMU_PER_POINT));
}

/** EMU per CSS pixel at a stage width, which is what keeps the overlay's chrome one pixel. */
export function unitAt(width: number, size: SlideSize): number {
  return size.cx / width;
}
