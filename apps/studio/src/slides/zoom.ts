/**
 * Zoom, as the page means it: CSS pixels per point, so 100 % draws a 13.333-inch slide 960 wide -
 * the width the fidelity harness scores at, and PowerPoint's own 75 %. The device pixel ratio
 * multiplies it for the strokes only: a 2x display at 100 % draws the strokes of 200 %.
 */

import { MIN_PX_PER_PT, type SlideSize } from '@pptx-studio/render-svg';

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

/** The zoom that fits the slide's width into `viewportWidthPx`, never under the renderer's floor. */
export function fitZoom(viewportWidthPx: number, size: SlideSize): number {
  return Math.max(MIN_PX_PER_PT, viewportWidthPx / (size.cx / EMU_PER_POINT));
}

/** EMU per CSS pixel at a stage width, which is what keeps the overlay's chrome one pixel. */
export function unitAt(width: number, size: SlideSize): number {
  return size.cx / width;
}

/** The part of `window` a display watcher needs. */
export interface Screen {
  readonly devicePixelRatio: number;
  matchMedia(query: string): {
    addEventListener(type: 'change', listener: () => void, options: { once: true }): void;
  };
}

/** Calls `onChange` each time the device pixel ratio changes, until the returned function is called. */
export function watchDevicePixelRatio(screen: Screen, onChange: () => void): () => void {
  let stopped = false;
  const listen = (): void => {
    screen.matchMedia(`(resolution: ${String(screen.devicePixelRatio)}dppx)`).addEventListener(
      'change',
      () => {
        if (stopped) return;
        onChange();
        listen();
      },
      { once: true },
    );
  };
  listen();
  return () => {
    stopped = true;
  };
}
