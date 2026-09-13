/**
 * Zoom as the viewer means it: CSS pixels per point, so 100 % draws a 13.333-inch
 * slide 960 wide - PowerPoint's own 75 %. The device pixel ratio multiplies it
 * for the strokes only: a 2x display at 100 % draws the strokes of 200 %.
 */

import { MIN_PX_PER_PT, type SlideSize } from '@pptx-studio/render-svg';

const EMU_PER_POINT = 12700;

export type ZoomChoice = '0.25' | '0.5' | '1' | '2' | '4' | 'fit';

export const ZOOM_CHOICES: readonly { value: ZoomChoice; label: string }[] = [
  { value: '0.25', label: '25%' },
  { value: '0.5', label: '50%' },
  { value: '1', label: '100%' },
  { value: '2', label: '200%' },
  { value: '4', label: '400%' },
  { value: 'fit', label: 'Fit' },
];

/** The zoom a choice means, given the frame the slide has to fit. */
export function zoomOf(choice: ZoomChoice, viewportWidthPx: number, size: SlideSize): number {
  return choice === 'fit' ? fitZoom(viewportWidthPx, size) : Number(choice);
}

/** The strip's thumbnails, 120 pixels wide. */
export const THUMB_ZOOM = 0.125;

/** The stage's CSS size at a zoom; the height follows the slide's own aspect. */
export function stageSizeAt(zoom: number, size: SlideSize): { width: number; height: number } {
  const width = Math.round((zoom * size.cx) / EMU_PER_POINT);
  return { width, height: Math.round((width * size.cy) / size.cx) };
}

/** The zoom that fits the slide's width into a viewport, never under the renderer's floor. */
export function fitZoom(viewportWidthPx: number, size: SlideSize): number {
  return Math.max(MIN_PX_PER_PT, viewportWidthPx / (size.cx / EMU_PER_POINT));
}

/** EMU per CSS pixel at a stage width, which keeps the overlay's chrome one pixel wide. */
export function unitAt(width: number, size: SlideSize): number {
  return size.cx / width;
}

/** The ratio the strokes are rounded at: the display's, or the floor's when the zoom is under it. */
export function deviceRatioAt(zoom: number, ratio: number): number {
  return Math.max(ratio, MIN_PX_PER_PT / zoom);
}

export function pointsOf(size: SlideSize): { widthPt: number; heightPt: number } {
  return { widthPt: size.cx / EMU_PER_POINT, heightPt: size.cy / EMU_PER_POINT };
}
