/**
 * Experiment C7 - what the `p:pic` container does, as questions with rivals.
 *
 * C6 settled what `a:blipFill` means. It says nothing about the shape wrapped
 * around it, and the corpus cannot: all 26 of its pictures are an unrotated,
 * unflipped, unoutlined `prst="rect"` with no competing fill, so every one of
 * the four questions below is unexercised by every deck we own.
 *
 * The image is four saturated quadrants on a black slide. Four distinct colours
 * break every symmetry a flip could hide, and none of the five colours is any
 * other, so a sample never has two readings.
 */

import { EMU_PER_INCH, SLIDE_HEIGHT, SLIDE_WIDTH } from '../../lib/pptx.ts';

/** One point, in EMU. Every rectangle here is a whole number of points. */
export const PT = 12700;

/** The probe image, in pixels: four quadrants of 32. */
export const QUAD_PX = 64;

/** Top-left, top-right, bottom-left, bottom-right - and nothing between them. */
export const TL = 0xff0000;
export const TR = 0x00ff00;
export const BL = 0x0000ff;
export const BR = 0xffff00;

/** The slide behind the shapes, so "painted nothing here" has its own colour. */
export const BACKGROUND = 0x000000;
/** The `p:spPr` fill that competes with the image, and the outline colour. */
export const SPPR_FILL = 0x00ffff;
export const LINE_COLOR = 0xff00ff;

/** The four quadrants, with no interpolation anywhere near a sample point. */
export function quadPixel(x: number, y: number): number {
  const half = QUAD_PX / 2;
  if (y < half) return x < half ? TL : TR;
  return x < half ? BL : BR;
}

export interface Rect {
  readonly x: number;
  readonly y: number;
  readonly cx: number;
  readonly cy: number;
}

/** Where a sample sits inside a probe's rectangle, as fractions of it. */
export interface Sample {
  readonly id: string;
  readonly u: number;
  readonly v: number;
}

/**
 * Six samples per probe, every one in flat colour.
 *
 * The quadrant samples are at a quarter and three quarters, as far from the
 * image's own seams as they can be - the centre of the shape is the one point
 * that must never be sampled, being where all four quadrants meet. `corner` is
 * inside the bounding box, well outside an inscribed ellipse, and inside where
 * a centred outline would fall, so it separates all three placements at once.
 */
export const SAMPLES: readonly Sample[] = [
  { id: 'tl', u: 0.25, v: 0.25 },
  { id: 'tr', u: 0.75, v: 0.25 },
  { id: 'bl', u: 0.25, v: 0.75 },
  { id: 'br', u: 0.75, v: 0.75 },
  { id: 'inner', u: 0.4, v: 0.6 },
  { id: 'corner', u: 0.04, v: 0.04 },
];

/** The probe side and the outline width, in points: the profile is read in these. */
export const SIDE_PT = 108;
export const OUTLINE_PT = 12;
/** The question a probe is asked to settle. One probe answers exactly one. */
export type Question = 'clip' | 'precedence' | 'outline' | 'mirror' | 'control';

export interface PictureProbe {
  readonly id: string;
  readonly question: Question;
  readonly rect: Rect;
  /** `ST_ShapeType` on the picture's own `p:spPr`. */
  readonly prst: string;
  readonly flipH: boolean;
  readonly flipV: boolean;
  /** Sixtieths of a degree. */
  readonly rot: number;
  /** A whole fill element in `p:spPr`, competing with the image. */
  readonly spPrFill: string | undefined;
  /** A whole `a:ln`. */
  readonly line: string | undefined;
}

const INCH = EMU_PER_INCH;
/** 1.5in square, so a 12pt outline is a twelfth of the width and easy to sample. */
const SIDE = INCH + INCH / 2;
const GAP = INCH / 2;
const COLUMNS = 4;

function at(index: number): Rect {
  const column = index % COLUMNS;
  const row = Math.floor(index / COLUMNS);
  return {
    x: GAP + column * (SIDE + GAP),
    y: GAP + row * (SIDE + GAP),
    cx: SIDE,
    cy: SIDE,
  };
}

const THICK_LINE = `<a:ln w="${String(12 * PT)}"><a:solidFill><a:srgbClr val="FF00FF"/></a:solidFill></a:ln>`;
const CYAN_FILL = '<a:solidFill><a:srgbClr val="00FFFF"/></a:solidFill>';

interface Spec {
  readonly id: string;
  readonly question: Question;
  readonly prst?: string;
  readonly flipH?: boolean;
  readonly flipV?: boolean;
  readonly rot?: number;
  readonly spPrFill?: string;
  readonly line?: string;
}

/**
 * The probes.
 *
 * `control` is the case the corpus already covers, and it is here so that a
 * reading which explains the four questions but breaks the plain picture is
 * refuted rather than merely unlucky.
 */
const SPECS: readonly Spec[] = [
  { id: 'control-rect', question: 'control' },
  { id: 'clip-ellipse', question: 'clip', prst: 'ellipse' },
  { id: 'clip-triangle', question: 'clip', prst: 'triangle' },
  { id: 'precedence-solid', question: 'precedence', spPrFill: CYAN_FILL },
  { id: 'precedence-ellipse', question: 'precedence', prst: 'ellipse', spPrFill: CYAN_FILL },
  { id: 'outline-thick', question: 'outline', line: THICK_LINE },
  { id: 'mirror-h', question: 'mirror', flipH: true },
  { id: 'mirror-v', question: 'mirror', flipV: true },
  { id: 'mirror-hv', question: 'mirror', flipH: true, flipV: true },
];

export function pictureProbes(): readonly PictureProbe[] {
  return SPECS.map((spec, index) => ({
    id: spec.id,
    question: spec.question,
    rect: at(index),
    prst: spec.prst ?? 'rect',
    flipH: spec.flipH ?? false,
    flipV: spec.flipV ?? false,
    rot: spec.rot ?? 0,
    spPrFill: spec.spPrFill,
    line: spec.line,
  }));
}

/**
 * The `p:sp` the outline question is measured against.
 *
 * 2.8 measured stroke placement on a shape; that a picture differs is only a
 * finding if both are read off the same export, so the contrast is a probe.
 */
export function shapeControl(): { readonly id: string; readonly rect: Rect } {
  return { id: 'outline-sp', rect: at(SPECS.length) };
}
/** Nothing may run off the slide: a clipped probe would be scored as background. */
export function assertOnSlide(probes: readonly PictureProbe[]): void {
  for (const probe of probes) {
    const { x, y, cx, cy } = probe.rect;
    if (x < 0 || y < 0 || x + cx > SLIDE_WIDTH || y + cy > SLIDE_HEIGHT) {
      throw new Error(`${probe.id} is off the slide`);
    }
  }
}
