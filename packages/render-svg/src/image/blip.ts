/**
 * An `a:blipFill` as SVG.
 *
 * The placement arithmetic is `@pptx-studio/paint`'s and was measured in C6;
 * this module only chooses the constructs that reproduce it. One choice is
 * load-bearing: a mirrored tiling repeats over TWO tiles per mirrored axis,
 * because `@flip` mirrors alternate tiles rather than every one, so the pattern
 * cell is 2x1, 1x2 or 2x2 and carries that many `<image>` elements. ADR 0036.
 */

import {
  blipPlacement,
  brightOffset,
  contrastScale,
  resolveColor,
  toHexColor,
  BLIP_LUMA,
  type BlipEffect,
  type BlipFill,
  type ColorContext,
  type ImageSize,
  type Rect,
} from '@pptx-studio/paint';

import { element, type SvgElement } from '../node.js';
import type { Box } from '../transform.js';
import type { Defs, Attrs } from '../paint.js';

/** What the caller hands back for an `r:embed`, having resolved it themselves. */
export interface MediaImage {
  readonly bytes: Uint8Array;
  readonly contentType: string;
}

/**
 * Resolves a relationship id against the part the fill was written in.
 *
 * The part is not optional: rIds are scoped to one .rels, so a fill inherited
 * from a layout means the LAYOUT rId2 and resolving it against the slide would
 * silently paint a different picture.
 */
export type MediaResolver = (embed: string, part: string) => MediaImage | undefined;

function transferFor(scale: number, offset: number): SvgElement[] {
  return (['feFuncR', 'feFuncG', 'feFuncB'] as const).map((name) =>
    element(name, { type: 'linear', slope: scale, intercept: offset }),
  );
}

/** A colour matrix that writes the BT.709 luminance into all three channels. */
function luminanceMatrix(): SvgElement {
  const { r, g, b } = BLIP_LUMA;
  const row = `${String(r)} ${String(g)} ${String(b)} 0 0`;
  return element('feColorMatrix', { type: 'matrix', values: `${row} ${row} ${row} 0 0 0 1 0` });
}

/** One effect as filter primitives, in the order they were written. */
function effectPrimitives(effect: BlipEffect, ctx: ColorContext): SvgElement[] {
  switch (effect.kind) {
    case 'grayscale':
      return [luminanceMatrix()];
    case 'biLevel': {
      // A step at the threshold, built as a slope steep enough to clip either
      // side of it; `discrete` can only split a table at even intervals.
      const thresh = effect.thresh / 100000;
      const slope = 1e5;
      return [
        luminanceMatrix(),
        element('feComponentTransfer', {}, transferFor(slope, 0.5 - slope * thresh)),
      ];
    }
    case 'lum': {
      const scale = contrastScale(effect.contrast);
      const offset = 0.5 * (1 - scale) + brightOffset(effect.bright, scale);
      return [element('feComponentTransfer', {}, transferFor(scale, offset))];
    }
    case 'duotone': {
      const from = resolveColor(effect.from, ctx);
      const to = resolveColor(effect.to, ctx);
      const channel = (name: 'feFuncR' | 'feFuncG' | 'feFuncB', a: number, b: number): SvgElement =>
        element(name, { type: 'linear', slope: (b - a) / 255, intercept: a / 255 });
      return [
        luminanceMatrix(),
        element('feComponentTransfer', {}, [
          channel('feFuncR', from.r, to.r),
          channel('feFuncG', from.g, to.g),
          channel('feFuncB', from.b, to.b),
        ]),
      ];
    }
    case 'alphaModFix':
      return [
        element('feComponentTransfer', {}, [
          element('feFuncA', { type: 'linear', slope: effect.amt / 100000, intercept: 0 }),
        ]),
      ];
    default:
      return colorChange(effect, ctx);
  }
}

/** A 256-entry `discrete` table that is 1 only at one channel value. */
function indicator(level: number): string {
  const values: string[] = [];
  for (let at = 0; at < 256; at++) values.push(at === level ? '1' : '0');
  return values.join(' ');
}

/**
 * `a:clrChange` as an exact three-channel match.
 *
 * No primitive tests colour equality, but a `discrete` transfer table does test
 * one channel exactly, and three of them summed reach 3 only where all three
 * matched - which the next transfer turns into a one-bit mask. Measured to
 * replace one colour and leave the value one level away alone.
 */
function colorChange(
  effect: Extract<BlipEffect, { kind: 'clrChange' }>,
  ctx: ColorContext,
): SvgElement[] {
  const from = resolveColor(effect.from, ctx);
  const to = resolveColor(effect.to, ctx);
  return [
    element('feComponentTransfer', { in: 'SourceGraphic', result: 'ind' }, [
      element('feFuncR', { type: 'discrete', tableValues: indicator(Math.round(from.r)) }),
      element('feFuncG', { type: 'discrete', tableValues: indicator(Math.round(from.g)) }),
      element('feFuncB', { type: 'discrete', tableValues: indicator(Math.round(from.b)) }),
      element('feFuncA', { type: 'linear', slope: 0, intercept: 1 }),
    ]),
    // The three indicators average to 1 only where every channel matched.
    element('feColorMatrix', {
      in: 'ind',
      result: 'sum',
      type: 'matrix',
      values: '0 0 0 0 0  0 0 0 0 0  0 0 0 0 0  0.3333 0.3333 0.3333 0 0',
    }),
    element('feComponentTransfer', { in: 'sum', result: 'mask' }, [
      element('feFuncA', { type: 'linear', slope: 3, intercept: -2 }),
    ]),
    element('feFlood', {
      'flood-color': `#${toHexColor(to)}`,
      'flood-opacity': effect.useAlpha ? to.a : 1,
      result: 'repl',
    }),
    element('feComposite', { in: 'repl', in2: 'mask', operator: 'in', result: 'painted' }),
    element('feComposite', { in: 'SourceGraphic', in2: 'mask', operator: 'out', result: 'kept' }),
    element('feMerge', {}, [
      element('feMergeNode', { in: 'kept' }),
      element('feMergeNode', { in: 'painted' }),
    ]),
  ];
}

function filterFor(
  effects: readonly BlipEffect[],
  ctx: ColorContext,
  defs: Defs,
): string | undefined {
  if (effects.length === 0) return undefined;
  const primitives = effects.flatMap((effect) => effectPrimitives(effect, ctx));
  const id = defs.id();
  defs.add(
    element(
      'filter',
      {
        id,
        x: '0%',
        y: '0%',
        width: '100%',
        height: '100%',
        'color-interpolation-filters': 'sRGB',
      },
      primitives,
    ),
  );
  return `url(#${id})`;
}

function imageNode(href: string, at: Rect, filter: string | undefined, flip: string): SvgElement {
  const attrs: Attrs = {
    href,
    x: at.x,
    y: at.y,
    width: at.w,
    height: at.h,
    preserveAspectRatio: 'none',
  };
  if (filter !== undefined) attrs['filter'] = filter;
  if (flip !== '') attrs['transform'] = flip;
  return element('image', attrs);
}

/**
 * An image fill as a `<pattern>` reference.
 *
 * A pattern rather than a clipped `<image>`, because the same construct then
 * serves both modes and because a pattern cell clips its own contents - which
 * is what crops a stretched image without a second clip path.
 */
export function blipPaint(
  fill: BlipFill,
  ctx: ColorContext,
  box: Box,
  defs: Defs,
  media: MediaResolver | undefined,
  size: (bytes: Uint8Array) => ImageSize,
  uri: (bytes: Uint8Array, contentType: string) => string,
): Attrs {
  // A fill whose image the caller cannot resolve paints nothing, exactly as a
  // shape whose blip is missing does in PowerPoint - and so does one whose only
  // image is the SVG original, until 10.7 draws it.
  const found = fill.embed === null ? undefined : media?.(fill.embed, fill.part);
  if (found === undefined) return { fill: 'none' };

  const image = size(found.bytes);
  const href = uri(found.bytes, found.contentType);
  const placement = blipPlacement(fill, { x: box.x, y: box.y, w: box.cx, h: box.cy }, image);
  const filter = filterFor(fill.effects, ctx, defs);
  const id = defs.id();

  if (placement.kind === 'stretch') {
    defs.add(
      element(
        'pattern',
        {
          id,
          patternUnits: 'userSpaceOnUse',
          x: placement.clip.x,
          y: placement.clip.y,
          width: placement.clip.w,
          height: placement.clip.h,
        },
        [
          imageNode(
            href,
            {
              x: placement.image.x - placement.clip.x,
              y: placement.image.y - placement.clip.y,
              w: placement.image.w,
              h: placement.image.h,
            },
            filter,
            '',
          ),
        ],
      ),
    );
    return { fill: `url(#${id})` };
  }

  const cols = placement.flipX ? 2 : 1;
  const rows = placement.flipY ? 2 : 1;
  const cells: SvgElement[] = [];
  for (let row = 0; row < rows; row++) {
    for (let col = 0; col < cols; col++) {
      const mirrorX = col === 1;
      const mirrorY = row === 1;
      if (!mirrorX && !mirrorY) {
        cells.push(imageNode(href, placement.image, filter, ''));
        continue;
      }
      const sx = mirrorX ? -1 : 1;
      const sy = mirrorY ? -1 : 1;
      const shiftX = mirrorX ? 2 * placement.tileW : 0;
      const shiftY = mirrorY ? 2 * placement.tileH : 0;
      const transform = `translate(${String(shiftX)} ${String(shiftY)}) scale(${String(sx)} ${String(sy)})`;
      cells.push(imageNode(href, placement.image, filter, transform));
    }
  }
  defs.add(
    element(
      'pattern',
      {
        id,
        patternUnits: 'userSpaceOnUse',
        x: placement.originX,
        y: placement.originY,
        width: placement.tileW * cols,
        height: placement.tileH * rows,
      },
      cells,
    ),
  );
  return { fill: `url(#${id})` };
}
