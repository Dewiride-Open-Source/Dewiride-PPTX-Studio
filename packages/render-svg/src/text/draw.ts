/**
 * The one entry point a renderer calls: a placed shape's text, as nodes.
 *
 * Resolving, laying out and emitting are three files because each is testable
 * on its own; this is the seam that joins them, and it is the only thing
 * `shape.ts` knows about text.
 */

import {
  APPROXIMATE_FACE_RULES,
  createCanvasMeasurer,
  createFaceBoxProbe,
  faceRules,
  hasFaceRules,
  type FaceBoxProbe,
  type FaceRules,
  type RunFont,
  type TextMeasurer,
} from '@pptx-studio/text';
import type { ListStyle } from '@pptx-studio/model';

import type { SvgNode } from '../node.js';
import { EMU_PER_POINT } from '../transform.js';
import type { Placed } from '../layout.js';

import { textNodes } from './emit.js';
import { askedFamily, layoutText, type TextBlock } from './layout.js';
import { resolveText } from './resolve.js';

/** How a caller supplies the three things text drawing cannot invent. */
export interface TextOptions {
  /** `p:defaultTextStyle`, the cascade's seventh source. */
  readonly defaultTextStyle?: ListStyle | undefined;
  /** Defaults to one over `OffscreenCanvas`, which needs a browser or a Worker. */
  readonly measurer?: TextMeasurer | undefined;
  readonly faceBox?: FaceBoxProbe | undefined;
  /**
   * What to do about a typeface whose rules nobody measured.
   *
   * Defaults to Arial's, because a renderer that refused to draw an underline on
   * an unmeasured face would refuse most decks. A caller that wants the
   * diagnostic passes `faceRules` itself and handles the throw.
   */
  readonly rulesFor?: ((typeface: string) => FaceRules) | undefined;
  /**
   * The CSS `font-family` a run is drawn in.
   *
   * Defaults to the run's own family, quoted, which is right where the caller
   * measures through that same name; a caller that measured in another face
   * names it here so the markup says what was drawn.
   */
  readonly cssFamilyFor?: ((font: RunFont) => string) | undefined;
}

/** Arial's rules for anything unmeasured, which is what a renderer wants. */
export function approximateRules(typeface: string): FaceRules {
  return hasFaceRules(typeface) ? faceRules(typeface) : APPROXIMATE_FACE_RULES;
}

/** Everything drawing needs, made once per slide rather than once per shape. */
export interface TextEngine {
  readonly defaultTextStyle: ListStyle | undefined;
  readonly measurer: TextMeasurer;
  readonly faceBox: FaceBoxProbe;
  readonly rulesFor: (typeface: string) => FaceRules;
  readonly cssFamilyFor: (font: RunFont) => string;
}

/**
 * The measurer and the probe, made once and not before they are needed.
 *
 * Once, because both hold a canvas context and a hundred-shape slide making a
 * hundred of them is a hundred allocations for one measurement each. Not before,
 * because a slide with no text at all must render where there is no
 * `OffscreenCanvas` - in Node, or in a test that only wants geometry.
 */
export function createTextEngine(options: TextOptions = {}): TextEngine {
  let measurer = options.measurer;
  let faceBox = options.faceBox;
  return {
    defaultTextStyle: options.defaultTextStyle,
    get measurer(): TextMeasurer {
      measurer ??= createCanvasMeasurer();
      return measurer;
    },
    get faceBox(): FaceBoxProbe {
      faceBox ??= createFaceBoxProbe();
      return faceBox;
    },
    rulesFor: options.rulesFor ?? approximateRules,
    cssFamilyFor: options.cssFamilyFor ?? askedFamily,
  };
}

/** The laid-out block of a placed shape, or `null` where it has no text. */
export function textBlockOf(placed: Placed, engine: TextEngine): TextBlock | null {
  const resolved = resolveText(placed, engine.defaultTextStyle);
  if (resolved === null) return null;
  return layoutText(resolved, {
    widthPt: placed.frame.cx / EMU_PER_POINT,
    heightPt: placed.frame.cy / EMU_PER_POINT,
    rot: placed.frame.rot,
    flipH: placed.frame.flipH,
    flipV: placed.frame.flipV,
    measurer: engine.measurer,
    faceBox: engine.faceBox,
    rulesFor: engine.rulesFor,
    cssFamilyFor: engine.cssFamilyFor,
  });
}

/** The text of a placed shape, as a sibling group of its geometry. */
export function shapeTextNodes(placed: Placed, engine: TextEngine): readonly SvgNode[] {
  const block = textBlockOf(placed, engine);
  if (block === null) return [];
  return textNodes(
    block,
    { x: placed.frame.x, y: placed.frame.y, cx: placed.frame.cx, cy: placed.frame.cy },
    {
      'data-text': String(placed.shape.cNvPrId),
      ...(placed.shape.name === '' ? {} : { 'data-name': placed.shape.name }),
    },
  );
}
