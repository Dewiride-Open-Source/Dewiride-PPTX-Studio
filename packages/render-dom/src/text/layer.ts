/**
 * The HTML text layer: the same lines as the SVG, as real text nodes.
 *
 * The plan's bet 4 asks for HTML text beside the SVG geometry, and the reason is
 * everything a `<text>` element cannot do: shaping, bidi, an IME, a selection a
 * user can drag, a screen reader. This draws from the identical `TextBlock`
 * `render-svg` emits, so the two cannot disagree about a line box - which is
 * sub-phase 3.8's own verification.
 *
 * ## Why the line carries an odd `line-height`
 *
 * A browser puts a baseline a half-leading plus an ascent below the line box and
 * PowerPoint puts it at the face's own share of it. `TextLine` carries the
 * `line-height` that reconciles the two, computed once in the layout pass. It
 * lands within one CSS pixel and no closer, because Chromium rounds a font's
 * ascent to a whole pixel before it lays a line out - which is why the SVG
 * emitter, where the baseline is a coordinate, is the one that thumbnails and
 * exports use.
 */

import type {
  PieceRule,
  TextBlock,
  TextLine,
  TextPiece,
  UprightGlyph,
} from '@pptx-studio/render-svg';
import { toCss, type Rgba } from '@pptx-studio/paint';

import { RenderDomError } from '../errors.js';

/** A slide's size in points, which is what a block's coordinates are in. */
export interface LayerSize {
  readonly widthPt: number;
  readonly heightPt: number;
}

/** One shape's text, with the frame it is drawn in. */
export interface LayerBlock {
  readonly block: TextBlock;
  /** The shape's frame, in points on the slide. */
  readonly leftPt: number;
  readonly topPt: number;
  readonly widthPt: number;
  readonly heightPt: number;
  /** `p:cNvPr/@id`, so a caller can find the element for a shape. */
  readonly cNvPrId: number;
  readonly name?: string | undefined;
}

export interface TextLayerOptions {
  readonly document?: Document | undefined;
  /** CSS pixels per point. One transform, never a re-layout. */
  readonly scale?: number | undefined;
  /** Let the layer take pointer events. Off by default: the SVG hit-tests. */
  readonly interactive?: boolean | undefined;
}

export interface MountedTextLayer {
  readonly root: HTMLElement;
  /** The element holding one shape's text. */
  element(cNvPrId: number): HTMLElement | null;
  /** Rescale without re-laying anything out. */
  resize(scale: number): void;
  unmount(): void;
}

function px(value: number): string {
  // Three decimals of a point is about a nanometre, and matches `render-svg`'s
  // own rounding, so the two layers round identically.
  return `${String(Math.round(value * 1000) / 1000)}px`;
}

/** Through `toCss`, because `Rgba` channels are 0..1 and not bytes. */
function css(color: Rgba | null): string {
  return color === null ? 'inherit' : toCss(color);
}

/** The dash pattern a rule is painted with, in multiples of its thickness. */
const DASHES: Readonly<Record<string, readonly number[]>> = {
  dotted: [1, 2],
  dashed: [4, 3],
  longDashed: [8, 3],
  dotDash: [4, 3, 1, 3],
  dotDotDash: [1, 3, 1, 3, 4, 3],
};

function ruleElement(
  document: Document,
  rule: PieceRule,
  x: number,
  baseline: number,
  color: string,
): HTMLElement {
  const el = document.createElement('div');
  const style = el.style;
  style.position = 'absolute';
  style.left = px(x + rule.leftPt);
  style.top = px(baseline + rule.topPt);
  style.width = px(rule.widthPt);
  style.height = px(rule.thickness);
  const dash = DASHES[rule.pattern];
  if (dash === undefined) {
    style.background = color;
    return el;
  }
  // A dashed rule as a repeating gradient, so it stays one element and scales
  // with the layer instead of being redrawn.
  const on = (dash[0] ?? 1) * rule.thickness;
  const off = (dash[1] ?? 1) * rule.thickness;
  style.background = `repeating-linear-gradient(to right, ${color} 0 ${px(on)}, transparent ${px(on)} ${px(on + off)})`;
  return el;
}

function pieceElement(document: Document, piece: TextPiece): HTMLElement {
  const el = document.createElement('span');
  const style = el.style;
  style.fontFamily = piece.cssFamily;
  style.fontSize = px(piece.font.sz / 100);
  if (piece.font.bold === true) style.fontWeight = 'bold';
  if (piece.font.italic === true) style.fontStyle = 'italic';
  if (piece.font.spc !== undefined && piece.font.spc !== 0) {
    style.letterSpacing = px(piece.font.spc / 100);
  }
  style.color = css(piece.color);
  if (piece.highlight !== null) style.background = css(piece.highlight);
  if (piece.risePt !== 0) {
    style.position = 'relative';
    style.top = px(-piece.risePt);
  }
  el.textContent = piece.text;
  return el;
}

/**
 * One glyph stood upright inside a turned line, placed from its own pen.
 *
 * The span is laid out unrotated with its baseline `sizePt` below its top - what
 * `UprightGlyph.strutPt` solves for - and then turned a quarter back about that
 * point, so the pen lands where the SVG emitter puts the same glyph.
 */
function uprightElement(document: Document, piece: TextPiece, glyph: UprightGlyph): HTMLElement {
  const el = pieceElement(document, piece);
  const style = el.style;
  const sizePt = piece.font.sz / 100;
  style.position = 'absolute';
  style.left = px(glyph.alongPt - piece.risePt);
  style.top = px(glyph.acrossPt - sizePt);
  style.lineHeight = px(glyph.strutPt);
  style.transformOrigin = `0 ${px(sizePt)}`;
  style.transform = 'rotate(-90deg)';
  el.textContent = glyph.text;
  return el;
}

function lineElement(document: Document, line: TextLine): HTMLElement {
  const el = document.createElement('div');
  const style = el.style;
  style.position = 'absolute';
  style.left = px(line.leftPt);
  style.top = px(line.topPt);
  style.width = px(line.widthPt);
  style.height = px(line.heightPt);
  style.lineHeight = px(line.strutPt);
  style.whiteSpace = 'pre';
  if (line.wordSpacingPt !== 0) style.wordSpacing = px(line.wordSpacingPt);

  for (const piece of line.pieces) {
    if (piece.upright.length > 0) {
      for (const glyph of piece.upright) el.appendChild(uprightElement(document, piece, glyph));
    } else el.appendChild(pieceElement(document, piece));
  }

  const baseline = line.baselinePt - line.topPt;
  for (const piece of line.pieces) {
    const color = css(piece.color);
    for (const rule of piece.rules) {
      el.appendChild(ruleElement(document, rule, piece.leftPt, baseline - piece.risePt, color));
    }
  }
  return el;
}

function blockElement(document: Document, entry: LayerBlock): HTMLElement {
  const el = document.createElement('div');
  const style = el.style;
  style.position = 'absolute';
  style.left = px(entry.leftPt);
  style.top = px(entry.topPt);
  style.width = px(entry.widthPt);
  style.height = px(entry.heightPt);
  if (entry.block.turnDeg !== 0) {
    style.transformOrigin = '50% 50%';
    style.transform = `rotate(${String(entry.block.turnDeg)}deg)`;
  }
  el.setAttribute('data-text', String(entry.cNvPrId));
  if (entry.name !== undefined && entry.name !== '') el.setAttribute('data-name', entry.name);
  for (const line of entry.block.lines) el.appendChild(lineElement(document, line));
  return el;
}

/**
 * Mount a text layer into a host, over a slide of `size` points.
 *
 * The host is emptied first. The layer is sized in points and scaled with one
 * transform, which is the plan's rule: font metrics are not linear in point
 * size, so a renderer that re-laid out per zoom would change its line breaks as
 * the user zoomed.
 */
export function mountTextLayer(
  host: Element,
  blocks: readonly LayerBlock[],
  size: LayerSize,
  options: TextLayerOptions = {},
): MountedTextLayer {
  if (host === null || typeof host !== 'object' || typeof host.appendChild !== 'function') {
    throw new RenderDomError('RENDER_DOM_NO_HOST', 'mountTextLayer needs an element to build into');
  }
  const document = options.document ?? host.ownerDocument;
  if (document === null || document === undefined) {
    throw new RenderDomError('RENDER_DOM_NO_DOCUMENT', 'the host element has no ownerDocument');
  }

  const root = document.createElement('div');
  const style = root.style;
  style.position = 'absolute';
  style.left = '0';
  style.top = '0';
  style.width = px(size.widthPt);
  style.height = px(size.heightPt);
  style.transformOrigin = '0 0';
  style.pointerEvents = options.interactive === true ? 'auto' : 'none';
  root.setAttribute('data-role', 'text-layer');

  const byId = new Map<number, HTMLElement>();
  for (const entry of blocks) {
    if (entry.block.lines.length === 0) continue;
    const el = blockElement(document, entry);
    byId.set(entry.cNvPrId, el);
    root.appendChild(el);
  }

  const apply = (scale: number): void => {
    style.transform = scale === 1 ? '' : `scale(${String(scale)})`;
  };
  apply(options.scale ?? 1);
  host.replaceChildren(root);

  return {
    root,
    element(cNvPrId: number): HTMLElement | null {
      return byId.get(cNvPrId) ?? null;
    },
    resize(scale: number): void {
      apply(scale);
    },
    unmount(): void {
      root.remove();
    },
  };
}
