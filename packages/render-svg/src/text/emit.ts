/**
 * A laid-out block as SVG: real `<text>` and `<tspan>`, never `foreignObject`.
 *
 * `foreignObject` taints a canvas, so a slide holding one cannot be rasterised
 * for a thumbnail, and Safari drops it inside an `<img>` without a word. The
 * cost of doing it properly is this file.
 *
 * ## One `<text>` per line, one `<tspan>` per piece, and no `x` on the tspans
 *
 * T8 measured that a line is shaped as one string and that a run boundary splits
 * only the drawing call - a kern crosses it, 6 of 6. A `<tspan>` carrying its own
 * `x` restarts shaping at that point and loses exactly that kern, so the pieces
 * flow instead and only the line is placed. Justification is `word-spacing` for
 * the same reason.
 */

import { toHexColor, type Rgba } from '@pptx-studio/paint';

import { element, text as textNode, type SvgElement, type SvgNode } from '../node.js';
import { EMU_PER_POINT } from '../transform.js';

import type { PieceRule, TextBlock, TextLine, TextPiece, UprightGlyph } from './layout.js';

/**
 * `#rrggbb`, with alpha handed over separately as SVG wants it.
 *
 * Through `toHexColor`, because `Rgba` channels are 0..1 and a renderer that
 * treats them as bytes paints everything black.
 */
function paint(color: Rgba | null): { fill: string; opacity: number | null } {
  if (color === null) return { fill: 'currentColor', opacity: null };
  return { fill: `#${toHexColor(color)}`, opacity: color.a >= 1 ? null : color.a };
}

/** The dash pattern a rule is painted with, in multiples of its thickness. */
const DASHES: Readonly<Record<string, readonly number[]>> = {
  dotted: [1, 2],
  dashed: [4, 3],
  longDashed: [8, 3],
  dotDash: [4, 3, 1, 3],
  dotDotDash: [1, 3, 1, 3, 4, 3],
};

function ruleNode(
  rule: PieceRule,
  x: number,
  baseline: number,
  fill: string,
  opacity: number | null,
): SvgElement {
  const dash = DASHES[rule.pattern];
  if (dash === undefined && !rule.wavy) {
    return element('rect', {
      x: x + rule.leftPt,
      y: baseline + rule.topPt,
      width: rule.widthPt,
      height: rule.thickness,
      fill,
      ...(opacity === null ? {} : { 'fill-opacity': opacity }),
    });
  }
  // A patterned or wavy rule is a stroked line: a dash array needs a stroke, and
  // a wave needs a path, and both keep the rule one element per stretch.
  const middle = baseline + rule.topPt + rule.thickness / 2;
  return element('path', {
    d: rule.wavy
      ? wave(x + rule.leftPt, middle, rule.widthPt, rule.thickness)
      : `M${String(x + rule.leftPt)} ${String(middle)}h${String(rule.widthPt)}`,
    fill: 'none',
    stroke: fill,
    ...(opacity === null ? {} : { 'stroke-opacity': opacity }),
    'stroke-width': rule.thickness,
    ...(dash === undefined
      ? {}
      : { 'stroke-dasharray': dash.map((step) => step * rule.thickness).join(' ') }),
  });
}

/** A wave of one period per two thicknesses, which is what a wavy rule looks like. */
function wave(x: number, y: number, width: number, thickness: number): string {
  const step = thickness * 2;
  const parts = [`M${String(x)} ${String(y)}`];
  for (let at = 0; at + step <= width; at += step * 2) {
    parts.push(`q${String(step / 2)} ${String(-step)} ${String(step)} 0`);
    parts.push(`q${String(step / 2)} ${String(step)} ${String(step)} 0`);
  }
  return parts.join('');
}

function pieceAttrs(piece: TextPiece): Record<string, string | number | null> {
  const { fill, opacity } = paint(piece.color);
  return {
    'font-family': piece.cssFamily,
    'font-size': piece.font.sz / 100,
    ...(piece.font.bold === true ? { 'font-weight': 'bold' } : {}),
    ...(piece.font.italic === true ? { 'font-style': 'italic' } : {}),
    ...(piece.font.spc === undefined || piece.font.spc === 0
      ? {}
      : { 'letter-spacing': piece.font.spc / 100 }),
    fill,
    ...(opacity === null ? {} : { 'fill-opacity': opacity }),
    ...(piece.risePt === 0 ? {} : { dy: -piece.risePt }),
  };
}

/**
 * The nodes of one line: highlights under it, the text, then the rules over it.
 *
 * A shifted piece carries a `dy` and the piece after it carries the opposite
 * one, so a superscript moves itself and nothing that follows.
 */
/**
 * One glyph stood upright inside a turned line.
 *
 * The line's own turn is undone about the glyph's pen, which is why this is a
 * `<text>` of its own rather than a `<tspan>`: a rotation is per element.
 */
function uprightNode(line: TextLine, piece: TextPiece, glyph: UprightGlyph): SvgElement {
  // A rise lifts the glyph along its own up, which the quarter turn maps to the
  // line's own left, so it moves the pen rather than becoming a `dy`.
  const x = line.leftPt + glyph.alongPt - piece.risePt;
  const y = line.topPt + glyph.acrossPt;
  return element(
    'text',
    {
      ...pieceAttrs(piece),
      x,
      y,
      dy: null,
      transform: `rotate(-90 ${String(x)} ${String(y)})`,
      'xml:space': 'preserve',
    },
    [textNode(glyph.text)],
  );
}

function lineNodes(line: TextLine): readonly SvgNode[] {
  const out: SvgNode[] = [];

  for (const piece of line.pieces) {
    if (piece.highlight === null) continue;
    const { fill, opacity } = paint(piece.highlight);
    out.push(
      element('rect', {
        x: line.leftPt + piece.leftPt,
        y: line.topPt,
        width: piece.widthPt,
        height: line.heightPt,
        fill,
        ...(opacity === null ? {} : { 'fill-opacity': opacity }),
      }),
    );
  }

  const spans: SvgNode[] = [];
  let carried = 0;
  line.pieces.forEach((piece, index) => {
    if (piece.upright.length > 0) {
      for (const glyph of piece.upright) out.push(uprightNode(line, piece, glyph));
      return;
    }
    const attrs = pieceAttrs(piece);
    const dy = -piece.risePt + carried;
    carried = piece.risePt;
    spans.push(
      element(
        'tspan',
        { ...attrs, ...(dy === 0 ? { dy: null } : { dy }), 'xml:space': 'preserve' },
        [textNode(piece.text)],
      ),
    );
    if (index === line.pieces.length - 1 && carried !== 0) {
      // Nothing follows to carry the shift back, and an unclosed `dy` would move
      // the next line in a renderer that reused the element.
      carried = 0;
    }
  });

  if (spans.length > 0) {
    out.push(
      element(
        'text',
        {
          x: line.leftPt,
          y: line.baselinePt,
          ...(line.wordSpacingPt === 0 ? {} : { 'word-spacing': line.wordSpacingPt }),
          'xml:space': 'preserve',
        },
        spans,
      ),
    );
  }

  for (const piece of line.pieces) {
    const { fill, opacity } = paint(piece.color);
    for (const rule of piece.rules) {
      out.push(
        ruleNode(rule, line.leftPt + piece.leftPt, line.baselinePt - piece.risePt, fill, opacity),
      );
    }
  }
  return out;
}

/**
 * A block as one `<g>`, turned about the frame's centre and scaled to points.
 *
 * The group is a **sibling** of the shape's own, never a child: the shape's
 * transform carries the mirror a `flipH` asks for, and text is never mirrored.
 */
export function textNodes(
  block: TextBlock,
  frame: { x: number; y: number; cx: number; cy: number },
  attrs: Record<string, string | number | null> = {},
): readonly SvgNode[] {
  if (block.lines.length === 0) return [];
  const halfX = frame.cx / 2;
  const halfY = frame.cy / 2;
  const parts = [`translate(${String(frame.x + halfX)} ${String(frame.y + halfY)})`];
  if (block.turnDeg !== 0) parts.push(`rotate(${String(block.turnDeg)})`);
  parts.push(`translate(${String(-halfX)} ${String(-halfY)})`);
  parts.push(`scale(${String(EMU_PER_POINT)})`);
  return [
    element(
      'g',
      { ...attrs, transform: parts.join(' ') },
      block.lines.flatMap((line) => lineNodes(line)),
    ),
  ];
}
