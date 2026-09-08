/**
 * A resolved text body into positioned lines: the one pass both renderers read.
 *
 * `render-svg` turns this into `<text>` and `<tspan>` and `render-dom` turns the
 * same object into an absolutely-positioned HTML layer. Neither of them decides
 * anything: by the time a `TextLine` exists the block has been turned, anchored,
 * broken, aligned and given a baseline, so the two cannot disagree about where a
 * word is. Every rule here was measured in experiment T8 -
 * `corpus/ground-truth/text-rendering.json`, ADR 0034.
 *
 * ## The turn is not a counter-flip
 *
 * `flipH` does nothing to text at all and `flipV` turns the whole block 180
 * degrees about the frame centre, so the two together are a half turn and the
 * glyphs are never mirrored. 20 of 20; counter-flipping both axes, which is what
 * the plan assumed, fits 10.
 */

import {
  blockOrigin,
  columnBox,
  contentBox,
  createFaceBoxProbe,
  createCanvasMeasurer,
  cssFamily,
  drawnLines,
  frameAxes,
  lineAdvance,
  scriptSlotOf,
  turnedInsets,
  strikeRules,
  toCodePoints,
  underlineRules,
  uprightPen,
  wrapText,
  type Box,
  type DrawnRule,
  type Edge,
  type FaceBox,
  type FaceBoxProbe,
  type FaceRules,
  type Insets,
  type LineSpacing,
  type RunFont,
  type ScriptSlot,
  type TextMeasurer,
} from '@pptx-studio/text';
import type { Rgba } from '@pptx-studio/paint';

import { RenderError } from '../errors.js';
import { swapsExtents } from '../transform.js';

import type { ResolvedFrame, ResolvedParagraph, ResolvedRun, ResolvedText } from './resolve.js';

/** One stretch of characters drawn in one font at one place. */
export interface TextPiece {
  /** The characters as the file has them, before `@cap`. */
  readonly source: string;
  /** The characters to draw, after `@cap` has been applied. */
  readonly text: string;
  readonly font: RunFont;
  /** The family quoted for CSS, computed once so two emitters cannot differ. */
  readonly cssFamily: string;
  readonly color: Rgba | null;
  readonly highlight: Rgba | null;
  /** Points from the line's own left edge. */
  readonly leftPt: number;
  readonly widthPt: number;
  /** Points above the line's baseline; a subscript is negative. */
  readonly risePt: number;
  /** The rules this piece draws, already placed. */
  readonly rules: readonly PieceRule[];
  /**
   * The glyphs of this piece, each placed and turned a quarter out of the line.
   *
   * Empty for every piece that flows, which is all of them outside an `eaVert`
   * frame. A piece with entries draws these and not `text`.
   */
  readonly upright: readonly UprightGlyph[];
}

/** One glyph stood upright inside a turned line, in the line's coordinates. */
export interface UprightGlyph {
  readonly text: string;
  /** The glyph's alphabetic baseline, from the line's own left edge. */
  readonly alongPt: number;
  /** The glyph's own x origin, from the line's own top edge. */
  readonly acrossPt: number;
  /**
   * The `line-height` that lands a browser's baseline `sizePt` below the top.
   *
   * The SVG emitter places a baseline as a coordinate and ignores this; the HTML
   * layer cannot, and a glyph drawn from its own pen is outside the line's own
   * strut. `sizePt` is the anchor because a single glyph has no other.
   */
  readonly strutPt: number;
}

/**
 * One drawn rule, in points from the line's left edge and its baseline.
 *
 * Placed here rather than in an emitter because `u="words"` needs the advance of
 * each word, and the measurer lives in this pass - two emitters measuring the
 * same word twice is exactly the drift the shared layout exists to prevent.
 */
export interface PieceRule {
  readonly leftPt: number;
  readonly widthPt: number;
  /** The rule's top edge below the baseline; a strikethrough's is negative. */
  readonly topPt: number;
  readonly thickness: number;
  readonly pattern: DrawnRule['pattern'];
  readonly wavy: boolean;
}

export interface TextLine {
  /** Points from the frame's top-left corner, before the block is turned. */
  readonly topPt: number;
  readonly heightPt: number;
  /** The baseline, from the same corner. */
  readonly baselinePt: number;
  readonly leftPt: number;
  readonly widthPt: number;
  /** The largest size on the line, which is what its box follows. */
  readonly sizePt: number;
  /**
   * The CSS `line-height` that lands a browser's baseline on `baselinePt`.
   *
   * A browser puts the baseline of an inline box a half-leading plus an ascent
   * below the line box; solving that for the share PowerPoint uses is one line
   * of arithmetic, done here so the HTML layer and the SVG cannot answer it
   * differently. It lands within one CSS pixel and no closer, because Chromium
   * rounds the font's ascent to a whole pixel before it lays a line out.
   */
  readonly strutPt: number;
  /**
   * Extra advance on every space, in points, from a justifying alignment.
   *
   * A width rather than a set of positions: an emitter that repositioned each
   * word would break the shaping of the line, and T8 measured that a line is
   * shaped whole.
   */
  readonly wordSpacingPt: number;
  readonly pieces: readonly TextPiece[];
}

export interface TextBlock {
  /** Degrees clockwise about the frame centre, after every turn is composed. */
  readonly turnDeg: number;
  /** The rectangle the block was laid out in, in the shape's own space. */
  readonly boxPt: Box;
  readonly lines: readonly TextLine[];
}

/** What `layoutText` needs besides the resolved text. */
export interface LayoutTextOptions {
  /** The shape's own extent, in points. */
  readonly widthPt: number;
  readonly heightPt: number;
  /** `a:xfrm/@rot`, in degrees clockwise, as `Frame` already has it. */
  readonly rot: number;
  readonly flipH: boolean;
  readonly flipV: boolean;
  readonly measurer?: TextMeasurer | undefined;
  readonly faceBox?: FaceBoxProbe | undefined;
  /** The rules for a typeface, so a caller decides what an unmeasured one gets. */
  readonly rulesFor: (typeface: string) => FaceRules;
}

/** A superscript or subscript is drawn at two thirds of the size it asked for. */
export const SHIFT_SIZE_RATIO = 2 / 3;

/** `cap="small"` draws the uppercased lowercase at four fifths of the size. */
export const SMALL_CAPS_RATIO = 0.8;

/**
 * The width a line is wrapped at when `a:bodyPr/@wrap="none"`.
 *
 * `wrapText` refuses an infinite width, and rightly - a width that is not a
 * number is a bug everywhere else - so this is a slide's diagonal times a
 * thousand, which no line of text reaches.
 */
const UNWRAPPED_WIDTH_PT = 1e6;

/** The alignments that stretch a line to the full column width. */
const JUSTIFYING = new Set(['just', 'justLow', 'dist', 'thaiDist']);

/** The one alignment that stretches the last line of a paragraph too. */
const JUSTIFIES_LAST = new Set(['dist']);

/** The directions this package lays out and draws. */
const DRAWN_DIRECTIONS = new Set(['horz', 'vert', 'vert270', 'eaVert', 'mongolianVert']);

/**
 * The quarter each direction turns the block, and the axes it is laid out on.
 *
 * The block is laid out horizontally in a frame with its extents swapped and
 * then turned, so these are the axes of the **unturned** layout. A quarter turn
 * clockwise sends its top edge to the right and its left edge to the top, which
 * is `vert` and `eaVert`; `vert270` sends top to left and left to bottom, which
 * is its own pair; `mongolianVert` runs downward like `vert` but stacks from the
 * other side, so it turns with `vert` and anchors from the opposite edge.
 */
const TURNS: Readonly<Record<string, { quarter: number; anchorFrom: Edge }>> = {
  horz: { quarter: 0, anchorFrom: 'top' },
  vert: { quarter: 90, anchorFrom: 'top' },
  eaVert: { quarter: 90, anchorFrom: 'top' },
  vert270: { quarter: 270, anchorFrom: 'top' },
  mongolianVert: { quarter: 90, anchorFrom: 'bottom' },
};

function turn(value: number): number {
  return ((value % 360) + 360) % 360;
}

/**
 * The angle the block is drawn at, and the box it is laid out in.
 *
 * `@upright="1"` drops the shape's own turn and lays the text out in the frame's
 * extents swapped, when that turn snaps to a quadrant - which is 2.10's
 * quadrant-snap rule reappearing on a different question.
 */
export function textTurn(
  frame: ResolvedFrame,
  options: Pick<LayoutTextOptions, 'widthPt' | 'heightPt' | 'rot' | 'flipH' | 'flipV'>,
): { turnDeg: number; boxPt: Box } {
  const shapeTurn = turn(options.rot + (options.flipV ? 180 : 0));
  const turnDeg = turn((frame.upright ? 0 : shapeTurn) + frame.bodyRotation);
  const swap = frame.upright && swapsExtents(shapeTurn);
  const widthPt = swap ? options.heightPt : options.widthPt;
  const heightPt = swap ? options.widthPt : options.heightPt;
  return {
    turnDeg,
    boxPt: {
      leftPt: (options.widthPt - widthPt) / 2,
      topPt: (options.heightPt - heightPt) / 2,
      widthPt,
      heightPt,
    },
  };
}

/** The characters a run draws, and the size each stretch is drawn at. */
interface CapStretch {
  readonly text: string;
  readonly sizeRatio: number;
}

/**
 * `@cap` applied to a run's text.
 *
 * `all` uppercases at the full size and `small` uppercases the lowercase at four
 * fifths, leaving what was already capital alone - which is why this returns a
 * list and not a string.
 */
export function capStretches(text: string, caps: string | undefined): readonly CapStretch[] {
  if (caps === undefined || caps === 'none') return [{ text, sizeRatio: 1 }];
  if (caps === 'all') return [{ text: text.toUpperCase(), sizeRatio: 1 }];
  const out: CapStretch[] = [];
  for (const char of text) {
    const upper = char.toUpperCase();
    const ratio = upper !== char ? SMALL_CAPS_RATIO : 1;
    const last = out.at(-1);
    if (last !== undefined && last.sizeRatio === ratio) {
      out[out.length - 1] = { text: last.text + upper, sizeRatio: ratio };
    } else out.push({ text: upper, sizeRatio: ratio });
  }
  return out.length === 0 ? [{ text: '', sizeRatio: 1 }] : out;
}

/** The font a piece is actually set in, after `@cap` and `@baseline`. */
function pieceFont(run: ResolvedRun, sizeRatio: number, scale: number): RunFont {
  const shift = run.baseline === 0 ? 1 : SHIFT_SIZE_RATIO;
  return { ...run.font, sz: run.font.sz * sizeRatio * shift * scale };
}

/** The characters of a paragraph, as one string, with a map back to the runs. */
interface Cell {
  readonly run: ResolvedRun;
  readonly sizeRatio: number;
  readonly text: string;
  /** `text` by code point, because `wrapText` returns code point indices. */
  readonly points: readonly string[];
  readonly source: string;
}

function cellsOf(paragraph: ResolvedParagraph): readonly Cell[] {
  const out: Cell[] = [];
  for (const run of paragraph.runs) {
    for (const stretch of capStretches(run.text, run.caps)) {
      if (stretch.text.length === 0) continue;
      out.push({
        run,
        sizeRatio: stretch.sizeRatio,
        text: stretch.text,
        points: toCodePoints(stretch.text),
        source: run.text,
      });
    }
  }
  return out;
}

/**
 * The measurer, wrapped so a range of the paragraph's text can be measured.
 *
 * A line is shaped as one string and a run boundary splits only the drawing
 * call, so a kern crosses it - 6 of 6. Measuring cell by cell would lose exactly
 * that kern, so a range spanning a boundary is measured as its parts and the
 * error is recorded rather than hidden: it is one kern pair per boundary.
 */
function rangeMeasurer(
  cells: readonly Cell[],
  measurer: TextMeasurer,
  scale: number,
): { text: string; measure: (from: number, to: number) => number } {
  const text = cells.map((cell) => cell.text).join('');
  const starts: number[] = [];
  let at = 0;
  for (const cell of cells) {
    starts.push(at);
    at += cell.points.length;
  }
  return {
    text,
    measure(from: number, to: number): number {
      let width = 0;
      cells.forEach((cell, index) => {
        const start = starts[index] ?? 0;
        const end = start + cell.points.length;
        const lo = Math.max(from, start);
        const hi = Math.min(to, end);
        if (hi <= lo) return;
        width += measurer.measure(
          cell.points.slice(lo - start, hi - start).join(''),
          pieceFont(cell.run, cell.sizeRatio, scale),
        ).width;
      });
      return width;
    },
  };
}

/**
 * Where one cell's share of a line splits into pieces, in code point indices.
 *
 * One stretch unless the frame stands its East Asian glyphs upright, which has
 * to draw each of them from its own pen and so cannot share a piece with the
 * Latin beside it. The split costs the kern across that boundary, which is the
 * same trade the cell boundary already makes.
 */
function stretchesOf(
  cell: Cell,
  lo: number,
  hi: number,
  start: number,
  splitScripts: boolean,
): readonly (readonly [number, number])[] {
  if (!splitScripts) return [[lo, hi]];
  const out: [number, number][] = [];
  let slot: ScriptSlot | null = null;
  for (let at = lo; at < hi; at += 1) {
    const point = cell.points[at - start] ?? '';
    const here = scriptSlotOf(point.codePointAt(0) ?? 0);
    const last = out.at(-1);
    if (last === undefined || here !== slot) out.push([at, at + 1]);
    else last[1] = at + 1;
    slot = here;
  }
  return out;
}

/** A piece whose characters all sit in the `a:ea` slot, so all stand upright. */
function standsUpright(text: string): boolean {
  const first = text.codePointAt(0);
  return first !== undefined && scriptSlotOf(first) === 'ea';
}

/** The pieces of one line, positioned from its own left edge. */
function piecesOf(
  cells: readonly Cell[],
  from: number,
  to: number,
  measure: (from: number, to: number) => number,
  scale: number,
  rulesFor: (typeface: string) => FaceRules,
  splitScripts: boolean,
): { pieces: readonly TextPiece[]; widthPt: number; sizePt: number } {
  const pieces: TextPiece[] = [];
  let leftPt = 0;
  let sizePt = 0;
  let at = 0;
  for (const cell of cells) {
    const start = at;
    const end = at + cell.points.length;
    at = end;
    const cellLo = Math.max(from, start);
    const cellHi = Math.min(to, end);
    if (cellHi <= cellLo) continue;
    const font = pieceFont(cell.run, cell.sizeRatio, scale);
    const sizeOfPiece = font.sz / 100;
    const faces = rulesFor(font.family);
    const drawn = [
      ...underlineRules(cell.run.underline, sizeOfPiece, faces),
      ...strikeRules(cell.run.strike, sizeOfPiece, faces),
    ];
    for (const [lo, hi] of stretchesOf(cell, cellLo, cellHi, start, splitScripts)) {
      const text = cell.points.slice(lo - start, hi - start).join('');
      const widthPt = measure(lo, hi);
      pieces.push({
        source: cell.source,
        text,
        font,
        cssFamily: cssFamily(font.family),
        color: cell.run.color,
        highlight: cell.run.highlight,
        leftPt,
        widthPt,
        risePt: ((cell.run.baseline / 100000) * cell.run.font.sz * scale) / 100,
        rules: placeRules(drawn, text, cell.run.underline, lo, hi, measure),
        upright: [],
      });
      leftPt += widthPt;
    }
    // The line box follows the largest size on the line, whatever the order,
    // and a shifted or small-capped run is measured at the size it asked for.
    sizePt = Math.max(sizePt, (cell.run.font.sz * scale) / 100);
  }
  return { pieces, widthPt: leftPt, sizePt };
}

/**
 * The stretches of a piece a rule runs under, in characters.
 *
 * A plain rule stops at the last glyph, so trailing spaces are dropped;
 * `u="words"` drops every space, which is why this is a list.
 */
function ruleStretches(
  text: string,
  underline: string | undefined,
): readonly (readonly [number, number])[] {
  if (underline === 'words') {
    const out: [number, number][] = [];
    let start: number | null = null;
    for (let i = 0; i < text.length; i += 1) {
      const isSpace = text[i] === ' ';
      if (!isSpace && start === null) start = i;
      if (isSpace && start !== null) {
        out.push([start, i]);
        start = null;
      }
    }
    if (start !== null) out.push([start, text.length]);
    return out;
  }
  const end = text.replace(TRAILING_SPACES, '').length;
  return end === 0 ? [] : [[0, end]];
}

/** Every rule of a piece, placed over the characters it actually covers. */
function placeRules(
  drawn: readonly DrawnRule[],
  text: string,
  underline: string | undefined,
  lo: number,
  hi: number,
  measure: (from: number, to: number) => number,
): readonly PieceRule[] {
  if (drawn.length === 0) return [];
  const out: PieceRule[] = [];
  for (const [from, to] of ruleStretches(text, underline)) {
    const leftPt = measure(lo, lo + from);
    const widthPt = measure(lo + from, Math.min(lo + to, hi));
    for (const rule of drawn) {
      out.push({
        leftPt,
        widthPt,
        topPt: rule.top,
        thickness: rule.thickness,
        pattern: rule.pattern,
        wavy: rule.wavy,
      });
    }
  }
  return out;
}

/** The spaces a run ends with, which no rule is drawn under. */
const TRAILING_SPACES = / +$/;

/**
 * Where a line starts inside its column, in points.
 *
 * Measured 21 of 21: the width an alignment anchors excludes whatever spaces the
 * line ends with, which hang past the edge. Counting them fits 19, and the two
 * it misses are the wrapped rows every real deck is made of.
 */
export function alignOffset(align: string, columnWidthPt: number, anchorWidthPt: number): number {
  if (align === 'ctr') return (columnWidthPt - anchorWidthPt) / 2;
  if (align === 'r') return columnWidthPt - anchorWidthPt;
  return 0;
}

/** Whether this alignment stretches this line to the full column. */
export function stretches(align: string, lastLine: boolean): boolean {
  if (!JUSTIFYING.has(align)) return false;
  return JUSTIFIES_LAST.has(align) || !lastLine;
}

/** `a:spcBef` and `a:spcAft`, which follow the same two rules as `a:lnSpc`. */
function spacingPoints(spacing: LineSpacing, sizePt: number): number {
  return sizePt <= 0 ? 0 : lineAdvance(sizePt * 100, spacing);
}

/**
 * Lay a resolved body out inside a shape.
 *
 * The measurer and the face-box probe are taken rather than made so that a test
 * can drive the whole pass from literals, which is the only way to assert a
 * layout without a browser and a font.
 */
export function layoutText(text: ResolvedText, options: LayoutTextOptions): TextBlock {
  const { frame } = text;
  if (!DRAWN_DIRECTIONS.has(frame.vertical)) {
    throw new RenderError(
      'RENDER_TEXT_UNSUPPORTED',
      `a:bodyPr/@vert="${frame.vertical}" is laid out but not yet drawn`,
      frame.vertical,
    );
  }

  const measurer = options.measurer ?? createCanvasMeasurer();
  const faceBox = options.faceBox ?? createFaceBoxProbe();
  const { turnDeg, boxPt } = textTurn(frame, options);

  const axes = frameAxes(frame.vertical);
  const uprightGlyphs = axes.verticalFace;
  const { quarter, anchorFrom } = TURNS[frame.vertical] ?? { quarter: 0, anchorFrom: 'top' };
  const swapped = quarter !== 0;
  const laidOut: Box = swapped
    ? {
        leftPt: boxPt.leftPt + (boxPt.widthPt - boxPt.heightPt) / 2,
        topPt: boxPt.topPt + (boxPt.heightPt - boxPt.widthPt) / 2,
        widthPt: boxPt.heightPt,
        heightPt: boxPt.widthPt,
      }
    : boxPt;

  const insets: Insets = turnedInsets(frame.insets, quarter);
  const content = contentBox({ widthPt: laidOut.widthPt, heightPt: laidOut.heightPt }, insets);
  const column = columnBox({
    content,
    index: 0,
    numCol: frame.columns,
    spcColPt: frame.columnSpacing,
    rtlCol: frame.rtlColumns,
  });

  const scale = frame.fontScale;
  const lines: TextLine[] = [];
  let cursor = 0;

  text.paragraphs.forEach((paragraph, index) => {
    const cells = cellsOf(paragraph);
    const { text: whole, measure } = rangeMeasurer(cells, measurer, scale);
    const marginLeft = paragraph.marginLeft;
    const firstIndent = paragraph.indent;

    if (index > 0) cursor += spacingPoints(paragraph.spaceBefore, lineSizeOf(cells, scale));

    const wrapWidth = frame.wrap === 'none' ? UNWRAPPED_WIDTH_PT : column.widthPt - marginLeft;
    const boxes =
      whole.length === 0
        ? [{ start: 0, end: 0, measuredEnd: 0 }]
        : wrapText({
            text: whole,
            widthPt: Math.max(wrapWidth, 1),
            measure,
            hyphenWidthPt: measure(0, 0),
          });

    boxes.forEach((box, lineIndex) => {
      const { pieces, widthPt, sizePt } = piecesOf(
        cells,
        box.start,
        box.end,
        measure,
        scale,
        options.rulesFor,
        uprightGlyphs,
      );
      const size = sizePt === 0 ? (paragraph.endRun.font.sz * scale) / 100 : sizePt;
      // `lineAdvance` takes the size in hundredths, as `a:rPr/@sz` writes it.
      const heightPt =
        lineAdvance(size * 100, paragraph.lineSpacing) * (1 - frame.lineSpaceReduction);
      const face = faceBox.box(dominantFamily(pieces, paragraph));
      // `measuredEnd` already ends before the spaces a line breaks at (3.3),
      // and an alignment does not count them either - 21 of 21.
      const anchorWidth = measure(box.start, box.measuredEnd);
      const indent = lineIndex === 0 ? firstIndent : 0;
      const lastLine = lineIndex === boxes.length - 1;
      const stretched = stretches(paragraph.align, lastLine);
      const available = column.widthPt - marginLeft - indent;
      const offset = stretched ? 0 : alignOffset(paragraph.align, available, anchorWidth);
      const drop = heightPt * baselineShareOf(face);
      lines.push({
        topPt: cursor,
        heightPt,
        baselinePt: cursor + drop,
        leftPt: marginLeft + indent + offset,
        widthPt: stretched ? available : widthPt,
        sizePt: size,
        strutPt: strutHeight(drop, face, size),
        wordSpacingPt: stretched ? wordSpacing(pieces, available, anchorWidth) : 0,
        pieces: uprightGlyphs ? stoodUpright(pieces, heightPt, measurer, faceBox) : pieces,
      });
      cursor += heightPt;
    });

    if (index < text.paragraphs.length - 1) {
      cursor += spacingPoints(paragraph.spaceAfter, lineSizeOf(cells, scale));
    }
  });

  const drawn = drawnLines(
    lines.map((line) => ({ topPt: line.topPt, heightPt: line.heightPt })),
    column.heightPt,
    frame.vertOverflow,
  );
  const kept = lines.slice(0, drawn.count);

  const origin = blockOrigin({
    content: column,
    axes: { ...frameAxes('horz'), anchorFrom },
    anchor: frame.anchor,
    anchorCtr: frame.anchorCtr,
    blockAcrossPt: cursor,
    blockAlongPt: Math.max(0, ...kept.map((line) => line.leftPt + line.widthPt)),
  });

  return {
    turnDeg: turn(turnDeg + quarter),
    boxPt: laidOut,
    lines: kept.map((line) => ({
      ...line,
      topPt: line.topPt + origin.topPt + laidOut.topPt,
      baselinePt: line.baselinePt + origin.topPt + laidOut.topPt,
      leftPt: line.leftPt + origin.leftPt + laidOut.leftPt,
    })),
  };
}

function baselineShareOf(face: FaceBox): number {
  return face.ascent / (face.ascent + face.descent);
}

/**
 * Every `a:ea` piece of a line given one pen per glyph, turned back upright.
 *
 * PowerPoint draws these from the `@`-prefixed vertical variant of the face, so
 * one drawing call carries the whole run; a browser has no such face and has to
 * place each glyph itself. The advance is the face's own and is not touched -
 * the block, the wrap and the line pitch are `vert`'s to the thousandth of a
 * point. ADR 0039.
 */
function stoodUpright(
  pieces: readonly TextPiece[],
  lineHeightPt: number,
  measurer: TextMeasurer,
  faceBox: FaceBoxProbe,
): readonly TextPiece[] {
  return pieces.map((piece) => {
    if (!standsUpright(piece.text)) return piece;
    if (piece.rules.length > 0) {
      throw new RenderError(
        'RENDER_TEXT_UNSUPPORTED',
        'an underline or strikethrough on upright East Asian text is not measured',
        piece.text,
      );
    }
    const sizePt = piece.font.sz / 100;
    const ideographic = faceBox.box(piece.font.family).ideographic;
    const glyphs: UprightGlyph[] = [];
    let alongPt = piece.leftPt;
    for (const glyph of toCodePoints(piece.text)) {
      const advancePt = measurer.measure(glyph, piece.font).width;
      const pen = uprightPen({ sizePt, advancePt, lineHeightPt, ideographic });
      glyphs.push({
        text: glyph,
        alongPt: alongPt + pen.alongPt,
        acrossPt: pen.acrossPt,
        strutPt: strutHeight(sizePt, faceBox.box(piece.font.family), sizePt),
      });
      alongPt += advancePt;
    }
    return { ...piece, upright: glyphs };
  });
}

/**
 * The `line-height` that puts a browser's baseline at `dropPt`.
 *
 * A browser's baseline sits `(L - (A + D)S) / 2 + AS` below the line box, so
 * `L = 2 drop + (D - A) S`. Clamped at zero: a face whose ascent is more than
 * twice its share of the box would otherwise ask for a negative height.
 */
export function strutHeight(dropPt: number, face: FaceBox, sizePt: number): number {
  return Math.max(0, 2 * dropPt + (face.descent - face.ascent) * sizePt);
}

/** The family the line's baseline follows: the one the largest run is set in. */
function dominantFamily(pieces: readonly TextPiece[], paragraph: ResolvedParagraph): string {
  let best: TextPiece | null = null;
  for (const piece of pieces) {
    if (best === null || piece.font.sz > best.font.sz) best = piece;
  }
  return best?.font.family ?? paragraph.endRun.font.family;
}

/** The largest size in a paragraph, for the spacing a paragraph break adds. */
function lineSizeOf(cells: readonly Cell[], scale: number): number {
  let size = 0;
  for (const cell of cells) size = Math.max(size, (cell.run.font.sz * scale) / 100);
  return size;
}

/**
 * The extra advance a justified line puts on each of its spaces.
 *
 * Only the spaces stretch, so a line with none - a single long word, or CJK -
 * is left alone rather than having its letters spread.
 */
function wordSpacing(
  pieces: readonly TextPiece[],
  availablePt: number,
  anchorWidthPt: number,
): number {
  let spaces = 0;
  for (const piece of pieces) {
    for (const char of piece.text) if (char === ' ') spaces += 1;
  }
  if (spaces === 0 || availablePt <= anchorWidthPt) return 0;
  return (availablePt - anchorWidthPt) / spaces;
}
