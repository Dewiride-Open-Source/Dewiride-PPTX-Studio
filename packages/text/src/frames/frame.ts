/**
 * The box a block of text goes in: insets, anchoring, columns and direction.
 *
 * Measured in experiment T6 - 595 probes, 21 questions, each fitting exactly one
 * candidate. `corpus/ground-truth/frames.json`, and
 * `docs/adr/phase-3-text/0032-anchors-insets-and-vertical-text.md`.
 */

import { TextError } from '../errors.js';

/** `ST_TextAnchoringType`. */
export type Anchor = 't' | 'ctr' | 'b' | 'just' | 'dist';

/** `ST_TextVerticalType`. */
export type VerticalText =
  'horz' | 'vert' | 'vert270' | 'wordArtVert' | 'eaVert' | 'mongolianVert' | 'wordArtVertRtl';

/** `ST_TextHorzOverflowType` and `ST_TextVertOverflowType`. */
export type VertOverflow = 'overflow' | 'ellipsis' | 'clip';

/** Which frame edge an axis runs from. */
export type Edge = 'left' | 'top' | 'right' | 'bottom';

/** `a:bodyPr`'s four insets, in points. */
export interface Insets {
  readonly left: number;
  readonly top: number;
  readonly right: number;
  readonly bottom: number;
}

/**
 * The inset defaults, per attribute rather than per element.
 *
 * Asymmetric, and each edge defaults on its own: a frame stating only `lIns`
 * still gets 3.6 above and below.
 */
export const DEFAULT_INSETS: Insets = { left: 7.2, top: 3.6, right: 7.2, bottom: 3.6 };

/** A rectangle in points, relative to the shape's own top-left corner. */
export interface Box {
  readonly leftPt: number;
  readonly topPt: number;
  readonly widthPt: number;
  readonly heightPt: number;
}

function checkSize(value: number, what: string): number {
  if (!Number.isFinite(value)) {
    throw new TextError('TEXT_FRAME', `${what} ${String(value)} is not a finite length`, what);
  }
  return value;
}

/**
 * The box the text is laid out in, after the insets.
 *
 * When two opposing insets exceed the frame the box collapses to the midpoint
 * between them rather than to the frame's centre - measured on an asymmetric
 * pair that separates the two.
 */
export function contentBox(frame: { widthPt: number; heightPt: number }, insets: Insets): Box {
  checkSize(frame.widthPt, 'frame width');
  checkSize(frame.heightPt, 'frame height');
  const near = (start: number, end: number, size: number): readonly [number, number] => {
    const far = size - end;
    if (start <= far) return [start, far - start];
    const middle = (start + far) / 2;
    return [middle, 0];
  };
  const [leftPt, widthPt] = near(insets.left, insets.right, frame.widthPt);
  const [topPt, heightPt] = near(insets.top, insets.bottom, frame.heightPt);
  return { leftPt, topPt, widthPt, heightPt };
}

/**
 * The insets as the frame sees them once the block has been turned.
 *
 * The insets do not turn with the text - `lIns` insets from the physical left
 * edge whatever `@vert` says, 7 of 7 - so a layout done in a frame that has been
 * turned a quarter has to be handed them the other way round. Without this a
 * `vert` frame with the default asymmetric insets is 17pt out at 200pt, which is
 * `vert-ins-vert` in the fixture. ADR 0039.
 */
export function turnedInsets(insets: Insets, quarterDeg: number): Insets {
  if (quarterDeg === 0) return insets;
  if (quarterDeg === 90) {
    return { left: insets.top, top: insets.right, right: insets.bottom, bottom: insets.left };
  }
  if (quarterDeg === 270) {
    return { left: insets.bottom, top: insets.left, right: insets.top, bottom: insets.right };
  }
  throw new TextError(
    'TEXT_FRAME',
    `${String(quarterDeg)} degrees is not a quarter turn a frame is laid out in`,
    String(quarterDeg),
  );
}

/* -------------------------------------------------------------------------- */
/* the two axes                                                               */
/* -------------------------------------------------------------------------- */

/**
 * Which axis the anchor and the alignment govern, and from which edge.
 *
 * `stacking` is the direction lines pile up in, which `@anchor` positions the
 * block along; the other axis is the direction text runs, which `@algn`
 * positions each line along.
 */
export interface FrameAxes {
  readonly stacking: 'x' | 'y';
  readonly anchorFrom: Edge;
  readonly alignFrom: Edge;
  /** GDI is asked for the `@`-prefixed vertical face. */
  readonly verticalFace: boolean;
  /** Each character occupies a line of its own. */
  readonly stackedGlyphs: boolean;
}

/**
 * The axes for each `ST_TextVerticalType`.
 *
 * The four that turn the text do not agree with each other: `vert` and `eaVert`
 * stack leftward from the right edge, `vert270` rightward from the left with the
 * text running upward, `mongolianVert` rightward with it running down.
 */
export const VERTICAL_AXES: Readonly<Record<VerticalText, FrameAxes>> = {
  horz: {
    stacking: 'y',
    anchorFrom: 'top',
    alignFrom: 'left',
    verticalFace: false,
    stackedGlyphs: false,
  },
  vert: {
    stacking: 'x',
    anchorFrom: 'right',
    alignFrom: 'top',
    verticalFace: false,
    stackedGlyphs: false,
  },
  vert270: {
    stacking: 'x',
    anchorFrom: 'left',
    alignFrom: 'bottom',
    verticalFace: false,
    stackedGlyphs: false,
  },
  eaVert: {
    stacking: 'x',
    anchorFrom: 'right',
    alignFrom: 'top',
    verticalFace: true,
    stackedGlyphs: false,
  },
  mongolianVert: {
    stacking: 'x',
    anchorFrom: 'left',
    alignFrom: 'top',
    verticalFace: true,
    stackedGlyphs: false,
  },
  wordArtVert: {
    stacking: 'x',
    anchorFrom: 'left',
    alignFrom: 'top',
    verticalFace: false,
    stackedGlyphs: true,
  },
  wordArtVertRtl: {
    stacking: 'x',
    anchorFrom: 'right',
    alignFrom: 'top',
    verticalFace: false,
    stackedGlyphs: true,
  },
};

/** The axes an absent or stated `@vert` selects. */
export function frameAxes(vert?: VerticalText): FrameAxes {
  const axes = VERTICAL_AXES[vert ?? 'horz'];
  if (axes === undefined) {
    throw new TextError(
      'TEXT_FRAME',
      `${String(vert)} is not an ST_TextVerticalType`,
      String(vert),
    );
  }
  return axes;
}

/* -------------------------------------------------------------------------- */
/* anchoring                                                                  */
/* -------------------------------------------------------------------------- */

/**
 * How far along its axis each anchor puts the block.
 *
 * `just` and `dist` are `b`. PowerPoint stretches nothing for either, lays both
 * out exactly where bottom anchoring puts them in all 190 readings, and its own
 * object model reports an anchor it has no name for.
 */
export function anchorFraction(anchor?: Anchor): number {
  switch (anchor ?? 't') {
    case 't':
      return 0;
    case 'ctr':
      return 0.5;
    case 'b':
    case 'just':
    case 'dist':
      return 1;
    default:
      throw new TextError(
        'TEXT_FRAME',
        `${String(anchor)} is not an ST_TextAnchoringType`,
        String(anchor),
      );
  }
}

/**
 * The block's offset from the start of its axis, in points.
 *
 * `from` is the edge the axis runs from, so a reversed axis is a reversed
 * fraction rather than a second formula.
 */
export function offsetAlong(
  from: Edge,
  fraction: number,
  availablePt: number,
  blockPt: number,
): number {
  if (!Number.isFinite(fraction) || fraction < 0 || fraction > 1) {
    throw new TextError('TEXT_FRAME', `${String(fraction)} is not a fraction`, String(fraction));
  }
  const room = availablePt - blockPt;
  return from === 'left' || from === 'top' ? fraction * room : (1 - fraction) * room;
}

/**
 * Where the block's near corner sits inside the content box, in points.
 *
 * `blockAcrossPt` is the block's thickness across the stacking axis and
 * `blockAlongPt` its extent along the line axis; `@anchorCtr` centres the block
 * along the line axis while `@anchor` positions it along the stacking one.
 */
export function blockOrigin(input: {
  readonly content: Box;
  readonly axes: FrameAxes;
  readonly anchor?: Anchor | undefined;
  readonly anchorCtr?: boolean | undefined;
  readonly blockAcrossPt: number;
  readonly blockAlongPt: number;
}): { readonly leftPt: number; readonly topPt: number } {
  const { content, axes } = input;
  const acrossBox = axes.stacking === 'x' ? content.widthPt : content.heightPt;
  const alongBox = axes.stacking === 'x' ? content.heightPt : content.widthPt;
  const across = offsetAlong(
    axes.anchorFrom,
    anchorFraction(input.anchor),
    acrossBox,
    input.blockAcrossPt,
  );
  const along =
    input.anchorCtr === true
      ? offsetAlong(axes.alignFrom, 0.5, alongBox, input.blockAlongPt)
      : offsetAlong(axes.alignFrom, 0, alongBox, input.blockAlongPt);
  return axes.stacking === 'x'
    ? { leftPt: content.leftPt + across, topPt: content.topPt + along }
    : { leftPt: content.leftPt + along, topPt: content.topPt + across };
}

/* -------------------------------------------------------------------------- */
/* the empty paragraph                                                        */
/* -------------------------------------------------------------------------- */

/**
 * The height of an `a:p` with no runs in it, in points.
 *
 * `a:endParaRPr/@sz` beats a `a:defRPr` that disagrees, both ways round.
 */
export function emptyParagraphHeight(sz: number, lnSpcMultiple = 1): number {
  if (!Number.isFinite(sz) || sz <= 0) {
    throw new TextError('TEXT_FRAME', `font size ${String(sz)} is not a positive size`, String(sz));
  }
  if (!Number.isFinite(lnSpcMultiple) || lnSpcMultiple < 0) {
    throw new TextError(
      'TEXT_FRAME',
      `line spacing ${String(lnSpcMultiple)} is not a non-negative multiple`,
      String(lnSpcMultiple),
    );
  }
  return lnSpcMultiple * 1.2 * (sz / 100);
}

/* -------------------------------------------------------------------------- */
/* columns                                                                    */
/* -------------------------------------------------------------------------- */

/** `ST_TextColumnCount`. PowerPoint repairs a package outside it. */
export const MAX_COLUMNS = 16;

/**
 * The box of one column, in points.
 *
 * `rtlCol` fills from the right, which is where the first paragraph lands.
 */
export function columnBox(input: {
  readonly content: Box;
  readonly index: number;
  readonly numCol?: number | undefined;
  readonly spcColPt?: number | undefined;
  readonly rtlCol?: boolean | undefined;
}): Box {
  const numCol = input.numCol ?? 1;
  if (!Number.isInteger(numCol) || numCol < 1 || numCol > MAX_COLUMNS) {
    throw new TextError(
      'TEXT_FRAME',
      `a:bodyPr/@numCol ${String(numCol)} is outside ST_TextColumnCount's 1..${String(MAX_COLUMNS)}`,
      String(numCol),
    );
  }
  const spcCol = input.spcColPt ?? 0;
  if (!Number.isFinite(spcCol) || spcCol < 0) {
    throw new TextError(
      'TEXT_FRAME',
      `a:bodyPr/@spcCol ${String(spcCol)} is not a non-negative length`,
      String(spcCol),
    );
  }
  if (!Number.isInteger(input.index) || input.index < 0 || input.index >= numCol) {
    throw new TextError(
      'TEXT_FRAME',
      `column ${String(input.index)} is outside the ${String(numCol)} this frame has`,
      String(input.index),
    );
  }
  const width = (input.content.widthPt - (numCol - 1) * spcCol) / numCol;
  const nth = input.rtlCol === true ? numCol - 1 - input.index : input.index;
  return {
    leftPt: input.content.leftPt + nth * (width + spcCol),
    topPt: input.content.topPt,
    widthPt: width,
    heightPt: input.content.heightPt,
  };
}

/* -------------------------------------------------------------------------- */
/* overflow                                                                   */
/* -------------------------------------------------------------------------- */

/** The character `@vertOverflow="ellipsis"` substitutes for a line that does not fit. */
export const ELLIPSIS = '…';

/** One laid-out line, as much of it as the overflow rule needs. */
export interface OverflowLine {
  readonly topPt: number;
  readonly heightPt: number;
}

/** What an overflow rule leaves on the slide. */
export interface DrawnLines {
  /** How many of the lines are drawn as they are. */
  readonly count: number;
  /** An ellipsis is drawn as the next line, in place of the one that did not fit. */
  readonly ellipsis: boolean;
}

/**
 * Which lines `@vertOverflow` draws.
 *
 * The block is laid out in full either way - neither `ellipsis` nor `clip`
 * changes its height or where the anchor puts it. `ellipsis` replaces the whole
 * first line that does not fit with a lone U+2026 rather than truncating it.
 */
export function drawnLines(
  lines: readonly OverflowLine[],
  availableHeightPt: number,
  overflow?: VertOverflow,
): DrawnLines {
  const kind = overflow ?? 'overflow';
  if (kind === 'overflow') return { count: lines.length, ellipsis: false };
  if (kind !== 'ellipsis' && kind !== 'clip') {
    throw new TextError(
      'TEXT_FRAME',
      `${String(overflow)} is not an ST_TextVertOverflowType`,
      String(overflow),
    );
  }
  // Strictly inside: a box exactly two line boxes tall draws one line, not two.
  let count = 0;
  for (const line of lines) {
    if (line.topPt + line.heightPt >= availableHeightPt) break;
    count++;
  }
  return { count, ellipsis: kind === 'ellipsis' && count < lines.length };
}

/* -------------------------------------------------------------------------- */
/* upright glyphs                                                             */
/* -------------------------------------------------------------------------- */

/** Where an upright glyph's pen goes, in the line's own coordinates. */
export interface UprightPen {
  /** The glyph's alphabetic baseline, from the leading edge of its cell. */
  readonly alongPt: number;
  /** The glyph's own x origin, from the leading edge of the line box. */
  readonly acrossPt: number;
}

/**
 * The pen an upright glyph is drawn from, turned a quarter back out of the line.
 *
 * Across the line, the glyph's advance box is centred: 1.0997 and 1.0998 of the
 * em measured on MS Gothic and SimSun against 1.1 predicted, Yu Gothic 1.0613.
 * Along it, the baseline is the em box's ascent from the cell's leading edge,
 * which is one less the ideographic baseline. ADR 0039.
 */
export function uprightPen(cell: {
  readonly sizePt: number;
  readonly advancePt: number;
  readonly lineHeightPt: number;
  readonly ideographic: number;
}): UprightPen {
  const { sizePt, advancePt, lineHeightPt, ideographic } = cell;
  if (!Number.isFinite(sizePt) || sizePt < 0 || !Number.isFinite(advancePt) || advancePt < 0) {
    throw new TextError(
      'TEXT_FRAME',
      `a cell of ${String(advancePt)}pt at ${String(sizePt)}pt is not a glyph cell`,
      String(sizePt),
    );
  }
  if (!Number.isFinite(ideographic) || ideographic < 0 || ideographic >= 1) {
    throw new TextError(
      'TEXT_FRAME',
      `an ideographic baseline of ${String(ideographic)} is not a fraction of the em`,
      String(ideographic),
    );
  }
  return {
    alongPt: (1 - ideographic) * sizePt,
    acrossPt: (lineHeightPt + advancePt) / 2,
  };
}
