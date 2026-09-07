/**
 * A shape's `p:txBody` into everything a layout pass needs, and nothing else.
 *
 * The cascade of 3.1 answers one property at a time; drawing a slide asks for
 * every property of every run, so this walks it once and hands over concrete
 * values. Nothing here measures and nothing here positions - that is
 * `layout.ts`, and keeping the two apart is what lets a test drive the layout
 * from a literal instead of from a package.
 */

import {
  resolveAnchor,
  resolveAnchorCtr,
  resolveBody,
  resolveColumns,
  resolveInsets,
  resolveIndent,
  resolveMarginLeft,
  resolveParagraph,
  resolveRun,
  resolveLatinTypeface,
  resolveSize,
  resolveVertical,
  resolveWrap,
  type Caps,
  type ListStyle,
  type Paragraph,
  type Spacing,
  type Strike,
  type TextAlign,
  type TextAnchor,
  type TextBody,
  type TextContent,
  type TextContext,
  type Underline,
  type VerticalText,
} from '@pptx-studio/model';
import { resolveColor, type ColorContext, type Fill, type Rgba } from '@pptx-studio/paint';
import type { LineSpacing, RunFont } from '@pptx-studio/text';

import { RenderError } from '../errors.js';
import type { Placed } from '../layout.js';

/** EMU to points, which is the unit everything in this directory works in. */
const EMU_PER_POINT = 12700;

/** One piece of a paragraph, with every property already answered. */
export interface ResolvedRun {
  readonly text: string;
  readonly font: RunFont;
  /** `null` where nothing in the cascade named a colour. */
  readonly color: Rgba | null;
  readonly highlight: Rgba | null;
  readonly underline: Underline | undefined;
  readonly strike: Strike | undefined;
  readonly caps: Caps | undefined;
  /** `a:rPr/@baseline`, thousandths of a percent. */
  readonly baseline: number;
  /** A hard `a:br` ends the line after this run. */
  readonly hardBreak: boolean;
}

export interface ResolvedParagraph {
  readonly level: number;
  readonly align: TextAlign;
  /** `@marL` and `@indent`, in points. */
  readonly marginLeft: number;
  readonly indent: number;
  readonly lineSpacing: LineSpacing;
  readonly spaceBefore: LineSpacing;
  readonly spaceAfter: LineSpacing;
  readonly runs: readonly ResolvedRun[];
  /** The properties of the paragraph mark, which set an empty line's height. */
  readonly endRun: ResolvedRun;
}

export interface ResolvedFrame {
  readonly anchor: TextAnchor;
  readonly anchorCtr: boolean;
  readonly vertical: VerticalText;
  readonly wrap: 'none' | 'square';
  /** In points, each edge on its own. */
  readonly insets: { left: number; top: number; right: number; bottom: number };
  readonly columns: number;
  readonly columnSpacing: number;
  readonly rtlColumns: boolean;
  /** `a:bodyPr/@rot` in degrees, positive clockwise. */
  readonly bodyRotation: number;
  readonly upright: boolean;
  readonly vertOverflow: 'overflow' | 'ellipsis' | 'clip';
  /** `a:normAutofit/@fontScale` as a multiplier, already quantised. */
  readonly fontScale: number;
  readonly lineSpaceReduction: number;
}

export interface ResolvedText {
  readonly frame: ResolvedFrame;
  readonly paragraphs: readonly ResolvedParagraph[];
}

/** Single spacing, in the thousandths of a percent `LineSpacing` holds. */
const DEFAULT_SPACING: LineSpacing = { kind: 'percent', value: 100000 };

/** No space at all, in the hundredths of a point `LineSpacing` holds. */
const NO_SPACING: LineSpacing = { kind: 'points', value: 0 };

/**
 * `a:lnSpc` and its siblings, as written.
 *
 * The value stays raw - thousandths of a percent, or hundredths of a point -
 * because `lineAdvance` is where the quantisation happens and doing it twice
 * rounds twice.
 */
function spacingOf(spacing: Spacing | undefined, fallback: LineSpacing): LineSpacing {
  if (spacing === undefined) return fallback;
  return { kind: spacing.kind === 'points' ? 'points' : 'percent', value: spacing.value };
}

/** A run's own fill, which is a colour or nothing a renderer can use here. */
function colorOf(fill: Fill | undefined, context: ColorContext): Rgba | null {
  return fill?.type === 'solid' ? resolveColor(fill.color, context) : null;
}

/**
 * The typeface a run asks for, theme reference followed.
 *
 * Only the Latin slot: splitting a run by script is `splitScriptRuns`, and it
 * happens at layout time where the characters are.
 */
function runOf(
  context: TextContext,
  paragraph: Paragraph,
  content: TextContent | undefined,
  colors: ColorContext,
  text: string,
  hardBreak: boolean,
): ResolvedRun {
  const ask = <T>(pick: Parameters<typeof resolveRun<T>>[3]): T | undefined =>
    resolveRun(context, paragraph, content, pick)?.value;
  const sz = resolveSize(context, paragraph, content).value;
  return {
    text,
    font: {
      family: resolveLatinTypeface(context, paragraph, content).value,
      sz,
      bold: ask((props) => props.b),
      italic: ask((props) => props.i),
      spc: ask((props) => props.spc),
      kern: ask((props) => props.kern),
    },
    color: colorOf(
      ask((props) => props.fill),
      colors,
    ),
    highlight: (() => {
      const highlight = ask((props) => props.highlight);
      return highlight === undefined ? null : resolveColor(highlight, colors);
    })(),
    underline: ask((props) => props.u),
    strike: ask((props) => props.strike),
    caps: ask((props) => props.cap),
    baseline: ask((props) => props.baseline) ?? 0,
    hardBreak,
  };
}

/**
 * The paragraphs of a body, split at every `a:br`.
 *
 * An `a:br` is a hard line break and not a paragraph break - no new bullet, no
 * `spcBef` - so it stays inside the paragraph and the layout ends the line on it.
 */
function paragraphOf(
  context: TextContext,
  paragraph: Paragraph,
  colors: ColorContext,
): ResolvedParagraph {
  const ask = <T>(pick: Parameters<typeof resolveParagraph<T>>[2]): T | undefined =>
    resolveParagraph(context, paragraph, pick)?.value;

  const runs: ResolvedRun[] = [];
  for (const content of paragraph.content) {
    if (content.kind === 'br') {
      const last = runs.at(-1);
      if (last === undefined) runs.push(runOf(context, paragraph, content, colors, '', true));
      else runs[runs.length - 1] = { ...last, hardBreak: true };
      continue;
    }
    runs.push(runOf(context, paragraph, content, colors, content.text, false));
  }

  const endContent =
    paragraph.endParaRPr === undefined
      ? undefined
      : ({
          kind: 'run',
          props: paragraph.endParaRPr,
          text: '',
          node: paragraph.node,
        } as TextContent);

  return {
    level: paragraph.level,
    align: ask((props) => props.algn) ?? 'l',
    marginLeft: resolveMarginLeft(context, paragraph).value / EMU_PER_POINT,
    indent: resolveIndent(context, paragraph).value / EMU_PER_POINT,
    lineSpacing: spacingOf(
      ask((props) => props.lnSpc),
      DEFAULT_SPACING,
    ),
    spaceBefore: spacingOf(
      ask((props) => props.spcBef),
      NO_SPACING,
    ),
    spaceAfter: spacingOf(
      ask((props) => props.spcAft),
      NO_SPACING,
    ),
    runs,
    endRun: runOf(context, paragraph, endContent, colors, '', false),
  };
}

/** `a:normAutofit`'s two stored scales, as multipliers. */
function autofitScales(body: TextBody | undefined): {
  fontScale: number;
  lineSpaceReduction: number;
} {
  const autofit = body?.bodyPr?.autofit;
  if (autofit?.kind !== 'normal') return { fontScale: 1, lineSpaceReduction: 0 };
  return {
    fontScale: (autofit.fontScale ?? 100000) / 100000,
    lineSpaceReduction: (autofit.lnSpcReduction ?? 0) / 100000,
  };
}

/**
 * Everything a shape's text needs, or `null` where it has none.
 *
 * A body with no paragraphs is `null` too: an empty `p:txBody` draws nothing,
 * and returning it as an empty layout would put a bullet on a blank line.
 */
export function resolveText(placed: Placed, defaultTextStyle?: ListStyle): ResolvedText | null {
  const body = placed.shape.text;
  if (body === undefined || body.paragraphs.length === 0) return null;

  const context: TextContext = {
    sheet: placed.sheet,
    shape: placed.shape,
    defaultTextStyle,
  };
  const colors: ColorContext = placed.colorContext;

  const insets = resolveInsets(placed.shape, placed.sheet);
  const vertical = resolveVertical(placed.shape, placed.sheet)?.value ?? 'horz';
  const bodyPr = resolveBody(placed.shape, placed.sheet, (props) => props.rot)?.value;
  const upright = resolveBody(placed.shape, placed.sheet, (props) => props.upright)?.value ?? false;
  const overflow =
    resolveBody(placed.shape, placed.sheet, (props) => props.vertOverflow)?.value ?? 'overflow';
  const spcCol = resolveBody(placed.shape, placed.sheet, (props) => props.spcCol)?.value ?? 0;
  const rtlCol = resolveBody(placed.shape, placed.sheet, (props) => props.rtlCol)?.value ?? false;

  const frame: ResolvedFrame = {
    anchor: resolveAnchor(placed.shape, placed.sheet)?.value ?? 't',
    anchorCtr: resolveAnchorCtr(placed.shape, placed.sheet)?.value ?? false,
    vertical,
    wrap: resolveWrap(placed.shape, placed.sheet)?.value ?? 'square',
    insets: insets,
    columns: resolveColumns(placed.shape, placed.sheet)?.value ?? 1,
    columnSpacing: spcCol / EMU_PER_POINT,
    rtlColumns: rtlCol,
    bodyRotation: (bodyPr ?? 0) / 60000,
    upright,
    vertOverflow: overflow,
    ...autofitScales(body),
  };

  if (!Number.isFinite(frame.bodyRotation)) {
    throw new RenderError(
      'RENDER_TEXT_FRAME',
      `a:bodyPr/@rot ${String(bodyPr)} is not an angle`,
      String(bodyPr),
    );
  }

  return {
    frame,
    paragraphs: body.paragraphs.map((paragraph) => paragraphOf(context, paragraph, colors)),
  };
}
