/**
 * Experiment F3 - where PowerPoint's export puts a whole-pixel edge on the device grid.
 *
 * Ten slides exported at `EXPORT_WIDTHS`: strokes, fills, outlines, pictures, inset strokes, line
 * ends, half-pixel pens and the shapes that are not axis-aligned, at up to four sub-pixel offsets.
 * Every model predicts a coverage profile; `assertSeparable` refuses a pair no case could tell apart.
 */

import { ln, LINE_GEOM, RECT_GEOM } from '../../paint/lines/probes.ts';

/** Agreement between the two exports, and between a model and a case, as coverage of one row. */
export const TOLERANCE = 0.15;

/** The least two models' profiles may differ by somewhere: a case on one then misses the other. */
export const SEPARABLE_BY = 0.2;

/** The slide, in points. */
export const SLIDE = { w: 960, h: 540 } as const;

export const EMU_PER_POINT = 12700;

/**
 * Export widths, in pixels: F2's seven, and seven more that are not an eighth of a pixel per
 * point with a whole-pixel height, so a rule read where every width is both is tried where the
 * gate's zooms never go - at a whole number of dots per inch and not, at a whole height and not.
 */
export const EXPORT_WIDTHS: readonly number[] = [
  120, 240, 480, 960, 1000, 1008, 1040, 1100, 1120, 1184, 1200, 1320, 1920, 3840,
];

/** The export's height for a width, half up, as `read.ps1` asks for it. */
export const heightOf = (width: number): number => Math.round((width * SLIDE.h) / SLIDE.w);

/** The three things a width is or is not, any of which could be what decides whether the export snaps. */
export interface ExportShape {
  readonly width: number;
  readonly height: number;
  /** The scale is a multiple of an eighth of a pixel per point. */
  readonly eighth: boolean;
  /** The height is a whole number of pixels before rounding. */
  readonly wholeHeight: boolean;
  /** The width is a whole number of dots per inch. */
  readonly wholeDpi: boolean;
}

export function exportShape(width: number): ExportShape {
  return {
    width,
    height: heightOf(width),
    eighth: Number.isInteger((width * 8) / SLIDE.w),
    wholeHeight: Number.isInteger((width * SLIDE.h) / SLIDE.w),
    wholeDpi: Number.isInteger((width * 72) / SLIDE.w),
  };
}

/** The widths the grid rule is read from: an eighth of a pixel per point with a whole-pixel height. */
export const GRID_WIDTHS: readonly number[] = EXPORT_WIDTHS.filter((width) => {
  const shape = exportShape(width);
  return shape.eighth && shape.wholeHeight;
});

/** Device pixels per point along an axis: the width's across columns, the rounded height's down rows. */
export const axisScale = (width: number, axis: 'x' | 'y'): number =>
  axis === 'x' ? width / SLIDE.w : heightOf(width) / SLIDE.h;

/** The sub-pixel offsets, in points, that a 960-px export lands on the grid as pixels. */
export const OFFSETS_PT: readonly number[] = [0, 0.25, 0.5, 0.75];

/** Nominal stroke widths for the axis-aligned strokes: a hairline, the fractions and the whole pixels. */
export const STROKE_WIDTHS_PT: readonly number[] = [0, 0.5, 1, 1.5, 2, 3, 4];

/** Pens that are an exact half pixel wide at 960 (the halves) or at 1200 (the tenths). */
export const TIE_WIDTHS_PT: readonly number[] = [
  1.5, 2.5, 3.5, 4.5, 5.5, 6.5, 1.2, 2, 2.8, 3.6, 4.4,
];

/**
 * The same question at 1920 (the quarters), 240 (6, 10, 14) and 120 (12, 20); no integer-EMU
 * width is a half pixel at 3840, where a half needs an eighth of a point.
 */
export const WIDE_TIE_WIDTHS_PT: readonly number[] = [
  0.75, 1.25, 1.75, 2.25, 2.75, 6, 10, 12, 14, 20,
];

/** The pens whose flat ends are read: both parities at 480 and 960, odd pens at 240 and 1200. */
export const END_WIDTHS_PT: readonly number[] = [1, 2, 4];

/** `algn="in"` widths, in points: whole pixels at 960, and at 480 and 1920 for the second. */
export const INSET_WIDTHS_PT: readonly number[] = [1, 2];

export type Family =
  | 'stroke'
  | 'outline'
  | 'fill'
  | 'picture'
  | 'border'
  | 'inset'
  | 'end'
  | 'ellipse'
  | 'roundRect'
  | 'triangle'
  | 'slant'
  | 'rotated'
  | 'tie';

/** What the ink is, across the read axis: a stroke band, or a fill's one edge with ink on one side. */
export type Shape = 'band' | 'from' | 'to';

export type Edge = 'top' | 'bottom' | 'left' | 'right';

/** Where the analysis reads: a window across one edge, averaged along it, all in points. */
export interface Read {
  /** The axis the profile runs down: `y` reads rows across a horizontal edge. */
  readonly axis: 'x' | 'y';
  /** The window's centre across the edge. */
  readonly at: number;
  /** Half the window, across. */
  readonly half: number;
  /** The span along the edge that is averaged. */
  readonly along0: number;
  readonly along1: number;
}

export interface Probe {
  readonly id: string;
  /** 1-based slide index in the deck. */
  readonly slide: number;
  readonly family: Family;
  readonly question: string;
  /** Already-built `p:spTree` children; empty for a second read of a shape another probe drew. */
  readonly markup: string;
  readonly read: Read;
  readonly shape: Shape;
  /** The stroke's nominal centre, or the fill's edge, in points along the read axis. */
  readonly centrePt: number;
  /** The nominal stroke width in points, for a band; a picture border's is the band outside its frame. */
  readonly widthPt?: number;
  /** Which edge of its shape a read crosses, where the shape has more than one. */
  readonly edge?: Edge;
  /** A slanted line's two endpoints along the read axis, and how far along it the read is. */
  readonly line?: { readonly from: number; readonly to: number; readonly t: number };
}

/** A device-space band `[top, bottom)` in pixels along the read axis, `Infinity` on a fill's open side. */
export interface Band {
  readonly top: number;
  readonly bottom: number;
}

/** A model of where one edge lands, from the nominal centre `c` and width `w`, both in device pixels. */
export interface Model {
  readonly name: string;
  readonly description: string;
  readonly predict: (c: number, w: number, probe: Probe, scale: number) => Band;
}

function emu(points: number): string {
  return String(Math.round(points * EMU_PER_POINT));
}

const BLACK = '<a:solidFill><a:srgbClr val="000000"/></a:solidFill>';
const NO_LINE = '<a:ln><a:noFill/></a:ln>';

interface ShapeSpec {
  readonly id: number;
  readonly name: string;
  readonly x: number;
  readonly y: number;
  readonly cx: number;
  readonly cy: number;
  readonly geom: string;
  readonly fill: string;
  readonly line: string;
  readonly rot?: number;
}

function shape(spec: ShapeSpec): string {
  const rot = spec.rot === undefined ? '' : ` rot="${String(spec.rot * 60000)}"`;
  return (
    '<p:sp><p:nvSpPr>' +
    `<p:cNvPr id="${String(spec.id)}" name="${spec.name}"/><p:cNvSpPr/><p:nvPr/>` +
    '</p:nvSpPr><p:spPr>' +
    `<a:xfrm${rot}><a:off x="${emu(spec.x)}" y="${emu(spec.y)}"/>` +
    `<a:ext cx="${emu(spec.cx)}" cy="${emu(spec.cy)}"/></a:xfrm>` +
    spec.geom +
    spec.fill +
    spec.line +
    '</p:spPr></p:sp>'
  );
}

function picture(
  id: number,
  name: string,
  x: number,
  y: number,
  cx: number,
  cy: number,
  embed: string,
  line: string,
): string {
  return (
    '<p:pic><p:nvPicPr>' +
    `<p:cNvPr id="${String(id)}" name="${name}"/><p:cNvPicPr><a:picLocks/></p:cNvPicPr><p:nvPr/>` +
    `</p:nvPicPr><p:blipFill><a:blip r:embed="${embed}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>` +
    `<p:spPr><a:xfrm><a:off x="${emu(x)}" y="${emu(y)}"/><a:ext cx="${emu(cx)}" cy="${emu(cy)}"/></a:xfrm>` +
    RECT_GEOM +
    line +
    '</p:spPr></p:pic>'
  );
}

const geom = (prst: string): string => `<a:prstGeom prst="${prst}"><a:avLst/></a:prstGeom>`;

/** An L: every edge axis-aligned, none of them a preset's. */
function ellGeom(): string {
  const pt = (x: number, y: number): string => `<a:pt x="${emu(x)}" y="${emu(y)}"/>`;
  return (
    '<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="0" t="0" r="r" b="b"/>' +
    `<a:pathLst><a:path w="${emu(200)}" h="${emu(80)}">` +
    `<a:moveTo>${pt(0, 0)}</a:moveTo><a:lnTo>${pt(200, 0)}</a:lnTo><a:lnTo>${pt(200, 40)}</a:lnTo>` +
    `<a:lnTo>${pt(80, 40)}</a:lnTo><a:lnTo>${pt(80, 80)}</a:lnTo><a:lnTo>${pt(0, 80)}</a:lnTo>` +
    '<a:close/></a:path></a:pathLst></a:custGeom>'
  );
}

const tag = (value: number): string => String(value).replace('.', '_');

/** The window across an edge, in points: wide enough for a 4-pt band and its neighbours' silence. */
const HALF_WINDOW = 8;

/* -------------------------------------------------------------------------- */
/* the probes                                                                 */
/* -------------------------------------------------------------------------- */

export function snapProbes(): Probe[] {
  const probes: Probe[] = [];
  let id = 2;

  /* --------------------------------------------- slides 1 and 2: strokes */
  for (const axis of ['y', 'x'] as const) {
    let k = 0;
    for (const w of STROKE_WIDTHS_PT) {
      for (const offset of OFFSETS_PT) {
        const centre = 24 + 18 * k + offset;
        k += 1;
        const name = `stroke-${axis === 'y' ? 'h' : 'v'}-${tag(w)}-${tag(offset)}`;
        probes.push({
          id: name,
          slide: axis === 'y' ? 1 : 2,
          family: 'stroke',
          question: `a ${axis === 'y' ? 'horizontal' : 'vertical'} ${String(w)}pt line centred ${String(offset)}pt past the grid`,
          markup: shape({
            id: id++,
            name,
            x: axis === 'y' ? 120 : centre,
            y: axis === 'y' ? centre : 60,
            cx: axis === 'y' ? 600 : 0,
            cy: axis === 'y' ? 0 : 400,
            geom: LINE_GEOM,
            fill: '<a:noFill/>',
            line: ln({ rawW: emu(w) }),
          }),
          read:
            axis === 'y'
              ? { axis, at: centre, half: HALF_WINDOW, along0: 200, along1: 640 }
              : { axis, at: centre, half: HALF_WINDOW, along0: 120, along1: 400 },
          shape: 'band',
          centrePt: centre,
          widthPt: w,
        });
      }
    }
  }

  /* ------------------------- slides 3, 4 and 6: four edges of one rectangle */
  const edges = (
    family: Family,
    stem: string,
    x: number,
    y: number,
    cx: number,
    cy: number,
    markup: string,
    slide: number,
    band?: number,
  ): void => {
    const sides: readonly [Edge, 'x' | 'y', number, Shape][] = [
      ['top', 'y', y, 'from'],
      ['bottom', 'y', y + cy, 'to'],
      ['left', 'x', x, 'from'],
      ['right', 'x', x + cx, 'to'],
    ];
    sides.forEach(([edge, axis, at, shape], i) => {
      const along: [number, number] =
        axis === 'y' ? [x + cx * 0.2, x + cx * 0.8] : [y + cy * 0.2, y + cy * 0.8];
      probes.push({
        id: `${stem}-${edge}`,
        slide,
        family,
        question: `the ${edge} edge of ${stem}`,
        markup: i === 0 ? markup : '',
        read: { axis, at, half: HALF_WINDOW, along0: along[0], along1: along[1] },
        shape: band === undefined ? shape : 'band',
        centrePt: at,
        edge,
        ...(band === undefined ? {} : { widthPt: band }),
      });
    });
  };

  OFFSETS_PT.forEach((offset, i) => {
    const x = 60 + offset + (i % 2) * 480;
    const y = 40 + offset + Math.floor(i / 2) * 250;
    const stem = `fill-${tag(offset)}`;
    edges(
      'fill',
      stem,
      x,
      y,
      300,
      100.5,
      shape({
        id: id++,
        name: stem,
        x,
        y,
        cx: 300,
        cy: 100.5,
        geom: RECT_GEOM,
        fill: BLACK,
        line: NO_LINE,
      }),
      3,
    );
  });

  const outlineWidths: readonly number[] = [1, 2];
  outlineWidths.forEach((w, row) => {
    OFFSETS_PT.forEach((offset, i) => {
      const x = 60 + offset + i * 220;
      const y = 60 + offset + row * 250;
      const stem = `outline-${tag(w)}-${tag(offset)}`;
      edges(
        'outline',
        stem,
        x,
        y,
        160,
        100.5,
        shape({
          id: id++,
          name: stem,
          x,
          y,
          cx: 160,
          cy: 100.5,
          geom: RECT_GEOM,
          fill: '<a:noFill/>',
          line: ln({ rawW: emu(w) }),
        }),
        4,
        w,
      );
    });
  });

  const ells: readonly [number, number][] = [
    [1, 0.25],
    [2, 0.75],
  ];
  ells.forEach(([w, offset], i) => {
    const x = 60 + offset + i * 300;
    const y = 430 + offset;
    const name = `outline-L-${tag(w)}-${tag(offset)}`;
    probes.push(
      {
        id: `${name}-top`,
        slide: 4,
        family: 'outline',
        question: `the top of a ${String(w)}pt L-shaped custom outline, ${String(offset)}pt past the grid`,
        markup: shape({
          id: id++,
          name,
          x,
          y,
          cx: 200,
          cy: 80,
          geom: ellGeom(),
          fill: '<a:noFill/>',
          line: ln({ rawW: emu(w) }),
        }),
        read: { axis: 'y', at: y, half: HALF_WINDOW, along0: x + 20, along1: x + 180 },
        shape: 'band',
        centrePt: y,
        widthPt: w,
        edge: 'top',
      },
      {
        id: `${name}-step`,
        slide: 4,
        family: 'outline',
        question: 'the inner step of the same outline',
        markup: '',
        read: { axis: 'y', at: y + 40, half: HALF_WINDOW, along0: x + 100, along1: x + 180 },
        shape: 'band',
        centrePt: y + 40,
        widthPt: w,
        edge: 'top',
      },
    );
  });

  /* --------------------------- slide 5: the shapes that are not rectangles */
  const curved: readonly [Family, number, number][] = [
    ['ellipse', 1, 0],
    ['ellipse', 1, 0.25],
    ['ellipse', 2, 0.5],
    ['ellipse', 2, 0.25],
    ['roundRect', 1, 0],
    ['roundRect', 1, 0.25],
    ['roundRect', 2, 0.5],
    ['roundRect', 2, 0.25],
    ['triangle', 1, 0],
    ['triangle', 1, 0.25],
    ['triangle', 2, 0.5],
    ['triangle', 2, 0.25],
    ['slant', 1, 0],
    ['slant', 1, 0.25],
    ['slant', 2, 0.5],
    ['slant', 2, 0.25],
    ['rotated', 1, 0],
    ['rotated', 1, 0.25],
    ['rotated', 2, 0.5],
    ['rotated', 2, 0.25],
  ];
  curved.forEach(([family, w, offset], i) => {
    const column = i % 4;
    const row = Math.floor(i / 4);
    const x = 40 + column * 230;
    const top = 30 + row * 102 + offset;
    const name = `${family}-${tag(w)}-${tag(offset)}`;
    const line = ln({ rawW: emu(w) });
    const across = (at: number, alongHalf: number, alongCentre: number): Read => ({
      axis: 'y',
      at,
      half: HALF_WINDOW,
      along0: alongCentre - alongHalf,
      along1: alongCentre + alongHalf,
    });
    const common = { slide: 5, family, shape: 'band' as const, widthPt: w };
    switch (family) {
      case 'ellipse':
        probes.push(
          {
            ...common,
            id: name,
            question: `the top of a ${String(w)}pt ellipse outline, ${String(offset)}pt past the grid`,
            markup: shape({
              id: id++,
              name,
              x,
              y: top,
              cx: 200,
              cy: 80,
              geom: geom('ellipse'),
              fill: '<a:noFill/>',
              line,
            }),
            read: across(top, 4, x + 100),
            centrePt: top,
            edge: 'top',
          },
          {
            ...common,
            id: `${name}-bottom`,
            question: 'the bottom of the same ellipse',
            markup: '',
            read: across(top + 80, 4, x + 100),
            centrePt: top + 80,
            edge: 'bottom',
          },
          {
            ...common,
            id: `${name}-left`,
            question: 'the left of the same ellipse',
            markup: '',
            // Half a point along: the side of a 200-by-80 ellipse bends away far faster than its top.
            read: { axis: 'x', at: x, half: HALF_WINDOW, along0: top + 39.5, along1: top + 40.5 },
            centrePt: x,
            edge: 'left',
          },
        );
        break;
      case 'roundRect':
        probes.push(
          {
            ...common,
            id: name,
            question: `the straight top of a ${String(w)}pt rounded rectangle, ${String(offset)}pt past the grid`,
            markup: shape({
              id: id++,
              name,
              x,
              y: top,
              cx: 200,
              cy: 80,
              geom: geom('roundRect'),
              fill: '<a:noFill/>',
              line,
            }),
            read: across(top, 20, x + 100),
            centrePt: top,
            edge: 'top',
          },
          {
            ...common,
            id: `${name}-bottom`,
            question: 'the straight bottom of the same rounded rectangle',
            markup: '',
            read: across(top + 80, 20, x + 100),
            centrePt: top + 80,
            edge: 'bottom',
          },
          {
            ...common,
            id: `${name}-left`,
            question: 'the straight left of the same rounded rectangle',
            markup: '',
            read: { axis: 'x', at: x, half: HALF_WINDOW, along0: top + 30, along1: top + 50 },
            centrePt: x,
            edge: 'left',
          },
        );
        break;
      case 'triangle':
        probes.push({
          ...common,
          id: name,
          question: `the base of a ${String(w)}pt triangle outline, ${String(offset)}pt past the grid`,
          markup: shape({
            id: id++,
            name,
            x,
            y: top,
            cx: 200,
            cy: 80,
            geom: geom('triangle'),
            fill: '<a:noFill/>',
            line,
          }),
          read: across(top + 80, 20, x + 100),
          centrePt: top + 80,
          edge: 'bottom',
        });
        break;
      case 'slant':
        // Two points of rise over two hundred: a line that is nearly, and not, horizontal.
        probes.push({
          ...common,
          id: name,
          question: `a ${String(w)}pt line rising 2pt over 200, read where it is ${String(offset)}pt past the grid`,
          markup: shape({
            id: id++,
            name,
            x,
            y: top + 40,
            cx: 200,
            cy: 2,
            geom: LINE_GEOM,
            fill: '<a:noFill/>',
            line,
          }),
          read: across(top + 41, 4, x + 100),
          centrePt: top + 41,
          line: { from: top + 40, to: top + 42, t: 0.5 },
        });
        break;
      case 'rotated':
        // Turned a quarter, the frame's left edge is the top edge on the slide, cx / 2 - cy / 2 above the frame's y.
        probes.push({
          ...common,
          id: name,
          question: `the top of a ${String(w)}pt rectangle outline turned 90 degrees, ${String(offset)}pt past the grid`,
          markup: shape({
            id: id++,
            name,
            x: x + 60,
            y: top + 30,
            cx: 80,
            cy: 20,
            geom: RECT_GEOM,
            fill: '<a:noFill/>',
            line,
            rot: 90,
          }),
          read: across(top, 6, x + 100),
          centrePt: top,
          edge: 'top',
        });
        break;
      default:
        throw new Error(`no probe for ${family}`);
    }
  });

  /* ------------------------------ slide 6: pictures, with and without a border */
  OFFSETS_PT.forEach((offset, i) => {
    const x = 60 + offset + (i % 2) * 480;
    const y = 30 + offset + Math.floor(i / 2) * 130;
    const stem = `picture-${tag(offset)}`;
    edges(
      'picture',
      stem,
      x,
      y,
      300,
      100.5,
      picture(id++, stem, x, y, 300, 100.5, 'rId2', NO_LINE),
      6,
    );
  });
  const borderWidths: readonly number[] = [1, 2];
  borderWidths.forEach((w, row) => {
    OFFSETS_PT.forEach((offset, i) => {
      const x = 60 + offset + i * 220;
      const y = 300 + offset + row * 110;
      const name = `border-${tag(w)}-${tag(offset)}`;
      probes.push({
        id: name,
        slide: 6,
        family: 'border',
        question: `a ${String(w)}pt border outside a picture whose top is ${String(offset)}pt past the grid`,
        markup: picture(id++, name, x, y, 160, 80, 'rId3', ln({ rawW: emu(w) })),
        read: { axis: 'y', at: y, half: HALF_WINDOW, along0: x + 30, along1: x + 130 },
        shape: 'band',
        centrePt: y,
        widthPt: w,
        edge: 'top',
      });
    });
  });

  /* --------------------------------------------------- slide 7: line ends */
  END_WIDTHS_PT.forEach((w, row) => {
    OFFSETS_PT.forEach((offset, i) => {
      const y = 40 + (row * OFFSETS_PT.length + i) * 40;
      const x0 = 100 + offset;
      const x1 = 700 + OFFSETS_PT[(i + 1) % OFFSETS_PT.length]!;
      const name = `end-${tag(w)}-${tag(offset)}`;
      probes.push(
        {
          id: `${name}-start`,
          slide: 7,
          family: 'end',
          question: `the flat start of a ${String(w)}pt line at ${String(offset)}pt past the grid`,
          markup: shape({
            id: id++,
            name,
            x: x0,
            y,
            cx: x1 - x0,
            cy: 0,
            geom: LINE_GEOM,
            fill: '<a:noFill/>',
            line: ln({ rawW: emu(w) }),
          }),
          read: { axis: 'x', at: x0, half: HALF_WINDOW, along0: y, along1: y },
          shape: 'from',
          centrePt: x0,
          widthPt: w,
          edge: 'left',
        },
        {
          id: `${name}-end`,
          slide: 7,
          family: 'end',
          question: `the flat end of the same line at ${String(x1 - 700)}pt past the grid`,
          markup: '',
          read: { axis: 'x', at: x1, half: HALF_WINDOW, along0: y, along1: y },
          shape: 'to',
          centrePt: x1,
          widthPt: w,
          edge: 'right',
        },
      );
    });
  });

  /* --------------------- slide 8: pens that are an exact half pixel at 960 or 1200 */
  TIE_WIDTHS_PT.forEach((w, i) => {
    OFFSETS_PT.slice(0, 2).forEach((offset, j) => {
      const y = 24 + (i * 2 + j) * 22 + offset;
      const name = `tie-${tag(w)}-${tag(offset)}`;
      probes.push({
        id: name,
        slide: 8,
        family: 'tie',
        question: `a ${String(w)}pt line, ${String(offset)}pt past the grid: which way does a half pixel round?`,
        markup: shape({
          id: id++,
          name,
          x: 120,
          y,
          cx: 600,
          cy: 0,
          geom: LINE_GEOM,
          fill: '<a:noFill/>',
          line: ln({ rawW: emu(w) }),
        }),
        read: { axis: 'y', at: y, half: HALF_WINDOW, along0: 200, along1: 640 },
        shape: 'band',
        centrePt: y,
        widthPt: w,
      });
    });
  });

  /* ------------------- slide 9: the half pixels at 1920, 240 and 120, wide pens among them */
  let top = 24;
  WIDE_TIE_WIDTHS_PT.forEach((w) => {
    // A wide pen needs a wider window, and its neighbours out of it.
    const half = Math.max(HALF_WINDOW, w / 2 + 4);
    OFFSETS_PT.slice(0, 2).forEach((offset) => {
      const y = top + offset;
      const name = `tie-${tag(w)}-${tag(offset)}`;
      probes.push({
        id: name,
        slide: 9,
        family: 'tie',
        question: `a ${String(w)}pt line, ${String(offset)}pt past the grid: which way does a half pixel round?`,
        markup: shape({
          id: id++,
          name,
          x: 120,
          y,
          cx: 600,
          cy: 0,
          geom: LINE_GEOM,
          fill: '<a:noFill/>',
          line: ln({ rawW: emu(w) }),
        }),
        read: { axis: 'y', at: y, half, along0: 200, along1: 640 },
        shape: 'band',
        centrePt: y,
        widthPt: w,
      });
      top += 2 * half + 6;
    });
  });

  /* ------------------------------------- slide 10: a stroke aligned inside its frame */
  INSET_WIDTHS_PT.forEach((w, row) => {
    OFFSETS_PT.forEach((offset, i) => {
      const x = 60 + offset + i * 220;
      const y = 60 + offset + row * 110;
      const name = `inset-${tag(w)}-${tag(offset)}`;
      probes.push({
        id: name,
        slide: 10,
        family: 'inset',
        question: `a ${String(w)}pt stroke aligned inside a frame whose top is ${String(offset)}pt past the grid`,
        markup: shape({
          id: id++,
          name,
          x,
          y,
          cx: 160,
          cy: 80,
          geom: RECT_GEOM,
          fill: '<a:noFill/>',
          line: ln({ rawW: emu(w), algn: 'in' }),
        }),
        read: { axis: 'y', at: y, half: HALF_WINDOW, along0: x + 30, along1: x + 130 },
        shape: 'band',
        centrePt: y,
        widthPt: w,
        edge: 'top',
      });
    });
  });

  return probes;
}

/* -------------------------------------------------------------------------- */
/* the candidate readings                                                     */
/* -------------------------------------------------------------------------- */

export const halfUp = (v: number): number => Math.floor(v + 0.5);

export const halfEven = (v: number): number => {
  const floor = Math.floor(v);
  const rest = v - floor;
  if (rest < 0.5) return floor;
  if (rest > 0.5) return floor + 1;
  return floor % 2 === 0 ? floor : floor + 1;
};

/** F2's pen, whole pixels rounded half up and never under one: what decides the parity. */
export const pen = (w: number): number => Math.max(1, halfUp(w));

/** The export scale at which an exact half pixel of width rounded down; at 960 it rounded up. */
export const HALF_ROUNDS_DOWN_AT = 1.25;

/** The pen the export draws: the width rounded half up, never under one, but a half down at 1200 wide. */
export const drawnPen = (w: number, scale: number): number => {
  const half = w - Math.floor(w) === 0.5;
  return Math.max(1, half && scale === HALF_ROUNDS_DOWN_AT ? Math.floor(w) : halfUp(w));
};

/** Whether the pen sits on pixel centres: when `w` rounded half up is odd. */
export const oddPen = (w: number): boolean => pen(w) % 2 === 1;

/** A coordinate on the export's grid: rounded half up, then half a pixel on for an odd pen. */
export const snapped = (v: number, w: number): number => halfUp(v) + (oddPen(w) ? 0.5 : 0);

function band(top: number, height: number): Band {
  return { top, bottom: top + height };
}

/** A fill's edge at `e`, with ink on the side `shape` says. */
function side(e: number, shape: Shape): Band {
  return shape === 'from' ? { top: e, bottom: Infinity } : { top: -Infinity, bottom: e };
}

/** A stroke band under the parity rule, centred at `centre`. */
function parityBand(centre: number, w: number, scale: number): Band {
  return band(centre - drawnPen(w, scale) / 2, drawnPen(w, scale));
}

/** Where a stroke band lands, from its nominal centre `c` and width `w` in pixels. */
export const STROKE_MODELS: readonly Model[] = [
  {
    name: 'S0',
    description: 'the true width, antialiased where it lies',
    predict: (c, w) => band(c - w / 2, w),
  },
  {
    name: 'S1',
    description: 'the whole-pixel pen, antialiased where it lies',
    predict: (c, w) => band(c - pen(w) / 2, pen(w)),
  },
  {
    name: 'SR',
    description: 'the whole-pixel pen, its top edge rounded half up',
    predict: (c, w) => band(halfUp(c - pen(w) / 2), pen(w)),
  },
  {
    name: 'SE',
    description: 'the whole-pixel pen, its top edge rounded half to even',
    predict: (c, w) => band(halfEven(c - pen(w) / 2), pen(w)),
  },
  {
    name: 'SF',
    description: 'the whole-pixel pen, its top edge rounded down',
    predict: (c, w) => band(Math.floor(c - pen(w) / 2), pen(w)),
  },
  {
    name: 'SC',
    description: 'the whole-pixel pen, its top edge rounded up',
    predict: (c, w) => band(Math.ceil(c - pen(w) / 2), pen(w)),
  },
  {
    name: 'SH',
    description:
      'the whole-pixel pen half a pixel down, antialiased: integer coordinates are pixel centres',
    predict: (c, w) => band(c - pen(w) / 2 + 0.5, pen(w)),
  },
  {
    name: 'SD',
    description: 'the whole-pixel pen half a pixel down, then its top edge rounded half up',
    predict: (c, w) => band(halfUp(c - pen(w) / 2 + 0.5), pen(w)),
  },
  {
    name: 'SG',
    description: "the centre rounded half up first, then the pen's top edge rounded half up",
    predict: (c, w) => band(halfUp(halfUp(c) - pen(w) / 2), pen(w)),
  },
  {
    name: 'SW',
    description: 'the true edges, each rounded half up on its own',
    predict: (c, w) => ({ top: halfUp(c - w / 2), bottom: halfUp(c + w / 2) }),
  },
  {
    name: 'SP',
    description:
      'the centre rounded half up, then on a pixel centre when the pen rounded half up is odd; the pen rounded half up, but a half down at 1200 wide',
    predict: (c, w, _probe, scale) => parityBand(snapped(c, w), w, scale),
  },
];

/** A pen of an exact half pixel: the two roundings the export could be using, and the one it is. */
export const TIE_MODELS: readonly Model[] = [
  ...STROKE_MODELS,
  {
    name: 'SE2',
    description:
      'the centre rounded half up, then on a pixel centre when the pen rounded half up is odd; the pen rounded half to even',
    predict: (c, w) => band(snapped(c, w) - Math.max(1, halfEven(w)) / 2, Math.max(1, halfEven(w))),
  },
  {
    name: 'SD2',
    description: 'SP with a half pixel rounded down at every width, never under one',
    predict: (c, w) => {
      const n = Math.max(1, w - Math.floor(w) === 0.5 ? Math.floor(w) : halfUp(w));
      return band(snapped(c, w) - n / 2, n);
    },
  },
];

/** A slanted line: SP at each endpoint, and the line between them antialiased. */
export const SLANT_MODELS: readonly Model[] = [
  ...STROKE_MODELS,
  {
    name: 'SL',
    description: 'both endpoints under SP, the stroke between them antialiased',
    predict: (c, w, probe, scale) => {
      if (probe.line === undefined) return parityBand(snapped(c, w), w, scale);
      const from = snapped(probe.line.from * scale, w);
      const to = snapped(probe.line.to * scale, w);
      return parityBand(from + (to - from) * probe.line.t, w, scale);
    },
  },
];

/** What the curve rasteriser adds along an axis: a quarter pixel down, an eighth right. */
const curveBias = (axis: 'x' | 'y'): number => (axis === 'y' ? 0.25 : 0.125);

/** A curved outline: the stroke rule, and what the curve rasteriser adds to it. */
export const CURVE_MODELS: readonly Model[] = [
  ...STROKE_MODELS,
  {
    name: 'SQ',
    description:
      'SP, then the top band a quarter pixel down and the left band an eighth of a pixel right; the bottom where SP put it',
    predict: (c, w, probe, scale) => {
      const b = parityBand(snapped(c, w), w, scale);
      return probe.edge === 'bottom'
        ? b
        : band(b.top + curveBias(probe.read.axis), drawnPen(w, scale));
    },
  },
  {
    name: 'SN',
    description:
      "SP, then a band's leading edge a quarter pixel in across rows and an eighth across columns, when the pen is wider than a pixel",
    predict: (c, w, probe, scale) => {
      const b = parityBand(snapped(c, w), w, scale);
      return drawnPen(w, scale) > 1
        ? { top: b.top + curveBias(probe.read.axis), bottom: b.bottom }
        : b;
    },
  },
];

/** Where a fill's or a picture's edge lands, from its nominal position `e`. */
export const EDGE_MODELS: readonly Model[] = [
  { name: 'E0', description: 'antialiased where it lies', predict: (e, _w, p) => side(e, p.shape) },
  { name: 'ER', description: 'rounded half up', predict: (e, _w, p) => side(halfUp(e), p.shape) },
  {
    name: 'EE',
    description: 'rounded half to even',
    predict: (e, _w, p) => side(halfEven(e), p.shape),
  },
  { name: 'EF', description: 'rounded down', predict: (e, _w, p) => side(Math.floor(e), p.shape) },
  { name: 'EC', description: 'rounded up', predict: (e, _w, p) => side(Math.ceil(e), p.shape) },
  {
    name: 'EH',
    description: 'half a pixel down, antialiased: integer coordinates are pixel centres',
    predict: (e, _w, p) => side(e + 0.5, p.shape),
  },
  {
    name: 'ED',
    description: 'half a pixel down, then rounded half up',
    predict: (e, _w, p) => side(halfUp(e + 0.5), p.shape),
  },
];

/** Where a flat line end lands, from its nominal position `e` and the line's width `w`. */
export const END_MODELS: readonly Model[] = [
  { name: 'E0', description: 'antialiased where it lies', predict: (e, _w, p) => side(e, p.shape) },
  { name: 'ER', description: 'rounded half up', predict: (e, _w, p) => side(halfUp(e), p.shape) },
  {
    name: 'EH',
    description: 'half a pixel on, antialiased: integer coordinates are pixel centres',
    predict: (e, _w, p) => side(e + 0.5, p.shape),
  },
  {
    name: 'EP',
    description: 'rounded half up, then half a pixel on when the pen rounded half up is odd',
    predict: (e, w, p) => side(snapped(e, w), p.shape),
  },
  {
    name: 'EX',
    description: 'EP, and a one-pixel pen a quarter pixel longer at each end',
    predict: (e, w, p, scale) => {
      const reach = drawnPen(w, scale) === 1 ? (p.shape === 'from' ? -0.25 : 0.25) : 0;
      return side(snapped(e, w) + reach, p.shape);
    },
  },
];

/** Where a picture border lands: a whole-pixel pen wholly outside a frame edge at `f` (ADR 0037). */
export const BORDER_MODELS: readonly Model[] = [
  {
    name: 'B0',
    description: 'the true width outside the frame, antialiased where it lies',
    predict: (f, w) => band(f - w, w),
  },
  {
    name: 'BR',
    description: 'the frame edge rounded half up, the pen wholly outside it',
    predict: (f, w) => band(halfUp(f) - pen(w), pen(w)),
  },
  {
    name: 'BE',
    description: 'the frame edge rounded half to even, the pen wholly outside it',
    predict: (f, w) => band(halfEven(f) - pen(w), pen(w)),
  },
  {
    name: 'BF',
    description: 'the frame edge rounded down, the pen wholly outside it',
    predict: (f, w) => band(Math.floor(f) - pen(w), pen(w)),
  },
  {
    name: 'BC',
    description: 'the frame edge rounded up, the pen wholly outside it',
    predict: (f, w) => band(Math.ceil(f) - pen(w), pen(w)),
  },
  {
    name: 'BH',
    description: 'the pen outside the frame, half a pixel down, antialiased',
    predict: (f, w) => band(f - pen(w) + 0.5, pen(w)),
  },
  {
    name: 'BS',
    description: 'a stroke band centred half the true width outside, its top rounded half up',
    predict: (f, w) => band(halfUp(f - w / 2 - pen(w) / 2), pen(w)),
  },
  {
    name: 'BP',
    description:
      'the frame edge rounded half up, then SP on a band centred half the true width outside it',
    predict: (f, w, _probe, scale) => parityBand(snapped(halfUp(f) - w / 2, w), w, scale),
  },
  {
    name: 'BT',
    description:
      'the frame edge rounded half up; the pen as SP draws it, centred half the true width outside it, antialiased',
    predict: (f, w, _probe, scale) => parityBand(halfUp(f) - w / 2, w, scale),
  },
];

/** Where an `algn="in"` stroke lands: the border family's readings, inside a frame edge at `f`. */
export const INSET_MODELS: readonly Model[] = [
  {
    name: 'I0',
    description: 'the true width inside the frame, antialiased where it lies',
    predict: (f, w) => band(f, w),
  },
  {
    name: 'IR',
    description: 'the frame edge rounded half up, the pen wholly inside it',
    predict: (f, w) => band(halfUp(f), pen(w)),
  },
  {
    name: 'IE',
    description: 'the frame edge rounded half to even, the pen wholly inside it',
    predict: (f, w) => band(halfEven(f), pen(w)),
  },
  {
    name: 'IF',
    description: 'the frame edge rounded down, the pen wholly inside it',
    predict: (f, w) => band(Math.floor(f), pen(w)),
  },
  {
    name: 'IC',
    description: 'the frame edge rounded up, the pen wholly inside it',
    predict: (f, w) => band(Math.ceil(f), pen(w)),
  },
  {
    name: 'IH',
    description: 'the pen inside the frame, half a pixel down, antialiased',
    predict: (f, w) => band(f + 0.5, pen(w)),
  },
  {
    name: 'IS',
    description: 'a stroke band centred half the true width inside, its top rounded half up',
    predict: (f, w) => band(halfUp(f + w / 2 - pen(w) / 2), pen(w)),
  },
  {
    name: 'IP',
    description: 'SP on a band centred half the true width inside the frame edge',
    predict: (f, w, _probe, scale) => parityBand(snapped(f + w / 2, w), w, scale),
  },
  {
    name: 'IT',
    description:
      'the frame edge rounded half up; the pen as SP draws it, centred half the true width inside it, antialiased',
    predict: (f, w, _probe, scale) => parityBand(halfUp(f) + w / 2, w, scale),
  },
  {
    name: 'IX',
    description: 'the alignment ignored: SP, centred on the frame edge',
    predict: (f, w, _probe, scale) => parityBand(snapped(f, w), w, scale),
  },
  {
    name: 'IZ',
    description: 'the pen as SP draws it, centred on the frame edge rounded half up, odd or even',
    predict: (f, w, _probe, scale) => parityBand(halfUp(f), w, scale),
  },
  {
    name: 'IW',
    description:
      'the pen as SP draws it, centred on the frame edge rounded half up, then inward by what the pen was widened by',
    predict: (f, w, _probe, scale) => parityBand(halfUp(f) + Math.max(0, pen(w) - w), w, scale),
  },
  {
    name: 'IN',
    description:
      'the pen as SP draws it, its outer edge a pixel outside the frame edge rounded half up - half a pixel for an odd pen - and the rest inside; a pen under a pixel wide slides inward by what it was widened by',
    predict: (f, w, _probe, scale) =>
      band(halfUp(f) - 1 + (oddPen(w) ? 0.5 : 0) + Math.max(0, 1 - w), drawnPen(w, scale)),
  },
];

export const MODELS_BY_FAMILY: Readonly<Record<Family, readonly Model[]>> = {
  stroke: STROKE_MODELS,
  outline: STROKE_MODELS,
  triangle: STROKE_MODELS,
  rotated: STROKE_MODELS,
  tie: TIE_MODELS,
  slant: SLANT_MODELS,
  ellipse: CURVE_MODELS,
  roundRect: CURVE_MODELS,
  fill: EDGE_MODELS,
  picture: EDGE_MODELS,
  end: END_MODELS,
  border: BORDER_MODELS,
  inset: INSET_MODELS,
};

/** The coverage of each pixel row from `from` to `to` inclusive under a band. */
export function profileOf(b: Band, from: number, to: number): number[] {
  const out: number[] = [];
  for (let r = from; r <= to; r++) {
    out.push(Math.min(1, Math.max(0, Math.min(b.bottom, r + 1) - Math.max(b.top, r))));
  }
  return out;
}

/** The rows a probe's window covers at a scale, and its nominal centre and width in pixels. */
export function windowOf(
  probe: Probe,
  scale: number,
): { from: number; to: number; c: number; w: number } {
  return {
    from: Math.round((probe.read.at - probe.read.half) * scale),
    to: Math.round((probe.read.at + probe.read.half) * scale),
    c: exact(probe.centrePt * scale),
    w: exact((probe.widthPt ?? 0) * scale),
  };
}

/** A device coordinate to the nanopixel, so an inexact scale cannot put an exact half a hair under it. */
export const exact = (v: number): number => Math.round(v * 1e9) / 1e9;

/** A model's profile over a probe's window at a scale. */
export function predictedProfile(model: Model, probe: Probe, scale: number): number[] {
  const { from, to, c, w } = windowOf(probe, scale);
  return profileOf(model.predict(c, w, probe, scale), from, to);
}

/** Refuse two models no scored width could tell apart on the probes that carry them. */
export function assertSeparable(
  probes: readonly Probe[],
  widths: readonly number[] = GRID_WIDTHS,
): void {
  for (const [family, models] of Object.entries(MODELS_BY_FAMILY)) {
    const carriers = probes.filter((p) => p.family === family);
    for (let a = 0; a < models.length; a++) {
      for (let b = a + 1; b < models.length; b++) {
        const separated = carriers.some((probe) =>
          widths.some((width) => {
            const pa = predictedProfile(models[a]!, probe, axisScale(width, probe.read.axis));
            const pb = predictedProfile(models[b]!, probe, axisScale(width, probe.read.axis));
            return pa.some((v, i) => Math.abs(v - pb[i]!) > SEPARABLE_BY);
          }),
        );
        if (!separated) {
          throw new Error(
            `${family}: ${models[a]!.name} and ${models[b]!.name} predict the same profile at every width - the experiment cannot separate them`,
          );
        }
      }
    }
  }
}
