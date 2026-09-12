/**
 * Experiment F2 - what PowerPoint's export does at another width that a linear scale of 960 does not.
 *
 * Seven slides exported at seven widths; each family names the readings a width can separate, and
 * `assertSeparable` refuses a pair none of them can. No `author.ps1`: nothing needs PowerPoint to author.
 */

import { ln, LINE_GEOM, prstDash, RECT_GEOM } from '../../paint/lines/probes.ts';

/** Export widths, in pixels: zoom 1/8 to 4 over a 960-point slide, and 1.25 so a 6-pt tile lands between pixels. */
export const EXPORT_WIDTHS: readonly number[] = [120, 240, 480, 960, 1200, 1920, 3840];

/** The slide, in points. */
export const SLIDE = { w: 960, h: 540 } as const;

export const EMU_PER_POINT = 12700;

export type Family =
  'hairline' | 'thin' | 'dash' | 'marker' | 'text' | 'pattern' | 'gradient' | 'border' | 'frame';

/** Where the analysis reads, in points on the slide. */
export type Read =
  /** Columns crossing a horizontal stroke: ink per column, averaged over the columns. */
  | {
      readonly kind: 'across-h';
      readonly x0: number;
      readonly x1: number;
      readonly y0: number;
      readonly y1: number;
    }
  /** Rows crossing a vertical stroke. */
  | {
      readonly kind: 'across-v';
      readonly y0: number;
      readonly y1: number;
      readonly x0: number;
      readonly x1: number;
    }
  /** One row along a horizontal stroke: the on/off runs. */
  | { readonly kind: 'along-h'; readonly y: number; readonly x0: number; readonly x1: number }
  /** The ink box inside a region. */
  | {
      readonly kind: 'box';
      readonly x0: number;
      readonly y0: number;
      readonly x1: number;
      readonly y1: number;
    }
  /** One column through a pattern: its period. */
  | { readonly kind: 'period-v'; readonly x: number; readonly y0: number; readonly y1: number }
  /** One row through a gradient: luma at ten positions. */
  | { readonly kind: 'ramp-h'; readonly y: number; readonly x0: number; readonly x1: number }
  /** The whole bitmap. */
  | { readonly kind: 'frame' };

export interface Probe {
  readonly id: string;
  /** 1-based slide index in the deck. */
  readonly slide: number;
  readonly family: Family;
  readonly question: string;
  /** Already-built `p:spTree` children, or nothing for a background-only probe. */
  readonly markup: string;
  readonly read: Read;
  /** Nominal stroke width in points, for the families a width parameterises. */
  readonly widthPt?: number;
  /** The period along the read column, in points, for a pattern tile of 6 pt (ADR 0022). */ readonly periodPt?: number;
}

/** A model of one measured number as a function of the export scale. */
export interface Model {
  readonly name: string;
  readonly description: string;
  /** The prediction in pixels, or `null` where the model makes none at that scale. */
  readonly predict: (scale: number, probe: Probe) => number | null;
}

const WHITE = '<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>';

function emu(points: number): string {
  return String(points * EMU_PER_POINT);
}

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
  readonly flipV?: boolean;
}

function shape(spec: ShapeSpec): string {
  const flip = spec.flipV === true ? ' flipV="1"' : '';
  return (
    '<p:sp><p:nvSpPr>' +
    `<p:cNvPr id="${String(spec.id)}" name="${spec.name}"/><p:cNvSpPr/><p:nvPr/>` +
    '</p:nvSpPr><p:spPr>' +
    `<a:xfrm${flip}><a:off x="${emu(spec.x)}" y="${emu(spec.y)}"/>` +
    `<a:ext cx="${emu(spec.cx)}" cy="${emu(spec.cy)}"/></a:xfrm>` +
    spec.geom +
    spec.fill +
    spec.line +
    '</p:spPr></p:sp>'
  );
}

function textBox(id: number, name: string, x: number, y: number, run: string): string {
  return (
    '<p:sp><p:nvSpPr>' +
    `<p:cNvPr id="${String(id)}" name="${name}"/><p:cNvSpPr txBox="1"/><p:nvPr/>` +
    '</p:nvSpPr><p:spPr>' +
    `<a:xfrm><a:off x="${emu(x)}" y="${emu(y)}"/><a:ext cx="${emu(400)}" cy="${emu(40)}"/></a:xfrm>` +
    RECT_GEOM +
    '<a:noFill/></p:spPr><p:txBody>' +
    '<a:bodyPr wrap="none" lIns="0" tIns="0" rIns="0" bIns="0"><a:noAutofit/></a:bodyPr><a:lstStyle/>' +
    `<a:p>${run}</a:p></p:txBody></p:sp>`
  );
}

function picture(id: number, name: string, x: number, y: number, line: string): string {
  return (
    '<p:pic><p:nvPicPr>' +
    `<p:cNvPr id="${String(id)}" name="${name}"/><p:cNvPicPr><a:picLocks/></p:cNvPicPr><p:nvPr/>` +
    '</p:nvPicPr><p:blipFill><a:blip r:embed="rId2"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>' +
    `<p:spPr><a:xfrm><a:off x="${emu(x)}" y="${emu(y)}"/><a:ext cx="${emu(300)}" cy="${emu(100)}"/></a:xfrm>` +
    RECT_GEOM +
    line +
    '</p:spPr></p:pic>'
  );
}

/* -------------------------------------------------------------------------- */
/* the probes                                                                 */
/* -------------------------------------------------------------------------- */

export const THIN_WIDTHS_PT: readonly number[] = [0.25, 0.3, 0.5, 0.75, 1, 1.1, 1.5, 2.25];

export const TEXT_SIZES_PT: readonly number[] = [6, 8, 10, 12, 18];

/** Stroke widths a large triangle head is measured on; C4 measured 4, 8 and 16 pt only. */
export const MARKER_WIDTHS_PT: readonly number[] = [0, 0.5, 1, 1.5, 2, 2.5, 3];

export function zoomProbes(): Probe[] {
  const probes: Probe[] = [];
  let id = 2;

  /* -------------------------------------------------------------- slide 1 */
  const hair = ln({ rawW: '0' });
  probes.push({
    id: 'hair-h',
    slide: 1,
    family: 'hairline',
    question: 'a horizontal w="0" line: one device pixel, or a fraction of a point?',
    markup: shape({
      id: id++,
      name: 'hair-h',
      x: 100,
      y: 100,
      cx: 400,
      cy: 0,
      geom: LINE_GEOM,
      fill: '<a:noFill/>',
      line: hair,
    }),
    read: { kind: 'across-h', x0: 150, x1: 450, y0: 84, y1: 116 },
    widthPt: 0,
  });
  probes.push({
    id: 'hair-v',
    slide: 1,
    family: 'hairline',
    question: 'the same hairline drawn vertically',
    markup: shape({
      id: id++,
      name: 'hair-v',
      x: 700,
      y: 60,
      cx: 0,
      cy: 300,
      geom: LINE_GEOM,
      fill: '<a:noFill/>',
      line: hair,
    }),
    read: { kind: 'across-v', y0: 100, y1: 320, x0: 684, x1: 716 },
    widthPt: 0,
  });
  probes.push({
    id: 'hair-d',
    slide: 1,
    family: 'hairline',
    question: 'the same hairline at 45 degrees: a pixel per column, or a pixel across?',
    markup: shape({
      id: id++,
      name: 'hair-d',
      x: 100,
      y: 180,
      cx: 200,
      cy: 200,
      geom: LINE_GEOM,
      fill: '<a:noFill/>',
      line: hair,
    }),
    read: { kind: 'across-h', x0: 150, x1: 250, y0: 200, y1: 380 },
    widthPt: 0,
  });
  probes.push({
    id: 'hair-rect',
    slide: 1,
    family: 'hairline',
    question: 'a hairline outline on an unfilled rectangle',
    markup: shape({
      id: id++,
      name: 'hair-rect',
      x: 400,
      y: 200,
      cx: 200,
      cy: 100,
      geom: RECT_GEOM,
      fill: '<a:noFill/>',
      line: hair,
    }),
    read: { kind: 'across-h', x0: 440, x1: 560, y0: 184, y1: 216 },
    widthPt: 0,
  });
  probes.push({
    id: 'hair-rect-filled',
    slide: 1,
    family: 'hairline',
    question: 'the same outline over a white fill',
    markup: shape({
      id: id++,
      name: 'hair-rect-filled',
      x: 400,
      y: 380,
      cx: 200,
      cy: 100,
      geom: RECT_GEOM,
      fill: WHITE,
      line: hair,
    }),
    read: { kind: 'across-h', x0: 440, x1: 560, y0: 364, y1: 396 },
    widthPt: 0,
  });

  /* -------------------------------------------------------------- slide 2 */
  THIN_WIDTHS_PT.forEach((w, i) => {
    const y = 60 + i * 60;
    const name = `thin-${String(w).replace('.', '_')}`;
    probes.push({
      id: name,
      slide: 2,
      family: 'thin',
      question: `a ${String(w)}pt line: true width, or a device-pixel minimum?`,
      markup: shape({
        id: id++,
        name,
        x: 100,
        y,
        cx: 600,
        cy: 0,
        geom: LINE_GEOM,
        fill: '<a:noFill/>',
        line: ln({ rawW: String(Math.round(w * EMU_PER_POINT)) }),
      }),
      read: { kind: 'across-h', x0: 200, x1: 600, y0: y - 16, y1: y + 16 },
      widthPt: w,
    });
  });

  /* -------------------------------------------------------------- slide 3 */
  const dashes: readonly [string, string][] = [
    ['dash', 'dash'],
    ['dot', 'dot'],
    ['sysDash', 'sysDash'],
  ];
  dashes.forEach(([name, val], i) => {
    const y = 80 + i * 80;
    probes.push({
      id: `dash-hair-${name}`,
      slide: 3,
      family: 'dash',
      question: `prstDash ${val} on a w="0" line: which width sets the period?`,
      markup: shape({
        id: id++,
        name: `dash-hair-${name}`,
        x: 100,
        y,
        cx: 700,
        cy: 0,
        geom: LINE_GEOM,
        fill: '<a:noFill/>',
        line: ln({ rawW: '0', dash: prstDash(val) }),
      }),
      read: { kind: 'along-h', y, x0: 120, x1: 780 },
      widthPt: 0,
    });
  });
  probes.push({
    id: 'dash-quarter',
    slide: 3,
    family: 'dash',
    question: 'prstDash dash on a 0.25pt line',
    markup: shape({
      id: id++,
      name: 'dash-quarter',
      x: 100,
      y: 320,
      cx: 700,
      cy: 0,
      geom: LINE_GEOM,
      fill: '<a:noFill/>',
      line: ln({ rawW: '3175', dash: prstDash('dash') }),
    }),
    read: { kind: 'along-h', y: 320, x0: 120, x1: 780 },
    widthPt: 0.25,
  });
  MARKER_WIDTHS_PT.forEach((w, i) => {
    const column = i % 3;
    const y = 360 + Math.floor(i / 3) * 60;
    const x = 100 + column * 300;
    const tip = x + 200;
    const name = `marker-${String(w).replace('.', '_')}`;
    probes.push({
      id: name,
      slide: 3,
      family: 'marker',
      question: `a large triangle head on a ${String(w)}pt line: sized from what width?`,
      markup: shape({
        id: id++,
        name,
        x,
        y,
        cx: 200,
        cy: 0,
        geom: LINE_GEOM,
        fill: '<a:noFill/>',
        line: ln({
          rawW: String(Math.round(w * EMU_PER_POINT)),
          tailEnd: '<a:tailEnd type="triangle" w="lg" len="lg"/>',
        }),
      }),
      read: { kind: 'box', x0: tip - 40, y0: y - 28, x1: tip + 40, y1: y + 28 },
      widthPt: w,
    });
  });

  /* -------------------------------------------------------------- slide 4 */
  const faces: readonly [string, number][] = [
    ...TEXT_SIZES_PT.map((sz): [string, number] => ['Arial', sz]),
    ['Calibri', 12],
    ['Cambria', 12],
  ];
  faces.forEach(([face, sz], i) => {
    const y = 40 + i * 60;
    const name = `text-${face.toLowerCase()}-${String(sz)}`;
    const run =
      `<a:r><a:rPr lang="en-US" sz="${String(sz * 100)}" dirty="0">` +
      `<a:solidFill><a:srgbClr val="000000"/></a:solidFill>` +
      `<a:latin typeface="${face}"/><a:cs typeface="${face}"/></a:rPr>` +
      '<a:t>Hamburgefonstiv 0123</a:t></a:r>';
    probes.push({
      id: name,
      slide: 4,
      family: 'text',
      question: `${face} at ${String(sz)}pt: does its extent scale with the export?`,
      markup: textBox(id++, name, 100, y, run),
      read: { kind: 'box', x0: 90, y0: y - 8, x1: 700, y1: y + 48 },
    });
  });

  /* -------------------------------------------------------------- slide 5 */
  probes.push({
    id: 'pattern-horz',
    slide: 5,
    family: 'pattern',
    question: 'pattFill horz: is the tile still six points at 1/8 and at 4x?',
    markup: shape({
      id: id++,
      name: 'pattern-horz',
      x: 100,
      y: 100,
      cx: 300,
      cy: 200,
      geom: RECT_GEOM,
      fill: '<a:pattFill prst="horz"><a:fgClr><a:srgbClr val="000000"/></a:fgClr><a:bgClr><a:srgbClr val="FFFFFF"/></a:bgClr></a:pattFill>',
      line: '<a:ln><a:noFill/></a:ln>',
    }),
    read: { kind: 'period-v', x: 250, y0: 110, y1: 290 },
    periodPt: 6,
  });
  probes.push({
    id: 'pattern-pct50',
    slide: 5,
    family: 'pattern',
    question: 'pattFill pct50: does its coverage hold where the tile cannot be resolved?',
    markup: shape({
      id: id++,
      name: 'pattern-pct50',
      x: 500,
      y: 100,
      cx: 300,
      cy: 200,
      geom: RECT_GEOM,
      fill: '<a:pattFill prst="pct50"><a:fgClr><a:srgbClr val="000000"/></a:fgClr><a:bgClr><a:srgbClr val="FFFFFF"/></a:bgClr></a:pattFill>',
      line: '<a:ln><a:noFill/></a:ln>',
    }),
    read: { kind: 'period-v', x: 650, y0: 110, y1: 290 },
    periodPt: 1.5,
  });
  probes.push({
    id: 'gradient-lin',
    slide: 5,
    family: 'gradient',
    question: 'a linear black-to-white ramp: the same curve at every width?',
    markup: shape({
      id: id++,
      name: 'gradient-lin',
      x: 100,
      y: 400,
      cx: 400,
      cy: 60,
      geom: RECT_GEOM,
      fill: '<a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="000000"/></a:gs><a:gs pos="100000"><a:srgbClr val="FFFFFF"/></a:gs></a:gsLst><a:lin ang="0" scaled="0"/></a:gradFill>',
      line: '<a:ln><a:noFill/></a:ln>',
    }),
    read: { kind: 'ramp-h', y: 430, x0: 100, x1: 500 },
  });

  probes.push({
    id: 'border-1pt',
    slide: 6,
    family: 'border',
    question: 'a 1pt picture border: thickness by width, and still wholly outside?',
    markup: picture(id++, 'border-1pt', 100, 30, ln({ rawW: '12700' })),
    read: { kind: 'across-h', x0: 150, x1: 350, y0: 14, y1: 46 },
    widthPt: 1,
  });
  probes.push({
    id: 'border-1_1pt',
    slide: 6,
    family: 'border',
    question: 'a 1.1pt picture border, which is 4.4 px at 4x: true, rounded, or a minimum?',
    markup: picture(id++, 'border-1_1pt', 100, 200, ln({ rawW: '13970' })),
    read: { kind: 'across-h', x0: 150, x1: 350, y0: 184, y1: 216 },
    widthPt: 1.1,
  });
  probes.push({
    id: 'border-hair',
    slide: 6,
    family: 'hairline',
    question: 'a w="0" border on a picture',
    markup: picture(id++, 'border-hair', 100, 370, hair),
    read: { kind: 'across-h', x0: 150, x1: 350, y0: 354, y1: 386 },
    widthPt: 0,
  });

  /* -------------------------------------------------------------- slide 7 */
  probes.push({
    id: 'frame',
    slide: 7,
    family: 'frame',
    question: 'a 16:9 slide exported at 120x68: stretched to the pixels, or letterboxed?',
    markup: '',
    read: { kind: 'frame' },
  });

  return probes;
}

/** The slide backgrounds, by 1-based slide: the frame probe is the only non-white one. */
export function backgroundOf(slide: number): string | undefined {
  return slide === 7 ? '<a:solidFill><a:srgbClr val="808080"/></a:solidFill>' : undefined;
}

/* -------------------------------------------------------------------------- */
/* the candidate readings                                                     */
/* -------------------------------------------------------------------------- */

/** Ink of a hairline, in pixels, as a function of the scale `s` = px per pt. */
export const HAIRLINE_MODELS: readonly Model[] = [
  { name: 'H1', description: 'one device pixel at every width', predict: () => 1 },
  { name: 'H2', description: 'half a point', predict: (s) => 0.5 * s },
  { name: 'H3', description: 'three quarters of a point', predict: (s) => 0.75 * s },
  {
    name: 'H4',
    description: 'half a point, never under one pixel',
    predict: (s) => Math.max(1, 0.5 * s),
  },
  {
    name: 'H5',
    description: 'three quarters of a point, never under one pixel',
    predict: (s) => Math.max(1, 0.75 * s),
  },
];

/** Ink of a thin stroke of `widthPt`, in pixels. */
export const THIN_MODELS: readonly Model[] = [
  {
    name: 'T1',
    description: 'the true width, antialiased',
    predict: (s, p) => (p.widthPt ?? 0) * s,
  },
  {
    name: 'T2',
    description: 'the true width, never under one pixel',
    predict: (s, p) => Math.max(1, (p.widthPt ?? 0) * s),
  },
  {
    name: 'T3',
    description: 'rounded to whole pixels',
    predict: (s, p) => Math.round((p.widthPt ?? 0) * s),
  },
  {
    name: 'T4',
    description: 'rounded to whole pixels, never under one',
    predict: (s, p) => Math.max(1, Math.round((p.widthPt ?? 0) * s)),
  },
  {
    name: 'T5',
    description: 'rounded up to whole pixels',
    predict: (s, p) => Math.ceil((p.widthPt ?? 0) * s),
  },
];

/** Multiples of the drawn width in one period of `dash`, `dot` and `sysDash`, from C4. */
export const DASH_PERIOD_WIDTHS: Readonly<Record<string, number>> = {
  dash: 7,
  dot: 4,
  sysDash: 4,
};

function periodWidths(probe: Probe): number {
  const name = probe.id.replace(/^dash-(hair-)?/, '');
  return DASH_PERIOD_WIDTHS[name === 'quarter' ? 'dash' : name] ?? 7;
}

/** The period of a dashed hairline, in pixels. */
export const DASH_MODELS: readonly Model[] = [
  {
    name: 'D1',
    description: 'the period of the drawn pixel width; a zero width has no dashes to draw',
    predict: (s, p) =>
      (p.widthPt ?? 0) === 0 ? null : periodWidths(p) * Math.max(1, (p.widthPt ?? 0) * s),
  },
  {
    name: 'D2',
    description: 'the period of a three-quarter-point line, whatever the width',
    predict: (s, p) => periodWidths(p) * 0.75 * s,
  },
  {
    name: 'D3',
    description: 'the period of a half-point line, whatever the width',
    predict: (s, p) => periodWidths(p) * 0.5 * s,
  },
  { name: 'D4', description: 'solid at every width', predict: () => null },
  {
    name: 'D5',
    description: 'the period of the true width, with no minimum',
    predict: (s, p) => periodWidths(p) * (p.widthPt ?? 0) * s,
  },
  {
    name: 'D6',
    description: 'the period of the drawn pixel width, a zero width drawn as one pixel',
    predict: (s, p) => periodWidths(p) * Math.max(1, (p.widthPt ?? 0) * s),
  },
];

/** The ink height at the back of a large triangle head, in pixels: lg is 5 widths (C4). */
export const MARKER_MODELS: readonly Model[] = [
  {
    name: 'M1',
    description: 'five times the nominal width, so nothing on a hairline',
    predict: (s, p) => Math.max(1, 5 * (p.widthPt ?? 0) * s),
  },
  {
    name: 'M2',
    description: 'five times the width, never under three quarters of a point',
    predict: (s, p) => Math.max(1, 5 * Math.max(0.75, p.widthPt ?? 0) * s),
  },
  {
    name: 'M3',
    description: 'five times the drawn pixel width',
    predict: (s, p) => 5 * Math.max(1, (p.widthPt ?? 0) * s),
  },
  {
    name: 'M4',
    description: 'five times the width, never under two points',
    predict: (s, p) => Math.max(1, 5 * Math.max(2, p.widthPt ?? 0) * s),
  },
  {
    name: 'M5',
    description: 'five times the width, never under two points, each never under a pixel',
    predict: (s, p) => 5 * Math.max(1, Math.max(2, p.widthPt ?? 0) * s),
  },
  {
    name: 'M6',
    description: 'five times a pen of max(2 pt, w), the pen rounded up to whole pixels',
    predict: (s, p) => 5 * Math.max(1, Math.ceil(Math.max(2, p.widthPt ?? 0) * s - 1e-9)),
  },
  {
    name: 'M7',
    description: 'five times a pen of max(2 pt, w), the pen rounded to whole pixels',
    predict: (s, p) => 5 * Math.max(1, Math.round(Math.max(2, p.widthPt ?? 0) * s)),
  },
  {
    name: 'M8',
    description: 'a ten-point head at or under two points, else five whole-pixel pens',
    predict: (s, p) =>
      (p.widthPt ?? 0) <= 2
        ? 5 * Math.max(1, 2 * s)
        : 5 * Math.max(1, Math.round((p.widthPt ?? 0) * s)),
  },
];

/** The period along the read column, in pixels; a tile is 6 pt (ADR 0022), so `periodPt` is its share. */
export const PATTERN_MODELS: readonly Model[] = [
  {
    name: 'P1',
    description: 'the tile is six points at every width',
    predict: (s, p) => (p.periodPt ?? 6) * s,
  },
  {
    name: 'P2',
    description: 'six points, snapped to whole pixels',
    predict: (s, p) => Math.max(1, Math.round((p.periodPt ?? 6) * s)),
  },
  {
    name: 'P3',
    description: 'the tile is eight device pixels at every width',
    predict: (s, p) => (8 * (p.periodPt ?? 6)) / 6,
  },
];

/** How far a border's inner edge sits inside the frame edge, in pixels (ADR 0037: wholly outside). */
export const BORDER_EDGE_MODELS: readonly Model[] = [
  { name: 'O1', description: 'wholly outside: inner edge on the frame', predict: () => 0 },
  {
    name: 'O2',
    description: 'centred on the frame',
    predict: (s, p) => ((p.widthPt ?? 1) / 2) * s,
  },
  { name: 'O3', description: 'wholly inside', predict: (s, p) => (p.widthPt ?? 1) * s },
  {
    name: 'O4',
    description: 'wholly outside, the whole-pixel stroke snapped to the grid, half up',
    predict: (s, p) => {
      if (p.read.kind !== 'across-h') return null;
      const edge = ((p.read.y0 + p.read.y1) / 2) * s;
      return Math.round(edge) - edge;
    },
  },
];

export const MODELS_BY_FAMILY: Readonly<Partial<Record<Family, readonly Model[]>>> = {
  hairline: HAIRLINE_MODELS,
  thin: THIN_MODELS,
  dash: DASH_MODELS,
  marker: MARKER_MODELS,
  pattern: PATTERN_MODELS,
  border: THIN_MODELS,
};

/** Refuse two models no width in the list could tell apart on the probes that carry them. */
export function assertSeparable(
  probes: readonly Probe[],
  widths: readonly number[] = EXPORT_WIDTHS,
): void {
  for (const [family, models] of Object.entries(MODELS_BY_FAMILY)) {
    const carriers = probes.filter((p) => p.family === family);
    for (let a = 0; a < models.length; a++) {
      for (let b = a + 1; b < models.length; b++) {
        const separated = carriers.some((probe) =>
          widths.some((w) => {
            const s = w / SLIDE.w;
            const pa = models[a]!.predict(s, probe);
            const pb = models[b]!.predict(s, probe);
            if (pa === null || pb === null) return pa !== pb;
            return Math.abs(pa - pb) > 0.3;
          }),
        );
        if (!separated) {
          throw new Error(
            `${family}: ${models[a]!.name} and ${models[b]!.name} predict the same number at every width - the experiment cannot separate them`,
          );
        }
      }
    }
  }
}
