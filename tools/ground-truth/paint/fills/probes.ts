/**
 * Experiment C3 - the fill questions: gradients, and the 54 pattern tiles.
 *
 * Sub-phase 2.6 measured what a *colour* resolves to. A fill is a field of
 * colours, and almost nothing about how DrawingML lays that field out is
 * written down in a way you could implement from. `a:lin/@ang` has a zero
 * direction and a sense the standard never states; `@scaled` changes the visual
 * angle on any shape that is not square and the standard's one sentence about it
 * does not say by how much; `a:fillToRect` describes a rectangle a browser has
 * no primitive for; and the 54 pattern tiles are not in the specification at
 * all - only their names are.
 *
 * ## What the probes are for
 *
 * Before any of this was written, PowerPoint was asked to *author* fills through
 * its own object model, and what it wrote settled three things for free
 * (`tools/ground-truth/README.md`, section C3):
 *
 *   - `Fill.Patterned(i)` over the whole `MsoPatternType` range is accepted for
 *     exactly 54 values, and PowerPoint spells each one into `a:pattFill/@prst`
 *     itself. `PRESET_PATTERNS` below is therefore PowerPoint's enumeration, not
 *     a transcription of one.
 *   - A two-colour gradient is written with **two** stops. Not 33. Whatever the
 *     "pre-sampled gamma ramp" is, PowerPoint's authoring UI does not do it, so
 *     if it is real it belongs to the renderer, and only a bitmap can say.
 *   - PowerPoint writes `a:gsLst` **out of `@pos` order** - its own from-centre
 *     variant 4 is `pos="50000"`, `pos="0"`, `pos="100000"`. A renderer that
 *     trusts document order is wrong on a file PowerPoint wrote.
 *
 * These probes ask what is left, which is everything geometric.
 *
 * ## One question per package
 *
 * Same rule as C2, same reason: PowerPoint's only diagnostic is "The file or
 * directory is corrupted and unreadable", so a deck that asks six questions and
 * comes back repaired has measured the repair. Every probe that might be refused
 * - an angle outside `ST_PositiveFixedAngle`, a stop position outside
 * `ST_PositiveFixedPercentage`, a `pattFill` with no `@prst` - is alone in its
 * own package, so a refusal names itself.
 */

/* -------------------------------------------------------------------------- */
/* colours                                                                    */
/* -------------------------------------------------------------------------- */

/** `<a:srgbClr val="RRGGBB"/>`, with any transforms nested inside. */
export function srgb(hex: string, transforms = ''): string {
  return transforms === ''
    ? `<a:srgbClr val="${hex}"/>`
    : `<a:srgbClr val="${hex}">${transforms}</a:srgbClr>`;
}

/** `<a:schemeClr val="name"/>`, with any transforms nested inside. */
export function scheme(name: string, transforms = ''): string {
  return transforms === ''
    ? `<a:schemeClr val="${name}"/>`
    : `<a:schemeClr val="${name}">${transforms}</a:schemeClr>`;
}

/* -------------------------------------------------------------------------- */
/* fills                                                                      */
/* -------------------------------------------------------------------------- */

/**
 * One gradient stop. `pos` is written verbatim, including the string spellings
 * and out-of-range values the hostile probes need.
 */
export function gs(pos: number | string, color: string): string {
  return `<a:gs pos="${String(pos)}">${color}</a:gs>`;
}

export interface GradOptions {
  readonly flip?: string;
  readonly rotWithShape?: 0 | 1;
  readonly tileRect?: string;
}

/**
 * `CT_GradientFill`. The child order is `gsLst`, then the shade properties
 * (`a:lin` or `a:path`), then `a:tileRect` - an `xsd:sequence`, so it is the
 * only order that opens.
 */
export function gradFill(
  stops: readonly string[],
  shade: string,
  options: GradOptions = {},
): string {
  const attrs =
    (options.flip === undefined ? '' : ` flip="${options.flip}"`) +
    (options.rotWithShape === undefined ? '' : ` rotWithShape="${String(options.rotWithShape)}"`);
  const gsLst = stops.length === 0 ? '<a:gsLst/>' : `<a:gsLst>${stops.join('')}</a:gsLst>`;
  return `<a:gradFill${attrs}>${gsLst}${shade}${options.tileRect ?? ''}</a:gradFill>`;
}

export function lin(ang: number | string, scaled: 0 | 1): string {
  return `<a:lin ang="${String(ang)}" scaled="${String(scaled)}"/>`;
}

interface Insets {
  readonly l?: number;
  readonly t?: number;
  readonly r?: number;
  readonly b?: number;
}

function insetAttrs(rect: Insets): string {
  return (['l', 't', 'r', 'b'] as const)
    .filter((k) => rect[k] !== undefined)
    .map((k) => `${k}="${String(rect[k])}"`)
    .join(' ');
}

/** `a:path` with an `a:fillToRect`. Omitted sides default to 0, as in the schema. */
export function path(kind: 'shape' | 'circle' | 'rect', rect: Insets = {}): string {
  return `<a:path path="${kind}"><a:fillToRect ${insetAttrs(rect)}/></a:path>`;
}

/** A point focus at `(fx, fy)` in shape fractions, written as four equal insets. */
export function focusAt(fx: number, fy: number): Insets {
  return {
    l: Math.round(fx * 100000),
    t: Math.round(fy * 100000),
    r: Math.round((1 - fx) * 100000),
    b: Math.round((1 - fy) * 100000),
  };
}

export function tileRect(rect: Insets = {}): string {
  const parts = insetAttrs(rect);
  return parts === '' ? '<a:tileRect/>' : `<a:tileRect ${parts}/>`;
}

/** `CT_PatternFill`. `fgClr` then `bgClr`, in that order. */
export function pattFill(prst: string | null, fg: string | null, bg: string | null): string {
  const attr = prst === null ? '' : ` prst="${prst}"`;
  const body =
    (fg === null ? '' : `<a:fgClr>${fg}</a:fgClr>`) +
    (bg === null ? '' : `<a:bgClr>${bg}</a:bgClr>`);
  return body === '' ? `<a:pattFill${attr}/>` : `<a:pattFill${attr}>${body}</a:pattFill>`;
}

/* -------------------------------------------------------------------------- */
/* the probes                                                                 */
/* -------------------------------------------------------------------------- */

/**
 * How the analysis reads this shape out of the exported bitmap.
 *
 * - `strip` - samples along the horizontal centre line. For a ramp, where what
 *   matters is the shape of the curve and one row of a 1920px export gives it at
 *   full resolution.
 * - `grid` - a lattice over the interior. For anything two-dimensional: an
 *   angle, a path gradient, a flip.
 * - `tile` - the pattern period, recovered from the pixels rather than assumed.
 * - `corner` - 129 pixels at 1:1, centred on the shape's horizontal midpoint.
 *   A 481-sample strip over a 1800 px shape is one sample every 3.7 pixels, which
 *   cannot resolve a feature 1.6 pixels wide.
 */
export type SampleMode = 'strip' | 'grid' | 'tile' | 'corner';

export interface Probe {
  /** Unique, and used verbatim as `p:cNvPr/@name` so the readback can find it. */
  readonly id: string;
  /** Which `.pptx` this goes in. One question per package; see the header. */
  readonly deck: string;
  readonly group: string;
  /** What this probe is for, in one sentence. Goes into the fixture. */
  readonly question: string;
  /** The whole fill element. */
  readonly fill: string;
  readonly sample: SampleMode;
  /** width:height of the shape inside its cell. Omitted means "fill the cell". */
  readonly aspect?: number;
  /** Fraction of the cell the shape occupies. Defaults to 0.88. */
  readonly scale?: number;
  /** An exact size in points, centred in the cell. Overrides `aspect` and `scale`. */
  readonly widthPt?: number;
  readonly heightPt?: number;
  /** The shape is the whole slide, whatever the deck's grid says. */
  readonly fullSlide?: boolean;
  readonly prst?: string;
  readonly rot?: number;
  readonly flipH?: boolean;
  readonly flipV?: boolean;
  /** A fill for an opaque shape drawn first at the same rectangle. Alpha needs one. */
  readonly backdrop?: string;
}

/**
 * Columns and rows *per slide*, per deck. A deck with more probes than cells
 * spills onto further slides.
 *
 * Every grid divides the slide exactly: 12192000 x 6858000 EMU is 960 x 540 pt,
 * and 12700 EMU is one point, so a cell whose edges are whole points lands on
 * exact pixel boundaries at 960, 1920 and 3840 wide. That matters more than it
 * sounds - a shape edge halfway through a pixel is an antialiased pixel, and an
 * antialiased pixel in a pattern tile is a bit we cannot read.
 */
export const DECK_GRID: Readonly<Record<string, { cols: number; rows: number }>> = {
  ramp: { cols: 1, rows: 15 },
  angle: { cols: 4, rows: 3 },
  path: { cols: 4, rows: 3 },
  tile: { cols: 4, rows: 2 },
  // Five across rather than nine: a 192 x 135 pt cell gives a 256 px square to
  // find an 8 px tile in, and 54 patterns over three slides costs nothing.
  pattern: { cols: 5, rows: 4 },
  patsize: { cols: 4, rows: 3 },
  patcolor: { cols: 4, rows: 2 },
  mixed: { cols: 4, rows: 2 },
  focus: { cols: 4, rows: 3 },
  flipasym: { cols: 4, rows: 2 },
  // One row per width, so the 900 pt shape has a row to itself.
  soften: { cols: 1, rows: 4 },
  // Quadrants, so the comparison shape sits away from the slide centre - where a
  // background focus and a shape focus would otherwise coincide.
  'bg-lin': { cols: 2, rows: 2 },
  'bg-lin-vert': { cols: 2, rows: 2 },
  'bg-path': { cols: 2, rows: 2 },
  'bg-path-off': { cols: 2, rows: 2 },
};

const BLACK = '000000';
const WHITE = 'FFFFFF';

/**
 * The 54 `ST_PresetPatternVal` values, in `MsoPatternType` order, as PowerPoint
 * spelled them into a file it wrote itself. Note both `dotDmnd` and `openDmnd`,
 * and that the twelve `pctNN` values are not every multiple of five.
 */
export const PRESET_PATTERNS: readonly string[] = [
  'pct5',
  'pct10',
  'pct20',
  'pct25',
  'pct30',
  'pct40',
  'pct50',
  'pct60',
  'pct70',
  'pct75',
  'pct80',
  'pct90',
  'dkHorz',
  'dkVert',
  'dkDnDiag',
  'dkUpDiag',
  'smCheck',
  'trellis',
  'ltHorz',
  'ltVert',
  'ltDnDiag',
  'ltUpDiag',
  'smGrid',
  'dotDmnd',
  'wdDnDiag',
  'wdUpDiag',
  'dashUpDiag',
  'dashDnDiag',
  'narVert',
  'narHorz',
  'dashVert',
  'dashHorz',
  'lgConfetti',
  'lgGrid',
  'horzBrick',
  'lgCheck',
  'smConfetti',
  'zigZag',
  'solidDmnd',
  'diagBrick',
  'openDmnd',
  'plaid',
  'sphere',
  'weave',
  'dotGrid',
  'divot',
  'shingle',
  'wave',
  'horz',
  'vert',
  'cross',
  'dnDiag',
  'upDiag',
  'diagCross',
];

/** A plain two-stop black-to-white ramp. The only fill whose answer we can read by eye. */
function bw(shade: string, options: GradOptions = {}): string {
  return gradFill([gs(0, srgb(BLACK)), gs(100000, srgb(WHITE))], shade, options);
}

/** A centred point focus - the from-centre fill PowerPoint's own gallery writes. */
const CTR: Insets = { l: 50000, t: 50000, r: 50000, b: 50000 };

export function fillProbes(): Probe[] {
  const out: Probe[] = [];
  const add = (p: Probe): void => {
    out.push(p);
  };

  /* ---- ramp: what curve joins two stops? -------------------------------- */
  //
  // The decisive probe of the sub-phase is the first one. A linear black-to-white
  // ramp is 128 at its midpoint if PowerPoint interpolates in sRGB, 188 if it
  // interpolates in linear light, and 176 if the 1.875 gamma ramp the plan
  // describes is applied by the renderer. Three hypotheses, sixty units apart,
  // one pixel.
  const R = 'ramp';
  add({
    id: 'ramp-bw',
    deck: R,
    group: 'ramp',
    sample: 'strip',
    question: 'Which curve joins two stops: sRGB, linear light, or a gamma ramp?',
    fill: bw(lin(0, 0)),
  });
  add({
    id: 'ramp-wb',
    deck: R,
    group: 'ramp',
    sample: 'strip',
    question: 'Is the ramp symmetric when the two stops are swapped?',
    fill: gradFill([gs(0, srgb(WHITE)), gs(100000, srgb(BLACK))], lin(0, 0)),
  });
  add({
    id: 'ramp-bw-scaled1',
    deck: R,
    group: 'ramp',
    sample: 'strip',
    question: 'At ang=0, does @scaled change anything? It must not.',
    fill: bw(lin(0, 1)),
  });
  add({
    id: 'ramp-rg',
    deck: R,
    group: 'ramp',
    sample: 'strip',
    question: 'Are the three channels interpolated independently?',
    fill: gradFill([gs(0, srgb('FF0000')), gs(100000, srgb('00FF00'))], lin(0, 0)),
  });
  add({
    id: 'ramp-by',
    deck: R,
    group: 'ramp',
    sample: 'strip',
    question: 'Blue to yellow: does any channel behave differently from the others?',
    fill: gradFill([gs(0, srgb('0000FF')), gs(100000, srgb('FFFF00'))], lin(0, 0)),
  });
  add({
    id: 'ramp-hue',
    deck: R,
    group: 'ramp',
    sample: 'strip',
    question: 'Blue to red - an HSL interpolation would take the hue path through magenta.',
    fill: gradFill([gs(0, srgb('0000FF')), gs(100000, srgb('FF0000'))], lin(0, 0)),
  });
  add({
    id: 'ramp-mid',
    deck: R,
    group: 'ramp',
    sample: 'strip',
    question: 'Two midtones, where sRGB and linear light differ least - a sensitivity control.',
    fill: gradFill([gs(0, srgb('404040')), gs(100000, srgb('C0C0C0'))], lin(0, 0)),
  });
  add({
    id: 'ramp-3stop',
    deck: R,
    group: 'ramp',
    sample: 'strip',
    question: 'With an explicit mid stop at 808080, do the halves match the two-stop curve?',
    fill: gradFill(
      [gs(0, srgb(BLACK)), gs(50000, srgb('808080')), gs(100000, srgb(WHITE))],
      lin(0, 0),
    ),
  });
  add({
    id: 'ramp-inset',
    deck: R,
    group: 'ramp',
    sample: 'strip',
    question: 'Outside the first and last stop, is the colour held flat or extrapolated?',
    fill: gradFill([gs(20000, srgb(BLACK)), gs(80000, srgb(WHITE))], lin(0, 0)),
  });
  add({
    id: 'ramp-steep',
    deck: R,
    group: 'ramp',
    sample: 'strip',
    question: 'A 2% ramp in the middle: is the transition hard, and is it centred?',
    fill: gradFill([gs(49000, srgb(BLACK)), gs(51000, srgb(WHITE))], lin(0, 0)),
  });
  add({
    id: 'ramp-revorder',
    deck: R,
    group: 'ramp',
    sample: 'strip',
    question: 'Stops out of document order - the shape PowerPoint writes. Sorted, or not?',
    fill: gradFill(
      [gs(50000, srgb(BLACK)), gs(0, srgb(WHITE)), gs(100000, srgb(WHITE))],
      lin(0, 0),
    ),
  });
  add({
    id: 'ramp-scheme',
    deck: R,
    group: 'ramp',
    sample: 'strip',
    question: 'Theme colours in stops: resolved and transformed before interpolation?',
    fill: gradFill(
      [gs(0, scheme('accent1')), gs(100000, scheme('accent1', '<a:lumMod val="50000"/>'))],
      lin(0, 0),
    ),
  });
  add({
    id: 'ramp-preset03',
    deck: R,
    group: 'ramp',
    sample: 'strip',
    question: 'A PowerPoint preset ramp, verbatim, including its 39999-style stop positions.',
    fill: gradFill(
      [
        gs(0, srgb('000000')),
        gs(39999, srgb('0A128C')),
        gs(70000, srgb('181CC7')),
        gs(88000, srgb('7005D4')),
        gs(100000, srgb('8C3D91')),
      ],
      lin(0, 0),
    ),
  });
  add({
    id: 'ramp-ang180',
    deck: R,
    group: 'ramp',
    sample: 'strip',
    question: 'ang=180 degrees: which end is stop 0 now? Fixes the sense of the angle.',
    fill: bw(lin(10800000, 0)),
  });
  add({
    id: 'ramp-alpha',
    deck: R,
    group: 'ramp',
    sample: 'strip',
    question: 'Is alpha interpolated along the ramp, and composited in which space?',
    backdrop: `<a:solidFill>${srgb('FF0000')}</a:solidFill>`,
    fill: gradFill(
      [
        gs(0, srgb(WHITE, '<a:alpha val="0"/>')),
        gs(100000, srgb(WHITE, '<a:alpha val="100000"/>')),
      ],
      lin(0, 0),
    ),
  });

  /* ---- angle: @ang and @scaled ------------------------------------------ */
  //
  // @scaled is one sentence in the standard and that sentence does not say what
  // the transform is. On a square shape the two readings agree, which is why a
  // renderer that gets it wrong looks correct until somebody draws a banner.
  const A = 'angle';
  const ANGLES: readonly (readonly [string, number])[] = [
    ['000', 0],
    ['045', 2700000],
    ['090', 5400000],
    ['135', 8100000],
    ['180', 10800000],
    ['270', 16200000],
  ];
  for (const [label, ang] of ANGLES) {
    add({
      id: `ang-${label}-sq`,
      deck: A,
      group: 'angle',
      sample: 'grid',
      aspect: 1,
      question: `ang=${label} degrees on a square: the zero direction and the sense.`,
      fill: bw(lin(ang, 0)),
    });
  }
  for (const scaled of [0, 1] as const) {
    const s = String(scaled);
    add({
      id: `ang-045-wide-s${s}`,
      deck: A,
      group: 'angle',
      sample: 'grid',
      aspect: 3,
      question: `45 degrees, scaled=${s}, on a 3:1 shape - where @scaled shows.`,
      fill: bw(lin(2700000, scaled)),
    });
    add({
      id: `ang-045-tall-s${s}`,
      deck: A,
      group: 'angle',
      sample: 'grid',
      aspect: 1 / 3,
      question: `45 degrees, scaled=${s}, on a 1:3 shape.`,
      fill: bw(lin(2700000, scaled)),
    });
    add({
      id: `ang-030-wide-s${s}`,
      deck: A,
      group: 'angle',
      sample: 'grid',
      aspect: 3,
      question: `30 degrees, scaled=${s}, on a 3:1 shape - an angle 45 cannot alias with.`,
      fill: bw(lin(1800000, scaled)),
    });
  }

  /* ---- path: the shapes SVG has no primitive for ------------------------ */
  const P = 'path';
  add({
    id: 'path-circle-ctr',
    deck: P,
    group: 'path',
    sample: 'grid',
    aspect: 1,
    question: 'A centred circle path: is stop 0 at the focus or at the edge?',
    fill: bw(path('circle', CTR)),
  });
  add({
    id: 'path-circle-rev',
    deck: P,
    group: 'path',
    sample: 'grid',
    aspect: 1,
    question: 'The same with the colours swapped - confirms the direction, not the colour.',
    fill: gradFill([gs(0, srgb(WHITE)), gs(100000, srgb(BLACK))], path('circle', CTR)),
  });
  add({
    id: 'path-rect-ctr',
    deck: P,
    group: 'path',
    sample: 'grid',
    aspect: 1,
    question: 'path="rect": a square ramp, for which SVG has no primitive.',
    fill: bw(path('rect', CTR)),
  });
  add({
    id: 'path-shape-ctr',
    deck: P,
    group: 'path',
    sample: 'grid',
    aspect: 1,
    question: 'path="shape" on a rectangle: does it differ from path="rect"?',
    fill: bw(path('shape', CTR)),
  });
  add({
    id: 'path-circle-whole',
    deck: P,
    group: 'path',
    sample: 'grid',
    aspect: 1,
    question: 'fillToRect with no insets: the focus is the whole box, not a point.',
    fill: bw(path('circle', { l: 0, t: 0, r: 0, b: 0 })),
  });
  add({
    id: 'path-circle-focusrect',
    deck: P,
    group: 'path',
    sample: 'grid',
    aspect: 1,
    question: 'A focus that is a real rectangle, half the box: is the middle flat?',
    fill: bw(path('circle', { l: 25000, t: 25000, r: 25000, b: 25000 })),
  });
  add({
    id: 'path-circle-offctr',
    deck: P,
    group: 'path',
    sample: 'grid',
    aspect: 1,
    question: 'An off-centre focus point at (0.2, 0.2): how is the far radius measured?',
    fill: bw(path('circle', { l: 20000, t: 20000, r: 80000, b: 80000 })),
  });
  add({
    id: 'path-rect-corner-pp',
    deck: P,
    group: 'path',
    sample: 'grid',
    aspect: 1,
    question: 'The from-corner fill PowerPoint writes, verbatim: focus top-left, tile doubled.',
    fill: gradFill(
      [gs(0, srgb(BLACK)), gs(100000, srgb(WHITE))],
      path('rect', { r: 100000, b: 100000 }),
      {
        flip: 'none',
        rotWithShape: 1,
        tileRect: tileRect({ l: -100000, t: -100000 }),
      },
    ),
  });
  add({
    id: 'path-rect-corner-notile',
    deck: P,
    group: 'path',
    sample: 'grid',
    aspect: 1,
    question: 'The same focus without the doubled tileRect - isolates what tileRect did.',
    fill: bw(path('rect', { r: 100000, b: 100000 })),
  });
  add({
    id: 'path-shape-ellipse',
    deck: P,
    group: 'path',
    sample: 'grid',
    aspect: 1,
    prst: 'ellipse',
    question: 'path="shape" on an ellipse: does the ramp follow the outline or the box?',
    fill: bw(path('shape', CTR)),
  });
  add({
    id: 'path-circle-wide',
    deck: P,
    group: 'path',
    sample: 'grid',
    aspect: 3,
    question: 'path="circle" on a 3:1 shape: a circle, or an ellipse fitted to the box?',
    fill: bw(path('circle', CTR)),
  });
  add({
    id: 'path-rect-wide',
    deck: P,
    group: 'path',
    sample: 'grid',
    aspect: 3,
    question: 'path="rect" on a 3:1 shape: does the square ramp stretch?',
    fill: bw(path('rect', CTR)),
  });

  /* ---- tile: tileRect and flip ------------------------------------------ */
  const T = 'tile';
  add({
    id: 'tile-lin-none',
    deck: T,
    group: 'tile',
    sample: 'grid',
    aspect: 1,
    question: 'Control: a plain ramp with an empty tileRect.',
    fill: bw(lin(0, 0), { tileRect: tileRect() }),
  });
  add({
    id: 'tile-lin-half',
    deck: T,
    group: 'tile',
    sample: 'grid',
    aspect: 1,
    question: 'A tileRect covering the left half: does the ramp repeat, or stretch?',
    fill: bw(lin(0, 0), { tileRect: tileRect({ r: 50000 }) }),
  });
  add({
    id: 'tile-lin-half-flipx',
    deck: T,
    group: 'tile',
    sample: 'grid',
    aspect: 1,
    question: 'The same tile with flip="x": mirrored on the repeat?',
    fill: bw(lin(0, 0), { flip: 'x', tileRect: tileRect({ r: 50000 }) }),
  });
  add({
    id: 'tile-lin-double',
    deck: T,
    group: 'tile',
    sample: 'grid',
    aspect: 1,
    question: 'A tileRect twice the shape: is only half the ramp visible?',
    fill: bw(lin(0, 0), { tileRect: tileRect({ r: -100000 }) }),
  });
  for (const flip of ['none', 'x', 'y', 'xy'] as const) {
    add({
      id: `tile-path-flip${flip}`,
      deck: T,
      group: 'tile',
      sample: 'grid',
      aspect: 1,
      question: `A quarter-size path tile with flip="${flip}".`,
      fill: bw(path('circle', CTR), { flip, tileRect: tileRect({ r: 50000, b: 50000 }) }),
    });
  }

  /* ---- pattern: 54 tiles, measured -------------------------------------- */
  //
  // The plan said to codegen these from Mono libgdiplus. Downloading it needs a
  // permission this session does not have, and measuring is the better answer
  // anyway: what a renderer needs is what PowerPoint paints, and libgdiplus is a
  // reimplementation of a different Microsoft product.
  for (const prst of PRESET_PATTERNS) {
    add({
      id: `pat-${prst}`,
      deck: 'pattern',
      group: 'pattern',
      sample: 'tile',
      aspect: 1,
      question: `The tile for prst="${prst}".`,
      fill: pattFill(prst, srgb(BLACK), srgb(WHITE)),
    });
  }

  /* ---- patsize: is the tile device pixels, or shape units? -------------- */
  //
  // Decisive for the renderer. An SVG <pattern> in objectBoundingBox units
  // scales with the shape; one in userSpaceOnUse does not. If PowerPoint's tile
  // is a fixed number of *device* pixels then neither is right at every zoom,
  // and the pattern has to be re-emitted per zoom level - a different design.
  for (const prst of ['lgCheck', 'smGrid'] as const) {
    for (const scale of [1, 0.7, 0.5, 0.35, 0.25, 0.15]) {
      const pct = String(Math.round(scale * 100));
      add({
        id: `patsize-${prst}-${pct}`,
        deck: 'patsize',
        group: 'patsize',
        sample: 'tile',
        aspect: 1,
        scale,
        question: `prst="${prst}" at ${pct}% of the cell - does the tile scale with the shape?`,
        fill: pattFill(prst, srgb(BLACK), srgb(WHITE)),
      });
    }
  }

  /* ---- patcolor: how a pattern's two colours resolve -------------------- */
  const PC = 'patcolor';
  add({
    id: 'patc-scheme',
    deck: PC,
    group: 'patcolor',
    sample: 'grid',
    aspect: 1,
    question: 'Themed fg with a transform and a themed bg: both resolved through the theme?',
    fill: pattFill('lgCheck', scheme('accent1', '<a:lumMod val="50000"/>'), scheme('bg1')),
  });
  add({
    id: 'patc-alpha-fg',
    deck: PC,
    group: 'patcolor',
    sample: 'grid',
    aspect: 1,
    backdrop: `<a:solidFill>${srgb('FF0000')}</a:solidFill>`,
    question: 'A translucent foreground over red: does the pattern composite per colour?',
    fill: pattFill('lgCheck', srgb(BLACK, '<a:alpha val="50000"/>'), srgb(WHITE)),
  });
  add({
    id: 'patc-alpha-bg',
    deck: PC,
    group: 'patcolor',
    sample: 'grid',
    aspect: 1,
    backdrop: `<a:solidFill>${srgb('FF0000')}</a:solidFill>`,
    question: 'A translucent background over red, with an opaque foreground.',
    fill: pattFill('lgCheck', srgb(BLACK), srgb(WHITE, '<a:alpha val="50000"/>')),
  });
  add({
    id: 'patc-same',
    deck: PC,
    group: 'patcolor',
    sample: 'grid',
    aspect: 1,
    question: 'fg and bg identical: a flat fill, or does antialiasing still show the tile?',
    fill: pattFill('lgCheck', srgb('4472C4'), srgb('4472C4')),
  });
  add({
    id: 'patc-prstclr',
    deck: PC,
    group: 'patcolor',
    sample: 'grid',
    aspect: 1,
    question: 'A prstClr inside a pattern - the 2.6 preset table, reached through a fill.',
    fill: pattFill('lgCheck', '<a:prstClr val="darkOliveGreen"/>', srgb(WHITE)),
  });
  add({
    id: 'patc-rot',
    deck: PC,
    group: 'patcolor',
    sample: 'grid',
    aspect: 1,
    rot: 2700000,
    question: 'A directional pattern on a shape rotated 45 degrees: does the tile rotate too?',
    fill: pattFill('ltHorz', srgb(BLACK), srgb(WHITE)),
  });
  add({
    id: 'patc-flip',
    deck: PC,
    group: 'patcolor',
    sample: 'grid',
    aspect: 1,
    flipH: true,
    question: 'A diagonal pattern on a horizontally flipped shape: does the tile mirror?',
    fill: pattFill('wdUpDiag', srgb(BLACK), srgb(WHITE)),
  });
  add({
    id: 'patc-ellipse',
    deck: PC,
    group: 'patcolor',
    sample: 'grid',
    aspect: 1,
    prst: 'ellipse',
    question: 'A pattern clipped to a non-rectangular geometry.',
    fill: pattFill('lgGrid', srgb(BLACK), srgb(WHITE)),
  });

  /* ---- mixed: the interactions nothing else covers ---------------------- */
  const M = 'mixed';
  for (const rws of [1, 0] as const) {
    const s = String(rws);
    add({
      id: `mix-rot45-rws${s}`,
      deck: M,
      group: 'mixed',
      sample: 'grid',
      aspect: 1,
      rot: 2700000,
      question: `A horizontal ramp on a shape rotated 45 degrees, rotWithShape=${s}.`,
      fill: bw(lin(0, 0), { rotWithShape: rws }),
    });
    add({
      id: `mix-fliph-rws${s}`,
      deck: M,
      group: 'mixed',
      sample: 'grid',
      aspect: 1,
      flipH: true,
      question: `A horizontal ramp on a horizontally flipped shape, rotWithShape=${s}.`,
      fill: bw(lin(0, 0), { rotWithShape: rws }),
    });
  }
  add({
    id: 'mix-flipv',
    deck: M,
    group: 'mixed',
    sample: 'grid',
    aspect: 1,
    flipV: true,
    question: 'A vertical ramp on a vertically flipped shape.',
    fill: bw(lin(5400000, 0)),
  });
  add({
    id: 'mix-alpha-path',
    deck: M,
    group: 'mixed',
    sample: 'grid',
    aspect: 1,
    backdrop: `<a:solidFill>${srgb('00A000')}</a:solidFill>`,
    question: 'A radial fade to transparent over green - the common vignette fill.',
    fill: gradFill(
      [
        gs(0, srgb(BLACK, '<a:alpha val="100000"/>')),
        gs(100000, srgb(BLACK, '<a:alpha val="0"/>')),
      ],
      path('circle', CTR),
    ),
  });
  add({
    id: 'mix-scheme-tintshade',
    deck: M,
    group: 'mixed',
    sample: 'grid',
    aspect: 1,
    question:
      'The fmtScheme gradient shape: one theme colour, tinted one end and shaded the other.',
    fill: gradFill(
      [
        gs(0, scheme('accent1', '<a:tint val="60000"/><a:satMod val="120000"/>')),
        gs(100000, scheme('accent1', '<a:shade val="80000"/>')),
      ],
      lin(5400000, 1),
    ),
  });
  add({
    id: 'mix-ellipse-grad',
    deck: M,
    group: 'mixed',
    sample: 'grid',
    aspect: 1,
    prst: 'ellipse',
    question: 'A linear ramp on an ellipse: laid out over the box, or over the path?',
    fill: bw(lin(0, 0)),
  });

  /* ---- focus: where an off-centre path focus puts the ramp -------------- */
  //
  // The first pass measured one off-centre focus, found that no concentric
  // normalisation fitted it, and shipped the farthest-corner reading with a
  // `centred: false` flag. One focus is an anecdote. These are fourteen, over
  // three aspect ratios and all three path kinds, chosen so that the readings
  // that fit the single (0.2, 0.2) sample to within a few bytes - concentric
  // circles scaled to the farthest corner, and the per-axis box rule - are
  // separated from each other everywhere except the shape's centre.
  const F = 'focus';
  const FOCI: readonly (readonly [string, number, number])[] = [
    ['1010', 0.1, 0.1],
    ['5015', 0.5, 0.15],
    ['8550', 0.85, 0.5],
    ['3070', 0.3, 0.7],
    ['7030', 0.7, 0.3],
    ['9090', 0.9, 0.9],
    ['1560', 0.15, 0.6],
    ['6085', 0.6, 0.85],
  ];
  for (const [label, fx, fy] of FOCI) {
    add({
      id: `focus-circle-${label}`,
      deck: F,
      group: 'focus',
      sample: 'grid',
      aspect: 1,
      question: `path="circle" with the focus at (${String(fx)}, ${String(fy)}).`,
      fill: bw(path('circle', focusAt(fx, fy))),
    });
  }
  // A circle on a 3:1 shape is isotropic in shape units, so an off-centre focus
  // on a wide shape is where a model written in shape *fractions* comes apart.
  for (const [label, aspect] of [
    ['wide', 3],
    ['tall', 1 / 3],
  ] as const) {
    add({
      id: `focus-circle-${label}`,
      deck: F,
      group: 'focus',
      sample: 'grid',
      aspect,
      question: `path="circle" at (0.25, 0.25) on a ${label} shape - shape units, or fractions?`,
      fill: bw(path('circle', focusAt(0.25, 0.25))),
    });
  }
  add({
    id: 'focus-rect-3070',
    deck: F,
    group: 'focus',
    sample: 'grid',
    aspect: 1,
    question: 'path="rect" off centre: does the box rule survive a focus that is not the centre?',
    fill: bw(path('rect', focusAt(0.3, 0.7))),
  });
  add({
    id: 'focus-rect-wide',
    deck: F,
    group: 'focus',
    sample: 'grid',
    aspect: 3,
    question: 'path="rect" at (0.25, 0.25) on a 3:1 shape.',
    fill: bw(path('rect', focusAt(0.25, 0.25))),
  });
  add({
    id: 'focus-shape-ellipse',
    deck: F,
    group: 'focus',
    sample: 'grid',
    aspect: 1,
    prst: 'ellipse',
    question: 'path="shape" off centre on an ellipse: does the outline still lead?',
    fill: bw(path('shape', focusAt(0.3, 0.3))),
  });
  add({
    id: 'focus-shape-rect',
    deck: F,
    group: 'focus',
    sample: 'grid',
    aspect: 1,
    question: 'path="shape" off centre on a rectangle: still identical to path="rect"?',
    fill: bw(path('shape', focusAt(0.3, 0.7))),
  });

  /* ---- flipasym: an asymmetric tile, which the first pass did not have --- */
  //
  // The first pass put a *centred* path tile under all four flip modes and found
  // them indistinguishable, which says nothing: a centred circle is its own
  // mirror image on both axes. These tiles have no symmetry to hide behind.
  const FA = 'flipasym';
  for (const flip of ['none', 'x', 'y', 'xy'] as const) {
    add({
      id: `flipasym-path-${flip}`,
      deck: FA,
      group: 'flipasym',
      sample: 'grid',
      aspect: 1,
      question: `A quarter-size path tile with an off-centre focus and flip="${flip}".`,
      fill: bw(path('circle', focusAt(0.25, 0.25)), {
        flip,
        tileRect: tileRect({ r: 50000, b: 50000 }),
      }),
    });
    add({
      id: `flipasym-lin-${flip}`,
      deck: FA,
      group: 'flipasym',
      sample: 'grid',
      aspect: 1,
      question: `A quarter-size 30-degree ramp tile with flip="${flip}" - asymmetric on both axes.`,
      fill: bw(lin(1800000, 0), { flip, tileRect: tileRect({ r: 50000, b: 50000 }) }),
    });
  }

  /* ---- soften: is the rounded corner device pixels or a share of the ramp? */
  //
  // A three-colour profile has a hard corner at its middle stop, and PowerPoint
  // rounds it off. What separates a fixed count of device pixels from a share of
  // the ramp is the *slope*: a rounding 1.6 px wide shows up as a deviation
  // proportional to bytes-per-pixel, while a rounding that is a share of the ramp
  // is the same number of bytes however steep the arms are. So the shape stays
  // one size and the arms are shortened instead, from 8% of the width down to 1%
  // - an eightfold change in slope - and `read.ps1` exports at 1920 and 3840 for
  // a further factor of two that no change to the markup can imitate.
  //
  // Three distinct colours, because two colours at 0 and 100% are joined by the
  // measured curve and a curve has no corner to find. The apex is E0 rather than
  // white: a corner at 255 is clipped by the export and the rounding vanishes
  // into the top two bytes, which is what a first attempt at this measured.
  const S = 'soften';
  for (const arm of [1, 2, 4, 8]) {
    const d = arm * 1000;
    add({
      id: `soften-arm${String(arm)}`,
      deck: S,
      group: 'soften',
      sample: 'corner',
      widthPt: 480,
      heightPt: 60,
      question: `A corner at 50% whose arms are ${String(arm)}% of the width each.`,
      fill: gradFill(
        [gs(50000 - d, srgb(BLACK)), gs(50000, srgb('E0E0E0')), gs(50000 + d, srgb('202020'))],
        lin(0, 0),
      ),
    });
  }

  return out;
}

/* -------------------------------------------------------------------------- */
/* the background probes                                                      */
/* -------------------------------------------------------------------------- */

/**
 * What a gradient in `p:bg` is laid out over.
 *
 * A slide background has no shape box, so every question the other probes ask
 * about a rectangle has to be asked again: a linear ramp needs an extent and a
 * path needs a focus, and the slide is the only rectangle in sight. Each deck
 * carries the fill on its slide background and the same fill on a shape covering
 * the right-hand third, so one bitmap answers both halves of the question - if
 * the background runs over the slide, the ramp under the shape and the ramp
 * inside it disagree everywhere except where they happen to cross.
 *
 * The probe named for the background is a `noFill` rectangle spanning the whole
 * slide. It paints nothing; it exists so the sampler has a named rectangle to
 * read the background through, which is the only way this harness can sample
 * something that is not a shape.
 */
export function backgroundProbes(): readonly {
  readonly deck: string;
  readonly background: string;
  readonly probes: readonly Probe[];
}[] {
  const cases: readonly (readonly [string, string, string])[] = [
    [
      'bg-lin',
      'A horizontal ramp on the slide background: is the extent the slide?',
      bw(lin(0, 0)),
    ],
    [
      'bg-lin-vert',
      'A vertical ramp on the slide background - the slide is 16:9, so the two axes differ.',
      bw(lin(5400000, 0)),
    ],
    [
      'bg-path',
      'A centred circle path on the slide background: is the focus the slide centre?',
      bw(path('circle', CTR)),
    ],
    [
      'bg-path-off',
      'An off-centre circle path on the background, at (0.25, 0.25) of the slide.',
      bw(path('circle', focusAt(0.25, 0.25))),
    ],
  ];
  return cases.map(([deck, question, fill]) => ({
    deck,
    background: fill,
    probes: [
      {
        id: `${deck}-slide`,
        deck,
        group: 'background',
        sample: 'grid',
        question,
        fill: '<a:noFill/>',
        fullSlide: true,
      },
      {
        id: `${deck}-shape`,
        deck,
        group: 'background',
        sample: 'grid',
        question: `The same fill on a shape, for comparison with ${deck}-slide.`,
        fill,
        widthPt: 200,
        heightPt: 200,
      },
    ],
  }));
}

/* -------------------------------------------------------------------------- */
/* the hostile probes, one package each                                       */
/* -------------------------------------------------------------------------- */

/**
 * Probes that may be refused. Each is alone in its own `.pptx`, so a refusal
 * names the thing that caused it - which is itself the measurement, and the
 * lesson C2 paid for.
 */
export function hostileProbes(): Probe[] {
  const one = (id: string, question: string, fill: string): Probe => ({
    id,
    deck: id,
    group: 'hostile',
    question,
    fill,
    sample: 'grid',
    aspect: 1,
    scale: 0.9,
  });
  // `@path` and `a:fillToRect` are both optional in `CT_PathShadeProperties`, so
  // these are legal files rather than hostile ones - but a renderer has to
  // default them to something, and an isolated package is what makes a refusal
  // legible if PowerPoint disagrees about the schema.
  const shapes = [
    one(
      'pathdef-nopath',
      'a:path with no @path, focus centred: which of the three kinds is the default?',
      bw('<a:path><a:fillToRect l="50000" t="50000" r="50000" b="50000"/></a:path>'),
    ),
    one(
      'pathdef-norect',
      'path="circle" with no a:fillToRect: is the focus the centre, or the top-left corner?',
      bw('<a:path path="circle"/>'),
    ),
    one(
      'pathdef-norect-rect',
      'path="rect" with no a:fillToRect - the same question, where the corner reading shows.',
      bw('<a:path path="rect"/>'),
    ),
    one('pathdef-bare', 'a:path with neither attribute nor child.', bw('<a:path/>')),
  ].map((p) => ({ ...p, group: 'pathdef' }));

  return [
    ...shapes,
    one(
      'h-angneg',
      'ST_PositiveFixedAngle excludes negatives. Is a negative @ang refused?',
      bw(lin(-2700000, 0)),
    ),
    one(
      'h-angbig',
      '@ang of 405 degrees, past the 21600000 ceiling. Refused, or wrapped?',
      bw(lin(24300000, 0)),
    ),
    one(
      'h-poshi',
      'A stop at pos=150000, past the ST_PositiveFixedPercentage ceiling.',
      gradFill([gs(0, srgb(BLACK)), gs(150000, srgb(WHITE))], lin(0, 0)),
    ),
    one(
      'h-posneg',
      'A stop at pos=-50000.',
      gradFill([gs(-50000, srgb(BLACK)), gs(100000, srgb(WHITE))], lin(0, 0)),
    ),
    one(
      'h-pospct',
      'The "50%" spelling of ST_Percentage on @pos, which 2.6 found legal on lumMod.',
      gradFill([gs(0, srgb(BLACK)), gs('50%', srgb('808080')), gs(100000, srgb(WHITE))], lin(0, 0)),
    ),
    one(
      'h-onestop',
      'A gradient with one stop. A flat fill, or nothing?',
      gradFill([gs(0, srgb('4472C4'))], lin(0, 0)),
    ),
    one('h-nostops', 'An empty gsLst.', gradFill([], lin(0, 0))),
    one(
      'h-nogslst',
      'No gsLst element at all.',
      '<a:gradFill><a:lin ang="0" scaled="0"/></a:gradFill>',
    ),
    one(
      'h-noshade',
      'Stops but no lin and no path: what is the default shade?',
      gradFill([gs(0, srgb(BLACK)), gs(100000, srgb(WHITE))], ''),
    ),
    one(
      'h-duppos',
      'Two stops at the same position: a hard edge, or does the second one win?',
      gradFill(
        [
          gs(0, srgb(BLACK)),
          gs(50000, srgb(BLACK)),
          gs(50000, srgb(WHITE)),
          gs(100000, srgb(WHITE)),
        ],
        lin(0, 0),
      ),
    ),
    one(
      'h-ftr-inv',
      'A fillToRect whose left inset is past its right inset.',
      bw(path('circle', { l: 80000, t: 50000, r: 80000, b: 50000 })),
    ),
    one(
      'h-patnoprst',
      'a:pattFill with no @prst. The attribute is optional in the schema.',
      pattFill(null, srgb(BLACK), srgb(WHITE)),
    ),
    one(
      'h-patbad',
      'a:pattFill with a @prst that is not in the enumeration.',
      pattFill('notAPattern', srgb(BLACK), srgb(WHITE)),
    ),
    one(
      'h-patnobg',
      'a:pattFill with a foreground and no background.',
      pattFill('lgCheck', srgb(BLACK), null),
    ),
    one(
      'h-patnofg',
      'a:pattFill with a background and no foreground.',
      pattFill('lgCheck', null, srgb(WHITE)),
    ),
  ];
}
