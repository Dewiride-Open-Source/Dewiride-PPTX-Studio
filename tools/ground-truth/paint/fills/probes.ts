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
 */
export type SampleMode = 'strip' | 'grid' | 'tile';

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

  return out;
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
  return [
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
