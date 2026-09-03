/**
 * Experiment C4 - the stroke and effect questions.
 *
 * Sub-phase 2.7 measured how a shape's *interior* is painted. This one measures
 * its edge, and the things drawn around it. Almost none of it is written down:
 * ECMA-376 names the eleven preset dashes and gives no arrays; it names six
 * arrowheads and gives no geometry; it defines `a:ln/@cap` with a default that
 * Office does not use; and `a:outerShdw/@blurRad` is "the blur radius" with no
 * statement of what radius means, which is the one number an SVG filter needs.
 *
 * ## What PowerPoint already told us, in writing
 *
 * Before any of this existed, PowerPoint was asked to *author* strokes and
 * effects through its own object model (`author-lines.ps1`), and reading the
 * files it saved settled several things for free:
 *
 *   - `Line.DashStyle` reaches all eleven `ST_PresetLineDashVal` names. Two of
 *     the twelve object-model values differ **only by `@cap`**: `msoLineSquareDot`
 *     is `val="dot" cap="sq"` and `msoLineRoundDot` is `val="dot" cap="rnd"`.
 *     The cap is not a detail of the dash - in PowerPoint's own UI it *is* the
 *     difference between two entries in the dash gallery.
 *   - Apart from those two, PowerPoint writes **no `@cap` at all**. So whatever
 *     the default cap is, it is what almost every line in every real deck uses,
 *     and getting it wrong is not an edge case.
 *   - A line the user has not styled has **no `a:ln` element whatsoever** - the
 *     stroke comes from `p:style/a:lnRef idx="2"` into the theme. "What is the
 *     default line width" is a theme question before it is a schema question.
 *   - All five `@cmpd` values and both `@algn` values are reachable and written
 *     explicitly. `InsetPen` is the object model's name for `algn="in"`.
 *   - Arrowheads are exactly six types x three widths x three lengths. The
 *     seventh type index is refused.
 *   - The 43 legacy shadow presets write **three different elements**:
 *     `a:prstShdw` (20 of them), `a:outerShdw` (13) and `a:innerShdw` (9). The
 *     plan for this sub-phase does not mention `a:prstShdw` and neither does any
 *     renderer this project has looked at; PowerPoint writes it today.
 *   - `a:reflection` has thirteen attributes and PowerPoint's nine presets name
 *     all of them, which is thirteen defaults nobody has to guess.
 *
 * What is left is everything a pixel can answer, and that is what is below.
 *
 * ## One question per package
 *
 * Same rule as C2 and C3. Fifteen probes here are hostile - a negative width, a
 * dash name outside the enumeration, a `custDash` with no segments - and each is
 * alone in its own `.pptx`, because PowerPoint's only diagnostic is a repair
 * prompt with no part name and a deck that asks two questions and comes back
 * repaired has answered neither.
 */

/* -------------------------------------------------------------------------- */
/* stroke markup                                                              */
/* -------------------------------------------------------------------------- */

/** Twelve thousand seven hundred EMU. Every dimension in this file is in points. */
export const EMU_PER_POINT = 12700;

export interface LineOptions {
  /** Points. Omitted means no `@w`, which is a question in itself. */
  readonly w?: number;
  readonly cap?: 'flat' | 'sq' | 'rnd' | string;
  readonly cmpd?: 'sng' | 'dbl' | 'thickThin' | 'thinThick' | 'tri' | string;
  readonly algn?: 'ctr' | 'in' | string;
  /** Defaults to solid black. Pass `'<a:noFill/>'` for an invisible stroke. */
  readonly fill?: string;
  /** A whole `a:prstDash` or `a:custDash`. */
  readonly dash?: string;
  /** A whole `a:round`, `a:bevel` or `a:miter`. */
  readonly join?: string;
  readonly headEnd?: string;
  readonly tailEnd?: string;
  /** Written verbatim in place of a computed `@w`, for the hostile probes. */
  readonly rawW?: string;
}

/**
 * `CT_LineProperties`. The child order is fill, then dash, then join, then
 * `a:headEnd`, then `a:tailEnd` - an `xsd:sequence`, so it is the only order
 * that opens.
 */
export function ln(options: LineOptions = {}): string {
  const attr = (name: string, value: string | undefined): string =>
    value === undefined ? '' : ` ${name}="${value}"`;
  const w =
    options.rawW !== undefined
      ? ` w="${options.rawW}"`
      : options.w === undefined
        ? ''
        : ` w="${String(Math.round(options.w * EMU_PER_POINT))}"`;
  const body =
    (options.fill ?? '<a:solidFill><a:srgbClr val="000000"/></a:solidFill>') +
    (options.dash ?? '') +
    (options.join ?? '') +
    (options.headEnd ?? '') +
    (options.tailEnd ?? '');
  return (
    `<a:ln${w}${attr('cap', options.cap)}${attr('cmpd', options.cmpd)}${attr('algn', options.algn)}>` +
    body +
    '</a:ln>'
  );
}

export function prstDash(val: string): string {
  return `<a:prstDash val="${val}"/>`;
}

/**
 * `a:custDash`. Each `a:ds` is a percentage **of the line width**, so a segment
 * of `400000` is four times the stroke thickness. That encoding is why the
 * preset arrays are quoted in width-multiples everywhere: a `custDash` is the
 * only way to write one down, and it can only be written that way.
 */
export function custDash(pairs: readonly (readonly [number, number])[]): string {
  if (pairs.length === 0) return '<a:custDash/>';
  const ds = pairs
    .map(([d, sp]) => `<a:ds d="${String(Math.round(d * 100000))}" sp="${String(Math.round(sp * 100000))}"/>`)
    .join('');
  return `<a:custDash>${ds}</a:custDash>`;
}

/** A whole `a:custDash` with the segment strings written verbatim. */
export function rawCustDash(inner: string): string {
  return `<a:custDash>${inner}</a:custDash>`;
}

export function lineEnd(
  which: 'headEnd' | 'tailEnd',
  type: string,
  w?: string,
  len?: string,
): string {
  const attrs =
    ` type="${type}"` +
    (w === undefined ? '' : ` w="${w}"`) +
    (len === undefined ? '' : ` len="${len}"`);
  return `<a:${which}${attrs}/>`;
}

/* -------------------------------------------------------------------------- */
/* effect markup                                                              */
/* -------------------------------------------------------------------------- */

/** `<a:srgbClr val="..."><a:alpha val="..."/></a:srgbClr>`, alpha in percent. */
export function shadowColor(hex = '000000', alphaPct = 100): string {
  const alpha = alphaPct === 100 ? '' : `<a:alpha val="${String(Math.round(alphaPct * 1000))}"/>`;
  return alpha === ''
    ? `<a:srgbClr val="${hex}"/>`
    : `<a:srgbClr val="${hex}">${alpha}</a:srgbClr>`;
}

export interface ShadowOptions {
  /** Points. */
  readonly blurRad?: number;
  /** Points. */
  readonly dist?: number;
  /** Degrees clockwise, converted to sixtieths of a degree. */
  readonly dirDeg?: number;
  /** Percent. */
  readonly sx?: number;
  readonly sy?: number;
  /** Degrees. */
  readonly kxDeg?: number;
  readonly kyDeg?: number;
  readonly algn?: string;
  readonly rotWithShape?: 0 | 1;
  readonly color?: string;
  /** Written verbatim in place of every computed attribute, for hostile probes. */
  readonly rawAttrs?: string;
}

function shadowAttrs(o: ShadowOptions): string {
  if (o.rawAttrs !== undefined) return o.rawAttrs;
  const pct = (v: number): string => String(Math.round(v * 1000));
  const emu = (v: number): string => String(Math.round(v * EMU_PER_POINT));
  const ang = (v: number): string => String(Math.round(v * 60000));
  return (
    (o.blurRad === undefined ? '' : ` blurRad="${emu(o.blurRad)}"`) +
    (o.dist === undefined ? '' : ` dist="${emu(o.dist)}"`) +
    (o.dirDeg === undefined ? '' : ` dir="${ang(o.dirDeg)}"`) +
    (o.sx === undefined ? '' : ` sx="${pct(o.sx)}"`) +
    (o.sy === undefined ? '' : ` sy="${pct(o.sy)}"`) +
    (o.kxDeg === undefined ? '' : ` kx="${ang(o.kxDeg)}"`) +
    (o.kyDeg === undefined ? '' : ` ky="${ang(o.kyDeg)}"`) +
    (o.algn === undefined ? '' : ` algn="${o.algn}"`) +
    (o.rotWithShape === undefined ? '' : ` rotWithShape="${String(o.rotWithShape)}"`)
  );
}

export function outerShdw(o: ShadowOptions = {}): string {
  return `<a:outerShdw${shadowAttrs(o)}>${o.color ?? shadowColor()}</a:outerShdw>`;
}

export function innerShdw(o: ShadowOptions = {}): string {
  return `<a:innerShdw${shadowAttrs(o)}>${o.color ?? shadowColor()}</a:innerShdw>`;
}

export function prstShdw(prst: string, o: ShadowOptions = {}): string {
  return `<a:prstShdw prst="${prst}"${shadowAttrs(o)}>${o.color ?? shadowColor()}</a:prstShdw>`;
}

/** `a:glow`. `rad` in points. */
export function glow(rad: number | string, color = shadowColor('FF0000')): string {
  const value = typeof rad === 'string' ? rad : String(Math.round(rad * EMU_PER_POINT));
  return `<a:glow rad="${value}">${color}</a:glow>`;
}

/** `a:softEdge`. `rad` in points. */
export function softEdge(rad: number | string | null): string {
  if (rad === null) return '<a:softEdge/>';
  const value = typeof rad === 'string' ? rad : String(Math.round(rad * EMU_PER_POINT));
  return `<a:softEdge rad="${value}"/>`;
}

/** `a:blur`. `rad` in points. */
export function blur(rad: number | string, grow: 0 | 1 = 1): string {
  const value = typeof rad === 'string' ? rad : String(Math.round(rad * EMU_PER_POINT));
  return `<a:blur rad="${value}" grow="${String(grow)}"/>`;
}

export function reflection(attrs: string): string {
  return `<a:reflection ${attrs}/>`;
}

/**
 * `a:effectLst`. The child order is blur, fillOverlay, glow, innerShdw,
 * outerShdw, prstShdw, reflection, softEdge - and it is a sequence, so a list
 * assembled in any other order does not open.
 */
export function effectLst(...children: readonly string[]): string {
  return children.length === 0 ? '<a:effectLst/>' : `<a:effectLst>${children.join('')}</a:effectLst>`;
}

/* -------------------------------------------------------------------------- */
/* probes                                                                     */
/* -------------------------------------------------------------------------- */

/**
 * How the analysis reads this probe out of the exported bitmap.
 *
 * Every region is stated in points, absolutely, on the slide - not relative to
 * the shape. A stroke is drawn *around* its geometry, so half of what is being
 * measured is outside the shape's own rectangle, and a region expressed
 * relative to the shape could not name it.
 *
 * - `row`   - coverage along one horizontal line. Dash runs, cap extents.
 * - `col`   - coverage along one vertical line. Stroke thickness, compound
 *             profiles, shadow ramps across a horizontal edge.
 * - `cols`  - per-column ink extent and total over a rectangle. Arrowhead
 *             outlines, and anything whose shape is the answer.
 */
export type ReadSpec =
  | { readonly mode: 'row'; readonly y: number; readonly x0: number; readonly x1: number }
  | { readonly mode: 'col'; readonly x: number; readonly y0: number; readonly y1: number }
  | {
      readonly mode: 'cols';
      readonly x0: number;
      readonly y0: number;
      readonly x1: number;
      readonly y1: number;
    };

export interface Probe {
  /** Unique, and used verbatim as `p:cNvPr/@name`. */
  readonly id: string;
  /** Which `.pptx` this goes in. One question per package for anything hostile. */
  readonly deck: string;
  readonly group: string;
  /** What this probe is for, in one sentence. Goes into the fixture. */
  readonly question: string;
  /** Points. */
  readonly x: number;
  readonly y: number;
  readonly cx: number;
  readonly cy: number;
  /** A whole `a:prstGeom` or `a:custGeom`. Defaults to `prst="line"`. */
  readonly geom?: string;
  /** Defaults to `a:noFill`. */
  readonly fill?: string;
  /** A whole `a:ln`. */
  readonly line?: string;
  /** A whole `a:effectLst`. */
  readonly effect?: string;
  readonly read: ReadSpec;
  /** Sixtieths of a degree on `a:xfrm/@rot`. */
  readonly rot?: number;
}

export const LINE_GEOM = '<a:prstGeom prst="line"><a:avLst/></a:prstGeom>';
export const RECT_GEOM = '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>';
const BLACK_FILL = '<a:solidFill><a:srgbClr val="000000"/></a:solidFill>';

/**
 * A path that turns a sharp corner, for the join probes.
 *
 * `a:custGeom` rather than a preset because no preset has a corner whose angle
 * we chose. The path is a V: down-right to the apex, then up-right, so the
 * interior angle at the apex is set by `rise` against `run` and a miter can be
 * made to exceed any limit we like.
 */
export function vGeom(run: number, rise: number): string {
  const w = run * 2;
  const h = rise;
  return (
    `<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="0" t="0" r="r" b="b"/>` +
    `<a:pathLst><a:path w="${String(w)}" h="${String(h)}">` +
    `<a:moveTo><a:pt x="0" y="0"/></a:moveTo>` +
    `<a:lnTo><a:pt x="${String(run)}" y="${String(h)}"/></a:lnTo>` +
    `<a:lnTo><a:pt x="${String(w)}" y="0"/></a:lnTo>` +
    `</a:path></a:pathLst></a:custGeom>`
  );
}

/** The eleven `ST_PresetLineDashVal` names, as PowerPoint spelled them itself. */
export const PRESET_DASHES: readonly string[] = [
  'solid',
  'dot',
  'sysDot',
  'dash',
  'sysDash',
  'lgDash',
  'dashDot',
  'sysDashDot',
  'lgDashDot',
  'lgDashDotDot',
  'sysDashDotDot',
];

/**
 * What everyone says the arrays are.
 *
 * These are the values quoted by essentially every open implementation and by
 * the non-normative notes people cite for them. They are written down here as a
 * *prediction*, emitted as `a:custDash` probes alongside the presets, and the
 * bitmap decides. If a prediction is right its row is identical to the preset's
 * row; if it is wrong we find out in the same run rather than shipping folklore.
 */
export const FOLKLORE_DASHES: Readonly<Record<string, readonly (readonly [number, number])[]>> = {
  dot: [[1, 3]],
  sysDot: [[1, 1]],
  dash: [[4, 3]],
  sysDash: [[3, 1]],
  lgDash: [[8, 3]],
  dashDot: [
    [4, 3],
    [1, 3],
  ],
  sysDashDot: [
    [3, 1],
    [1, 1],
  ],
  lgDashDot: [
    [8, 3],
    [1, 3],
  ],
  lgDashDotDot: [
    [8, 3],
    [1, 3],
    [1, 3],
  ],
  sysDashDotDot: [
    [3, 1],
    [1, 1],
    [1, 1],
  ],
};

export const ARROW_TYPES: readonly string[] = [
  'none',
  'triangle',
  'stealth',
  'diamond',
  'oval',
  'arrow',
];
export const ARROW_SIZES: readonly string[] = ['sm', 'med', 'lg'];

/** The left margin every full-width probe starts at, and the right edge it stops at. */
const X0 = 20;
const X1 = 940;

/** A horizontal line probe, plus the row that reads it. */
function hline(
  base: Omit<Probe, 'x' | 'y' | 'cx' | 'cy' | 'read' | 'geom'> & { geom?: string },
  y: number,
  x0 = X0,
  x1 = X1,
  pad = 0,
): Probe {
  return {
    ...base,
    geom: base.geom ?? LINE_GEOM,
    x: x0,
    y,
    cx: x1 - x0,
    cy: 0,
    read: { mode: 'row', y, x0: x0 - pad, x1: x1 + pad },
  };
}

export function lineProbes(): Probe[] {
  const probes: Probe[] = [];

  /* ---------------------------------------------------------------- dashes */
  // Twelve points of stroke is twenty-four pixels at the 1920-wide export and
  // forty-eight at 3840. A `dash` segment of four widths is then 192 px, which
  // is measurable to better than one percent - the whole reason the dash deck
  // is exported at 3840 and the others are not.
  const DASH_W = 12;
  PRESET_DASHES.forEach((name, i) => {
    probes.push(
      hline(
        {
          id: `dash-${name}`,
          deck: 'dash',
          group: 'dash',
          question: `what on/off run lengths does prstDash val="${name}" paint?`,
          line: ln({ w: DASH_W, dash: prstDash(name) }),
        },
        24 + i * 22,
      ),
    );
  });
  // The same eleven, written as the arrays everyone quotes. A prediction that
  // is right produces an identical row.
  Object.entries(FOLKLORE_DASHES).forEach(([name, pairs], i) => {
    probes.push(
      hline(
        {
          id: `folk-${name}`,
          deck: 'dash',
          group: 'dash',
          question: `does custDash ${JSON.stringify(pairs)} paint the same as prstDash val="${name}"?`,
          line: ln({ w: DASH_W, dash: custDash(pairs) }),
        },
        290 + i * 22,
      ),
    );
  });

  /* -------------------------------------------------- dash x width, x cap */
  // If the arrays are in width-multiples the run lengths scale with `@w`; if
  // they are absolute lengths they do not. Three widths settle it.
  [3, 6, 24].forEach((w, i) => {
    probes.push(
      hline(
        {
          id: `dashw-${String(w)}`,
          deck: 'dashw',
          group: 'dash',
          question: `is the dash array proportional to @w? (w=${String(w)}pt)`,
          line: ln({ w, dash: prstDash('dash') }),
        },
        60 + i * 120,
      ),
    );
  });
  // PowerPoint distinguishes square-dot from round-dot by cap alone, so the cap
  // must change what a dash looks like. The question is whether it lengthens
  // each dash (and shortens each gap) or whether the array is compensated.
  (['flat', 'sq', 'rnd'] as const).forEach((cap, i) => {
    (['dot', 'dash'] as const).forEach((name, j) => {
      probes.push(
        hline(
          {
            id: `dashcap-${name}-${cap}`,
            deck: 'dashcap',
            group: 'cap',
            question: `does cap="${cap}" lengthen each ${name} segment, or is the array compensated?`,
            line: ln({ w: DASH_W, cap, dash: prstDash(name) }),
          },
          40 + (i * 2 + j) * 80,
        ),
      );
    });
  });

  /* ------------------------------------------------------------------ caps */
  // A short line whose ink extent past its own endpoint is the whole answer.
  // 200pt long at 24pt wide: a square cap adds 12pt at each end, which is 48px
  // at the 1920-wide export and impossible to miss.
  const CAP_W = 24;
  ([undefined, 'flat', 'sq', 'rnd'] as const).forEach((cap, i) => {
    const y = 80 + i * 110;
    probes.push({
      id: `cap-${cap ?? 'absent'}`,
      deck: 'cap',
      group: 'cap',
      question:
        cap === undefined
          ? 'with no @cap at all, does the ink stop at the endpoint (flat) or overshoot by half a width (square)?'
          : `how far past the endpoint does cap="${cap}" paint?`,
      x: 300,
      y,
      cx: 200,
      cy: 0,
      geom: LINE_GEOM,
      line: ln({ w: CAP_W, cap }),
      read: { mode: 'row', y, x0: 260, x1: 540 },
    });
  });
  // The same question again with a round cap, read across the line rather than
  // along it, so the cap's *profile* is on record and not just its extent.
  probes.push({
    id: 'cap-rnd-profile',
    deck: 'cap',
    group: 'cap',
    question: 'is a round cap a half-disc? sampled as a column just past the endpoint',
    x: 300,
    y: 460,
    cx: 200,
    cy: 0,
    geom: LINE_GEOM,
    line: ln({ w: CAP_W, cap: 'rnd' }),
    read: { mode: 'col', x: 506, y0: 440, y1: 480 },
  });

  /* ---------------------------------------------------------------- widths */
  // What does a line with no @w measure? What does w="0" measure? Both are
  // ordinary in real files and neither has a stated answer that Office follows.
  const widthCases: { id: string; line: string; question: string }[] = [
    { id: 'w-absent', line: ln({}), question: 'a:ln with no @w - how thick is it?' },
    { id: 'w-zero', line: ln({ rawW: '0' }), question: 'w="0" - hairline, or nothing at all?' },
    { id: 'w-1pt', line: ln({ w: 1 }), question: 'w=1pt, as a scale check' },
    { id: 'w-quarter', line: ln({ rawW: '3175' }), question: 'w=0.25pt - is a sub-pixel width painted?' },
    { id: 'w-9525', line: ln({ rawW: '9525' }), question: 'w=0.75pt, the width Office calls "hairline" in its UI' },
  ];
  widthCases.forEach((c, i) => {
    const y = 60 + i * 100;
    probes.push({
      id: c.id,
      deck: 'width',
      group: 'width',
      question: c.question,
      x: 200,
      y,
      cx: 560,
      cy: 0,
      geom: LINE_GEOM,
      line: c.line,
      read: { mode: 'col', x: 480, y0: y - 40, y1: y + 40 },
    });
  });
  // The maximum is 1584 points, which is three times the height of the slide,
  // so it gets a package to itself or it paints over every other answer in one.
  probes.push({
    id: 'w-max',
    deck: 'wmax',
    group: 'width',
    question: 'what does a stroke at the ST_LineWidth maximum of 1584pt actually cover?',
    x: 200,
    y: 270,
    cx: 560,
    cy: 0,
    geom: LINE_GEOM,
    line: ln({ rawW: '20116800' }),
    read: { mode: 'col', x: 480, y0: 120, y1: 420 },
  });

  /* ------------------------------------------------------------- alignment */
  // A thick outline on a rectangle. `algn="ctr"` straddles the geometry;
  // `algn="in"` is supposed to sit wholly inside it. The measurement is where
  // the ink starts and stops relative to an edge we placed to the point.
  const ALGN_W = 20;
  ([undefined, 'ctr', 'in'] as const).forEach((algn, i) => {
    const x = 120 + i * 260;
    probes.push({
      id: `algn-${algn ?? 'absent'}`,
      deck: 'algn',
      group: 'algn',
      question:
        algn === undefined
          ? 'with no @algn, is the stroke centred on the geometry?'
          : `where does algn="${algn}" put a ${String(ALGN_W)}pt stroke relative to the shape rectangle?`,
      x,
      y: 100,
      cx: 180,
      cy: 180,
      geom: RECT_GEOM,
      fill: '<a:noFill/>',
      line: ln({ w: ALGN_W, algn }),
      read: { mode: 'row', y: 190, x0: x - 40, x1: x + 40 },
    });
  });
  // The same three again on a *filled* shape, because an inset stroke on a
  // filled shape is where the difference is visible to a user and where a
  // clip-and-double-stroke emulation would show its seam.
  ([undefined, 'ctr', 'in'] as const).forEach((algn, i) => {
    const x = 120 + i * 260;
    probes.push({
      id: `algnfill-${algn ?? 'absent'}`,
      deck: 'algn',
      group: 'algn',
      question: `does algn="${algn ?? 'absent'}" change where the fill ends on a filled shape?`,
      x,
      y: 340,
      cx: 180,
      cy: 140,
      geom: RECT_GEOM,
      // Dark grey, not mid grey: coverage of 808080 over white is 0.498, which
      // sits just under the half-coverage threshold every crossing in this
      // experiment is found by, and a fill that never crosses is a fill that
      // cannot be measured.
      fill: '<a:solidFill><a:srgbClr val="404040"/></a:solidFill>',
      line: ln({ w: ALGN_W, algn }),
      read: { mode: 'row', y: 410, x0: x - 40, x1: x + 40 },
    });
  });

  /* -------------------------------------------------------------- compound */
  // Five compound strokes at 36pt, which is 72 pixels of profile to divide up.
  // The question is what fractions of @w each sub-stroke and each gap takes.
  const CMPD_W = 36;
  // Three across and two down, 320pt apart. The first attempt used five across
  // at a 180pt pitch and measured the *neighbour*: a 36pt compound stroke
  // reaches 22pt outside its own rectangle at each side, so a 40pt gap between
  // two shapes puts one shape's right rail inside the next shape's read window,
  // and every profile carried a stray pair of crossings that looked like an
  // extra rail.
  const cmpdCases: { id: string; line: string; question: string }[] = [
    ...(['sng', 'dbl', 'thickThin', 'thinThick', 'tri'] as const).map((cmpd) => ({
      id: `cmpd-${cmpd}`,
      line: ln({ w: CMPD_W, cmpd }),
      question: `how does cmpd="${cmpd}" divide a ${String(CMPD_W)}pt stroke?`,
    })),
    {
      id: 'cmpd-dbl-in',
      line: ln({ w: CMPD_W, cmpd: 'dbl', algn: 'in' }),
      question: 'where does a double stroke sit when algn="in"?',
    },
  ];
  cmpdCases.forEach((c, i) => {
    const x = 60 + (i % 3) * 320;
    const y = 100 + Math.floor(i / 3) * 240;
    probes.push({
      id: c.id,
      deck: 'cmpd',
      group: 'cmpd',
      question: c.question,
      x,
      y,
      cx: 140,
      cy: 140,
      geom: RECT_GEOM,
      fill: '<a:noFill/>',
      line: c.line,
      // Wide enough on the inside to hold an inset stroke's far rail, which
      // sits a whole stroke width in from the edge.
      read: { mode: 'row', y: y + 70, x0: x - 30, x1: x + 50 },
    });
  });
  // A compound dash, which is the combination every renderer gets wrong: is the
  // dash applied to the compound stroke as a whole, or to each rail?
  probes.push({
    id: 'cmpd-dbl-dash',
    deck: 'dashw',
    group: 'cmpd',
    question: 'is a dash applied to the compound stroke as a whole, or to each rail?',
    x: 200,
    y: 440,
    cx: 560,
    cy: 0,
    geom: LINE_GEOM,
    line: ln({ w: CMPD_W, cmpd: 'dbl', dash: prstDash('dash') }),
    read: { mode: 'row', y: 440, x0: 200, x1: 760 },
  });

  /* ------------------------------------------------------------------ join */
  // A V with a 30-degree included angle, so the miter is long enough that the
  // default limit matters. `run` and `rise` are path units; the shape's own
  // extent sets the scale.
  const JOIN_W = 20;
  const joinCases: { id: string; join: string | undefined; question: string }[] = [
    { id: 'join-absent', join: undefined, question: 'with no join element, which join does Office use?' },
    { id: 'join-round', join: '<a:round/>', question: 'a:round on a sharp corner' },
    { id: 'join-bevel', join: '<a:bevel/>', question: 'a:bevel on a sharp corner' },
    { id: 'join-miter', join: '<a:miter/>', question: 'a:miter with no @lim - what is the default limit?' },
    {
      id: 'join-miter-lim',
      join: '<a:miter lim="100000"/>',
      question: 'a:miter with lim=100% - does a limit of one width fall back to a bevel?',
    },
    {
      id: 'join-miter-big',
      join: '<a:miter lim="800000"/>',
      question: 'a:miter with lim=800% - how far does the spike reach?',
    },
  ];
  joinCases.forEach((c, i) => {
    const x = 40 + (i % 3) * 300;
    const y = 60 + Math.floor(i / 3) * 240;
    probes.push({
      id: c.id,
      deck: 'join',
      group: 'join',
      question: c.question,
      x,
      y,
      cx: 200,
      cy: 120,
      geom: vGeom(1000, 300),
      fill: '<a:noFill/>',
      line: ln({ w: JOIN_W, join: c.join }),
      // A column through the apex: the spike, if there is one, points down.
      read: { mode: 'col', x: x + 100, y0: y + 100, y1: y + 220 },
    });
  });

  /* -------------------------------------------- how far will a miter reach? */
  // The V above has an 80-degree apex, whose miter is only 1.56 stroke widths
  // long - so it proves a default limit of *at least* that and no more. Five
  // sharper corners bracket it. The miter ratio at an included angle t is
  // 1/sin(t/2), and for a V of half-run r and rise h that is
  // sqrt(r^2 + h^2) / r, so the shape's own aspect chooses the ratio: these
  // five are 2, 4, 8, 12 and 20, and the ink stops growing at the limit.
  const miterShapes: { halfRun: number; rise: number; ratio: number }[] = [
    { halfRun: 60, rise: 104, ratio: 2 },
    { halfRun: 40, rise: 155, ratio: 4 },
    { halfRun: 30, rise: 147, ratio: 5 },
    { halfRun: 25, rise: 148, ratio: 6 },
    { halfRun: 22, rise: 152, ratio: 7 },
    // The bracket. Everything up to 7 came back with a full miter and 8.01 came
    // back bevelled, so the default limit is somewhere between - and these two
    // straddle 8 by a tenth either way.
    { halfRun: 20, rise: 157, ratio: 7.91 },
    { halfRun: 20, rise: 161, ratio: 8.11 },
    { halfRun: 20, rise: 159, ratio: 8 },
    { halfRun: 15, rise: 179, ratio: 12 },
    { halfRun: 10, rise: 200, ratio: 20 },
  ];
  // Two packages, because ten of these do not fit across one slide: the read
  // window is a column through the apex and the shapes have to stand clear of
  // one another by more than a stroke width.
  const miterX = [20, 20];
  miterShapes.forEach((s, i) => {
    const which = i < 5 ? 0 : 1;
    const x = miterX[which]!;
    miterX[which] = x + s.halfRun * 2 + 60;
    const y = 20;
    probes.push({
      id: `miter-r${String(s.ratio)}`,
      deck: which === 0 ? 'miter' : 'miter2',
      group: 'join',
      question: `with a miter ratio of ${String(s.ratio)} and no @lim, does the spike reach its full length?`,
      x,
      y,
      cx: s.halfRun * 2,
      cy: s.rise,
      geom: vGeom(1000, Math.round((1000 * s.rise) / s.halfRun)),
      fill: '<a:noFill/>',
      line: ln({ w: JOIN_W, join: '<a:miter/>' }),
      read: { mode: 'col', x: x + s.halfRun, y0: y + s.rise, y1: y + s.rise + 240 },
    });
  });

  /* ---------------------------------------- what unit is a:miter/@lim in? ---*/
  // A limit that clips at eight would be an odd number to choose in units of
  // the stroke width, and a very ordinary one in units of *half* the width -
  // which is how the miter's own extension past the corner is naturally
  // measured, and how several drawing APIs express it. These four put an
  // explicit @lim on a corner whose miter ratio is 5: if @lim is a ratio in
  // whole widths, 400% clips and 600% does not, and if it is in half widths the
  // change happens between 800% and 1200% instead.
  [200, 400, 600, 1000].forEach((pct, i) => {
    const x = 30 + i * 240;
    probes.push({
      id: `mlim-${String(pct)}`,
      deck: 'mlim',
      group: 'join',
      question: `at a corner whose miter ratio is 5, does lim=${String(pct)}% clip?`,
      x,
      y: 20,
      cx: 60,
      cy: 147,
      geom: vGeom(1000, Math.round((1000 * 147) / 30)),
      fill: '<a:noFill/>',
      line: ln({ w: JOIN_W, join: `<a:miter lim="${String(pct * 1000)}"/>` }),
      read: { mode: 'col', x: x + 30, y0: 167, y1: 400 },
    });
  });

  /* --------------------------------------------------- the preset shadows ---*/
  // `a:prstShdw` is not in this sub-phase's plan and is not in any renderer this
  // project has read. PowerPoint writes it for twenty of its forty-three shadow
  // presets, which is to say it is in real files today. Nothing states what any
  // of the twenty look like, so this records what they paint: a red shadow of a
  // black block, thrown far enough that its own box can be measured.
  for (let n = 1; n <= 20; n++) {
    const col = (n - 1) % 5;
    const row = Math.floor((n - 1) / 5);
    const x = 20 + col * 190;
    const y = 20 + row * 130;
    probes.push({
      id: `prst-shdw${String(n)}`,
      deck: `prstshdw${String(row)}`,
      group: 'prstShdw',
      question: `what does a:prstShdw prst="shdw${String(n)}" paint?`,
      x,
      y,
      cx: 70,
      cy: 60,
      geom: RECT_GEOM,
      fill: BLACK_FILL,
      line: '<a:ln><a:noFill/></a:ln>',
      effect: effectLst(
        prstShdw(`shdw${String(n)}`, { dist: 60, dirDeg: 0, color: shadowColor('FF0000') }),
      ),
      read: { mode: 'cols', x0: x - 40, y0: y - 40, x1: x + 180, y1: y + 120 },
    });
  }

  /* ------------------------------------------- does the shadow rotate too? ---*/
  ([0, 1] as const).forEach((rot, i) => {
    const x = 120 + i * 400;
    probes.push({
      id: `shdw-rotwith${String(rot)}`,
      deck: 'shrot',
      group: 'shadow',
      question: `on a shape rotated 45 degrees, does rotWithShape="${String(rot)}" turn the shadow with it?`,
      x,
      y: 160,
      cx: 120,
      cy: 60,
      rot: 45 * 60000,
      geom: RECT_GEOM,
      fill: BLACK_FILL,
      line: '<a:ln><a:noFill/></a:ln>',
      effect: effectLst(
        outerShdw({
          blurRad: 0,
          dist: 80,
          dirDeg: 0,
          rotWithShape: rot,
          color: shadowColor('FF0000'),
        }),
      ),
      read: { mode: 'cols', x0: x - 120, y0: 20, x1: x + 250, y1: 380 },
    });
  });

  /* ------------------------------------------------------------ arrowheads */
  // Six types x three widths x three lengths, each on its own short line, read
  // as a per-column ink profile around the tip. That profile *is* the marker
  // outline: for a triangle the half-height falls linearly, for a stealth it
  // has a notch, for an oval it is an ellipse. Recovering it by measurement is
  // also how this sub-phase avoids transcribing LibreOffice's vertex tables,
  // which are MPL-2.0 and would put file-level copyleft in an Apache-2.0 tree.
  const AH_W = 8;
  const AH_LEN = 60;
  ARROW_TYPES.forEach((type, ti) => {
    ARROW_SIZES.forEach((len, li) => {
      ARROW_SIZES.forEach((wid, wi) => {
        const index = li * 3 + wi;
        const col = index % 3;
        const row = Math.floor(index / 3);
        // Three columns of 300pt and three rows of 90pt, one type per deck, so
        // a head can be 40pt long and 26pt to either side without any two
        // probes' ink meeting.
        const x = 30 + col * 300;
        const y = 90 + row * 130;
        probes.push({
          id: `ah-${type}-l${len}-w${wid}`,
          deck: `arrow${String(ti)}`,
          group: 'arrowhead',
          question: `what outline does tailEnd type="${type}" len="${len}" w="${wid}" paint at ${String(AH_W)}pt?`,
          x,
          y,
          cx: AH_LEN,
          cy: 0,
          geom: LINE_GEOM,
          line: ln({ w: AH_W, tailEnd: lineEnd('tailEnd', type, wid, len) }),
          // A window centred on the endpoint, wide enough to hold the whole
          // head and 40pt of shaft behind it.
          read: {
            mode: 'cols',
            x0: x + AH_LEN - 60,
            x1: x + AH_LEN + 40,
            y0: y - 26,
            y1: y + 26,
          },
        });
      });
    });
  });
  // Does an arrowhead shorten the shaft? Read the whole line, not just the tip.
  probes.push({
    id: 'ah-shaft',
    deck: 'arrow0',
    group: 'arrowhead',
    question: 'does the shaft still reach the endpoint under a triangle head, or is it pulled back?',
    x: 300,
    y: 500,
    cx: 400,
    cy: 0,
    geom: LINE_GEOM,
    line: ln({ w: 24, tailEnd: lineEnd('tailEnd', 'triangle', 'med', 'med') }),
    read: { mode: 'row', y: 500, x0: 640, x1: 760 },
  });
  // Does the head scale with the stroke? Every length in section H is quoted in
  // stroke widths, which is only the right unit if the same head at a different
  // `@w` is the same number of widths. Two more widths settle it.
  [4, 16].forEach((w, i) => {
    const y = 250 + i * 70;
    probes.push({
      id: `ah-scale${String(w)}`,
      deck: 'ahscale',
      group: 'arrowhead',
      question: `a triangle head at w=${String(w)}pt - is it still three stroke widths long?`,
      x: 640,
      y,
      cx: 120,
      cy: 0,
      geom: LINE_GEOM,
      line: ln({ w, tailEnd: lineEnd('tailEnd', 'triangle', 'med', 'med') }),
      read: { mode: 'cols', x0: 700, y0: y - 26, x1: 800, y1: y + 26 },
    });
  });
  // Both ends at once, to confirm headEnd is the start and tailEnd is the end.
  probes.push({
    id: 'ah-both',
    deck: 'arrow0',
    group: 'arrowhead',
    question: 'is headEnd at the a:off end of the line and tailEnd at the far end?',
    x: 300,
    y: 430,
    cx: 400,
    cy: 0,
    geom: LINE_GEOM,
    line: ln({
      w: 16,
      headEnd: lineEnd('headEnd', 'oval', 'lg', 'lg'),
      tailEnd: lineEnd('tailEnd', 'triangle', 'sm', 'sm'),
    }),
    read: { mode: 'row', y: 430, x0: 260, x1: 760 },
  });

  /* --------------------------------------------------------------- shadows */
  // Offset first, with no blur, so the vector can be read exactly. Eight
  // directions settle both the zero direction and the sense of rotation.
  [0, 45, 90, 135, 180, 225, 270, 315].forEach((deg, i) => {
    const col = i % 4;
    const row = Math.floor(i / 4);
    const x = 60 + col * 230;
    const y = 60 + row * 240;
    probes.push({
      id: `shdw-dir${String(deg)}`,
      deck: 'shdir',
      group: 'shadow',
      question: `where does an outerShdw with dir=${String(deg)} degrees and dist=30pt land?`,
      x,
      y,
      cx: 100,
      cy: 100,
      geom: RECT_GEOM,
      fill: BLACK_FILL,
      line: '<a:ln><a:noFill/></a:ln>',
      effect: effectLst(outerShdw({ blurRad: 0, dist: 30, dirDeg: deg, color: shadowColor('FF0000') })),
      read: { mode: 'cols', x0: x - 50, y0: y - 50, x1: x + 150, y1: y + 150 },
    });
  });
  // Then blur. The shadow has to be somewhere the shape is not, or there is
  // nothing to read: with dist=0 it sits exactly behind an opaque shape and is
  // invisible. So each row is a filled 120pt block whose shadow is thrown 200pt
  // to the right, and the measurement is the profile across the shadow's own
  // far edge - a step function convolved with whatever Office's blur is.
  [0, 4, 8, 16, 24].forEach((rad, i) => {
    const y = 30 + i * 100;
    probes.push({
      id: `shdw-blur${String(rad)}`,
      deck: 'shblur',
      group: 'shadow',
      question: `what edge profile does blurRad=${String(rad)}pt give a shadow thrown clear of its shape?`,
      x: 60,
      y,
      cx: 120,
      cy: 50,
      geom: RECT_GEOM,
      fill: BLACK_FILL,
      line: '<a:ln><a:noFill/></a:ln>',
      effect: effectLst(
        outerShdw({ blurRad: rad, dist: 200, dirDeg: 0, color: shadowColor('000000') }),
      ),
      // The shadow spans x 260..380. This window straddles its right edge.
      read: { mode: 'row', y: y + 25, x0: 300, x1: 460 },
    });
  });
  // The same radii read across the shadow's *horizontal* edge. A separable
  // Gaussian and a radial one differ here and nowhere else, and the shadow is
  // still thrown sideways so this does not depend on the direction convention
  // the block above is still in the middle of establishing.
  [0, 8, 24].forEach((rad, i) => {
    const y = 30 + i * 170;
    probes.push({
      id: `shdw-vblur${String(rad)}`,
      deck: 'shblur',
      group: 'shadow',
      question: `at blurRad=${String(rad)}pt, is the horizontal edge blurred by the same amount as the vertical one?`,
      x: 560,
      y,
      cx: 140,
      cy: 60,
      geom: RECT_GEOM,
      fill: BLACK_FILL,
      line: '<a:ln><a:noFill/></a:ln>',
      effect: effectLst(
        outerShdw({ blurRad: rad, dist: 200, dirDeg: 0, color: shadowColor('000000') }),
      ),
      // The shadow spans x 760..900 and is 60pt tall. This column straddles its
      // top edge and stops short of its bottom one - a window that held both
      // would be a slab, not an edge, and no single-edge fit converges on one.
      read: { mode: 'col', x: 820, y0: y - 40, y1: y + 30 },
    });
  });
  // Scale, skew and alignment - the attributes that make a shadow an affine
  // transform of the shape rather than a translation of it.
  //
  // The first attempt at this block used dist=0 and a shadow *smaller* than the
  // shape, and measured nothing at all: a shrunken shadow with no offset is
  // entirely behind an opaque shape. Every probe here throws the shadow 200pt
  // clear first, so its box can be read whole and compared against where an
  // untransformed shadow would have landed.
  // Four to a package, two columns by two rows, and the first slot in every
  // package is the untransformed control - every number here is quoted against
  // it, and a control in a different file is not one.
  //
  // The vertical windows are 255pt for a shape 55pt tall. That is not generous,
  // it is necessary: a doubled shadow is 110pt tall and may grow in either
  // direction, and the first attempt packed four rows into the slide with 145pt
  // windows, so the row below's shadow grew up into the row above's window and
  // the *control* came back 115pt tall against a 55pt shape.
  const affine: { deck: string; id: string; o: ShadowOptions; question: string }[] = [
    {
      deck: 'shaffine',
      id: 'shdw-plain',
      o: {},
      question: 'the control: an untransformed shadow thrown 200pt to the right',
    },
    {
      deck: 'shaffine',
      id: 'shdw-sx',
      o: { sx: 200 },
      question: 'sx=200% - which edge of the shadow box stays put?',
    },
    { deck: 'shaffine', id: 'shdw-sy', o: { sy: 200 }, question: 'sy=200% - which edge stays put?' },
    {
      deck: 'shaffine',
      id: 'shdw-syneg',
      o: { sy: -100 },
      question: 'sy=-100% - a mirror, and about which line?',
    },
    { deck: 'shaffine2', id: 'shdw-plain2', o: {}, question: 'the control, repeated' },
    {
      deck: 'shaffine2',
      id: 'shdw-kx',
      o: { kxDeg: 20 },
      question: 'kx=20 degrees - which way does it lean?',
    },
    {
      deck: 'shaffine2',
      id: 'shdw-ky',
      o: { kyDeg: 20 },
      question: 'ky=20 degrees - and which way does that lean?',
    },
    {
      deck: 'shaffine2',
      id: 'shdw-algn-bl',
      o: { sx: 200, sy: 200, algn: 'bl' },
      question: 'algn="bl" with the shadow doubled - which corner is the anchor?',
    },
    { deck: 'shaffine3', id: 'shdw-plain3', o: {}, question: 'the control, repeated again' },
    {
      deck: 'shaffine3',
      id: 'shdw-algn-ctr',
      o: { sx: 200, sy: 200, algn: 'ctr' },
      question: 'algn="ctr" with the shadow doubled - centred on what?',
    },
    {
      deck: 'shaffine3',
      id: 'shdw-algn-tl',
      o: { sx: 200, sy: 200, algn: 'tl' },
      question: 'algn="tl" with the shadow doubled - top-left anchored?',
    },
    {
      deck: 'shaffine3',
      id: 'shdw-algn-absent',
      o: { sx: 200, sy: 200 },
      question: 'no @algn with the shadow doubled - what is the default anchor?',
    },
  ];
  const affineSlot = new Map<string, number>();
  affine.forEach((c) => {
    const slot = affineSlot.get(c.deck) ?? 0;
    affineSlot.set(c.deck, slot + 1);
    const x = 20 + (slot % 2) * 470;
    const y = 40 + Math.floor(slot / 2) * 260;
    probes.push({
      id: c.id,
      deck: c.deck,
      group: 'shadow',
      question: c.question,
      x,
      y,
      cx: 100,
      cy: 55,
      geom: RECT_GEOM,
      fill: BLACK_FILL,
      line: '<a:ln><a:noFill/></a:ln>',
      effect: effectLst(
        outerShdw({ blurRad: 0, dist: 200, dirDeg: 0, color: shadowColor('FF0000'), ...c.o }),
      ),
      read: { mode: 'cols', x0: x + 130, y0: y - 100, x1: x + 450, y1: y + 155 },
    });
  });

  /* ------------------------------------------- glow, soft edge, inner, blur */
  [4, 8, 16].forEach((rad, i) => {
    const x = 60 + i * 300;
    probes.push({
      id: `glow-${String(rad)}`,
      deck: 'glow',
      group: 'glow',
      question: `how far does a glow of rad=${String(rad)}pt reach, and with what profile?`,
      x,
      y: 80,
      cx: 120,
      cy: 120,
      geom: RECT_GEOM,
      fill: BLACK_FILL,
      line: '<a:ln><a:noFill/></a:ln>',
      effect: effectLst(glow(rad, shadowColor('FF0000'))),
      read: { mode: 'row', y: 140, x0: x - 60, x1: x + 60 },
    });
  });
  // Five radii, not two. Two showed the ramp is a clean Gaussian of sigma
  // rad/3 sitting some distance *inside* the geometric edge, and put that
  // distance at 0.875 and 0.906 of the radius - which is either one constant
  // measured twice with a pixel of noise, or two points on a curve. Five
  // decides, and this is exactly the shape of question 2.7 got wrong by
  // fitting a constant to too little data.
  [2, 4, 8, 16, 32].forEach((rad, i) => {
    const x = 40 + (i % 3) * 300;
    const y = 60 + Math.floor(i / 3) * 240;
    probes.push({
      id: `soft-${String(rad)}`,
      deck: 'soft',
      group: 'softEdge',
      question: `does softEdge rad=${String(rad)}pt fade inward from the edge, outward, or both?`,
      x,
      y,
      cx: 200,
      cy: 160,
      geom: RECT_GEOM,
      fill: BLACK_FILL,
      line: '<a:ln><a:noFill/></a:ln>',
      effect: effectLst(softEdge(rad)),
      read: { mode: 'row', y: y + 80, x0: x - 50, x1: x + 90 },
    });
  });
  [4, 16].forEach((rad, i) => {
    const x = 660;
    const y = 300 + i * 120;
    probes.push({
      id: `blur-${String(rad)}`,
      deck: 'glow',
      group: 'blur',
      question: `does a:blur rad=${String(rad)}pt blur the shape itself, and does grow=1 expand it?`,
      x,
      y,
      cx: 160,
      cy: 90,
      geom: RECT_GEOM,
      fill: BLACK_FILL,
      line: '<a:ln><a:noFill/></a:ln>',
      effect: effectLst(blur(rad)),
      read: { mode: 'row', y: y + 45, x0: x - 50, x1: x + 60 },
    });
  });
  [8, 24].forEach((rad, i) => {
    const x = 60 + i * 300;
    probes.push({
      id: `inner-${String(rad)}`,
      deck: 'inner',
      group: 'innerShdw',
      question: `an innerShdw of blurRad=${String(rad)}pt at dist=0 - how far in does it reach?`,
      x,
      y: 100,
      cx: 200,
      cy: 160,
      geom: RECT_GEOM,
      fill: '<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>',
      line: '<a:ln><a:noFill/></a:ln>',
      effect: effectLst(innerShdw({ blurRad: rad, dist: 0, color: shadowColor('000000') })),
      read: { mode: 'row', y: 180, x0: x - 20, x1: x + 120 },
    });
  });
  // Which inside edge does an offset inner shadow darken? An outer shadow at
  // dir=0 lands to the right of the shape, and the obvious guess is that an
  // inner one darkens the inside of the left edge. Both directions are probed
  // because the first run said otherwise and one reading is not a rule.
  probes.push({
    id: 'inner-dist',
    deck: 'inner',
    group: 'innerShdw',
    question: 'an innerShdw with dist=20pt dir=0 - which inside edge does it darken?',
    x: 60,
    y: 320,
    cx: 200,
    cy: 160,
    geom: RECT_GEOM,
    fill: '<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>',
    line: '<a:ln><a:noFill/></a:ln>',
    effect: effectLst(innerShdw({ blurRad: 0, dist: 20, dirDeg: 0, color: shadowColor('000000') })),
    read: { mode: 'row', y: 400, x0: 40, x1: 280 },
  });
  probes.push({
    id: 'inner-dist180',
    deck: 'inner',
    group: 'innerShdw',
    question: 'the same at dir=180 - does the darkened edge swap over?',
    x: 340,
    y: 320,
    cx: 200,
    cy: 160,
    geom: RECT_GEOM,
    fill: '<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>',
    line: '<a:ln><a:noFill/></a:ln>',
    effect: effectLst(
      innerShdw({ blurRad: 0, dist: 20, dirDeg: 180, color: shadowColor('000000') }),
    ),
    read: { mode: 'row', y: 400, x0: 320, x1: 560 },
  });
  // The nine reflection presets PowerPoint authored, replayed verbatim. Not to
  // model them - to record what they paint, and to find out whether a
  // reflection is a mirror of the shape or a mirror of the shape *plus* its
  // effects.
  probes.push({
    id: 'refl-tight',
    deck: 'inner',
    group: 'reflection',
    question: 'what does the reflection preset PowerPoint calls "tight" actually paint?',
    x: 460,
    y: 300,
    cx: 160,
    cy: 100,
    geom: RECT_GEOM,
    fill: BLACK_FILL,
    line: '<a:ln><a:noFill/></a:ln>',
    effect: effectLst(
      reflection(
        'blurRad="6350" stA="50000" endA="300" endPos="55000" dir="5400000" sy="-100000" algn="bl" rotWithShape="0"',
      ),
    ),
    read: { mode: 'col', x: 540, y0: 290, y1: 520 },
  });

  /* ----------------------------------------------------- order and stacking */
  // Two effects at once. `a:effectLst` is a sequence, so the file order is
  // fixed by the schema and cannot express intent - which means the painting
  // order is a fact about Office, not about the document.
  probes.push({
    id: 'stack-glow-shadow',
    deck: 'inner',
    group: 'order',
    question: 'with both a glow and an outer shadow, which is painted on top?',
    x: 700,
    y: 80,
    cx: 140,
    cy: 120,
    geom: RECT_GEOM,
    fill: BLACK_FILL,
    line: '<a:ln><a:noFill/></a:ln>',
    effect: effectLst(
      glow(16, shadowColor('00FF00')),
      outerShdw({ blurRad: 0, dist: 30, dirDeg: 0, color: shadowColor('FF0000') }),
    ),
    read: { mode: 'cols', x0: 650, y0: 30, x1: 900, y1: 250 },
  });

  return probes;
}

/* -------------------------------------------------------------------------- */
/* the hostile probes                                                         */
/* -------------------------------------------------------------------------- */

/**
 * Fifteen ways to be wrong, each alone in its own package.
 *
 * The split is not tidiness. PowerPoint's only diagnostic is a repair prompt
 * that names nothing, so a deck carrying two of these and coming back repaired
 * has answered neither question. C2 learned that the expensive way.
 */
export function hostileProbes(): Probe[] {
  const cases: { id: string; question: string; line?: string; effect?: string }[] = [
    {
      id: 'h-wneg',
      question: 'is a negative @w refused? ST_LineWidth is a positive coordinate',
      line: ln({ rawW: '-12700' }),
    },
    {
      id: 'h-whuge',
      question: 'is a @w above the ST_LineWidth maximum refused?',
      line: ln({ rawW: '20116801' }),
    },
    {
      id: 'h-dashbad',
      question: 'is a prstDash val outside the eleven refused?',
      line: ln({ w: 12, dash: prstDash('dotDotDash') }),
    },
    {
      id: 'h-dashempty',
      question: 'is an empty a:custDash refused, or does it mean solid?',
      line: ln({ w: 12, dash: '<a:custDash/>' }),
    },
    {
      id: 'h-dashzero',
      question: 'is a custDash segment of zero length refused?',
      line: ln({ w: 12, dash: rawCustDash('<a:ds d="0" sp="300000"/>') }),
    },
    {
      id: 'h-dashneg',
      question: 'is a negative custDash segment refused?',
      line: ln({ w: 12, dash: rawCustDash('<a:ds d="-400000" sp="300000"/>') }),
    },
    {
      id: 'h-capbad',
      question: 'is a @cap outside ST_LineCap refused?',
      line: ln({ w: 12, cap: 'butt' }),
    },
    {
      id: 'h-cmpdbad',
      question: 'is a @cmpd outside ST_CompoundLine refused?',
      line: ln({ w: 12, cmpd: 'quad' }),
    },
    {
      id: 'h-algnbad',
      question: 'is an @algn outside ST_PenAlignment refused?',
      line: ln({ w: 12, algn: 'out' }),
    },
    {
      id: 'h-endbad',
      question: 'is a lineEnd @type outside the six refused?',
      line: ln({ w: 12, tailEnd: lineEnd('tailEnd', 'chevron', 'med', 'med') }),
    },
    {
      id: 'h-endsize',
      question: 'is a lineEnd @w outside sm/med/lg refused?',
      line: ln({ w: 12, tailEnd: lineEnd('tailEnd', 'triangle', 'xl', 'med') }),
    },
    {
      id: 'h-miterneg',
      question: 'is a negative a:miter/@lim refused?',
      line: ln({ w: 12, join: '<a:miter lim="-100000"/>' }),
    },
    {
      id: 'h-blurneg',
      question: 'is a negative outerShdw/@blurRad refused?',
      effect: effectLst(outerShdw({ rawAttrs: ' blurRad="-50800" dist="0"' })),
    },
    {
      id: 'h-glowneg',
      question: 'is a negative a:glow/@rad refused?',
      effect: effectLst(glow('-50800')),
    },
    {
      id: 'h-softnorad',
      question: 'is a:softEdge with no @rad refused? the attribute is required in the schema',
      effect: effectLst(softEdge(null)),
    },
  ];

  return cases.map((c) => ({
    id: c.id,
    deck: c.id,
    group: 'hostile',
    question: c.question,
    x: 200,
    y: 200,
    cx: 400,
    cy: c.effect === undefined ? 0 : 200,
    geom: c.effect === undefined ? LINE_GEOM : RECT_GEOM,
    fill: c.effect === undefined ? '<a:noFill/>' : BLACK_FILL,
    line: c.line ?? '<a:ln><a:noFill/></a:ln>',
    effect: c.effect,
    read: { mode: 'row' as const, y: 200, x0: 160, x1: 640 },
  }));
}
