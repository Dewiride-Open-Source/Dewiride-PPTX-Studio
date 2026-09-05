import type { ProbeSlide } from '../../markup/chassis.ts';
import {
  connector,
  grid,
  line,
  prstGeom,
  scheme,
  shape,
  solidFill,
  srgb,
} from '../../markup/shapes.ts';
import type { ProbeDeck } from '../../markup/types.ts';

/**
 * `a:ln` in every shape it takes, and `p:cxnSp`.
 *
 * Two facts this deck exists to hold, both of which are places the
 * specification and Office disagree or where the obvious reading is wrong:
 *
 * **`a:ln/@cap` defaults to flat in Office**, not `sq` as ECMA-376 says. Follow
 * the specification and every uncapped line renders one line-width too long at
 * each end - a difference nobody notices on a border and everybody notices on a
 * dashed leader line. Slide 1 writes all three caps explicitly *and* leaves one
 * shape with no `@cap` at all, which is the case that decides it.
 *
 * **A connector is not a line-shaped `p:sp`.** `CT_Connector` has its own
 * `p:nvCxnSpPr`, and `a:stCxn`/`a:endCxn` inside `p:cNvCxnSpPr` are what glue
 * an end to a connection site on another shape - the sites `a05-geometry`'s
 * `a:cxnLst` declares. It also has no `p:txBody` at all, so a renderer that
 * treats `p:cxnSp` as `p:sp` will look right until someone moves a box.
 *
 * `@cmpd` is preserved but rendered as `sng` until a later phase: the double
 * and triple compound lines are on slide 1 so that the preservation claim has
 * something to be true about.
 */

const ACCENT = solidFill(scheme('accent1', '<a:shade val="80000"/>'));

/** A diagonal line preset, which is the clearest way to look at a stroke. */
function stroke(
  id: number,
  name: string,
  cell: { x: number; y: number; cx: number; cy: number },
  ln: string,
): string {
  return shape({
    id,
    name,
    ...cell,
    geometry: prstGeom('line'),
    fill: '<a:noFill/>',
    line: ln,
    caption: false,
  });
}

/** A captioned rectangle behind each stroke, so the slide reads without a legend. */
function labelled(
  id: number,
  name: string,
  cell: { x: number; y: number; cx: number; cy: number },
  ln: string,
): string {
  return shape({
    id,
    name,
    ...cell,
    geometry: prstGeom('rect'),
    fill: '<a:noFill/>',
    line: ln,
  });
}

// ------------------------------------------------- slide 1: widths and shapes

function widths(): ProbeSlide {
  const at = grid(5, 3);
  let id = 10;
  const shapes: string[] = [];
  const add = (name: string, ln: string): void => {
    shapes.push(labelled(id++, name, at(shapes.length), ln));
  };

  for (const width of [3175, 12700, 28575, 57150]) {
    add('w=' + String(width), line({ width, fill: ACCENT }));
  }
  // The one with no `@cap`. ECMA says the default is square; Office draws flat.
  add('no @cap', line({ width: 38100, fill: ACCENT }));
  for (const cap of ['rnd', 'sq', 'flat'] as const) {
    add('cap=' + cap, line({ width: 38100, cap, fill: ACCENT }));
  }
  for (const compound of ['sng', 'dbl', 'thickThin', 'thinThick', 'tri'] as const) {
    add('cmpd=' + compound, line({ width: 38100, compound, fill: ACCENT }));
  }
  // `algn="in"` insets the whole stroke, which is a clip-and-double-stroke in
  // SVG and has no direct equivalent anywhere else.
  add('algn=in', line({ width: 38100, align: 'in', fill: ACCENT }));
  add(
    'gradient line',
    line({
      width: 38100,
      fill:
        '<a:gradFill><a:gsLst>' +
        `<a:gs pos="0">${scheme('accent1')}</a:gs>` +
        `<a:gs pos="100000">${scheme('accent6')}</a:gs>` +
        '</a:gsLst><a:lin ang="0" scaled="1"/></a:gradFill>',
    }),
  );

  return { title: 'a06 — a:ln widths, caps, compounds and alignment', body: shapes.join('') };
}

// -------------------------------------- slide 2: dashes, joins and arrowheads

/** All eleven values of `ST_PresetLineDashVal`, in the enumeration's order. */
const DASHES = [
  'solid',
  'dot',
  'dash',
  'lgDash',
  'dashDot',
  'lgDashDot',
  'lgDashDotDot',
  'sysDash',
  'sysDot',
  'sysDashDot',
  'sysDashDotDot',
] as const;

function dashes(): ProbeSlide {
  // Nineteen shapes: eleven dashes, a:custDash, three joins, four arrowhead
  // pairs. A 6 x 3 grid is eighteen cells and hangs the last one off the slide.
  const at = grid(7, 3);
  let id = 40;
  const shapes: string[] = [];

  for (const dash of DASHES) {
    shapes.push(
      stroke(
        id++,
        dash,
        at(shapes.length),
        line({
          width: 28575,
          cap: 'flat',
          fill: ACCENT,
          dash: `<a:prstDash val="${dash}"/>`,
        }),
      ),
    );
  }
  // `a:custDash` is dash-space pairs as percentages of the line width, and it
  // is what a preset dash means underneath.
  shapes.push(
    stroke(
      id++,
      'custDash',
      at(shapes.length),
      line({
        width: 28575,
        cap: 'flat',
        fill: ACCENT,
        dash: '<a:custDash><a:ds d="400000" sp="200000"/><a:ds d="100000" sp="200000"/></a:custDash>',
      }),
    ),
  );

  for (const join of ['bevel', 'miter', 'round'] as const) {
    shapes.push(
      labelled(
        id++,
        'join ' + join,
        at(shapes.length),
        line({
          width: 57150,
          fill: ACCENT,
          join: join === 'miter' ? '<a:miter lim="800000"/>' : `<a:${join}/>`,
        }),
      ),
    );
  }

  // Arrowheads. `@type` is the shape, `@w`/`@len` scale it in line widths.
  const ends: readonly (readonly [string, string])[] = [
    ['triangle', 'stealth'],
    ['arrow', 'diamond'],
    ['oval', 'none'],
    ['diamond', 'triangle'],
  ];
  for (const [head, tail] of ends) {
    shapes.push(
      stroke(
        id++,
        head + ' / ' + tail,
        at(shapes.length),
        line({
          width: 28575,
          fill: ACCENT,
          headEnd: `<a:headEnd type="${head}" w="lg" len="lg"/>`,
          tailEnd: `<a:tailEnd type="${tail}" w="med" len="med"/>`,
        }),
      ),
    );
  }

  return {
    title: 'a06 — all 11 prstDash values, a:custDash, the three joins and arrowheads',
    body: shapes.join(''),
  };
}

// ------------------------------------------------------- slide 3: connectors

function connectors(): ProbeSlide {
  // Four boxes, each with the four default connection sites a rectangle has,
  // and connectors glued between them by `a:stCxn`/`a:endCxn`.
  const boxes = [
    { id: 100, x: 700000, y: 1600000 },
    { id: 101, x: 4600000, y: 1600000 },
    { id: 102, x: 700000, y: 4300000 },
    { id: 103, x: 4600000, y: 4300000 },
  ];
  const box = (spec: { id: number; x: number; y: number }, name: string): string =>
    shape({
      id: spec.id,
      name,
      x: spec.x,
      y: spec.y,
      cx: 1800000,
      cy: 1000000,
      geometry: prstGeom('roundRect', { adj: 20000 }),
      fill: solidFill(scheme('accent2', '<a:lumMod val="40000"/><a:lumOff val="60000"/>')),
      line: line({ width: 12700, fill: ACCENT }),
    });

  const arrow = line({
    width: 19050,
    fill: solidFill(srgb('44546A')),
    tailEnd: '<a:tailEnd type="triangle" w="med" len="med"/>',
  });

  return {
    title: 'a06 — p:cxnSp glued to connection sites with a:stCxn and a:endCxn',
    body:
      boxes.map((spec, index) => box(spec, 'Box ' + String(index + 1))).join('') +
      // Site 3 of a rectangle is its right edge, site 1 its left. Straight.
      connector({
        id: 110,
        name: 'straightConnector1',
        x: 2500000,
        y: 2100000,
        cx: 2100000,
        cy: 0,
        geometry: prstGeom('straightConnector1'),
        line: arrow,
        start: { id: 100, idx: 3 },
        end: { id: 101, idx: 1 },
      }) +
      // Site 2 is the bottom edge, site 0 the top. Bent, with three segments.
      connector({
        id: 111,
        name: 'bentConnector3',
        x: 1600000,
        y: 2600000,
        cx: 0,
        cy: 1700000,
        geometry: prstGeom('bentConnector3', { adj1: 50000 }),
        line: arrow,
        start: { id: 100, idx: 2 },
        end: { id: 102, idx: 0 },
      }) +
      connector({
        id: 112,
        name: 'curvedConnector3',
        x: 5500000,
        y: 2600000,
        cx: 0,
        cy: 1700000,
        geometry: prstGeom('curvedConnector3', { adj1: 50000 }),
        line: arrow,
        start: { id: 101, idx: 2 },
        end: { id: 103, idx: 0 },
      }) +
      // An unglued connector: valid, common, and the case that proves
      // `a:stCxn` is optional rather than structural.
      connector({
        id: 113,
        name: 'line, unglued',
        x: 2500000,
        y: 4800000,
        cx: 2100000,
        cy: 0,
        geometry: prstGeom('line'),
        line: line({
          width: 19050,
          fill: solidFill(srgb('44546A')),
          dash: '<a:prstDash val="dash"/>',
        }),
      }),
  };
}

export const a06Lines: ProbeDeck = {
  id: 'a06-lines',
  title: 'PPTX Studio corpus: a06 lines and connectors',
  description:
    'a:ln across four widths, all three caps and a shape with none, all five @cmpd values, ' +
    'algn=in, a gradient stroke, all 11 ST_PresetLineDashVal values, a:custDash, the three line ' +
    'joins and four arrowhead pairs; plus p:cxnSp connectors glued to connection sites with ' +
    'a:stCxn and a:endCxn, and one deliberately unglued. First deck to emit connector.',
  features: {
    // Chassis 3 + 3 titles + 15 stroke shapes + 19 dash/join/arrow shapes + 4
    // boxes. The four connectors are `connector`, not `shape`.
    shape: 44,
    placeholder: 6,
    // 38 shapes with geometry plus all four connectors, which have geometry
    // too - a connector's `a:prstGeom` is what gives it its route.
    presetGeom: 42,
    connector: 4,
    // The theme's two, plus the gradient stroke on slide 1. `a:gradFill`
    // inside an `a:ln` is the same element and the census counts it as one.
    gradientFill: 3,
  },
  build: () => ({
    title: 'PPTX Studio corpus: a06 lines and connectors',
    slides: [widths(), dashes(), connectors()],
  }),
};
