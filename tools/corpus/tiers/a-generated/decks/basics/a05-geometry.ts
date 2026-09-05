import type { ProbeSlide } from '../../markup/chassis.ts';
import { grid, line, prstGeom, scheme, shape, solidFill } from '../../markup/shapes.ts';
import type { ProbeDeck } from '../../markup/types.ts';

/**
 * `a:custGeom`, and preset geometry with its adjust values moved off default.
 *
 * `a:custGeom` is an `xsd:sequence` of six children -
 * `avLst, gdLst, ahLst, cxnLst, rect, pathLst` - and only the last is
 * interesting to a renderer. The other five are what make a custom shape
 * *editable*, and they are exactly what every importer drops: adjust handles
 * with their ranges, connection sites with their angles, and the text rectangle
 * that is not the bounding box.
 *
 * ## What this deck deliberately is not
 *
 * It is not a 187-preset gallery. Sub-phase 2.1 transcodes Apache POI's
 * `presetShapeDefinitions.xml` and 2.4 gets a free conformance test out of it -
 * rewrite each `prstGeom` as the equivalent `custGeom` and assert identical
 * output. Enumerating 187 presets here would duplicate that generated data in a
 * hand-written fixture, and the hand-written one would be the wrong one. What
 * the corpus owes 2.x is a deck exercising the *branches*: the six path
 * commands, both adjust-handle kinds, per-path fill and stroke, and a path
 * space that is not the shape's.
 *
 * The 24 presets on slide 3 are chosen for that reason. Eight of them are the
 * ones sub-phase 2.3 names as covering every branch of the arc code, where
 * `a:arcTo`'s angles are **not** parametric and feeding `stAng` straight to an
 * SVG ellipse is wrong for every one.
 */

const stroke = line({ width: 9525, fill: solidFill(scheme('tx1')) });
const fill = solidFill(scheme('accent1', '<a:lumMod val="60000"/><a:lumOff val="40000"/>'));

/**
 * `a:custGeom` with its six children in the only order the schema allows.
 *
 * Every field is the **contents** of its wrapper and the wrapper is built here.
 * The first version took the wrapper for four of them and the contents for the
 * other two, and the callers below passed contents for all six - so `a:ahXY`
 * and `a:cxn` went in as direct children of `a:custGeom`. PowerPoint's answer
 * was to refuse the entire package: "PowerPoint could not open the file", no
 * part name, no element name, nothing. Making the wrapper the helper's job is
 * the fix that cannot be got wrong again.
 */
function custGeom(parts: {
  readonly adjustValues?: string;
  readonly guides?: string;
  readonly handles?: string;
  readonly sites?: string;
  readonly textRect?: string;
  readonly paths: string;
}): string {
  const wrap = (name: string, contents: string | undefined): string =>
    contents === undefined || contents === ''
      ? `<a:${name}/>`
      : `<a:${name}>${contents}</a:${name}>`;
  return (
    '<a:custGeom>' +
    wrap('avLst', parts.adjustValues) +
    wrap('gdLst', parts.guides) +
    wrap('ahLst', parts.handles) +
    wrap('cxnLst', parts.sites) +
    (parts.textRect ?? '') +
    wrap('pathLst', parts.paths) +
    '</a:custGeom>'
  );
}

const pt = (x: string | number, y: string | number): string =>
  `<a:pt x="${String(x)}" y="${String(y)}"/>`;
const moveTo = (x: string | number, y: string | number): string =>
  `<a:moveTo>${pt(x, y)}</a:moveTo>`;
const lnTo = (x: string | number, y: string | number): string => `<a:lnTo>${pt(x, y)}</a:lnTo>`;

// -------------------------------------------------- slide 1: the path commands

function pathCommands(): ProbeSlide {
  const at = grid(3, 2);
  let id = 10;
  const shapes: string[] = [];
  const add = (name: string, geometry: string): void => {
    shapes.push(shape({ id: id++, name, ...at(shapes.length), geometry, fill, line: stroke }));
  };

  // A path space of 100 x 100 is mapped onto the shape's `a:ext`, so these
  // coordinates are proportions and the shape resizes correctly.
  const p = (body: string, attributes = ' w="100" h="100"'): string =>
    `<a:path${attributes}>${body}</a:path>`;

  add(
    'moveTo lnTo close',
    custGeom({ paths: p(moveTo(50, 0) + lnTo(100, 100) + lnTo(0, 100) + '<a:close/>') }),
  );
  add(
    'cubicBezTo',
    custGeom({
      paths: p(
        moveTo(0, 100) +
          `<a:cubicBezTo>${pt(0, 0)}${pt(100, 100)}${pt(100, 0)}</a:cubicBezTo>` +
          lnTo(100, 100) +
          '<a:close/>',
      ),
    }),
  );
  add(
    'quadBezTo',
    custGeom({
      paths: p(
        moveTo(0, 100) + `<a:quadBezTo>${pt(50, -40)}${pt(100, 100)}</a:quadBezTo>` + '<a:close/>',
      ),
    }),
  );
  // `a:arcTo` is placed by its **start point**, not a centre, and `stAng`/
  // `swAng` are 60000ths of a degree. Office clamps `swAng` to +-360 degrees.
  add(
    'arcTo',
    custGeom({
      paths: p(
        moveTo(50, 50) +
          lnTo(100, 50) +
          '<a:arcTo wR="50" hR="50" stAng="0" swAng="10800000"/>' +
          '<a:close/>',
      ),
    }),
  );
  // Two paths in one geometry, each with its own fill and stroke. This is why
  // a renderer must emit one `<path>` per `a:path` and never merge them.
  add(
    'two paths, per-path fill and stroke',
    custGeom({
      paths:
        p(
          moveTo(0, 0) + lnTo(100, 0) + lnTo(100, 100) + lnTo(0, 100) + '<a:close/>',
          ' w="100" h="100" fill="norm" stroke="1"',
        ) +
        p(
          moveTo(50, 20) + lnTo(80, 50) + lnTo(50, 80) + lnTo(20, 50) + '<a:close/>',
          ' w="100" h="100" fill="darken" stroke="0"',
        ),
    }),
  );
  // `@w`/`@h` omitted means no path-space scaling: the coordinates are EMU in
  // the shape's own space. Not a divide-by-zero, and not a resize bug either.
  const c = at(5);
  add(
    'no path space, EMU coordinates',
    custGeom({
      paths: p(moveTo(0, 0) + lnTo(c.cx, 0) + lnTo(Math.round(c.cx / 2), c.cy) + '<a:close/>', ''),
    }),
  );

  return { title: 'a05 — the six a:path commands, and path space', body: shapes.join('') };
}

// --------------------------------- slide 2: guides, handles, sites, text rect

function editableGeometry(): ProbeSlide {
  const at = grid(2, 1);
  const id = 40;

  // An adjustable arrow, written the way the built-in `rightArrow` is: adjust
  // values pinned into range by `gdLst`, an XY handle bound to both by name,
  // four connection sites, and a text rectangle that is the arrow's shaft
  // rather than its bounding box.
  const arrow = custGeom({
    adjustValues: '<a:gd name="adj1" fmla="val 50000"/><a:gd name="adj2" fmla="val 40000"/>',
    guides:
      // `ss` is min(w,h) and one of the 44 seeded built-in guides; so are
      // `l`, `r`, `t`, `b`, `hc`, `vc`, `w` and `h`.
      '<a:gd name="maxAdj2" fmla="*/ 100000 w ss"/>' +
      '<a:gd name="a1" fmla="pin 0 adj1 100000"/>' +
      '<a:gd name="a2" fmla="pin 0 adj2 maxAdj2"/>' +
      '<a:gd name="dx1" fmla="*/ ss a2 100000"/>' +
      '<a:gd name="x1" fmla="+- r 0 dx1"/>' +
      '<a:gd name="dy1" fmla="*/ h a1 200000"/>' +
      '<a:gd name="y1" fmla="+- vc 0 dy1"/>' +
      '<a:gd name="y2" fmla="+- vc dy1 0"/>',
    handles:
      '<a:ahXY gdRefX="adj2" minX="0" maxX="maxAdj2" gdRefY="adj1" minY="0" maxY="100000">' +
      '<a:pos x="x1" y="y1"/></a:ahXY>',
    sites:
      '<a:cxn ang="16200000"><a:pos x="hc" y="t"/></a:cxn>' +
      '<a:cxn ang="10800000"><a:pos x="l" y="vc"/></a:cxn>' +
      '<a:cxn ang="5400000"><a:pos x="hc" y="b"/></a:cxn>' +
      '<a:cxn ang="0"><a:pos x="r" y="vc"/></a:cxn>',
    // Not the bounding box. Text in this shape belongs in the shaft.
    textRect: '<a:rect l="l" t="y1" r="x1" b="y2"/>',
    paths:
      '<a:path>' +
      moveTo('l', 'y1') +
      lnTo('x1', 'y1') +
      lnTo('x1', 't') +
      lnTo('r', 'vc') +
      lnTo('x1', 'b') +
      lnTo('x1', 'y2') +
      lnTo('l', 'y2') +
      '<a:close/></a:path>',
  });

  // A polar handle. `ahPolar` binds a radius guide and an angle guide, and is
  // the kind every circular preset uses; nothing in the corpus had one.
  const dial = custGeom({
    adjustValues: '<a:gd name="adj1" fmla="val 25000"/><a:gd name="adj2" fmla="val 3600000"/>',
    guides:
      '<a:gd name="r1" fmla="pin 0 adj1 50000"/>' +
      '<a:gd name="ang1" fmla="pin 0 adj2 21600000"/>' +
      '<a:gd name="dr" fmla="*/ ss r1 100000"/>' +
      '<a:gd name="ry" fmla="+- vc 0 dr"/>',
    handles:
      '<a:ahPolar gdRefR="adj1" minR="0" maxR="50000" gdRefAng="adj2" minAng="0" maxAng="21600000">' +
      '<a:pos x="hc" y="ry"/></a:ahPolar>',
    sites: '<a:cxn ang="0"><a:pos x="r" y="vc"/></a:cxn>',
    textRect: '<a:rect l="l" t="t" r="r" b="b"/>',
    paths:
      '<a:path>' +
      moveTo('l', 'vc') +
      '<a:arcTo wR="hc" hR="vc" stAng="10800000" swAng="21600000"/>' +
      '<a:close/></a:path>' +
      '<a:path fill="none" stroke="1">' +
      moveTo('hc', 'vc') +
      lnTo('hc', 'ry') +
      '</a:path>',
  });

  return {
    title: 'a05 — gdLst, ahXY, ahPolar, cxnLst and the a:rect text rectangle',
    body:
      shape({
        id: id,
        name: 'custGeom arrow with an XY handle',
        ...at(0),
        geometry: arrow,
        fill,
        line: stroke,
      }) +
      shape({
        id: id + 1,
        name: 'custGeom dial with a polar handle',
        ...at(1),
        geometry: dial,
        fill,
        line: stroke,
      }),
  };
}

// ---------------------------------------------- slide 3: presets off default

/**
 * Presets with their adjust values moved off default.
 *
 * The eight marked `arc` are sub-phase 2.3's list: between them they take every
 * branch of the arc code. The adjust names are the ones POI's definitions use;
 * an unrecognised name in `avLst` is harmless - it merges over `gdLst` by name
 * and an unmatched one is simply unused - so being wrong here renders a default
 * shape rather than a broken package.
 */
const PRESETS: readonly (readonly [string, Readonly<Record<string, number>>])[] = [
  ['roundRect', { adj: 40000 }],
  ['round1Rect', { adj: 35000 }],
  ['snip1Rect', { adj: 35000 }],
  ['plaque', { adj: 25000 }],
  ['bevel', { adj: 20000 }],
  ['frame', { adj1: 20000 }],
  ['halfFrame', { adj1: 40000, adj2: 30000 }],
  ['corner', { adj1: 40000, adj2: 40000 }],
  ['triangle', { adj: 20000 }],
  ['parallelogram', { adj: 35000 }],
  ['trapezoid', { adj: 30000 }],
  ['hexagon', { adj: 30000, vf: 115470 }],
  ['pentagon', { hf: 105146, vf: 110557 }],
  ['star5', { adj: 25000, hf: 105146, vf: 110557 }],
  ['donut', { adj: 30000 }], // arc
  ['pie', { adj1: 1800000, adj2: 14400000 }], // arc
  ['arc', { adj1: 16200000, adj2: 5400000 }], // arc
  ['chord', { adj1: 2700000, adj2: 16200000 }], // arc
  ['blockArc', { adj1: 10800000, adj2: 5400000, adj3: 25000 }], // arc
  ['moon', { adj: 60000 }], // arc
  ['smileyFace', { adj: -2000 }], // arc
  ['circularArrow', { adj1: 12500, adj2: 1142319, adj3: 20457681, adj4: 10800000, adj5: 12500 }], // arc
  ['actionButtonSound', {}], // arc
  ['wedgeRoundRectCallout', { adj1: -20833, adj2: 62500, adj3: 16667 }],
];

function presets(): ProbeSlide {
  const at = grid(8, 3);
  let id = 70;
  return {
    title: 'a05 — 24 presets with their a:avLst moved off default',
    body: PRESETS.map(([preset, adjustments], index) =>
      shape({
        id: id++,
        name: preset,
        ...at(index),
        geometry: prstGeom(preset, adjustments),
        fill,
        line: line({ width: 6350, fill: solidFill(scheme('tx1')) }),
      }),
    ).join(''),
  };
}

export const a05Geometry: ProbeDeck = {
  id: 'a05-geometry',
  title: 'PPTX Studio corpus: a05 geometry',
  description:
    'a:custGeom across all six path commands, both path-space modes, per-path fill and stroke, ' +
    'a:gdLst formulas, both adjust-handle kinds, a:cxnLst connection sites and an a:rect text ' +
    'rectangle that is not the bounding box; plus 24 a:prstGeom presets with non-default adjust ' +
    'values, eight of them sub-phase 2.3 arc cases. First deck to emit customGeom.',
  features: {
    // Chassis 3 + 3 titles + 6 path shapes + 2 editable shapes + 24 presets.
    shape: 38,
    placeholder: 6,
    // Only slide 3. The eight custom-geometry shapes carry no `a:prstGeom` -
    // the two are a choice inside `a:spPr`, never both.
    presetGeom: 24,
    customGeom: 8,
    gradientFill: 2,
  },
  build: () => ({
    title: 'PPTX Studio corpus: a05 geometry',
    slides: [pathCommands(), editableGeometry(), presets()],
  }),
};
