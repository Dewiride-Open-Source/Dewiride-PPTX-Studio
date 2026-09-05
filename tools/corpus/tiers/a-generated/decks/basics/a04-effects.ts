import type { ProbeSlide } from '../package.ts';
import { grid, line, prstGeom, scheme, shape, solidFill, srgb } from '../shapes.ts';
import type { ProbeDeck } from '../types.ts';

/**
 * The effect vocabulary, and the 3-D scene that is preserved but never rendered.
 *
 * `a:effectLst` is an `xsd:sequence`, not a bag, and its order is not
 * alphabetical by accident of the spelling - it is
 *
 * ```
 * blur, fillOverlay, glow, innerShdw, outerShdw, prstShdw, reflection, softEdge
 * ```
 *
 * taken from `schema-order.gen.ts` rather than from memory. A shape carrying
 * several effects is where that matters, so slide 1 ends with two shapes that
 * carry four and six.
 *
 * Slide 3 is an experiment rather than a fixture, and is alone on its slide for
 * that reason. `a:effectDag` is the other half of `a:spPr`'s effect choice: a
 * directed graph of effect containers, fully specified in ECMA-376 and never
 * written by PowerPoint, which emits `a:effectLst` for everything a user can
 * author. Whether PowerPoint *reads* one is not answerable from the schema, and
 * a deck that opens without a repair prompt is the only way to find out. If it
 * repairs, the finding is that our writer must never synthesize one - and being
 * on its own slide is what makes that one bisection step instead of nine.
 */

const stroke = line({ width: 6350, fill: solidFill(scheme('tx1', '<a:alpha val="35000"/>')) });
const shadowColor = srgb('000000', '<a:alpha val="45000"/>');

const BLUR = '<a:blur rad="63500" grow="1"/>';
const FILL_OVERLAY = `<a:fillOverlay blend="over">${solidFill(srgb('F1C40F', '<a:alpha val="40000"/>'))}</a:fillOverlay>`;
const GLOW = `<a:glow rad="101600">${scheme('accent4', '<a:alpha val="70000"/>')}</a:glow>`;
const INNER_SHDW = `<a:innerShdw blurRad="63500" dist="50800" dir="2700000">${shadowColor}</a:innerShdw>`;
const OUTER_SHDW =
  '<a:outerShdw blurRad="57150" dist="38100" dir="5400000" sx="102000" sy="102000"' +
  ` algn="ctr" rotWithShape="0">${shadowColor}</a:outerShdw>`;
// ST_PresetShadowVal is shdw1..shdw20; there is no other spelling.
const PRST_SHDW = `<a:prstShdw prst="shdw13" dist="38100" dir="2700000">${shadowColor}</a:prstShdw>`;
const REFLECTION =
  '<a:reflection blurRad="6350" stA="52000" stPos="0" endA="300" endPos="35000"' +
  ' dist="0" dir="5400000" fadeDir="5400000" sx="100000" sy="-100000"' +
  ' algn="bl" rotWithShape="0"/>';
const SOFT_EDGE = '<a:softEdge rad="63500"/>';

function effects(): ProbeSlide {
  const at = grid(5, 2);
  let id = 10;
  const shapes: string[] = [];
  const add = (name: string, effectLst: string, geometry = prstGeom('roundRect')): void => {
    shapes.push(
      shape({
        id: id++,
        name,
        ...at(shapes.length),
        geometry,
        fill: solidFill(scheme('accent1', '<a:lumMod val="75000"/>')),
        line: stroke,
        effects: `<a:effectLst>${effectLst}</a:effectLst>`,
      }),
    );
  };

  add('blur', BLUR);
  add('fillOverlay', FILL_OVERLAY);
  add('glow', GLOW);
  add('innerShdw', INNER_SHDW);
  add('outerShdw', OUTER_SHDW);
  add('prstShdw', PRST_SHDW);
  add('reflection', REFLECTION);
  add('softEdge', SOFT_EDGE);
  // The order below is the schema's, and the reason these two shapes exist.
  add('glow + outerShdw + reflection + softEdge', GLOW + OUTER_SHDW + REFLECTION + SOFT_EDGE);
  add(
    'all eight, in sequence order',
    BLUR + FILL_OVERLAY + GLOW + INNER_SHDW + OUTER_SHDW + PRST_SHDW + REFLECTION + SOFT_EDGE,
    prstGeom('ellipse'),
  );

  return {
    title: 'a04 — a:effectLst, one effect at a time and then all eight',
    body: shapes.join(''),
  };
}

function scene3d(): ProbeSlide {
  const at = grid(4, 1);
  let id = 40;
  const shapes: string[] = [];
  const add = (name: string, scene: string, sp3d?: string): void => {
    shapes.push(
      shape({
        id: id++,
        name,
        ...at(shapes.length),
        geometry: prstGeom('roundRect'),
        fill: solidFill(scheme('accent3')),
        line: stroke,
        scene3d: scene,
        ...(sp3d === undefined ? {} : { sp3d }),
      }),
    );
  };

  // `a:scene3d` is camera, lightRig, backdrop, extLst - and `a:camera`'s only
  // child is `a:rot`, which is easy to hang off the wrong element.
  const camera = (prst: string, rot = ''): string => `<a:camera prst="${prst}">${rot}</a:camera>`;
  const lightRig = (rig: string, dir: string, rot = ''): string =>
    `<a:lightRig rig="${rig}" dir="${dir}">${rot}</a:lightRig>`;

  add(
    'scene3d orthographicFront',
    `<a:scene3d>${camera('orthographicFront')}${lightRig('threePt', 't')}</a:scene3d>`,
  );
  add(
    'scene3d perspective + camera rot',
    '<a:scene3d>' +
      camera('perspectiveFront', '<a:rot lat="1200000" lon="20400000" rev="0"/>') +
      lightRig('balanced', 'tl', '<a:rot lat="0" lon="0" rev="1200000"/>') +
      '</a:scene3d>',
  );
  add(
    'scene3d + backdrop',
    '<a:scene3d>' +
      camera('isometricOffAxis1Left') +
      lightRig('soft', 'tr') +
      '<a:backdrop><a:anchor x="0" y="0" z="0"/>' +
      '<a:norm dx="0" dy="0" dz="100000"/><a:up dx="0" dy="100000" dz="0"/></a:backdrop>' +
      '</a:scene3d>',
  );
  add(
    'scene3d + sp3d bevel and extrusion',
    `<a:scene3d>${camera('orthographicFront')}${lightRig('threePt', 't')}</a:scene3d>`,
    // `a:sp3d` is bevelT, bevelB, extrusionClr, contourClr, extLst.
    '<a:sp3d extrusionH="76200" contourW="12700" prstMaterial="metal">' +
      '<a:bevelT w="63500" h="25400" prst="circle"/>' +
      '<a:bevelB w="50800" h="19050" prst="angle"/>' +
      `<a:extrusionClr>${scheme('accent3', '<a:shade val="60000"/>')}</a:extrusionClr>` +
      `<a:contourClr>${scheme('accent3', '<a:shade val="40000"/>')}</a:contourClr>` +
      '</a:sp3d>',
  );

  return { title: 'a04 — a:scene3d and a:sp3d (phase: preserve)', body: shapes.join('') };
}

/** The experiment. One shape, alone, so a repair prompt localises in one step. */
function effectDag(): ProbeSlide {
  const at = grid(2, 1);
  return {
    title: 'a04 — a:effectDag, the half of the effect choice PowerPoint never writes',
    body: shape({
      id: 70,
      name: 'effectDag tree',
      ...at(0),
      geometry: prstGeom('roundRect'),
      fill: solidFill(scheme('accent2')),
      line: stroke,
      effects:
        '<a:effectDag name="probe" type="tree">' +
        `<a:cont name="branch" type="sib">${OUTER_SHDW}${GLOW}</a:cont>` +
        SOFT_EDGE +
        '</a:effectDag>',
    }),
  };
}

export const a04Effects: ProbeDeck = {
  id: 'a04-effects',
  title: 'PPTX Studio corpus: a04 effects',
  description:
    'All eight a:effectLst children one at a time and then together in sequence order; a:scene3d ' +
    'with three camera rigs and a backdrop, plus a:sp3d bevels and extrusion; and an a:effectDag ' +
    'alone on its own slide, which is the half of a:spPr effect choice PowerPoint never authors. ' +
    'First deck to emit innerShadow, glow, softEdge, reflection and scene3d.',
  features: {
    // Chassis 3 + 3 titles + 10 effect shapes + 4 3-D shapes + 1 effectDag.
    shape: 21,
    placeholder: 6,
    presetGeom: 15,
    gradientFill: 2,
    // Each effect is counted once alone, again inside the four-effect shape
    // where it appears, again inside the all-eight shape, and - for the two
    // the graph carries - once more inside `a:effectDag`. The census scans
    // tokens, so nesting inside a container it does not model is invisible to
    // it, and that is the right answer: the element is in the file.
    shadow: 4,
    glow: 4,
    softEdge: 4,
    reflection: 3,
    innerShadow: 2,
    scene3d: 4,
    // `a:blur`, `a:fillOverlay`, `a:prstShdw` and `a:effectDag` itself are in
    // this deck and are not census keys, so they cannot appear here. That is a
    // gap in the census rather than in the deck, and no rule catches it:
    // `C-COV` runs the other way, from key to deck, so a census key with no
    // deck fails and markup with no census key is invisible to everything.
  },
  build: () => ({
    title: 'PPTX Studio corpus: a04 effects',
    slides: [effects(), scene3d(), effectDag()],
  }),
};
