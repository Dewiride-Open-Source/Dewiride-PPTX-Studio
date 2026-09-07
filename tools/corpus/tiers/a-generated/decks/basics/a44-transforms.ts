import type { ProbeSlide } from '../../markup/chassis.ts';
import { grid, group, line, prstGeom, scheme, shape, solidFill } from '../../markup/shapes.ts';
import type { ProbeDeck } from '../../markup/types.ts';

/**
 * Rotation, mirroring, and the coordinate space a group's children live in.
 *
 * Every other deck in the corpus writes `a:xfrm` with no `@rot` and no
 * `@flipH`, nests no group inside another and never resizes one, so the three
 * rules sub-phase 2.10 rests on had no corpus slide to be wrong on. Gate 2
 * names them. ADR 0038.
 *
 * The probe geometry is `rtTriangle` throughout: it is the cheapest shape that
 * is asymmetric in both axes at once, where a square hides a flip and an arrow
 * hides a vertical one. Each probe is square and centred in its cell, so a
 * shape at 45 degrees stays inside the cell it belongs to.
 */

const STROKE = line({ width: 12700, fill: solidFill(scheme('tx1')) });
const FILL = solidFill(scheme('accent1'));

/** Sixtieths of a degree, which is what `@rot` counts in. */
const DEG = 60000;

interface Square {
  readonly x: number;
  readonly y: number;
  readonly cx: number;
  readonly cy: number;
}

/** A square of `side`, centred in the cell, so any rotation stays inside it. */
function centred(cell: { x: number; y: number; cx: number; cy: number }, side: number): Square {
  return {
    x: cell.x + Math.floor((cell.cx - side) / 2),
    y: cell.y + Math.floor((cell.cy - side) / 2),
    cx: side,
    cy: side,
  };
}

// -------------------------------------------------- slide 1: rotate and flip

/** Side 1150000: its diagonal is 1626000, inside the 1727200-tall cell. */
const PROBE_SIDE = 1150000;

/** The three `a:xfrm` attributes, spread whole so an absent one stays absent. */
interface Xfrm {
  readonly rotation?: number;
  readonly flipH?: boolean;
  readonly flipV?: boolean;
}

interface Transform {
  readonly name: string;
  readonly xfrm: Xfrm;
}

/**
 * The twelve readings, ordered so the separating cases come last.
 *
 * The final pair is what a renderer composing rotation before the mirror draws
 * at the wrong angle; the ten before them read the same under either order.
 */
const TRANSFORMS: readonly Transform[] = [
  { name: 'rot 0', xfrm: {} },
  { name: 'rot 30', xfrm: { rotation: 30 * DEG } },
  { name: 'rot 45', xfrm: { rotation: 45 * DEG } },
  { name: 'rot 90', xfrm: { rotation: 90 * DEG } },
  { name: 'rot 135', xfrm: { rotation: 135 * DEG } },
  { name: 'rot 180', xfrm: { rotation: 180 * DEG } },
  { name: 'rot 270', xfrm: { rotation: 270 * DEG } },
  { name: 'flipH', xfrm: { flipH: true } },
  { name: 'flipV', xfrm: { flipV: true } },
  { name: 'flipH flipV', xfrm: { flipH: true, flipV: true } },
  { name: 'flipH rot 30', xfrm: { rotation: 30 * DEG, flipH: true } },
  { name: 'flipV rot 30', xfrm: { rotation: 30 * DEG, flipV: true } },
];

function rotateAndFlip(): ProbeSlide {
  const at = grid(4, 3);
  return {
    title: 'a44 - rotation and mirroring on a shape',
    body: TRANSFORMS.map((transform, index) =>
      shape({
        id: 10 + index,
        name: transform.name,
        ...centred(at(index), PROBE_SIDE),
        geometry: prstGeom('rtTriangle'),
        fill: FILL,
        line: STROKE,
        ...transform.xfrm,
      }),
    ).join(''),
  };
}

// ----------------------------------------------- slide 2: groups, and nesting

/** Side 1750000: its diagonal is 2475000, inside the 2590800-tall cell. */
const GROUP_SIDE = 1750000;

const LEAF_PRESETS = ['rtTriangle', 'ellipse', 'rect'] as const;
const LEAF_FILLS = ['accent2', 'accent4', 'accent6'] as const;

/** Three leaves stepped down the diagonal of a `side` square at `origin`. */
function leaves(startId: number, originX: number, originY: number, side: number): string {
  const unit = Math.floor(side / 3);
  return LEAF_PRESETS.map((preset, n) =>
    shape({
      id: startId + n,
      name: preset,
      x: originX + n * unit,
      y: originY + n * unit,
      cx: unit,
      cy: unit,
      geometry: prstGeom(preset),
      fill: solidFill(scheme(LEAF_FILLS[n] ?? 'accent1')),
      line: STROKE,
      caption: false,
    }),
  ).join('');
}

function groups(): ProbeSlide {
  const at = grid(3, 2);
  const box = (index: number): Square => centred(at(index), GROUP_SIDE);
  const shapes: string[] = [];

  // The control: child space is the group's own, so nothing is scaled.
  const plain = box(0);
  shapes.push(
    group({
      id: 40,
      name: 'group, chExt = ext',
      ...plain,
      children: leaves(41, plain.x, plain.y, GROUP_SIDE),
    }),
  );

  // A `chExt` of half the extent draws every child at twice its written size;
  // dropping the scale factor renders this quarter-size in one corner.
  const resized = box(1);
  const half = Math.floor(GROUP_SIDE / 2);
  shapes.push(
    group({
      id: 50,
      name: 'group, chExt = ext / 2',
      ...resized,
      childOffsetX: 0,
      childOffsetY: 0,
      childWidth: half,
      childHeight: half,
      children: leaves(51, 0, 0, half),
    }),
  );

  const rotated = box(2);
  shapes.push(
    group({
      id: 60,
      name: 'group, rot 30',
      ...rotated,
      rotation: 30 * DEG,
      children: leaves(61, rotated.x, rotated.y, GROUP_SIDE),
    }),
  );

  // A mirrored group mirrors each child about the group centre, not its own.
  const flipped = box(3);
  shapes.push(
    group({
      id: 70,
      name: 'group, flipH',
      ...flipped,
      flipH: true,
      children: leaves(71, flipped.x, flipped.y, GROUP_SIDE),
    }),
  );

  // The case Gate 2 names: the only place two transforms compose across a level.
  const nested = box(4);
  const innerSide = Math.floor(GROUP_SIDE * 0.6);
  shapes.push(
    group({
      id: 80,
      name: 'group rot 30 > group flipH',
      ...nested,
      rotation: 30 * DEG,
      children:
        group({
          id: 81,
          name: 'inner, flipH',
          x: nested.x,
          y: nested.y,
          cx: innerSide,
          cy: innerSide,
          flipH: true,
          children: leaves(82, nested.x, nested.y, innerSide),
        }) +
        shape({
          id: 85,
          name: 'sibling',
          x: nested.x + innerSide,
          y: nested.y + innerSide,
          cx: GROUP_SIDE - innerSide,
          cy: GROUP_SIDE - innerSide,
          geometry: prstGeom('rtTriangle'),
          fill: FILL,
          line: STROKE,
          caption: false,
        }),
    }),
  );

  // A child space far from the group's own origin: a renderer that subtracts
  // nothing puts every child off the slide.
  const shifted = box(5);
  const ORIGIN = 5000000;
  shapes.push(
    group({
      id: 90,
      name: 'group, chOff far from off',
      ...shifted,
      childOffsetX: ORIGIN,
      childOffsetY: ORIGIN,
      childWidth: GROUP_SIDE,
      childHeight: GROUP_SIDE,
      children: leaves(91, ORIGIN, ORIGIN, GROUP_SIDE),
    }),
  );

  return {
    title: 'a44 - groups: resized, rotated, mirrored and nested',
    body: shapes.join(''),
  };
}

export const a44Transforms: ProbeDeck = {
  id: 'a44-transforms',
  title: 'PPTX Studio corpus: a44 transforms',
  description:
    'Rotation and mirroring on a shape, including the flip-with-rotation pair that separates the ' +
    'two composition orders, and six groups: the identity, a halved child space, a rotated group, ' +
    'a mirrored one, a rotated group holding a mirrored one, and a child space far from the ' +
    'origin. First deck in the corpus to write @rot, @flipH or a nested p:grpSp.',
  addedIn: '2.14',
  features: {
    // Chassis 3 + 2 slide titles + 12 transforms + 19 group leaves.
    shape: 36,
    placeholder: 5,
    presetGeom: 31,
    gradientFill: 2,
    group: 7,
  },
  build: () => ({
    title: 'PPTX Studio corpus: a44 transforms',
    slides: [rotateAndFlip(), groups()],
  }),
};
