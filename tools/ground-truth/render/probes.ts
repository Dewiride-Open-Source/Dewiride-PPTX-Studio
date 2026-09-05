/**
 * Experiment C6 - what a group does to the shapes inside it.
 *
 * Sub-phase 2.10 renders geometry, and everything hard about that is a
 * transform: the order a flip and a rotation compose in, the coordinate system
 * a group's children are written in, and what `a:grpFill` reaches for.
 *
 * ## The order question was settled before this file existed
 *
 * `tools/ground-truth/render/author.ps1` asked PowerPoint to mirror shapes it had already
 * rotated, and read back what it wrote. A shape at 30 degrees, mirrored, comes
 * out as `rot="19800000" flipH="1"` - that is **minus** thirty degrees with a
 * flip, on all four of 30, 45, 120 and 200, and on both axes. Since
 * `F R(t) F = R(-t)` for any reflection, writing `-t` is what a renderer that
 * flips *before* it rotates has to write to reproduce a mirrored figure, and
 * writing `+t` is what one that rotates first has to write. PowerPoint's writer
 * and its renderer agree with each other by construction, so this is the
 * renderer's answer: **`M = translate . rotate . flip`, flip innermost.**
 *
 * Both flips together came back with the rotation *unchanged*, which is the
 * same fact from the other side: `flipH flipV` is a 180-degree rotation, and
 * rotations commute.
 *
 * So this file does not re-ask that. It asks the questions the authored decks
 * could not reach, because COM cannot write a `chOff` that disagrees with an
 * `off`, cannot write an `a:grpFill`, and cannot write a `chExt` of zero.
 *
 * ## The measurement, and why it costs nothing
 *
 * A leaf shape inside a group has no slide position of its own - the group's
 * transform gives it one, and `Shape.Left` reports that position in points.
 * Measured on PowerPoint's own output: a child of a group rotated 30 degrees
 * and mirrored reported `L=321.2435302734375`, and the same number appears in
 * the file as `x="4079793"` once the group is ungrouped. So the object model is
 * the renderer's own arithmetic, to four decimal places, with no bitmap in the
 * way.
 *
 * Two readings per leaf, therefore:
 *
 * - **In place**, where `Left`/`Top`/`Width`/`Height` are the composed frame.
 *   Rotation is composed too; `HorizontalFlip` is *not* - measured, it reports
 *   the child's own attribute while `Rotation` reports the total. That
 *   inconsistency is why there is a second pass.
 * - **After an in-memory `Ungroup`**, which is PowerPoint composing the whole
 *   chain and writing it down. Nothing is saved; the file on disk is untouched.
 *
 * `a:grpFill` is the exception and needs the bitmap, for the same reason C3
 * needed it: a gradient has no single colour, and the whole question is whether
 * a child shows a *slice* of the group's ramp or a copy of it.
 *
 * ## Every probe declares its chain
 *
 * A probe carries the group chain as data and the analysis computes each
 * candidate model from it, rather than comparing against a number typed in
 * here. The catalogue and the markup are generated from the same field, so a
 * probe cannot ask one question and measure another - which is the mistake C5
 * made twice.
 */

import { groupXml, shapeXml } from '../lib/pptx.ts';
import { buildSheetPackage, SCHEME_ONE, shape as sheetShape } from '../lib/sheet-pptx.ts';

export const EMU_PER_POINT = 12700;
export const SLIDE_WIDTH_PT = 960;
export const SLIDE_HEIGHT_PT = 540;
/** `Slide.Export` width. Two device pixels per point, so every whole-point edge
 *  lands on a pixel boundary and no sampled interior is a blend. */
export const EXPORT_WIDTH = 1920;
export const EXPORT_HEIGHT = 1080;

/** Degrees to `ST_Angle`. */
export function deg(value: number): number {
  return Math.round(value * 60000);
}

function emu(points: number, what: string): number {
  const value = points * EMU_PER_POINT;
  if (!Number.isInteger(value)) {
    throw new Error(`${what}: ${String(points)}pt is not a whole number of EMU`);
  }
  return value;
}

export interface Rect {
  readonly x: number;
  readonly y: number;
  readonly w: number;
  readonly h: number;
}

/** One link of a transform chain: a `p:grpSp`'s `a:xfrm`, in points. */
export interface Frame {
  /** `a:off` and `a:ext`, in the *parent's* child space. */
  readonly rect: Rect;
  /** `a:chOff` and `a:chExt`: the space this group's own children are written in. */
  readonly child: Rect;
  /** Degrees, clockwise on screen. */
  readonly rot?: number | undefined;
  readonly flipH?: boolean | undefined;
  readonly flipV?: boolean | undefined;
  /** A whole fill element on `p:grpSpPr`. */
  readonly fill?: string | undefined;
  /**
   * Write no `a:chOff` and no `a:chExt`. `child` is then what the *probe*
   * claims, and the point is to find out what PowerPoint assumes instead.
   */
  readonly noChildTransform?: boolean | undefined;
}

/** A point on the exported bitmap, in slide points, and what it asks. */
export interface Sample {
  readonly id: string;
  readonly x: number;
  readonly y: number;
  readonly asks: string;
  /** The colour the "slice" reading predicts, when the probe has one. */
  readonly expect?: string | undefined;
}

/**
 * A row of the exported bitmap carrying a reference ramp.
 *
 * The gradient probes read a colour and need a *position* out of it. Scanning
 * a shape that has the identical gradient over a known rectangle turns the
 * colour back into a fraction along the ramp, which is what the question is
 * actually about - and it does that without this file having to know anything
 * about the 1.875 blend curve, which is 2.7's business and measured there.
 */
export interface ReferenceRow {
  readonly y: number;
  readonly x0: number;
  readonly x1: number;
  /** Scan down a column instead, for a vertical ramp. */
  readonly vertical?: boolean | undefined;
}

export interface TransformProbe {
  readonly id: string;
  readonly deck: string;
  /** 1-based slide index within the deck. */
  readonly slide: number;
  readonly kind: 'map' | 'turn' | 'fill' | 'width';
  readonly question: string;
  /** The leaf shape whose composed frame is the answer. */
  readonly shape: string;
  /** The leaf's `a:xfrm`, in the innermost group's child space. */
  readonly leaf: Rect;
  readonly leafRot: number;
  readonly leafFlipH: boolean;
  readonly leafFlipV: boolean;
  /** Outermost group first. Empty for a shape that is not in a group. */
  readonly chain: readonly Frame[];
  /** `a:ln/@w` in points, when the probe is about stroke width. */
  readonly lineWidth?: number | undefined;
  readonly samples?: readonly Sample[] | undefined;
  readonly reference?: ReferenceRow | undefined;
  /** Other shapes in the innermost group. The analysis needs them to compute
   *  what the group actually contains, which is the answer to the fill question. */
  readonly siblings?: readonly SiblingSpec[] | undefined;
}

export interface TransformDeck {
  readonly deck: string;
  readonly hostile: boolean;
  /** One entry per slide: the children of `p:spTree` after `p:grpSpPr`. */
  readonly slides: readonly string[];
  readonly probes: readonly TransformProbe[];
  /**
   * A whole package, for the decks a one-master builder cannot express.
   *
   * The compositing probes need a master, a layout and a slide that disagree,
   * which `buildPptx` has no way to say. When this is present the builder uses
   * it and ignores `slides`.
   */
  readonly bytes?: (() => Uint8Array) | undefined;
}

/* -------------------------------------------------------------------------- */
/* colours                                                                    */
/* -------------------------------------------------------------------------- */

/** Literal `a:srgbClr`s throughout, so no reading depends on a theme. */
export const LEAF_COLOR = '1E88E5';
export const GROUP_COLOR = 'E53935';
export const GRADIENT_START = 'FF0000';
export const GRADIENT_END = '0000FF';
export const CONTROL_COLOR = '00A000';

const solid = (hex: string): string => `<a:solidFill><a:srgbClr val="${hex}"/></a:solidFill>`;
const GRP_FILL = '<a:grpFill/>';
const NO_LINE = '<a:ln><a:noFill/></a:ln>';

/**
 * A horizontal ramp across the whole shape, red at the left edge and blue at
 * the right.
 *
 * Two stops rather than the 33 the 2.7 ramp uses, because this probe reads
 * *position*, not colour accuracy: the question is whether a child shows the
 * slice of the group's ramp under it or a copy of the whole ramp, and at 8% and
 * 92% of the way across those two readings are red-versus-blue apart.
 */
const GRADIENT_FILL =
  '<a:gradFill flip="none" rotWithShape="1"><a:gsLst>' +
  `<a:gs pos="0"><a:srgbClr val="${GRADIENT_START}"/></a:gs>` +
  `<a:gs pos="100000"><a:srgbClr val="${GRADIENT_END}"/></a:gs>` +
  '</a:gsLst><a:lin ang="0" scaled="0"/></a:gradFill>';

/**
 * `pct50` rather than one of the line patterns.
 *
 * A `ltHorz` covers one pixel row in eight, so a sampled block is white seven
 * times out of eight and the reading says nothing. Half of `pct50`'s pixels
 * carry ink at every scale, so "is the pattern painted at all" is answerable
 * from a block of pixels without also solving for the tile phase.
 */
const PATTERN_FILL =
  `<a:pattFill prst="pct50"><a:fgClr><a:srgbClr val="${GRADIENT_START}"/></a:fgClr>` +
  `<a:bgClr><a:srgbClr val="FFFFFF"/></a:bgClr></a:pattFill>`;

/* -------------------------------------------------------------------------- */
/* markup                                                                     */
/* -------------------------------------------------------------------------- */

interface LeafOptions {
  readonly fill?: string | undefined;
  readonly line?: string | undefined;
  readonly prst?: string | undefined;
  /** Further shapes to place beside the leaf *inside* the innermost group. */
  readonly siblings?: readonly SiblingSpec[] | undefined;
}

/** A second shape in the same group, for the probes that need two readings. */
export interface SiblingSpec {
  readonly name: string;
  readonly rect: Rect;
  readonly fill: string;
}

let nextId = 100;

/** Wrap a leaf in its chain, outermost group last to be built. */
function nest(probe: TransformProbe, options: LeafOptions = {}): string {
  let markup = shapeXml({
    id: nextId++,
    name: probe.shape,
    x: emu(probe.leaf.x, probe.shape),
    y: emu(probe.leaf.y, probe.shape),
    cx: emu(probe.leaf.w, probe.shape),
    cy: emu(probe.leaf.h, probe.shape),
    fill: options.fill ?? solid(LEAF_COLOR),
    prst: options.prst ?? 'rect',
    rot: deg(probe.leafRot),
    flipH: probe.leafFlipH,
    flipV: probe.leafFlipV,
    line: options.line ?? NO_LINE,
  });

  for (const sibling of options.siblings ?? []) {
    markup += shapeXml({
      id: nextId++,
      name: sibling.name,
      x: emu(sibling.rect.x, sibling.name),
      y: emu(sibling.rect.y, sibling.name),
      cx: emu(sibling.rect.w, sibling.name),
      cy: emu(sibling.rect.h, sibling.name),
      fill: sibling.fill,
      line: NO_LINE,
    });
  }

  for (let i = probe.chain.length - 1; i >= 0; i--) {
    const frame = probe.chain[i]!;
    const name = `${probe.shape}-g${String(i)}`;
    markup = groupXml({
      id: nextId++,
      name,
      x: emu(frame.rect.x, name),
      y: emu(frame.rect.y, name),
      cx: emu(frame.rect.w, name),
      cy: emu(frame.rect.h, name),
      chX: emu(frame.child.x, name),
      chY: emu(frame.child.y, name),
      chCx: emu(frame.child.w, name),
      chCy: emu(frame.child.h, name),
      rot: deg(frame.rot ?? 0),
      flipH: frame.flipH,
      flipV: frame.flipV,
      fill: frame.fill,
      noChild: frame.noChildTransform,
      children: markup,
    });
  }
  return markup;
}

/* -------------------------------------------------------------------------- */
/* the catalogue                                                              */
/* -------------------------------------------------------------------------- */

interface Draft {
  readonly id: string;
  readonly kind: TransformProbe['kind'];
  readonly question: string;
  readonly leaf: Rect;
  readonly rot?: number | undefined;
  readonly flipH?: boolean | undefined;
  readonly flipV?: boolean | undefined;
  readonly chain: readonly Frame[];
  readonly leafFill?: string | undefined;
  readonly leafLine?: string | undefined;
  readonly leafPrst?: string | undefined;
  readonly siblings?: readonly SiblingSpec[] | undefined;
  readonly lineWidth?: number | undefined;
  readonly samples?: readonly Sample[] | undefined;
  readonly reference?: ReferenceRow | undefined;
  /** Extra markup on the same slide, drawn *after* - so, on top. */
  readonly extra?: string | undefined;
  /** Extra markup drawn *before*, so behind everything: reference swatches. */
  readonly before?: string | undefined;
}

function build(deck: string, hostile: boolean, drafts: readonly Draft[]): TransformDeck {
  const slides: string[] = [];
  const probes: TransformProbe[] = [];
  drafts.forEach((draft, i) => {
    const probe: TransformProbe = {
      id: draft.id,
      deck,
      slide: i + 1,
      kind: draft.kind,
      question: draft.question,
      shape: draft.id,
      leaf: draft.leaf,
      leafRot: draft.rot ?? 0,
      leafFlipH: draft.flipH ?? false,
      leafFlipV: draft.flipV ?? false,
      chain: draft.chain,
      lineWidth: draft.lineWidth,
      samples: draft.samples,
      reference: draft.reference,
      siblings: draft.siblings,
    };
    probes.push(probe);
    slides.push(
      (draft.before ?? '') +
        nest(probe, {
          fill: draft.leafFill,
          line: draft.leafLine,
          prst: draft.leafPrst,
          siblings: draft.siblings,
        }) +
        (draft.extra ?? ''),
    );
  });
  return { deck, hostile, slides, probes };
}

/** `off`/`ext` and `chOff`/`chExt` equal: the identity map a fresh group has. */
function identity(rect: Rect, extra: Partial<Frame> = {}): Frame {
  return { rect, child: rect, ...extra };
}

/* ---- the coordinate map -------------------------------------------------- */

const MAP: Draft[] = [
  {
    id: 'm-identity',
    kind: 'map',
    question: 'a group whose chOff/chExt equal its off/ext changes nothing',
    leaf: { x: 150, y: 120, w: 80, h: 40 },
    chain: [identity({ x: 100, y: 100, w: 400, h: 200 })],
  },
  {
    id: 'm-scale2',
    kind: 'map',
    question: 'ext twice chExt doubles the child',
    leaf: { x: 150, y: 120, w: 80, h: 40 },
    chain: [
      { rect: { x: 100, y: 100, w: 800, h: 400 }, child: { x: 100, y: 100, w: 400, h: 200 } },
    ],
  },
  {
    id: 'm-half',
    kind: 'map',
    question: 'ext half of chExt halves the child',
    leaf: { x: 150, y: 120, w: 80, h: 40 },
    chain: [
      { rect: { x: 100, y: 100, w: 200, h: 100 }, child: { x: 100, y: 100, w: 400, h: 200 } },
    ],
  },
  {
    id: 'm-nonuniform',
    kind: 'map',
    question: 'the two axes scale independently',
    leaf: { x: 150, y: 120, w: 80, h: 40 },
    chain: [
      { rect: { x: 100, y: 100, w: 800, h: 100 }, child: { x: 100, y: 100, w: 400, h: 200 } },
    ],
  },
  {
    id: 'm-choff-zero',
    kind: 'map',
    question: 'chOff of zero: the children are written relative to the origin',
    leaf: { x: 50, y: 20, w: 80, h: 40 },
    chain: [{ rect: { x: 100, y: 100, w: 400, h: 200 }, child: { x: 0, y: 0, w: 400, h: 200 } }],
  },
  {
    id: 'm-choff-neg',
    kind: 'map',
    question: 'a negative chOff',
    leaf: { x: -150, y: -80, w: 80, h: 40 },
    chain: [
      { rect: { x: 100, y: 100, w: 400, h: 200 }, child: { x: -200, y: -100, w: 400, h: 200 } },
    ],
  },
  {
    id: 'm-choff-far',
    kind: 'map',
    question: 'a child space nowhere near the slide, scaled a tenth',
    leaf: { x: 10500, y: 20200, w: 800, h: 400 },
    chain: [
      {
        rect: { x: 100, y: 100, w: 400, h: 200 },
        child: { x: 10000, y: 20000, w: 4000, h: 2000 },
      },
    ],
  },
  {
    id: 'm-nested',
    kind: 'map',
    question: 'two groups, and the scales multiply',
    leaf: { x: 200, y: 130, w: 50, h: 15 },
    chain: [
      { rect: { x: 100, y: 100, w: 800, h: 400 }, child: { x: 100, y: 100, w: 400, h: 200 } },
      { rect: { x: 150, y: 120, w: 200, h: 60 }, child: { x: 150, y: 120, w: 100, h: 30 } },
    ],
  },
  {
    id: 'm-three',
    kind: 'map',
    question: 'three levels of group',
    leaf: { x: 40, y: 40, w: 10, h: 10 },
    chain: [
      { rect: { x: 100, y: 100, w: 400, h: 400 }, child: { x: 0, y: 0, w: 200, h: 200 } },
      { rect: { x: 20, y: 20, w: 100, h: 100 }, child: { x: 0, y: 0, w: 50, h: 50 } },
      { rect: { x: 10, y: 10, w: 60, h: 60 }, child: { x: 20, y: 20, w: 30, h: 30 } },
    ],
  },
];

/* ---- orientation --------------------------------------------------------- */

/** Every orientation probe uses this frame, so only the angles vary. */
const TURN_FRAME: Rect = { x: 300, y: 150, w: 300, h: 200 };
const TURN_LEAF: Rect = { x: 350, y: 200, w: 100, h: 60 };

function turn(
  id: string,
  question: string,
  frame: Partial<Frame>,
  leaf: { rot?: number; flipH?: boolean; flipV?: boolean },
): Draft {
  return {
    id,
    kind: 'turn',
    question,
    leaf: TURN_LEAF,
    rot: leaf.rot,
    flipH: leaf.flipH,
    flipV: leaf.flipV,
    chain: [identity(TURN_FRAME, frame)],
    // A right triangle rather than a rectangle. The reading is the object
    // model's, which a rectangle would answer just as well, but a rectangle
    // cannot tell a mirror from an identity in the exported bitmap - and the
    // bitmap is the only check the object model has.
    leafPrst: 'rtTriangle',
  };
}

const TURN: Draft[] = [
  turn('t-grot', 'a group rotation, with an unturned child', { rot: 30 }, {}),
  turn('t-grot-crot', 'a group rotation and a child rotation', { rot: 30 }, { rot: 40 }),
  turn('t-gfliph-crot', 'a flipped group holding a rotated child', { flipH: true }, { rot: 40 }),
  turn('t-gfliph-cfliph', 'two flips on the same axis', { flipH: true }, { flipH: true }),
  turn('t-gflipv-crot', 'a vertically flipped group, rotated child', { flipV: true }, { rot: 40 }),
  turn('t-gfliph-cflipv', 'one flip on each axis', { flipH: true }, { flipV: true }),
  turn(
    't-gboth-crot',
    'a group flipped on both axes, which is a half turn',
    { flipH: true, flipV: true },
    { rot: 40 },
  ),
  turn(
    't-grot-gflip-crot-cflip',
    'everything at once',
    { rot: 30, flipH: true },
    { rot: 40, flipH: true },
  ),
  {
    id: 't-nested-rot',
    kind: 'turn',
    question: 'three rotations down a chain',
    leaf: { x: 350, y: 200, w: 100, h: 60 },
    rot: 40,
    chain: [
      identity({ x: 250, y: 100, w: 400, h: 300 }, { rot: 20 }),
      identity({ x: 300, y: 150, w: 300, h: 200 }, { rot: 30 }),
    ],
  },
  {
    id: 't-nested-flip',
    kind: 'turn',
    question: 'a flip above two rotations',
    leaf: { x: 350, y: 200, w: 100, h: 60 },
    rot: 40,
    chain: [
      identity({ x: 250, y: 100, w: 400, h: 300 }, { flipH: true }),
      identity({ x: 300, y: 150, w: 300, h: 200 }, { rot: 30 }),
    ],
  },
  {
    id: 't-scale-rot',
    kind: 'turn',
    question: 'a rotation about the centre of a rectangle the group scaled',
    leaf: { x: 150, y: 120, w: 50, h: 25 },
    chain: [
      {
        rect: { x: 100, y: 100, w: 400, h: 200 },
        child: { x: 100, y: 100, w: 200, h: 100 },
        rot: 90,
      },
    ],
  },
  {
    id: 't-shear',
    kind: 'turn',
    question: 'a rotated child inside a non-uniformly scaled group - where does the shear go?',
    leaf: { x: 200, y: 150, w: 100, h: 100 },
    rot: 45,
    chain: [
      { rect: { x: 100, y: 100, w: 800, h: 200 }, child: { x: 100, y: 100, w: 400, h: 200 } },
    ],
    samples: [
      // The 45-degree square's corners, under the two readings. A sheared
      // parallelogram reaches further in x than a rotated rectangle does.
      { id: 't-shear-left', x: 265, y: 200, asks: 'inside a sheared parallelogram only' },
      { id: 't-shear-centre', x: 400, y: 200, asks: 'inside under either reading' },
      { id: 't-shear-right', x: 535, y: 200, asks: 'inside a sheared parallelogram only' },
      // The rotated rectangle's topmost corner is at (364.6, 93.9) and its
      // span at y=110 is x 348.6..380.7. The parallelogram never rises above
      // y=129.3, so this point is inside one reading and outside the other.
      { id: 't-shear-top', x: 365, y: 110, asks: 'inside a rotated rectangle only' },
    ],
  },
  {
    id: 't-rot-negative',
    kind: 'turn',
    question: 'a negative rot attribute',
    leaf: TURN_LEAF,
    rot: -30,
    chain: [identity(TURN_FRAME)],
  },
  {
    id: 't-rot-over',
    kind: 'turn',
    question: 'a rot past a full turn',
    leaf: TURN_LEAF,
    rot: 390,
    chain: [identity(TURN_FRAME)],
  },
  {
    id: 't-width',
    kind: 'width',
    question: 'does a group scale the stroke of the shapes inside it?',
    leaf: { x: 150, y: 120, w: 100, h: 50 },
    lineWidth: 4,
    leafLine: `<a:ln w="${String(4 * EMU_PER_POINT)}"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln>`,
    chain: [
      { rect: { x: 100, y: 100, w: 800, h: 400 }, child: { x: 100, y: 100, w: 400, h: 200 } },
    ],
    samples: [
      // The mapped rectangle is (200,140)-(400,240) and the stroke is centred
      // on its edge, so an unscaled 4pt stroke covers y 138..142 and a scaled
      // 8pt one covers 136..144. One pixel row tells them apart.
      { id: 't-width-scaled', x: 300, y: 137, asks: 'black only if the group scaled the stroke' },
      { id: 't-width-stroke', x: 300, y: 141, asks: 'on the stroke under either reading' },
      { id: 't-width-fill', x: 300, y: 150, asks: 'the interior', expect: LEAF_COLOR },
    ],
  },
];

/* ---- a:grpFill ----------------------------------------------------------- */

/** The group every fill probe uses: 600 x 300 points at (100, 100). */
const FILL_FRAME: Rect = { x: 100, y: 100, w: 600, h: 300 };

function fillProbe(
  id: string,
  question: string,
  groupFill: string | undefined,
  leafFill: string,
  samples: readonly Sample[],
  siblings?: readonly SiblingSpec[],
  before?: string,
  ref?: ReferenceRow,
): Draft {
  return {
    id,
    kind: 'fill',
    question,
    leaf: { x: 150, y: 180, w: 100, h: 140 },
    chain: [identity(FILL_FRAME, { fill: groupFill })],
    leafFill,
    samples,
    siblings,
    before,
    reference: ref,
  };
}

/**
 * The same gradient on an ordinary shape covering exactly the group rectangle,
 * drawn behind everything.
 *
 * The first run of this experiment compared a `grpFill` child against the
 * group's own ground and found white, because - measured, and the surprise of
 * the deck - **a group's fill is not painted**. It exists only to be asked for.
 * So the reference the slice hypothesis needs has to be a shape.
 */
const GRADIENT_REFERENCE = (): string =>
  shapeXml({
    id: nextId++,
    name: 'f-gradient-reference',
    x: FILL_FRAME.x * EMU_PER_POINT,
    y: FILL_FRAME.y * EMU_PER_POINT,
    cx: FILL_FRAME.w * EMU_PER_POINT,
    cy: FILL_FRAME.h * EMU_PER_POINT,
    fill: GRADIENT_FILL,
    line: NO_LINE,
  });

/** A second child near the group's right edge, inside the same group. */
const RIGHT_LEAF: SiblingSpec = {
  name: 'f-gradient-right',
  rect: { x: 550, y: 180, w: 100, h: 140 },
  fill: GRP_FILL,
};

const FILLS: Draft[] = [
  {
    // The control. If this reading is wrong nothing else on the deck means
    // anything, and C's own write-up says to read the controls first.
    id: 'f-control',
    kind: 'fill',
    question: 'an ordinary solid fill inside a group, so the sampling is proved',
    leaf: { x: 150, y: 180, w: 100, h: 140 },
    chain: [identity(FILL_FRAME)],
    leafFill: solid(CONTROL_COLOR),
    samples: [{ id: 'f-control-centre', x: 200, y: 250, asks: 'the leaf', expect: CONTROL_COLOR }],
  },
  fillProbe('f-solid', 'a:grpFill against a solid group fill', solid(GROUP_COLOR), GRP_FILL, [
    { id: 'f-solid-leaf', x: 200, y: 250, asks: 'the grpFill child', expect: GROUP_COLOR },
    { id: 'f-solid-bare', x: 400, y: 250, asks: 'the group rectangle where no child is' },
  ]),
  fillProbe(
    'f-gradient',
    'a slice of the group gradient, or a copy of it?',
    GRADIENT_FILL,
    GRP_FILL,
    [
      // The group spans x 100..700. The left child's centre is 8.3% along it
      // and the right child's 91.7%, so a slice reads red and blue while a copy
      // reads the same middling purple twice.
      { id: 'f-grad-left', x: 200, y: 250, asks: 'the left grpFill child' },
      { id: 'f-grad-right', x: 600, y: 250, asks: 'the right grpFill child' },
      // Directly above each child, on the group's own ground, so a slice
      // reading is the same colour at the same x and a copy reading is not.
      // The reference shape behind everything, carrying the identical gradient
      // over the identical rectangle. A slice reads the same colour at the same
      // x; a copy does not.
      { id: 'f-grad-ref-left', x: 200, y: 130, asks: 'the reference ramp at the left child x' },
      { id: 'f-grad-ref-right', x: 600, y: 130, asks: 'the reference ramp at the right child x' },
    ],
    [RIGHT_LEAF],
    GRADIENT_REFERENCE(),
    { y: 130, x0: FILL_FRAME.x, x1: FILL_FRAME.x + FILL_FRAME.w },
  ),
  fillProbe(
    'f-pattern',
    'does a:grpFill reach a pattern fill?',
    PATTERN_FILL,
    GRP_FILL,
    [
      { id: 'f-patt-leaf', x: 200, y: 250, asks: 'the grpFill child' },
      { id: 'f-patt-group', x: 400, y: 130, asks: 'the group fill itself' },
      { id: 'f-patt-ref', x: 400, y: 390, asks: 'the same pattern on an ordinary shape' },
    ],
    undefined,
    // Below the children, so it is not covered, and inside the group rectangle,
    // so its presence also says whether that rectangle paints anything itself.
    reference({ x: 100, y: 380, w: 600, h: 20 }, PATTERN_FILL),
  ),
  fillProbe('f-nofill', 'a:grpFill over an explicit a:noFill group', '<a:noFill/>', GRP_FILL, [
    { id: 'f-nofill-leaf', x: 200, y: 250, asks: 'the grpFill child', expect: 'FFFFFF' },
  ]),
  fillProbe('f-silent', 'a:grpFill over a group that states no fill at all', undefined, GRP_FILL, [
    { id: 'f-silent-leaf', x: 200, y: 250, asks: 'the grpFill child', expect: 'FFFFFF' },
  ]),
  {
    id: 'f-toplevel',
    kind: 'fill',
    question: 'a:grpFill on a shape that is in no group',
    leaf: { x: 150, y: 180, w: 100, h: 140 },
    chain: [],
    leafFill: GRP_FILL,
    samples: [{ id: 'f-top-leaf', x: 200, y: 250, asks: 'the shape', expect: 'FFFFFF' }],
  },
  {
    id: 'f-nested',
    kind: 'fill',
    question: 'does a:grpFill walk past an inner group that also says a:grpFill?',
    leaf: { x: 150, y: 180, w: 100, h: 140 },
    chain: [
      identity(FILL_FRAME, { fill: GRADIENT_FILL }),
      identity({ x: 120, y: 150, w: 300, h: 200 }, { fill: GRP_FILL }),
    ],
    leafFill: GRP_FILL,
    samples: [{ id: 'f-nested-leaf', x: 200, y: 250, asks: 'the doubly-nested grpFill child' }],
  },
  {
    id: 'f-group-alone',
    kind: 'fill',
    question: 'is a group fill painted at all when nothing asks for it?',
    // The leaf is parked outside the group's rectangle so the sample under test
    // is the group's own ground and nothing else.
    leaf: { x: 750, y: 420, w: 60, h: 60 },
    chain: [identity(FILL_FRAME, { fill: solid(GROUP_COLOR) })],
    leafFill: solid(CONTROL_COLOR),
    samples: [
      { id: 'f-alone-inside', x: 400, y: 250, asks: 'inside the group rectangle, no child there' },
      { id: 'f-alone-leaf', x: 780, y: 450, asks: 'the leaf', expect: CONTROL_COLOR },
    ],
  },
];

/* ---- over what rectangle is a group's fill laid out? --------------------- */

/**
 * The first fill deck established that `a:grpFill` shows a **slice** of the
 * group's gradient rather than a copy of it. It also showed that the slice does
 * not run across the group's declared rectangle: a child centred at x=200 in a
 * group spanning x 100..700 read the colour the ramp has one tenth of the way
 * along, not one sixth. One tenth of the way along *what* is this deck.
 *
 * Every slide carries a reference shape with the identical gradient over a
 * known rectangle, drawn behind everything, so a sampled colour turns back into
 * a position without this file knowing anything about the blend curve.
 */
const VERT_GRADIENT =
  '<a:gradFill flip="none" rotWithShape="1"><a:gsLst>' +
  `<a:gs pos="0"><a:srgbClr val="${GRADIENT_START}"/></a:gs>` +
  `<a:gs pos="100000"><a:srgbClr val="${GRADIENT_END}"/></a:gs>` +
  '</a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill>';

function reference(rect: Rect, fill: string): string {
  return shapeXml({
    id: nextId++,
    name: 'reference',
    x: rect.x * EMU_PER_POINT,
    y: rect.y * EMU_PER_POINT,
    cx: rect.w * EMU_PER_POINT,
    cy: rect.h * EMU_PER_POINT,
    fill,
    line: NO_LINE,
  });
}

/** The reference ramp runs the full width of the group's declared rectangle. */
const REF_ROW: ReferenceRow = { y: 130, x0: FILL_FRAME.x, x1: FILL_FRAME.x + FILL_FRAME.w };

const EXTENT: Draft[] = [
  {
    id: 'g-one',
    kind: 'fill',
    question: 'one grpFill child: does it see the whole ramp or a sliver of it?',
    leaf: { x: 300, y: 180, w: 100, h: 140 },
    chain: [identity(FILL_FRAME, { fill: GRADIENT_FILL })],
    leafFill: GRP_FILL,
    before: reference(FILL_FRAME, GRADIENT_FILL),
    reference: REF_ROW,
    samples: [
      { id: 'g-one-a', x: 310, y: 250, asks: 'near the left edge of the only child' },
      { id: 'g-one-b', x: 350, y: 250, asks: 'the middle of the only child' },
      { id: 'g-one-c', x: 390, y: 250, asks: 'near the right edge of the only child' },
    ],
  },
  {
    id: 'g-solid-sibling',
    kind: 'fill',
    question: 'does a sibling that does not use grpFill still set the extent?',
    leaf: { x: 300, y: 180, w: 100, h: 140 },
    chain: [identity(FILL_FRAME, { fill: GRADIENT_FILL })],
    leafFill: GRP_FILL,
    siblings: [
      {
        name: 'g-solid-sibling-b',
        rect: { x: 600, y: 180, w: 100, h: 140 },
        fill: solid(CONTROL_COLOR),
      },
    ],
    before: reference(FILL_FRAME, GRADIENT_FILL),
    reference: REF_ROW,
    samples: [
      { id: 'g-sib-a', x: 350, y: 250, asks: 'the middle of the grpFill child' },
      { id: 'g-sib-b', x: 650, y: 250, asks: 'the solid sibling', expect: CONTROL_COLOR },
    ],
  },
  {
    id: 'g-outside',
    kind: 'fill',
    question: 'a sibling outside the group rectangle: is the extent clipped to it?',
    leaf: { x: 300, y: 180, w: 100, h: 140 },
    chain: [identity(FILL_FRAME, { fill: GRADIENT_FILL })],
    leafFill: GRP_FILL,
    siblings: [
      { name: 'g-outside-b', rect: { x: 840, y: 180, w: 100, h: 140 }, fill: solid(CONTROL_COLOR) },
    ],
    before: reference(FILL_FRAME, GRADIENT_FILL),
    reference: REF_ROW,
    samples: [{ id: 'g-out-a', x: 350, y: 250, asks: 'the middle of the grpFill child' }],
  },
  {
    id: 'g-deep',
    kind: 'fill',
    question: 'three groups deep, each saying grpFill',
    leaf: { x: 300, y: 180, w: 100, h: 140 },
    chain: [
      identity(FILL_FRAME, { fill: GRADIENT_FILL }),
      identity({ x: 200, y: 150, w: 400, h: 220 }, { fill: GRP_FILL }),
      identity({ x: 250, y: 160, w: 300, h: 200 }, { fill: GRP_FILL }),
    ],
    leafFill: GRP_FILL,
    before: reference(FILL_FRAME, GRADIENT_FILL),
    reference: REF_ROW,
    samples: [
      { id: 'g-deep-a', x: 310, y: 250, asks: 'near the left edge of the leaf' },
      { id: 'g-deep-b', x: 350, y: 250, asks: 'the middle of the leaf' },
      { id: 'g-deep-c', x: 390, y: 250, asks: 'near the right edge of the leaf' },
    ],
  },
  {
    id: 'g-vert',
    kind: 'fill',
    question: 'the same question on the other axis',
    leaf: { x: 300, y: 130, w: 100, h: 60 },
    chain: [identity(FILL_FRAME, { fill: VERT_GRADIENT })],
    leafFill: GRP_FILL,
    siblings: [{ name: 'g-vert-b', rect: { x: 300, y: 320, w: 100, h: 60 }, fill: GRP_FILL }],
    before: reference(FILL_FRAME, VERT_GRADIENT),
    reference: { y: 150, x0: FILL_FRAME.y, x1: FILL_FRAME.y + FILL_FRAME.h, vertical: true },
    samples: [
      { id: 'g-vert-top', x: 350, y: 160, asks: 'the middle of the upper child' },
      { id: 'g-vert-bot', x: 350, y: 350, asks: 'the middle of the lower child' },
    ],
  },
];

/* ---- the shear, at every angle ------------------------------------------- */

/**
 * A rotated child inside a group that scales the two axes differently.
 *
 * The image of a rotated rectangle under `diag(sx, sy)` is a parallelogram, and
 * `a:xfrm` has nowhere to put one - there is an offset, an extent, an angle and
 * two mirrors, and no shear. So PowerPoint must approximate, and the first run
 * of C6 found that at 45 degrees with `sx=2, sy=1` it reports a 100 x 200
 * rectangle where scaling the rectangle would give 200 x 100: the two extents
 * come out swapped. One angle cannot say why. This deck asks eleven.
 */
function shearProbe(id: string, rot: number, ext: Rect, child: Rect): Draft {
  return {
    id,
    kind: 'turn',
    question: `a child at ${String(rot)} degrees in a group scaling x by ${String(
      ext.w / child.w,
    )} and y by ${String(ext.h / child.h)}`,
    leaf: { x: 200, y: 150, w: 100, h: 100 },
    rot,
    chain: [{ rect: ext, child }],
    leafPrst: 'rtTriangle',
  };
}

/** x doubled, y untouched. */
const WIDE: Rect = { x: 100, y: 100, w: 800, h: 200 };
const WIDE_CHILD: Rect = { x: 100, y: 100, w: 400, h: 200 };
/** y doubled, x untouched. */
const TALL: Rect = { x: 100, y: 100, w: 400, h: 400 };
const TALL_CHILD: Rect = { x: 100, y: 100, w: 400, h: 200 };

const SHEAR: Draft[] = [
  ...[0, 15, 30, 44, 45, 46, 60, 75, 90, 135, 180].map((a) =>
    shearProbe(`s-wide-${String(a)}`, a, WIDE, WIDE_CHILD),
  ),
  ...[30, 45, 60].map((a) => shearProbe(`s-tall-${String(a)}`, a, TALL, TALL_CHILD)),
  // Neither axis left alone, so a rule stated as "swap the factors" and one
  // stated as "use the other axis's factor" stop agreeing.
  shearProbe(
    's-both-30',
    30,
    { x: 100, y: 100, w: 800, h: 600 },
    { x: 100, y: 100, w: 400, h: 200 },
  ),
  shearProbe(
    's-both-60',
    60,
    { x: 100, y: 100, w: 800, h: 600 },
    { x: 100, y: 100, w: 400, h: 200 },
  ),
  // A mirror is axis-aligned, so a flipped child in the same group should not
  // need approximating at all.
  {
    id: 's-flip',
    kind: 'turn',
    question: 'a flipped, unrotated child in a non-uniformly scaled group',
    leaf: { x: 200, y: 150, w: 100, h: 100 },
    flipH: true,
    chain: [{ rect: WIDE, child: WIDE_CHILD }],
    leafPrst: 'rtTriangle',
  },
];

/* ---- what a slide shows that it did not write ---------------------------- */

/**
 * A slide draws its master's shapes and its layout's, and nobody writes down
 * which ones or in what order.
 *
 * The object model is no help here: `Slide.Shapes` lists only what the slide
 * part itself contains, and the inherited furniture is on the screen without
 * being in the collection. So this is a bitmap question - three squares in
 * three colours, one on each sheet, and a sample in each.
 *
 * `@showMasterSp` is the second half of it. It defaults to true, it exists on
 * both a slide and a layout, and its name says master while the thing it
 * governs may well be the layout's. Sub-phase 7.9 sells it as a "hide layout
 * graphics" toggle, so being wrong about which sheet it reaches is not a
 * detail.
 */
const MASTER_COLOR = 'D32F2F';
const LAYOUT_COLOR = '388E3C';
const SLIDE_COLOR = '1976D2';
const MASTER_PH_COLOR = '8E24AA';
const LAYOUT_PH_COLOR = 'F9A825';

const INHERIT_BOXES = {
  master: { x: 60, y: 60, w: 120, h: 120 },
  layout: { x: 220, y: 60, w: 120, h: 120 },
  slide: { x: 380, y: 60, w: 120, h: 120 },
  masterPh: { x: 60, y: 240, w: 120, h: 120 },
  layoutPh: { x: 220, y: 240, w: 120, h: 120 },
} as const;

function inheritSample(id: string, key: keyof typeof INHERIT_BOXES, asks: string): Sample {
  const box = INHERIT_BOXES[key];
  return { id, x: box.x + box.w / 2, y: box.y + box.h / 2, asks };
}

function inheritDeck(
  name: string,
  slideShowMasterSp: boolean | undefined,
  layoutShowMasterSp: boolean | undefined,
): TransformDeck {
  const square = (
    id: number,
    label: string,
    key: keyof typeof INHERIT_BOXES,
    hex: string,
    ph?: string,
  ) =>
    sheetShape({
      id,
      name: label,
      rect: INHERIT_BOXES[key],
      fill: `<a:solidFill><a:srgbClr val="${hex}"/></a:solidFill>`,
      line: '<a:ln><a:noFill/></a:ln>',
      ph,
    });

  const probe: TransformProbe = {
    id: name,
    deck: name,
    slide: 1,
    kind: 'fill',
    question: `what a slide shows with slide showMasterSp=${String(slideShowMasterSp ?? true)} and layout showMasterSp=${String(layoutShowMasterSp ?? true)}`,
    shape: 'slide-square',
    leaf: INHERIT_BOXES.slide,
    leafRot: 0,
    leafFlipH: false,
    leafFlipV: false,
    chain: [],
    samples: [
      inheritSample(`${name}-master`, 'master', "the master's own shape"),
      inheritSample(`${name}-layout`, 'layout', "the layout's own shape"),
      inheritSample(`${name}-slide`, 'slide', "the slide's own shape"),
      inheritSample(
        `${name}-master-ph`,
        'masterPh',
        'a master placeholder the slide never matched',
      ),
      inheritSample(
        `${name}-layout-ph`,
        'layoutPh',
        'a layout placeholder the slide never matched',
      ),
    ],
  };

  return {
    deck: name,
    hostile: false,
    slides: [],
    probes: [probe],
    bytes: () =>
      buildSheetPackage({
        themes: [{ scheme: SCHEME_ONE }],
        masters: [
          {
            theme: 0,
            bg: '<p:bg><p:bgPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>',
            shapes: [
              square(10, 'master-square', 'master', MASTER_COLOR),
              square(11, 'master-ph', 'masterPh', MASTER_PH_COLOR, 'type="body" idx="1"'),
            ],
          },
        ],
        layouts: [
          {
            master: 0,
            ...(layoutShowMasterSp === undefined ? {} : { showMasterSp: layoutShowMasterSp }),
            shapes: [
              square(20, 'layout-square', 'layout', LAYOUT_COLOR),
              square(21, 'layout-ph', 'layoutPh', LAYOUT_PH_COLOR, 'type="body" idx="7"'),
            ],
          },
        ],
        slides: [
          {
            layout: 0,
            ...(slideShowMasterSp === undefined ? {} : { showMasterSp: slideShowMasterSp }),
            shapes: [square(30, 'slide-square', 'slide', SLIDE_COLOR)],
          },
        ],
      }),
  };
}

/* ---- one hostile package per question ------------------------------------ */

function hostile(id: string, question: string, frame: Frame, leaf: Rect): TransformDeck {
  return build(id, true, [{ id, kind: 'map', question, leaf, chain: [frame] }]);
}

/* -------------------------------------------------------------------------- */

export function transformDecks(): readonly TransformDeck[] {
  nextId = 100;
  const decks: TransformDeck[] = [
    build('map', false, MAP),
    build('turn', false, TURN),
    build('fill', false, FILLS),
    build('extent', false, EXTENT),
    build('shear', false, SHEAR),
    inheritDeck('inherit-default', undefined, undefined),
    inheritDeck('inherit-slide-off', false, undefined),
    inheritDeck('inherit-layout-off', undefined, false),
    inheritDeck('inherit-both-off', false, false),
    hostile(
      'h-chext-zero-x',
      'a chExt of zero on one axis: no scaling, or a collapse?',
      { rect: { x: 100, y: 100, w: 400, h: 200 }, child: { x: 100, y: 100, w: 0, h: 200 } },
      { x: 150, y: 120, w: 80, h: 40 },
    ),
    hostile(
      'h-chext-zero-both',
      'a chExt of zero on both axes',
      { rect: { x: 100, y: 100, w: 400, h: 200 }, child: { x: 100, y: 100, w: 0, h: 0 } },
      { x: 150, y: 120, w: 80, h: 40 },
    ),
    hostile(
      'h-ext-zero-x',
      'an ext of zero: a group with no width',
      { rect: { x: 100, y: 100, w: 0, h: 200 }, child: { x: 100, y: 100, w: 400, h: 200 } },
      { x: 150, y: 120, w: 80, h: 40 },
    ),
    hostile(
      'h-chext-tiny',
      'a chExt of four points against an ext of four hundred',
      { rect: { x: 100, y: 100, w: 400, h: 200 }, child: { x: 100, y: 100, w: 4, h: 2 } },
      { x: 102, y: 101, w: 1, h: 1 },
    ),
    hostile(
      'h-chext-negative',
      'a negative chExt',
      { rect: { x: 100, y: 100, w: 400, h: 200 }, child: { x: 100, y: 100, w: -400, h: 200 } },
      { x: 150, y: 120, w: 80, h: 40 },
    ),
    hostile(
      'h-no-chxfrm',
      'a group whose a:xfrm has off and ext but no chOff or chExt',
      {
        rect: { x: 100, y: 100, w: 400, h: 200 },
        // What the probe *claims*, which is exactly what the file does not say.
        child: { x: 100, y: 100, w: 400, h: 200 },
        noChildTransform: true,
      },
      { x: 150, y: 120, w: 80, h: 40 },
    ),
  ];

  return decks;
}

export const ALL_SAMPLES = (): readonly Sample[] =>
  transformDecks().flatMap((d) => d.probes.flatMap((p) => p.samples ?? []));
