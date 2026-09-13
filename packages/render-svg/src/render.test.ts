/**
 * The renderer, against what PowerPoint actually did.
 *
 * As in 2.9, the fixture is **re-derived** rather than compared against a
 * summary of itself. Every C6 probe carries the whole group chain it was built
 * from and the frame PowerPoint reported; these tests feed the chain through
 * this package's own `placeChild` and check the prediction. A transform that
 * drifts fails on 46 measurements, not on one sentence somebody wrote about
 * them.
 *
 * The three rules that are easiest to get plausibly wrong each have their own
 * adversarial check, computing what the *wrong* model would say and asserting
 * that the reading refutes it. A test that only confirms the right answer
 * cannot tell a correct renderer from one that has never been exercised.
 */

import { parseSheet, parseTheme, type Sheet, type Shape, type Xfrm } from '@pptx-studio/model';
import { effectFilter, type BlipEffect, type Color } from '@pptx-studio/paint';
import { parseXmlString } from '@pptx-studio/xml';
import { describe, expect, it } from 'vitest';

import fixture from '../../../corpus/ground-truth/transforms.json' with { type: 'json' };
import pictures from '../../../corpus/ground-truth/pictures.json' with { type: 'json' };
import snap from '../../../corpus/ground-truth/snap.json' with { type: 'json' };
import zoom from '../../../corpus/ground-truth/zoom.json' with { type: 'json' };

import { layoutSheet, layoutSlide, inheritedSheets, flatten, type Placed } from './layout.js';
import { RenderError } from './errors.js';
import { blipPaint } from './image/blip.js';
import { Defs } from './paint.js';
import { num, serializeSvg } from './node.js';
import { renderSlide, slideNode } from './slide.js';
import { inverseFramePoint } from './transform.js';
import {
  UNIT_CHILD_SPACE,
  childSpace,
  composeTurn,
  frameOf,
  frameTransform,
  placeChild,
  swapsExtents,
  type ChildSpace,
  type Frame,
} from './transform.js';

/* -------------------------------------------------------------------------- */
/* the fixture, in this package's own terms                                   */
/* -------------------------------------------------------------------------- */

interface FixtureRect {
  readonly x: number;
  readonly y: number;
  readonly w: number;
  readonly h: number;
}

interface FixtureFrame {
  readonly rect: FixtureRect;
  readonly child: FixtureRect;
  readonly rot?: number;
  readonly flipH?: boolean;
  readonly flipV?: boolean;
  readonly noChildTransform?: boolean;
}

const DEGREE = 60000;

/**
 * The probes are in points and this package works in EMU.
 *
 * Everything in the transform is linear and every scale is a ratio, so the two
 * agree exactly at any unit - which is worth using rather than multiplying
 * through, because a unit conversion inside a test is one more thing that can be
 * wrong in the same direction as the code it checks.
 */
function toXfrm(rect: FixtureRect, turn: Partial<FixtureFrame>, child: FixtureRect | null): Xfrm {
  return {
    x: rect.x,
    y: rect.y,
    cx: rect.w,
    cy: rect.h,
    rot: Math.round((turn.rot ?? 0) * DEGREE),
    flipH: turn.flipH ?? false,
    flipV: turn.flipV ?? false,
    child: child === null ? null : { x: child.x, y: child.y, cx: child.w, cy: child.h },
  };
}

const ROOT: Frame = { x: 0, y: 0, cx: 0, cy: 0, rot: 0, flipH: false, flipV: false };

/** Walk a probe's chain the way the renderer walks a group tree: downwards. */
function composeProbe(probe: {
  chain: readonly FixtureFrame[];
  leaf: FixtureRect;
  leafRot: number;
  leafFlipH: boolean;
  leafFlipV: boolean;
}): Frame {
  let frame = ROOT;
  let space: ChildSpace = UNIT_CHILD_SPACE;
  for (const link of probe.chain) {
    const xfrm = toXfrm(link.rect, link, link.noChildTransform === true ? null : link.child);
    frame = placeChild(frame, space, frameOf(xfrm));
    space = childSpace(frame, xfrm);
  }
  const leaf = toXfrm(
    probe.leaf,
    { rot: probe.leafRot, flipH: probe.leafFlipH, flipV: probe.leafFlipV },
    null,
  );
  return placeChild(frame, space, frameOf(leaf));
}

const PROBES = fixture.probes as unknown as readonly (Parameters<typeof composeProbe>[0] & {
  id: string;
  deck: string;
  kind: string;
  fromRepairedDeck: boolean;
  reading: {
    frame: { x: number; y: number; w: number; h: number };
    rotation: number;
    flipH: boolean;
    flipV: boolean;
    lineWeight: number | null;
  };
})[];

/** Everything but the shear deck, whose extents are a separate finding. */
const PLACED = PROBES.filter(
  (p) => !p.fromRepairedDeck && p.deck !== 'shear' && p.id !== 't-shear',
);

function close(a: number, b: number, tol = 0.01): boolean {
  return Math.abs(a - b) < tol;
}

/* -------------------------------------------------------------------------- */

describe('the child coordinate map, re-derived', () => {
  it('places every probe where PowerPoint put it', () => {
    const misses: string[] = [];
    for (const probe of PLACED) {
      const frame = composeProbe(probe);
      const want = probe.reading.frame;
      if (
        !close(frame.x, want.x) ||
        !close(frame.y, want.y) ||
        !close(frame.cx, want.w) ||
        !close(frame.cy, want.h)
      ) {
        misses.push(
          `${probe.id}: got ${String(frame.x)},${String(frame.y)},${String(frame.cx)},${String(frame.cy)}` +
            ` want ${String(want.x)},${String(want.y)},${String(want.w)},${String(want.h)}`,
        );
      }
    }
    expect(misses).toEqual([]);
    // The fixture's own scoring agrees, and says how many cases that is.
    expect(fixture.findings.childMap.winner).toBe('reference');
    expect(PLACED.length).toBeGreaterThanOrEqual(46);
  });

  it('refutes the model that drops the scale factor, which is right three times in four', () => {
    let wrong = 0;
    for (const probe of PLACED) {
      // The classic group bug: honour `chOff` and ignore `ext/chExt`.
      let frame = ROOT;
      for (const link of probe.chain) {
        const xfrm = toXfrm(link.rect, link, link.child);
        frame = placeChild(frame, UNIT_CHILD_SPACE, frameOf(xfrm));
      }
      const leaf = toXfrm(
        probe.leaf,
        { rot: probe.leafRot, flipH: probe.leafFlipH, flipV: probe.leafFlipV },
        null,
      );
      const got = placeChild(frame, UNIT_CHILD_SPACE, frameOf(leaf));
      if (!close(got.x, probe.reading.frame.x) || !close(got.cx, probe.reading.frame.w)) wrong += 1;
    }
    expect(wrong).toBeGreaterThan(0);
    const scored = fixture.findings.childMap.scores.find((s) => s.model === 'noScale');
    expect(scored?.right).toBeLessThan(scored?.of ?? 0);
  });

  it('reads a chExt of zero as no scaling, and an ext of zero as zero', () => {
    const zero = PROBES.find((p) => p.id === 'h-chext-zero-x');
    expect(zero).toBeDefined();
    expect(composeProbe(zero!).cx).toBeCloseTo(zero!.reading.frame.w, 6);

    const collapsed = PROBES.find((p) => p.id === 'h-ext-zero-x');
    expect(collapsed).toBeDefined();
    // The group has no width, so neither does the child. The two zeros are not
    // the same zero and this is the one that really means it.
    expect(collapsed!.reading.frame.w).toBe(0);
    expect(composeProbe(collapsed!).cx).toBe(0);
  });

  it('treats a group with no chOff or chExt as an origin at zero, not at off', () => {
    const probe = PROBES.find((p) => p.id === 'h-no-chxfrm');
    expect(probe).toBeDefined();
    const frame = composeProbe(probe!);
    expect(frame.x).toBeCloseTo(probe!.reading.frame.x, 6);
    // And the reading is not what `chOff = off` would give, which is the leaf's
    // own coordinate unchanged.
    expect(probe!.reading.frame.x).not.toBe(probe!.leaf.x);
  });
});

/* -------------------------------------------------------------------------- */

/** A turn as a 2x2, so two spellings of the same transform compare equal. */
function matrixOf(turn: { rot: number; flipH: boolean; flipV: boolean }): number[] {
  const radians = (turn.rot * Math.PI) / 180;
  const cos = Math.cos(radians);
  const sin = Math.sin(radians);
  const fx = turn.flipH ? -1 : 1;
  const fy = turn.flipV ? -1 : 1;
  return [cos * fx, sin * fx, -sin * fy, cos * fy];
}

/** The fixture spells the composed angle `rotation`; this package spells it `rot`. */
function readingTurn(reading: { rotation: number; flipH: boolean; flipV: boolean }): {
  rot: number;
  flipH: boolean;
  flipV: boolean;
} {
  return { rot: reading.rotation, flipH: reading.flipH, flipV: reading.flipV };
}

function sameMatrix(a: readonly number[], b: readonly number[]): boolean {
  return a.every((value, i) => Math.abs(value - (b[i] ?? 0)) < 1e-6);
}

describe('the orientation composition, re-derived', () => {
  const turns = PROBES.filter(
    (p) => !p.fromRepairedDeck && p.kind === 'turn' && p.deck !== 'shear' && p.id !== 't-shear',
  );

  it('composes every turn the way PowerPoint reported it', () => {
    expect(turns.length).toBeGreaterThanOrEqual(13);
    for (const probe of turns) {
      const frame = composeProbe(probe);
      expect(
        sameMatrix(matrixOf(frame), matrixOf(readingTurn(probe.reading))),
        `${probe.id}: ${String(frame.rot)} H${String(frame.flipH)} V${String(frame.flipV)}` +
          ` vs ${String(probe.reading.rotation)} H${String(probe.reading.flipH)} V${String(probe.reading.flipV)}`,
      ).toBe(true);
    }
  });

  it('refutes summing the angles, which is right on nine of the thirteen', () => {
    const wrong = turns.filter((probe) => {
      let rot = probe.leafRot;
      let flipH = probe.leafFlipH;
      let flipV = probe.leafFlipV;
      for (const link of probe.chain) {
        rot += link.rot ?? 0;
        if (link.flipH === true) flipH = !flipH;
        if (link.flipV === true) flipV = !flipV;
      }
      return !sameMatrix(matrixOf({ rot, flipH, flipV }), matrixOf(readingTurn(probe.reading)));
    });
    expect(wrong.length).toBeGreaterThan(0);
    expect(wrong.map((p) => p.id)).toContain('t-gfliph-crot');
  });

  it('subtracts the inner angle under a group that mirrors one axis', () => {
    // The whole finding in one line: `X R(b) = R(-b) X`.
    expect(
      composeTurn({ rot: 0, flipH: true, flipV: false }, { rot: 40, flipH: false, flipV: false }),
    ).toEqual({ rot: 320, flipH: true, flipV: false });
    // Two mirrors are a half turn, which commutes, so the angle survives.
    expect(
      composeTurn({ rot: 0, flipH: true, flipV: true }, { rot: 40, flipH: false, flipV: false }),
    ).toEqual({ rot: 40, flipH: true, flipV: true });
  });
});

/* -------------------------------------------------------------------------- */

describe('a rotated child in a non-uniformly scaled group', () => {
  const cases = fixture.findings.shear.cases;

  it('swaps the two scale factors at exactly the angles PowerPoint does', () => {
    expect(cases.length).toBeGreaterThanOrEqual(18);
    for (const one of cases) {
      expect(swapsExtents(one.rot), `${one.id} at ${String(one.rot)} degrees`).toBe(one.swapped);
    }
  });

  it('rounds the quadrant up at both 45 and 135, which no sin/cos test does', () => {
    expect(swapsExtents(44)).toBe(false);
    expect(swapsExtents(45)).toBe(true);
    expect(swapsExtents(134)).toBe(true);
    expect(swapsExtents(135)).toBe(false);
    expect(fixture.findings.shear.swapRule.winner).toBe('quadrantRoundHalfUp');
  });

  it('reproduces the frame PowerPoint reported for every shear case', () => {
    for (const one of cases) {
      const probe = PROBES.find((p) => p.id === one.id);
      expect(probe, one.id).toBeDefined();
      const frame = composeProbe(probe!);
      expect(frame.cx, `${one.id} width`).toBeCloseTo(one.reported.w, 4);
      expect(frame.cy, `${one.id} height`).toBeCloseTo(one.reported.h, 4);
      expect(frame.x, `${one.id} x`).toBeCloseTo(one.reported.x, 4);
      expect(frame.y, `${one.id} y`).toBeCloseTo(one.reported.y, 4);
    }
  });

  it('keeps the centre where a plain scaling would put it', () => {
    expect(fixture.findings.shear.centreAlwaysAgrees).toBe(true);
    expect(fixture.findings.shear.rotationAlwaysKept).toBe(true);
  });
});

/* -------------------------------------------------------------------------- */

describe('frameTransform', () => {
  it('puts the mirror inside the rotation', () => {
    const transform = frameTransform({
      x: 100,
      y: 200,
      cx: 40,
      cy: 20,
      rot: 30,
      flipH: true,
      flipV: false,
    });
    // Right to left: to the origin, mirror, rotate, out to the slide.
    expect(transform).toBe('translate(120 210) rotate(30) scale(-1 1) translate(-20 -10)');
  });

  it('says nothing at all about a shape that needs no transform', () => {
    expect(frameTransform({ x: 0, y: 0, cx: 0, cy: 0, rot: 0, flipH: false, flipV: false })).toBe(
      '',
    );
  });
});

/* -------------------------------------------------------------------------- */
/* sheets, built by hand                                                      */
/* -------------------------------------------------------------------------- */

const NS =
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"' +
  ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"' +
  ' xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"';

const SP_TREE_HEAD =
  '<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
  '<p:grpSpPr/>';

const CLR_MAP =
  '<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2"' +
  ' accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6"' +
  ' hlink="hlink" folHlink="folHlink"/>';

let nextId = 100;

interface SpecOptions {
  readonly rect?: { x: number; y: number; cx: number; cy: number };
  readonly rot?: number;
  readonly flipH?: boolean;
  readonly flipV?: boolean;
  readonly fill?: string;
  readonly line?: string;
  readonly prst?: string;
  readonly ph?: string;
  readonly name?: string;
  /** State no geometry at all, which is how a placeholder inherits one. */
  readonly noGeom?: boolean;
  readonly hidden?: boolean;
}

function sp(options: SpecOptions = {}): string {
  const id = nextId++;
  const rect = options.rect;
  const turn =
    (options.rot === undefined ? '' : ` rot="${String(Math.round(options.rot * DEGREE))}"`) +
    (options.flipH === true ? ' flipH="1"' : '') +
    (options.flipV === true ? ' flipV="1"' : '');
  const xfrm =
    rect === undefined
      ? ''
      : `<a:xfrm${turn}><a:off x="${String(rect.x)}" y="${String(rect.y)}"/>` +
        `<a:ext cx="${String(rect.cx)}" cy="${String(rect.cy)}"/></a:xfrm>`;
  return (
    '<p:sp><p:nvSpPr>' +
    `<p:cNvPr id="${String(id)}" name="${options.name ?? `s${String(id)}`}"${options.hidden === true ? ' hidden="1"' : ''}/><p:cNvSpPr/>` +
    `<p:nvPr>${options.ph === undefined ? '' : `<p:ph ${options.ph}/>`}</p:nvPr>` +
    '</p:nvSpPr><p:spPr>' +
    xfrm +
    (options.noGeom === true
      ? ''
      : `<a:prstGeom prst="${options.prst ?? 'rect'}"><a:avLst/></a:prstGeom>`) +
    (options.fill ?? '') +
    (options.line ?? '') +
    '</p:spPr></p:sp>'
  );
}

function grp(
  frame: { x: number; y: number; cx: number; cy: number },
  child: { x: number; y: number; cx: number; cy: number },
  children: readonly string[],
  options: { rot?: number; flipH?: boolean; fill?: string; name?: string } = {},
): string {
  const id = nextId++;
  const turn =
    (options.rot === undefined ? '' : ` rot="${String(Math.round(options.rot * DEGREE))}"`) +
    (options.flipH === true ? ' flipH="1"' : '');
  return (
    '<p:grpSp><p:nvGrpSpPr>' +
    `<p:cNvPr id="${String(id)}" name="${options.name ?? `g${String(id)}`}"/><p:cNvGrpSpPr/><p:nvPr/>` +
    '</p:nvGrpSpPr><p:grpSpPr>' +
    `<a:xfrm${turn}><a:off x="${String(frame.x)}" y="${String(frame.y)}"/>` +
    `<a:ext cx="${String(frame.cx)}" cy="${String(frame.cy)}"/>` +
    `<a:chOff x="${String(child.x)}" y="${String(child.y)}"/>` +
    `<a:chExt cx="${String(child.cx)}" cy="${String(child.cy)}"/></a:xfrm>` +
    (options.fill ?? '') +
    '</p:grpSpPr>' +
    children.join('') +
    '</p:grpSp>'
  );
}

const THEME_XML =
  '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="T">' +
  '<a:themeElements><a:clrScheme name="T">' +
  '<a:dk1><a:srgbClr val="000000"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>' +
  '<a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>' +
  '<a:accent1><a:srgbClr val="4472C4"/></a:accent1><a:accent2><a:srgbClr val="ED7D31"/></a:accent2>' +
  '<a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4>' +
  '<a:accent5><a:srgbClr val="5B9BD5"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6>' +
  '<a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink>' +
  '</a:clrScheme><a:fontScheme name="T"><a:majorFont><a:latin typeface="Calibri Light"/></a:majorFont>' +
  '<a:minorFont><a:latin typeface="Calibri"/></a:minorFont></a:fontScheme>' +
  '<a:fmtScheme name="T"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill>' +
  '</a:fillStyleLst><a:lnStyleLst><a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill>' +
  '</a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>' +
  '<a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst>' +
  '</a:fmtScheme></a:themeElements></a:theme>';

interface SheetOptions {
  readonly shapes?: readonly string[];
  readonly showMasterSp?: boolean;
  readonly bg?: string;
}

function buildChain(
  slide: SheetOptions = {},
  layout: SheetOptions = {},
  master: SheetOptions = {},
): { slide: Sheet; layout: Sheet; master: Sheet } {
  const body = (kind: 'slide' | 'layout' | 'master', options: SheetOptions): string => {
    const tag = kind === 'slide' ? 'p:sld' : kind === 'layout' ? 'p:sldLayout' : 'p:sldMaster';
    const attrs =
      (options.showMasterSp === false ? ' showMasterSp="0"' : '') +
      (kind === 'layout' ? ' type="obj"' : '');
    const tail = kind === 'master' ? CLR_MAP : '<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>';
    return (
      `<${tag} ${NS}${attrs}><p:cSld name="${kind}">${options.bg ?? ''}${SP_TREE_HEAD}` +
      `${(options.shapes ?? []).join('')}</p:spTree></p:cSld>${tail}</${tag}>`
    );
  };
  const parsed = (kind: 'slide' | 'layout' | 'master', options: SheetOptions) =>
    parseSheet(parseXmlString(body(kind, options)).root, `/ppt/${kind}.xml`);

  const theme = parseTheme(parseXmlString(THEME_XML).root, '/ppt/theme/theme1.xml');
  const masterSheet: Sheet = { ...parsed('master', master), parent: null, theme };
  const layoutSheet_: Sheet = { ...parsed('layout', layout), parent: masterSheet, theme: null };
  const slideSheet: Sheet = { ...parsed('slide', slide), parent: layoutSheet_, theme: null };
  return { slide: slideSheet, layout: layoutSheet_, master: masterSheet };
}

const SOLID = (hex: string): string => `<a:solidFill><a:srgbClr val="${hex}"/></a:solidFill>`;
const SIZE = { cx: 12192000, cy: 6858000 };

/* -------------------------------------------------------------------------- */

describe('layout', () => {
  it('places a group child through the child coordinate system', () => {
    const { slide } = buildChain({
      shapes: [
        grp({ x: 100, y: 100, cx: 800, cy: 400 }, { x: 100, y: 100, cx: 400, cy: 200 }, [
          sp({ rect: { x: 150, y: 120, cx: 80, cy: 40 }, name: 'leaf' }),
        ]),
      ],
    });
    const leaf = flatten(layoutSheet(slide)).find((p) => p.shape.name === 'leaf');
    expect(leaf).toBeDefined();
    // The same arithmetic as probe `m-scale2`, which PowerPoint answered
    // 200,140,160,80.
    expect(leaf!.frame).toMatchObject({ x: 200, y: 140, cx: 160, cy: 80 });
  });

  it('never paints a group fill', () => {
    const { slide } = buildChain({
      shapes: [
        grp(
          { x: 0, y: 0, cx: 100, cy: 100 },
          { x: 0, y: 0, cx: 100, cy: 100 },
          [sp({ rect: { x: 10, y: 10, cx: 20, cy: 20 }, name: 'leaf' })],
          { fill: SOLID('FF0000'), name: 'group' },
        ),
      ],
    });
    const group = layoutSheet(slide)[0]!;
    expect(group.shape.kind).toBe('grpSp');
    expect(group.fill).toBeNull();
    // But it is still there to be asked for.
    expect(group.appearance.fill).not.toBeNull();
  });

  it('lays a grpFill over the group content bounds, not the group rectangle', () => {
    const { slide } = buildChain({
      shapes: [
        grp(
          { x: 100, y: 100, cx: 600, cy: 300 },
          { x: 100, y: 100, cx: 600, cy: 300 },
          [
            sp({ rect: { x: 150, y: 180, cx: 100, cy: 140 }, fill: '<a:grpFill/>', name: 'left' }),
            sp({ rect: { x: 550, y: 180, cx: 100, cy: 140 }, fill: '<a:grpFill/>', name: 'right' }),
          ],
          { fill: SOLID('FF0000') },
        ),
      ],
    });
    const placed = flatten(layoutSheet(slide));
    const left = placed.find((p) => p.shape.name === 'left')!;
    // The measured span for exactly this arrangement was 150..650 - the two
    // children's own bounds - against a group that declares 100..700.
    const check = fixture.findings.groupFill.spanChecks.find((c) => c.probe === 'f-gradient')!;
    expect(check.matchesContent).toBe(true);
    expect(check.matchesDeclared).toBe(false);
    expect(left.fillBox.x).toBe(check.contentBounds.x0);
    expect(left.fillBox.x + left.fillBox.cx).toBe(check.contentBounds.x1);
    expect(left.fill).not.toBeNull();
    expect(left.fill?.type).toBe('solid');
  });

  it('walks a grpFill up through a group that also says grpFill', () => {
    const { slide } = buildChain({
      shapes: [
        grp(
          { x: 0, y: 0, cx: 600, cy: 300 },
          { x: 0, y: 0, cx: 600, cy: 300 },
          [
            grp(
              { x: 0, y: 0, cx: 300, cy: 200 },
              { x: 0, y: 0, cx: 300, cy: 200 },
              [sp({ rect: { x: 10, y: 10, cx: 40, cy: 40 }, fill: '<a:grpFill/>', name: 'leaf' })],
              { fill: '<a:grpFill/>' },
            ),
          ],
          { fill: SOLID('00FF00') },
        ),
      ],
    });
    const leaf = flatten(layoutSheet(slide)).find((p) => p.shape.name === 'leaf')!;
    expect(leaf.fill?.type).toBe('solid');
    // And over the outermost group's content bounds, which is the leaf itself.
    expect(leaf.fillBox).toMatchObject({ x: 10, y: 10, cx: 40, cy: 40 });
  });

  it('paints nothing for a grpFill that reached no group', () => {
    const { slide } = buildChain({
      shapes: [sp({ rect: { x: 0, y: 0, cx: 10, cy: 10 }, fill: '<a:grpFill/>', name: 'lonely' })],
    });
    const lonely = layoutSheet(slide)[0]!;
    // The shape asked for a group fill and there is no group, so the answer is
    // that nothing is painted - not that the request survives unresolved.
    expect(lonely.appearance.fill?.type).toBe('group');
    expect(lonely.fill).toBeNull();
    const markup = renderSlide(slide, SIZE);
    expect(markup).toContain('fill="none"');
  });
});

/* -------------------------------------------------------------------------- */

describe('what a slide shows that it did not write', () => {
  const compositing = fixture.findings.compositing;

  function visible(slideOff: boolean, layoutOff: boolean): { master: boolean; layout: boolean } {
    const { slide } = buildChain(
      {
        shapes: [sp({ rect: { x: 380, y: 60, cx: 120, cy: 120 }, name: 'slide' })],
        ...(slideOff ? { showMasterSp: false } : {}),
      },
      {
        shapes: [sp({ rect: { x: 220, y: 60, cx: 120, cy: 120 }, name: 'layout' })],
        ...(layoutOff ? { showMasterSp: false } : {}),
      },
      { shapes: [sp({ rect: { x: 60, y: 60, cx: 120, cy: 120 }, name: 'master' })] },
    );
    const names = new Set(flatten(layoutSlide(slide)).map((p) => p.shape.name));
    return { master: names.has('master'), layout: names.has('layout') };
  }

  it('matches PowerPoint on all four combinations', () => {
    const table: Record<string, [boolean, boolean]> = {
      'inherit-default': [false, false],
      'inherit-slide-off': [true, false],
      'inherit-layout-off': [false, true],
      'inherit-both-off': [true, true],
    };
    for (const row of compositing) {
      const [slideOff, layoutOff] = table[row.deck]!;
      const got = visible(slideOff, layoutOff);
      expect(got.master, `${row.deck} master`).toBe(row.masterVisible);
      expect(got.layout, `${row.deck} layout`).toBe(row.layoutVisible);
    }
  });

  it('never draws an unmatched placeholder from a layout or a master', () => {
    const { slide } = buildChain(
      { shapes: [] },
      {
        shapes: [
          sp({ rect: { x: 0, y: 0, cx: 10, cy: 10 }, ph: 'type="body" idx="7"', name: 'lph' }),
        ],
      },
      {
        shapes: [
          sp({ rect: { x: 0, y: 0, cx: 10, cy: 10 }, ph: 'type="body" idx="1"', name: 'mph' }),
        ],
      },
    );
    const names = flatten(layoutSlide(slide)).map((p) => p.shape.name);
    expect(names).not.toContain('lph');
    expect(names).not.toContain('mph');
    expect(compositing.every((row) => !row.placeholdersVisible)).toBe(true);
  });

  it('orders the inherited sheets furthest first', () => {
    const { slide, layout, master } = buildChain();
    expect(inheritedSheets(slide)).toEqual([master, layout]);
    expect(inheritedSheets(layout)).toEqual([master]);
    expect(inheritedSheets(master)).toEqual([]);
  });
});

/* -------------------------------------------------------------------------- */

describe('emission', () => {
  it('renders a solid rectangle as one path with one fill', () => {
    const { slide } = buildChain({
      shapes: [sp({ rect: { x: 0, y: 0, cx: 914400, cy: 457200 }, fill: SOLID('FF8800') })],
    });
    const markup = renderSlide(slide, SIZE, { idPrefix: 't' });
    expect(markup).toContain('viewBox="0 0 12192000 6858000"');
    expect(markup).toContain('fill="#FF8800"');
    expect(markup.match(/<path /g)?.length).toBe(1);
  });

  it('emits one path per a:path, so a preset can fill one and stroke another', () => {
    const { slide } = buildChain({
      shapes: [
        sp({
          rect: { x: 0, y: 0, cx: 914400, cy: 914400 },
          prst: 'smileyFace',
          fill: SOLID('FFCC00'),
          line: '<a:ln w="12700"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln>',
        }),
      ],
    });
    const markup = renderSlide(slide, SIZE, { idPrefix: 't' });
    // A filled face and an unfilled mouth: two paths, and one of them says so.
    expect((markup.match(/<path /g) ?? []).length).toBeGreaterThan(1);
    expect(markup).toContain('fill="none"');
  });

  it('turns a flipped, rotated shape the way PowerPoint wrote it', () => {
    const { slide } = buildChain({
      shapes: [
        sp({
          rect: { x: 0, y: 0, cx: 200, cy: 100 },
          rot: 330,
          flipH: true,
          fill: SOLID('112233'),
        }),
      ],
    });
    const markup = renderSlide(slide, SIZE, { idPrefix: 't' });
    expect(markup).toContain('rotate(330) scale(-1 1)');
  });

  it('gives every slide its own id space', () => {
    const { slide } = buildChain({
      shapes: [
        sp({
          rect: { x: 0, y: 0, cx: 100, cy: 100 },
          fill:
            '<a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs>' +
            '<a:gs pos="100000"><a:srgbClr val="0000FF"/></a:gs></a:gsLst>' +
            '<a:lin ang="0"/></a:gradFill>',
        }),
      ],
    });
    const a = renderSlide(slide, SIZE, { idPrefix: 'a' });
    const b = renderSlide(slide, SIZE, { idPrefix: 'b' });
    expect(a).toContain('url(#a-1)');
    expect(b).toContain('url(#b-1)');
    expect(a).not.toEqual(b);
  });

  it('spans a gradient across the shape rather than collapsing it at one end', () => {
    // The bug this replaces: `svgStops` returns a fraction and this package
    // divided it by a hundred thousand again, putting every stop at zero and
    // painting the whole shape in the last stop's colour. Nothing in the unit
    // tests noticed; the pixel comparison against PowerPoint failed on
    // seventeen samples at once.
    const { slide } = buildChain({
      shapes: [
        sp({
          rect: { x: 0, y: 0, cx: 400, cy: 200 },
          fill:
            '<a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs>' +
            '<a:gs pos="100000"><a:srgbClr val="0000FF"/></a:gs></a:gsLst>' +
            '<a:lin ang="0" scaled="0"/></a:gradFill>',
        }),
      ],
    });
    const markup = renderSlide(slide, SIZE, { idPrefix: 'g' });
    const offsets = [...markup.matchAll(/<stop offset="([^"]+)"/g)].map((m) => Number(m[1]));
    expect(offsets.length).toBeGreaterThan(1);
    expect(Math.min(...offsets)).toBe(0);
    expect(Math.max(...offsets)).toBe(1);
    // And the ramp runs left to right across the shape, in user units.
    expect(markup).toContain('gradientUnits="userSpaceOnUse"');
    expect(markup).toMatch(/x1="0" y1="100" x2="400" y2="100"/);
  });

  it('anchors a pattern tile in user space at the box the fill is laid over', () => {
    const { slide } = buildChain({
      shapes: [
        sp({
          rect: { x: 0, y: 0, cx: 400, cy: 200 },
          fill:
            '<a:pattFill prst="pct50"><a:fgClr><a:srgbClr val="FF0000"/></a:fgClr>' +
            '<a:bgClr><a:srgbClr val="FFFFFF"/></a:bgClr></a:pattFill>',
        }),
      ],
    });
    const markup = renderSlide(slide, SIZE, { idPrefix: 'p' });
    // Never objectBoundingBox: a physical tile does not scale with the shape.
    expect(markup).toContain('patternUnits="userSpaceOnUse"');
    // `resolvePattern` returns the *minimal* period, which for a half-tone is
    // two pixels rather than the full eight-pixel tile - 19050 EMU, or one and
    // a half points. Same picture, a quarter of the nodes.
    expect(markup).toContain('width="19050" height="19050"');
    expect(markup).toContain('fill="#FF0000"');
  });

  it('escapes a shape name that would otherwise close the tag', () => {
    const { slide } = buildChain({
      shapes: [sp({ rect: { x: 0, y: 0, cx: 10, cy: 10 }, name: 'a&quot;b' })],
    });
    const markup = renderSlide(slide, SIZE, { idPrefix: 't' });
    expect(markup).toContain('data-name="a&quot;b"');
    expect(markup).not.toContain('data-name="a"b"');
  });

  it('paints a background from the master when the slide states none', () => {
    const { slide } = buildChain(
      {},
      {},
      { bg: `<p:bg><p:bgPr>${SOLID('ABCDEF')}<a:effectLst/></p:bgPr></p:bg>` },
    );
    const markup = renderSlide(slide, SIZE, { idPrefix: 't' });
    expect(markup).toContain('data-role="background"');
    expect(markup).toContain('fill="#ABCDEF"');
  });

  it('refuses a preset nobody has heard of', () => {
    const { slide } = buildChain({
      shapes: [sp({ rect: { x: 0, y: 0, cx: 10, cy: 10 }, prst: 'notAShape' })],
    });
    expect(() => renderSlide(slide, SIZE)).toThrow(RenderError);
    try {
      renderSlide(slide, SIZE);
    } catch (error) {
      expect((error as RenderError).code).toBe('RENDER_UNKNOWN_PRESET');
    }
  });

  it('serializes the same tree slideNode builds', () => {
    const { slide } = buildChain({
      shapes: [sp({ rect: { x: 0, y: 0, cx: 10, cy: 10 }, fill: SOLID('010203') })],
    });
    const rendered = slideNode(slide, SIZE, { idPrefix: 't' });
    expect(serializeSvg(rendered.node)).toBe(renderSlide(slide, SIZE, { idPrefix: 't' }));
  });
});

/* -------------------------------------------------------------------------- */

describe('geometry the model now carries', () => {
  it('honours an a:avLst override rather than the preset default', () => {
    const withDefault = buildChain({
      shapes: [sp({ rect: { x: 0, y: 0, cx: 400, cy: 400 }, prst: 'roundRect' })],
    }).slide;
    const overridden = parseSheet(
      parseXmlString(
        `<p:sld ${NS}><p:cSld name="s">${SP_TREE_HEAD}` +
          '<p:sp><p:nvSpPr><p:cNvPr id="9" name="r"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr>' +
          '<a:xfrm><a:off x="0" y="0"/><a:ext cx="400" cy="400"/></a:xfrm>' +
          '<a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 5000"/></a:avLst>' +
          '</a:prstGeom></p:spPr></p:sp></p:spTree></p:cSld>' +
          '<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>',
      ).root,
      '/ppt/slides/slide1.xml',
    );
    const a = layoutSheet(withDefault)[0]!;
    const b = layoutSheet({ ...overridden, parent: null, theme: null })[0]!;
    expect(a.geometry?.paths[0]?.d).not.toBe(b.geometry?.paths[0]?.d);
  });

  it('lets a slide placeholder inherit its outline from the layout', () => {
    const { slide } = buildChain(
      { shapes: [sp({ ph: 'idx="1"', name: 'ph', noGeom: true })] },
      {
        shapes: [
          sp({
            rect: { x: 100, y: 200, cx: 300, cy: 150 },
            ph: 'idx="1"',
            prst: 'ellipse',
            fill: SOLID('123456'),
          }),
        ],
      },
    );
    const placed = layoutSheet(slide)[0]!;
    expect(placed.frame).toMatchObject({ x: 100, y: 200, cx: 300, cy: 150 });
    expect(placed.geometry?.name).toBe('ellipse');
  });

  it('renders an orphan placeholder at the origin with no size', () => {
    const { slide } = buildChain(
      { shapes: [sp({ ph: 'idx="9"', name: 'orphan' })] },
      { shapes: [] },
    );
    const placed: Placed = layoutSheet(slide)[0]!;
    expect(placed.frame).toMatchObject({ x: 0, y: 0, cx: 0, cy: 0 });
  });

  it('reads an a:custGeom the same way it reads a preset', () => {
    const shape: Shape = layoutSheet(
      buildChain({
        shapes: [
          '<p:sp><p:nvSpPr><p:cNvPr id="7" name="cg"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr>' +
            '<a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></a:xfrm>' +
            '<a:custGeom><a:avLst/><a:gdLst/><a:pathLst><a:path w="100" h="100">' +
            '<a:moveTo><a:pt x="0" y="0"/></a:moveTo>' +
            '<a:lnTo><a:pt x="100" y="0"/></a:lnTo>' +
            '<a:lnTo><a:pt x="100" y="100"/></a:lnTo>' +
            '<a:close/></a:path></a:pathLst></a:custGeom>' +
            '</p:spPr></p:sp>',
        ],
      }).slide,
    )[0]!.shape;
    expect(shape.geometry?.kind).toBe('custom');
    const placed = layoutSheet(buildChain({ shapes: [] }).slide);
    expect(placed).toEqual([]);
  });
});

/* -------------------------------------------------------------------------- */

describe('strokes', () => {
  it('does not scale a stroke with the group around it', () => {
    const { slide } = buildChain({
      shapes: [
        grp({ x: 0, y: 0, cx: 800, cy: 400 }, { x: 0, y: 0, cx: 400, cy: 200 }, [
          sp({
            rect: { x: 0, y: 0, cx: 100, cy: 50 },
            line: '<a:ln w="50800"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln>',
            name: 'leaf',
          }),
        ]),
      ],
    });
    const markup = renderSlide(slide, SIZE, { idPrefix: 't' });
    // Four points, doubled by the group in position and not in weight.
    expect(markup).toContain('stroke-width="50800"');
    expect(fixture.findings.strokeScale.weightAfterUngroup).toBe(4);
    expect(fixture.findings.strokeScale.strokeIsScaled).toBe(false);
  });

  it('draws an inset stroke double width and clips it to the shape', () => {
    const { slide } = buildChain({
      shapes: [
        sp({
          rect: { x: 0, y: 0, cx: 100, cy: 100 },
          line: '<a:ln w="12700" algn="in"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln>',
        }),
      ],
    });
    const markup = renderSlide(slide, SIZE, { idPrefix: 'i' });
    // Half of a doubled band, clipped to the outline, is the whole band inside.
    expect(markup).toContain('stroke-width="25400"');
    expect(markup).toContain('<clipPath id="i-1" clipPathUnits="userSpaceOnUse">');
    expect(markup).toContain('clip-path="url(#i-1)"');
  });

  it('draws nothing for a shape marked hidden', () => {
    const { slide } = buildChain({
      shapes: [
        sp({
          rect: { x: 0, y: 0, cx: 100, cy: 100 },
          fill: SOLID('FF0000'),
          name: 'gone',
          hidden: true,
        }),
        sp({ rect: { x: 200, y: 0, cx: 100, cy: 100 }, fill: SOLID('00FF00'), name: 'here' }),
      ],
    });
    expect(layoutSheet(slide).map((p) => p.shape.name)).toEqual(['here']);
    const markup = renderSlide(slide, SIZE, { idPrefix: 'h' });
    expect(markup).not.toContain('data-name="gone"');
    expect(markup).toContain('data-name="here"');
  });

  it('writes the join and the miter limit rather than leaving SVG its own', () => {
    const { slide } = buildChain({
      shapes: [
        sp({
          rect: { x: 0, y: 0, cx: 100, cy: 100 },
          line: '<a:ln w="12700"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln>',
        }),
      ],
    });
    const markup = renderSlide(slide, SIZE, { idPrefix: 't' });
    // Office is round and 8; SVG's own defaults are miter and 4.
    expect(markup).toContain('stroke-linejoin="round"');
    expect(markup).toContain('stroke-miterlimit="8"');
    expect(markup).toContain('stroke-linecap="butt"');
  });
});

/* -------------------------------------------------------------------------- */

/**
 * A picture's outline sits wholly outside its box - 12pt out and none in, of a
 * 12pt line - where a shape's default band straddles the geometry at 6 and 6.
 * Both numbers are read off `corpus/ground-truth/pictures.json`, so the two
 * constructions below are asserted against PowerPoint rather than each other.
 */
describe('strokes at a named width, re-derived from F2', () => {
  const BLACK = '<a:solidFill><a:srgbClr val="000000"/></a:solidFill>';
  const ln = (w: string, extra = ''): string => `<a:ln w="${w}">${BLACK}${extra}</a:ln>`;
  // A 400-pt line across the slide at y = 100 pt, the F2 hairline probe's own geometry.
  const LINE_RECT = { x: 1270000, y: 1270000, cx: 5080000, cy: 0 };
  const lineSlide = (line: string): Sheet =>
    buildChain({ shapes: [sp({ rect: LINE_RECT, prst: 'line', line, name: 'probe' })] }).slide;
  const strokeWidth = (markup: string): string =>
    /stroke-width="([^"]+)"/.exec(markup)?.[1] ?? 'none';

  it('draws a w="0" line as one device pixel at the width it is asked for', () => {
    const slide = lineSlide(ln('0'));
    expect(strokeWidth(renderSlide(slide, SIZE, { idPrefix: 'h', width: 960 }))).toBe('12700');
    expect(strokeWidth(renderSlide(slide, SIZE, { idPrefix: 'h', width: 3840 }))).toBe('3175');
    expect(strokeWidth(renderSlide(slide, SIZE, { idPrefix: 'h', width: 120 }))).toBe('101600');
    expect(zoom.findings.hairline).toBe('H1');
  });

  it('draws a w="0" line as one screen pixel that no transform scales when no width is named', () => {
    const markup = renderSlide(lineSlide(ln('0')), SIZE, { idPrefix: 'h' });
    expect(markup).toContain('stroke-width="1"');
    expect(markup).toContain('vector-effect="non-scaling-stroke"');
  });

  it('draws a dashed hairline solid', () => {
    const slide = lineSlide(ln('0', '<a:prstDash val="dash"/>'));
    expect(renderSlide(slide, SIZE, { idPrefix: 'd', width: 960 })).not.toContain(
      'stroke-dasharray',
    );
    expect(renderSlide(slide, SIZE, { idPrefix: 'd' })).not.toContain('stroke-dasharray');
    expect(zoom.findings.dashOnHairline).toBe('D1');
  });

  it('rounds a stroke to whole pixels at the named width, never under one', () => {
    const quarter = lineSlide(ln('3175'));
    const oneAndAHalf = lineSlide(ln('19050'));
    const elevenTenths = lineSlide(ln('13970'));
    expect(strokeWidth(renderSlide(quarter, SIZE, { idPrefix: 'q', width: 960 }))).toBe('12700');
    expect(strokeWidth(renderSlide(oneAndAHalf, SIZE, { idPrefix: 'q', width: 960 }))).toBe(
      '25400',
    );
    expect(strokeWidth(renderSlide(elevenTenths, SIZE, { idPrefix: 'q', width: 3840 }))).toBe(
      '12700',
    );
    // No width names no device, and the true width stays.
    expect(strokeWidth(renderSlide(oneAndAHalf, SIZE, { idPrefix: 'q' }))).toBe('19050');
    expect(zoom.findings.thin).toBe('T4');
  });

  it('rounds to the device pixels of a 2x display: the strokes of twice the width', () => {
    const slide = lineSlide(ln('3175'));
    const rootSize = /\swidth="\d+" height="\d+"/;
    const doubled = renderSlide(slide, SIZE, {
      idPrefix: 'r',
      width: 960,
      height: 540,
      devicePixelRatio: 2,
    });
    const twice = renderSlide(slide, SIZE, { idPrefix: 'r', width: 1920, height: 1080 });
    expect(strokeWidth(doubled)).toBe('6350');
    expect(doubled.replace(rootSize, '')).toBe(twice.replace(rootSize, ''));
    expect(doubled).toContain('width="960"');
    // The ratio is a device: without one there is nothing for it to name.
    expect(() => renderSlide(slide, SIZE, { idPrefix: 'r', devicePixelRatio: 2 })).toThrow(
      RenderError,
    );
    expect(() =>
      renderSlide(slide, SIZE, { idPrefix: 'r', width: 960, devicePixelRatio: 0 }),
    ).toThrow(RenderError);
  });

  it('draws a picture border at the drawn width, under a band clip the zoom never moves', () => {
    const id = nextId++;
    const pic =
      `<p:pic><p:nvPicPr><p:cNvPr id="${String(id)}" name="picture"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>` +
      '<p:blipFill><a:blip r:embed="rId9"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>' +
      '<p:spPr><a:xfrm><a:off x="1270000" y="1270000"/><a:ext cx="3810000" cy="1270000"/></a:xfrm>' +
      `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>${ln('3175')}</p:spPr></p:pic>`;
    const { slide } = buildChain({ shapes: [pic] });
    const at960 = renderSlide(slide, SIZE, { idPrefix: 'p', width: 960 });
    const at240 = renderSlide(slide, SIZE, { idPrefix: 'p', width: 240 });
    // A quarter point is one pixel at 960 and four points at 240, doubled for the one-sided band.
    expect(at960).toContain('stroke-width="25400"');
    expect(at240).toContain('stroke-width="101600"');
    // The clip keeping the outer half reaches the same constant distance at both.
    const clipOf = (markup: string): string =>
      /<clipPath[^>]*><path d="(M[^ ]+ [^H]+H[^V]+V[^H]+H[^Z]+Z)/.exec(markup)?.[1] ?? '';
    expect(clipOf(at960)).toBe(clipOf(at240));
    // Twice a quarter point plus the twenty-point pixel of the 5 % floor: 40.5 pt out.
    expect(clipOf(at960).startsWith(`M-${String(40.5 * 12700)} -${String(40.5 * 12700)}H`)).toBe(
      true,
    );
    expect(zoom.findings.border).toBe('T4');
  });

  /** Ink per column across a horizontal line, from the SVG drawn at `width` in this browser. */
  async function inkAcross(markup: string, width: number, height: number): Promise<number> {
    const image = new Image();
    const href = URL.createObjectURL(new Blob([markup], { type: 'image/svg+xml' }));
    image.src = href;
    await image.decode();
    const canvas = document.createElement('canvas');
    canvas.width = width;
    canvas.height = height;
    const context = canvas.getContext('2d', { willReadFrequently: true });
    if (context === null) throw new Error('no 2d context');
    context.fillStyle = '#FFFFFF';
    context.fillRect(0, 0, width, height);
    context.drawImage(image, 0, 0, width, height);
    URL.revokeObjectURL(href);
    const data = context.getImageData(0, 0, width, height).data;
    const scale = width / 960;
    let total = 0;
    let columns = 0;
    for (
      let x = Math.round(150 * scale);
      x <= Math.round(450 * scale);
      x += Math.max(1, Math.round(10 * scale))
    ) {
      for (let y = Math.round(84 * scale); y <= Math.round(116 * scale); y++) {
        const at = (y * width + x) * 4;
        total += 1 - Math.min(data[at]!, data[at + 1]!, data[at + 2]!) / 255;
      }
      columns += 1;
    }
    return total / columns;
  }

  it('inks one pixel per column at 240 and at 3840 wide, as PowerPoint does', async () => {
    const slide = lineSlide(ln('0'));
    const at = (width: number): Promise<number> =>
      inkAcross(
        renderSlide(slide, SIZE, { idPrefix: 'r', width, height: (width * 9) / 16 }),
        width,
        (width * 9) / 16,
      );
    const [small, reference, large] = await Promise.all([at(240), at(960), at(3840)]);
    expect(Math.abs(small - 1)).toBeLessThanOrEqual(zoom.tolerancePx);
    expect(Math.abs(reference - 1)).toBeLessThanOrEqual(zoom.tolerancePx);
    expect(Math.abs(large - 1)).toBeLessThanOrEqual(zoom.tolerancePx);
    // A stroke that scaled with the slide would ink four times as much at 4x.
    expect(Math.abs(large - 4 * reference)).toBeGreaterThan(1);
  });
});

describe('line ends, re-derived from C4 and F2', () => {
  const BLACK = '<a:solidFill><a:srgbClr val="000000"/></a:solidFill>';
  const LINE_RECT = { x: 1270000, y: 1270000, cx: 5080000, cy: 0 };
  const ended = (w: string, ends: string, prst = 'line'): Sheet =>
    buildChain({
      shapes: [
        sp({ rect: LINE_RECT, prst, line: `<a:ln w="${w}">${BLACK}${ends}</a:ln>`, name: 'probe' }),
      ],
    }).slide;
  const TAIL = '<a:tailEnd type="triangle" w="lg" len="lg"/>';
  const markerOf = (markup: string): string => /<marker [^>]*>/.exec(markup)?.[0] ?? '';
  const attr = (tag: string, name: string): string =>
    new RegExp(` ${name}="([^"]*)"`).exec(tag)?.[1] ?? '';

  it('draws a tail end at the end of an open path and a head end at its start, oriented outward', () => {
    const both = renderSlide(
      ended('12700', '<a:headEnd type="oval" w="med" len="med"/>' + TAIL),
      SIZE,
      { idPrefix: 'e', width: 960 },
    );
    expect(both).toContain('marker-start="url(#e-1)"');
    expect(both).toContain('marker-end="url(#e-2)"');
    const [head, tail] = [...both.matchAll(/<marker [^>]*>/g)].map((m) => m[0]);
    expect(attr(head ?? '', 'orient')).toBe('auto-start-reverse');
    expect(attr(tail ?? '', 'orient')).toBe('auto');
    expect(attr(head ?? '', 'markerUnits')).toBe('userSpaceOnUse');
    // The head is the med oval, three pens; the tail the lg triangle, five.
    expect(attr(head ?? '', 'markerWidth')).toBe(String(6 * 12700));
    expect(attr(tail ?? '', 'markerWidth')).toBe(String(10 * 12700));
  });

  it('sizes a head from a two-point pen on a one-point line: lg is five pens, so ten points', () => {
    const tag = markerOf(renderSlide(ended('12700', TAIL), SIZE, { idPrefix: 'e', width: 960 }));
    expect(attr(tag, 'markerWidth')).toBe(String(10 * 12700));
    expect(attr(tag, 'markerHeight')).toBe(String(10 * 12700));
    // A triangle's tip sits on the endpoint: the reference point is the whole length in.
    expect(attr(tag, 'refX')).toBe(String(10 * 12700));
    expect(attr(tag, 'refY')).toBe(String(5 * 12700));
    // Above two points the pen is the drawn stroke: 2.5 pt is three pixels at 960, so fifteen.
    const thick = markerOf(renderSlide(ended('31750', TAIL), SIZE, { idPrefix: 'e', width: 960 }));
    expect(attr(thick, 'markerWidth')).toBe(String(15 * 12700));
    // No width names no device: the pen is two points, or the width above it.
    const bare = markerOf(renderSlide(ended('31750', TAIL), SIZE, { idPrefix: 'e' }));
    expect(attr(bare, 'markerWidth')).toBe(String(12.5 * 12700));
    expect(zoom.findings.markerOnHairline).toBe('M8');
  });

  it('centres a diamond and an oval on the endpoint, and strokes an open arrow with the pen', () => {
    const oval = markerOf(
      renderSlide(ended('12700', '<a:tailEnd type="oval" w="lg" len="lg"/>'), SIZE, {
        idPrefix: 'e',
        width: 960,
      }),
    );
    expect(attr(oval, 'refX')).toBe(String(5 * 12700));
    const arrow = renderSlide(ended('12700', '<a:tailEnd type="arrow" w="lg" len="lg"/>'), SIZE, {
      idPrefix: 'e',
      width: 960,
    });
    const v = /<marker [^>]*><path [^>]*>/.exec(arrow)?.[0] ?? '';
    expect(v).toContain('fill="none"');
    expect(v).toContain(`stroke-width="${String(2 * 12700)}"`);
    expect(v).toContain('stroke-linejoin="round"');
    expect(v).not.toContain('Z"');
    // The V's vertex sits half a pen back, so the round join's edge is on the endpoint.
    expect(v).toContain('d="M-12700,0 L114300,63500 L-12700,127000"');
  });

  it("asks a gradient-stroked line's head for the line's own paint, not the marker's space", () => {
    const gradient =
      '<a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="000000"/></a:gs>' +
      '<a:gs pos="100000"><a:srgbClr val="FFFFFF"/></a:gs></a:gsLst><a:lin ang="0"/></a:gradFill>';
    const slide = buildChain({
      shapes: [
        sp({
          rect: LINE_RECT,
          prst: 'line',
          line: `<a:ln w="12700">${gradient}${TAIL}</a:ln>`,
          name: 'probe',
        }),
      ],
    }).slide;
    const markup = renderSlide(slide, SIZE, { idPrefix: 'e', width: 960 });
    expect(markup).toMatch(/<path [^>]*stroke="url\(#e-\d+\)"[^>]*marker-end=/);
    expect(/<marker [^>]*><path [^>]*>/.exec(markup)?.[0]).toContain('fill="context-stroke"');
  });

  it('draws no end on a closed path, and none where the file names none', () => {
    const closed = renderSlide(ended('12700', TAIL, 'rect'), SIZE, { idPrefix: 'e', width: 960 });
    expect(closed).not.toContain('<marker');
    expect(closed).not.toContain('marker-end');
    const none = renderSlide(ended('12700', '<a:tailEnd type="none"/>'), SIZE, {
      idPrefix: 'e',
      width: 960,
    });
    expect(none).not.toContain('<marker');
  });

  /** The inked rows at the widest column of a large triangle head whose tip is at 500 pt, 100 pt. */
  async function headHeight(
    markup: string,
    width: number,
  ): Promise<{ rows: number; centre: number }> {
    const height = (width * 9) / 16;
    const image = new Image();
    const href = URL.createObjectURL(new Blob([markup], { type: 'image/svg+xml' }));
    image.src = href;
    await image.decode();
    const canvas = document.createElement('canvas');
    canvas.width = width;
    canvas.height = height;
    const context = canvas.getContext('2d', { willReadFrequently: true });
    if (context === null) throw new Error('no 2d context');
    context.fillStyle = '#FFFFFF';
    context.fillRect(0, 0, width, height);
    context.drawImage(image, 0, 0, width, height);
    URL.revokeObjectURL(href);
    const data = context.getImageData(0, 0, width, height).data;
    const scale = width / 960;
    let widest = { rows: 0, centre: 0 };
    for (let x = Math.round(480 * scale); x <= Math.round(500 * scale); x++) {
      let rows = 0;
      let sum = 0;
      for (let y = Math.round(84 * scale); y <= Math.round(116 * scale); y++) {
        const at = (y * width + x) * 4;
        if (1 - Math.min(data[at]!, data[at + 1]!, data[at + 2]!) / 255 >= 0.5) {
          rows += 1;
          sum += y + 0.5;
        }
      }
      if (rows > widest.rows) widest = { rows, centre: sum / rows };
    }
    return widest;
  }

  it('ends an open arrow on the endpoint, with no ink past it', async () => {
    const markup = renderSlide(ended('0', '<a:tailEnd type="arrow" w="lg" len="lg"/>'), SIZE, {
      idPrefix: 'r',
      width: 960,
      height: 540,
    });
    const image = new Image();
    const href = URL.createObjectURL(new Blob([markup], { type: 'image/svg+xml' }));
    image.src = href;
    await image.decode();
    const canvas = document.createElement('canvas');
    canvas.width = 960;
    canvas.height = 540;
    const context = canvas.getContext('2d', { willReadFrequently: true });
    if (context === null) throw new Error('no 2d context');
    context.fillStyle = '#FFFFFF';
    context.fillRect(0, 0, 960, 540);
    context.drawImage(image, 0, 0, 960, 540);
    URL.revokeObjectURL(href);
    const data = context.getImageData(0, 0, 960, 540).data;
    const inkAt = (x: number): number => {
      let ink = 0;
      for (let y = 84; y <= 116; y++) {
        const at = (y * 960 + x) * 4;
        ink += 1 - Math.min(data[at]!, data[at + 1]!, data[at + 2]!) / 255;
      }
      return ink;
    };
    // The tip is at x = 500 px: ink in the column before it, none in the column after.
    expect(inkAt(498)).toBeGreaterThan(0.5);
    expect(inkAt(501)).toBeLessThan(0.05);
  });

  it('inks a ten-point head at 960 and a five-pixel one at 240, as PowerPoint did', async () => {
    const slide = ended('0', TAIL);
    const at = (width: number): Promise<{ rows: number; centre: number }> =>
      headHeight(
        renderSlide(slide, SIZE, { idPrefix: 'r', width, height: (width * 9) / 16 }),
        width,
      );
    const [small, reference] = await Promise.all([at(240), at(960)]);
    const seen = zoom.measured['marker-0'];
    expect(Math.abs(reference.rows - seen['960'].boxH)).toBeLessThanOrEqual(zoom.boxTolerancePx);
    expect(Math.abs(small.rows - seen['240'].boxH)).toBeLessThanOrEqual(zoom.boxTolerancePx);
    // A head that scaled with the line would be a quarter the size at 240; the pen's floor holds it.
    expect(small.rows).toBeGreaterThan(reference.rows / 4);
    // And it sits on the line, at y = 100 pt, not beside it.
    expect(Math.abs(reference.centre - 100)).toBeLessThanOrEqual(1);
    expect(Math.abs(small.centre - 25)).toBeLessThanOrEqual(1);
  });
});

describe('the device grid, re-derived from F3', () => {
  const BLACK = '<a:solidFill><a:srgbClr val="000000"/></a:solidFill>';
  const ln = (pt: number): string => `<a:ln w="${String(Math.round(pt * 12700))}">${BLACK}</a:ln>`;
  const PT = 12700;
  // A rectangle a quarter point past the grid on both axes, so nothing about it is a pixel edge.
  const RECT = { x: 100.25 * PT, y: 60.25 * PT, cx: 160 * PT, cy: 100 * PT };
  const at = (markup: string, width: number | undefined, extra = {}): string =>
    renderSlide(buildChain({ shapes: [markup] }).slide, SIZE, {
      idPrefix: 'g',
      ...(width === undefined ? {} : { width }),
      ...extra,
    });
  const pathTags = (markup: string): string[] => markup.match(/<path\b[^>]*>/g) ?? [];
  const crisp = (tag: string): boolean => tag.includes('shape-rendering="crispEdges"');
  const shift = (tag: string): string | null =>
    /transform="translate\(([^)]*)\)"/.exec(tag)?.[1] ?? null;

  it('draws an axis-aligned outline crisp, an odd pen half a device pixel on and an even pen where it lies', () => {
    const one = pathTags(at(sp({ rect: RECT, line: ln(1) }), 960));
    expect(one).toHaveLength(1);
    expect(crisp(one[0]!)).toBe(true);
    expect(shift(one[0]!)).toBe('6350 6350');
    // Two pixels at 960: even, no shift. Three at 1.5 pt and 1920: odd, half of 1/2 pt.
    expect(shift(pathTags(at(sp({ rect: RECT, line: ln(2) }), 960))[0]!)).toBeNull();
    expect(shift(pathTags(at(sp({ rect: RECT, line: ln(1.5) }), 1920))[0]!)).toBe('3175 3175');
    // A hairline is one pixel, so it moves too; at 240 that half pixel is two points.
    expect(shift(pathTags(at(sp({ rect: RECT, line: ln(0) }), 240))[0]!)).toBe('25400 25400');
    // The 2x display doubles the pen: what was odd is even.
    expect(
      shift(pathTags(at(sp({ rect: RECT, line: ln(1) }), 960, { devicePixelRatio: 2 }))[0]!),
    ).toBeNull();
    expect(snap.findings.stroke).toBe('SP');
  });

  it('draws a fill crisp with no shift, and a curve or a turned edge antialiased where it lies', () => {
    const fill = pathTags(at(sp({ rect: RECT, fill: BLACK }), 960))[0]!;
    expect(crisp(fill)).toBe(true);
    expect(shift(fill)).toBeNull();
    // Under an odd pen the fill keeps its rows and the pen alone moves: two paths, not one.
    const underOdd = pathTags(at(sp({ rect: RECT, fill: BLACK, line: ln(1) }), 960));
    expect(underOdd).toHaveLength(2);
    expect(underOdd[0]).toContain('fill="#000000"');
    expect(crisp(underOdd[0]!)).toBe(true);
    expect(shift(underOdd[0]!)).toBeNull();
    expect(underOdd[1]).toContain('fill="none"');
    expect(shift(underOdd[1]!)).toBe('6350 6350');
    expect(underOdd[1]).toContain('stroke="#000000"');
    // Under an even pen nothing moves, and the two paths stay two so a zoom changes no structure.
    const underEven = pathTags(at(sp({ rect: RECT, fill: BLACK, line: ln(2) }), 960));
    expect(underEven).toHaveLength(2);
    expect(shift(underEven[1]!)).toBeNull();
    expect(pathTags(at(sp({ rect: RECT, fill: BLACK, line: ln(2) }), undefined))).toHaveLength(1);
    for (const turned of [
      sp({ rect: RECT, line: ln(1), prst: 'ellipse' }),
      sp({ rect: RECT, line: ln(1), prst: 'roundRect' }),
      sp({ rect: RECT, line: ln(1), rot: 30 }),
      sp({ rect: RECT, line: ln(1), prst: 'triangle' }),
      sp({ rect: RECT, fill: BLACK, prst: 'ellipse' }),
    ]) {
      const tags = pathTags(at(turned, 960));
      expect(tags.length).toBeGreaterThan(0);
      for (const tag of tags) {
        expect(crisp(tag), tag).toBe(false);
        expect(shift(tag), tag).toBeNull();
      }
    }
    expect(snap.findings.fill).toBe('ER');
    expect(snap.findings.ellipse).toBe('SQ');
  });

  it('turns the half pixel with the shape, so it is still down and right on the slide', () => {
    const turned = (rot: number, flipH = false, flipV = false): string | null =>
      shift(pathTags(at(sp({ rect: RECT, line: ln(1), rot, flipH, flipV }), 960))[0]!);
    // The local vector the frame's own inverse maps half a device pixel down and right onto.
    const expected = (rot: number, flipH = false, flipV = false): string => {
      const frame = { ...RECT, rot, flipH, flipV };
      const origin = inverseFramePoint(frame, { x: 0, y: 0 });
      const moved = inverseFramePoint(frame, { x: 6350, y: 6350 });
      return `${num(Math.round(moved.x - origin.x))} ${num(Math.round(moved.y - origin.y))}`;
    };
    for (const [rot, flipH, flipV] of [
      [0, false, false],
      [90, false, false],
      [180, false, false],
      [270, false, false],
      [0, true, false],
      [0, false, true],
      [90, true, false],
      [270, true, true],
    ] as const) {
      expect(turned(rot, flipH, flipV), `${String(rot)} ${String(flipH)} ${String(flipV)}`).toBe(
        expected(rot, flipH, flipV),
      );
    }
    expect(turned(90)).toBe('6350 -6350');
    // A quarter turn keeps every edge on an axis, and the export snaps it: 24 of 24.
    expect(crisp(pathTags(at(sp({ rect: RECT, line: ln(1), rot: 90 }), 960))[0]!)).toBe(true);
    expect(snap.findings.rotated).toBe('SP');
  });

  it('snaps a rectilinear custom geometry and not one with a diagonal', () => {
    const pt = (x: number, y: number): string =>
      `<a:pt x="${String(x * PT)}" y="${String(y * PT)}"/>`;
    const custom = (points: readonly [number, number][]): string =>
      sp({
        rect: RECT,
        line: ln(1),
        noGeom: true,
      }).replace(
        '</a:xfrm>',
        `</a:xfrm><a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="0" t="0" r="r" b="b"/>` +
          `<a:pathLst><a:path w="${String(160 * PT)}" h="${String(100 * PT)}"><a:moveTo>${pt(...points[0]!)}</a:moveTo>` +
          points
            .slice(1)
            .map(([x, y]) => `<a:lnTo>${pt(x, y)}</a:lnTo>`)
            .join('') +
          '<a:close/></a:path></a:pathLst></a:custGeom>',
      );
    const ell = custom([
      [0, 0],
      [160, 0],
      [160, 40],
      [80, 40],
      [80, 100],
      [0, 100],
    ]);
    expect(crisp(pathTags(at(ell, 960))[0]!)).toBe(true);
    const chevron = custom([
      [0, 0],
      [160, 0],
      [120, 50],
      [160, 100],
      [0, 100],
    ]);
    expect(crisp(pathTags(at(chevron, 960))[0]!)).toBe(false);
  });

  it('keeps a clipped band antialiased under its antialiased clip, and snaps the picture under it', () => {
    const id = nextId++;
    const pic =
      `<p:pic><p:nvPicPr><p:cNvPr id="${String(id)}" name="picture"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>` +
      '<p:blipFill><a:blip r:embed="rId9"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>' +
      `<p:spPr><a:xfrm><a:off x="${String(RECT.x)}" y="${String(RECT.y)}"/><a:ext cx="${String(RECT.cx)}" cy="${String(RECT.cy)}"/></a:xfrm>` +
      `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>${ln(1)}</p:spPr></p:pic>`;
    const tags = pathTags(at(pic, 960)).filter((tag) => !tag.includes('clipPath'));
    const fill = tags.find((tag) => tag.includes('stroke="none"'))!;
    const band = tags.find((tag) => tag.includes('clip-path="url('))!;
    expect(crisp(fill)).toBe(true);
    expect(crisp(band)).toBe(false);
    expect(shift(band)).toBeNull();
    expect(snap.findings.picture).toBe('ER');
    expect(snap.findings.border).toBe('BT');
  });

  it('clips a band antialiased whatever its clip path asks, which is why the band stays so', async () => {
    // A crisp clip child at a quarter-pixel edge: the clipped fill's edge row is still partial.
    const clipped = (rendering: string): string =>
      '<svg xmlns="http://www.w3.org/2000/svg" width="100" height="100" viewBox="0 0 100 100">' +
      '<defs><clipPath id="c" clipPathUnits="userSpaceOnUse">' +
      `<path d="M0 20.25H100V100H0Z" shape-rendering="${rendering}"/></clipPath></defs>` +
      '<rect x="10" y="0" width="80" height="60" fill="#000" clip-path="url(#c)" shape-rendering="crispEdges"/></svg>';
    for (const rendering of ['crispEdges', 'auto']) {
      const rows = await rowsAcross(clipped(rendering), 100, 100, [40, 60], [19, 22]);
      expect(
        rows.map((v) => Math.round(v * 100) / 100),
        rendering,
      ).toEqual([0, 0.75, 1, 1]);
    }
    // The same edge on a crisp rect with no clip is whole: the antialiasing is the clip's.
    const bare =
      '<svg xmlns="http://www.w3.org/2000/svg" width="100" height="100" viewBox="0 0 100 100">' +
      '<rect x="10" y="20.25" width="80" height="40" fill="#000" shape-rendering="crispEdges"/></svg>';
    expect(await rowsAcross(bare, 100, 100, [40, 60], [19, 22])).toEqual([0, 1, 1, 1]);
  });

  it('snaps nothing when no width names a device', () => {
    for (const tag of pathTags(at(sp({ rect: RECT, line: ln(1), fill: BLACK }), undefined))) {
      expect(crisp(tag)).toBe(false);
      expect(shift(tag)).toBeNull();
    }
  });

  it('stretches into a named box as the export does, and keeps the aspect with only a width', () => {
    expect(at(sp({ rect: RECT, fill: BLACK }), 960, { height: 541 })).toContain(
      'preserveAspectRatio="none"',
    );
    expect(at(sp({ rect: RECT, fill: BLACK }), 960)).toContain(
      'preserveAspectRatio="xMidYMid meet"',
    );
    expect(zoom.findings.frameStretched).toBe(true);
  });

  /** The rows of one column band across a horizontal edge, as this browser draws the markup. */
  async function rowsAcross(
    markup: string,
    width: number,
    height: number,
    columns: [number, number],
    rows: [number, number],
  ): Promise<number[]> {
    const image = new Image();
    const href = URL.createObjectURL(new Blob([markup], { type: 'image/svg+xml' }));
    image.src = href;
    await image.decode();
    const canvas = document.createElement('canvas');
    canvas.width = width;
    canvas.height = height;
    const context = canvas.getContext('2d', { willReadFrequently: true });
    if (context === null) throw new Error('no 2d context');
    context.fillStyle = '#FFFFFF';
    context.fillRect(0, 0, width, height);
    context.drawImage(image, 0, 0, width, height);
    URL.revokeObjectURL(href);
    const data = context.getImageData(0, 0, width, height).data;
    const out: number[] = [];
    for (let y = rows[0]; y <= rows[1]; y++) {
      let ink = 0;
      for (let x = columns[0]; x <= columns[1]; x++) {
        const px = (y * width + x) * 4;
        ink += 1 - Math.min(data[px]!, data[px + 1]!, data[px + 2]!) / 255;
      }
      out.push(ink / (columns[1] - columns[0] + 1));
    }
    return out;
  }

  interface SnapProbe {
    readonly id: string;
    readonly family: string;
    readonly centrePt: number;
    readonly widthPt?: number;
    readonly read: { readonly axis: 'x' | 'y'; readonly at: number; readonly half: number };
  }

  /** F3's own probe slide: every horizontal stroke, at the coordinates the export was asked for. */
  function strokeSlide(): { sheet: Sheet; probes: SnapProbe[] } {
    const probes = (snap.probes as SnapProbe[]).filter(
      (probe) => probe.family === 'stroke' && probe.read.axis === 'y',
    );
    const shapes = probes.map((probe) =>
      sp({
        rect: { x: 120 * PT, y: Math.round(probe.centrePt * PT), cx: 600 * PT, cy: 0 },
        prst: 'line',
        line: ln(probe.widthPt ?? 0),
        name: probe.id,
      }),
    );
    return { sheet: buildChain({ shapes }).slide, probes };
  }

  it('inks the rows PowerPoint inked, on every horizontal stroke probe at 960 and at 1920', async () => {
    const { sheet, probes } = strokeSlide();
    const measured = snap.measured as Record<
      string,
      Record<string, { from: number; cover: number[] }>
    >;
    let compared = 0;
    let worst = 0;
    for (const width of [960, 1920]) {
      const scale = width / 960;
      const height = (width * 9) / 16;
      const markup = renderSlide(sheet, SIZE, { idPrefix: 'f', width, height });
      for (const probe of probes) {
        const from = Math.round((probe.read.at - probe.read.half) * scale);
        const to = Math.round((probe.read.at + probe.read.half) * scale);
        const ours = await rowsAcross(
          markup,
          width,
          height,
          [Math.round(200 * scale), Math.round(640 * scale)],
          [from, to],
        );
        const seen = measured[probe.id]![String(width)]!;
        const theirs = new Array<number>(to - from + 1).fill(0);
        seen.cover.forEach((v, i) => {
          theirs[seen.from - from + i] = v;
        });
        const error = Math.max(...ours.map((v, i) => Math.abs(v - theirs[i]!)));
        worst = Math.max(worst, error);
        expect(error, `${probe.id}@${String(width)}`).toBeLessThanOrEqual(snap.tolerance);
        compared += 1;
      }
    }
    expect(compared).toBe(56);
    expect(worst).toBeLessThanOrEqual(snap.tolerance);
  });

  it('would not, drawn antialiased where it lies: the same probes straddle two rows', async () => {
    // The markup with the snap taken out of it: the device pen, antialiased where it lies.
    const { sheet, probes } = strokeSlide();
    const markup = renderSlide(sheet, SIZE, { idPrefix: 'u', width: 960, height: 540 }).replace(
      /<path\b[^>]*shape-rendering="crispEdges"[^>]*>/g,
      (path) => path.replace(/ shape-rendering="crispEdges"| transform="translate\([^)]*\)"/g, ''),
    );
    const measured = snap.measured as Record<
      string,
      Record<string, { from: number; cover: number[] }>
    >;
    let misses = 0;
    for (const probe of probes) {
      const from = Math.round(probe.read.at - probe.read.half);
      const to = Math.round(probe.read.at + probe.read.half);
      const ours = await rowsAcross(markup, 960, 540, [200, 640], [from, to]);
      const seen = measured[probe.id]!['960']!;
      const theirs = new Array<number>(to - from + 1).fill(0);
      seen.cover.forEach((v, i) => {
        theirs[seen.from - from + i] = v;
      });
      if (Math.max(...ours.map((v, i) => Math.abs(v - theirs[i]!))) > snap.tolerance) misses += 1;
    }
    // Every odd pen, and every even pen off the grid: 25 of 28.
    expect(misses).toBe(25);
  });
});

describe('the outline band, re-derived', () => {
  const WIDTH = 152400;
  const LINE = `<a:ln w="${String(WIDTH)}"><a:solidFill><a:srgbClr val="FF00FF"/></a:solidFill></a:ln>`;
  const RECT = { x: 0, y: 0, cx: 1371600, cy: 1371600 };

  function pic(): string {
    const id = nextId++;
    return (
      `<p:pic><p:nvPicPr><p:cNvPr id="${String(id)}" name="picture"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>` +
      '<p:blipFill><a:blip r:embed="rId9"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>' +
      `<p:spPr><a:xfrm><a:off x="${String(RECT.x)}" y="${String(RECT.y)}"/>` +
      `<a:ext cx="${String(RECT.cx)}" cy="${String(RECT.cy)}"/></a:xfrm>` +
      `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>${LINE}</p:spPr></p:pic>`
    );
  }

  const band = pictures.outlineBands;

  it('draws a one-sided band at double width, so half of it survives the clip', () => {
    expect(band.picture.outsidePt + band.picture.insidePt).toBe(pictures.outlinePt);
    const markup = renderSlide(buildChain({ shapes: [pic()] }).slide, SIZE, { idPrefix: 'o' });
    expect(markup).toContain(`stroke-width="${String(WIDTH * 2)}"`);
  });

  it('clips a picture band to the complement of its outline', () => {
    const markup = renderSlide(buildChain({ shapes: [pic()] }).slide, SIZE, { idPrefix: 'o' });
    // Nothing of the band falls inside the picture, so the clip is everything
    // the outline does not cover: one path, two subpaths, evenodd.
    expect(band.picture.insidePt).toBe(0);
    expect(markup).toContain('clip-rule="evenodd"');
    expect(markup).toContain('clipPath');
  });

  it('keeps the picture under a clipped band: the fill is one path and the band another', async () => {
    // A one-pixel red PNG, so the picture is a colour and not a gap.
    const red =
      'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8DwHwAFBQIAX8jx0gAAAABJRU5ErkJggg==';
    const media = (): { bytes: Uint8Array; contentType: string } => ({
      bytes: Uint8Array.from(atob(red), (c) => c.charCodeAt(0)),
      contentType: 'image/png',
    });
    const width = 480;
    const markup = renderSlide(buildChain({ shapes: [pic()] }).slide, SIZE, {
      idPrefix: 'k',
      media,
      width,
      height: 270,
    });
    const paths = [...markup.matchAll(/<path [^>]*>/g)].map((m) => m[0]);
    expect(paths).toHaveLength(3);
    expect(paths[1]).toContain('fill="url(#k-1)"');
    expect(paths[1]).not.toContain('clip-path');
    expect(paths[2]).toContain('fill="none"');
    expect(paths[2]).toContain('clip-path="url(#k-2)"');
    // And the pixel inside the frame is the picture's, not the page's white.
    const image = new Image();
    const href = URL.createObjectURL(new Blob([markup], { type: 'image/svg+xml' }));
    image.src = href;
    await image.decode();
    const canvas = document.createElement('canvas');
    canvas.width = width;
    canvas.height = 270;
    const context = canvas.getContext('2d', { willReadFrequently: true });
    if (context === null) throw new Error('no 2d context');
    context.fillStyle = '#FFFFFF';
    context.fillRect(0, 0, width, 270);
    context.drawImage(image, 0, 0, width, 270);
    URL.revokeObjectURL(href);
    // The picture is 108 pt square at 0.5 px/pt; sample its centre.
    const [r, g, b] = context.getImageData(27, 27, 1, 1).data;
    expect([r, g, b]).toEqual([255, 0, 0]);
  });

  it('leaves a shape band centred, with no clip at all', () => {
    expect(band.shape.insidePt).toBe(pictures.outlinePt / 2);
    expect(band.shape.outsidePt).toBe(pictures.outlinePt / 2);
    const markup = renderSlide(
      buildChain({
        shapes: [
          sp({
            rect: RECT,
            fill: '<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>',
            line: LINE,
          }),
        ],
      }).slide,
      SIZE,
      { idPrefix: 'c' },
    );
    expect(markup).toContain(`stroke-width="${String(WIDTH)}"`);
    expect(markup).not.toContain('clip-rule="evenodd"');
  });

  it('still clips an algn="in" shape band to the outline itself', () => {
    const inset = `<a:ln w="${String(WIDTH)}" algn="in"><a:solidFill><a:srgbClr val="FF00FF"/></a:solidFill></a:ln>`;
    const markup = renderSlide(
      buildChain({
        shapes: [
          sp({
            rect: RECT,
            fill: '<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>',
            line: inset,
          }),
        ],
      }).slide,
      SIZE,
      { idPrefix: 'i' },
    );
    expect(markup).toContain(`stroke-width="${String(WIDTH * 2)}"`);
    expect(markup).toContain('clipPath');
    expect(markup).not.toContain('clip-rule="evenodd"');
  });
});

/* -------------------------------------------------------------------------- */

/**
 * The `a:blip` colour effects, as arithmetic rather than as markup.
 *
 * `blips.json` states the rules in words - duotone interpolates linearly
 * between two colours by the BT.709 luminance, `a:clrChange` matches a colour
 * exactly - so these evaluate the emitted filter and check it computes that.
 * Asserting the attribute text instead would have passed while both of these
 * were emitting a colour 255 times too dark. ADR 0038.
 */
describe('the blip colour effects, evaluated', () => {
  const BOX = { x: 0, y: 0, cx: 100, cy: 100 };
  const IMAGE = { bytes: new Uint8Array([0]), contentType: 'image/png' };

  function filterMarkup(effects: readonly BlipEffect[]): string {
    const defs = new Defs('b');
    blipPaint(
      {
        type: 'blip',
        embed: 'rId1',
        svgEmbed: null,
        part: 'ppt/slides/slide1.xml',
        srcRect: { l: 0, t: 0, r: 0, b: 0 },
        mode: { kind: 'stretch', fillRect: { l: 0, t: 0, r: 0, b: 0 } },
        effects,
        dpi: 0,
        rotWithShape: false,
      },
      {},
      BOX,
      defs,
      () => IMAGE,
      () => ({ widthPx: 32, heightPx: 32, dpi: 96 }),
      () => 'data:image/png;base64,AA==',
    );
    const node = defs.toNode();
    expect(node).not.toBeNull();
    return node === null ? '' : serializeSvg(node);
  }

  /** `slope` and `intercept` off one `feFunc`, which is what the filter runs. */
  function transfer(markup: string, channel: 'R' | 'G' | 'B'): [number, number] {
    const found = new RegExp(
      `<feFunc${channel} type="linear" slope="([-0-9.e]+)" intercept="([-0-9.e]+)"`,
    ).exec(markup);
    expect(found, `no linear feFunc${channel} in ${markup.slice(0, 200)}`).not.toBeNull();
    return [Number(found?.[1]), Number(found?.[2])];
  }

  const srgb = (hex: string): Color => ({ space: 'srgb', hex, transforms: [] });

  it('interpolates a duotone between its two colours, in 0..1 space', () => {
    // Black to FFC000, which is the pair `a03-fills-02` carries.
    const markup = filterMarkup([{ kind: 'duotone', from: srgb('000000'), to: srgb('FFC000') }]);

    // `feFunc` is evaluated as `slope * luma + intercept` on 0..1 channels, so
    // luma 0 must give the `from` colour and luma 1 the `to` colour.
    for (const [channel, level] of [
      ['R', 0xff],
      ['G', 0xc0],
      ['B', 0x00],
    ] as const) {
      const [slope, intercept] = transfer(markup, channel);
      expect(intercept * 255).toBeCloseTo(0, 6);
      // Within half a level: the attribute itself is written rounded.
      expect((slope * 1 + intercept) * 255).toBeCloseTo(level, 0);
    }
  });

  it('refutes the reading that divides the endpoints by 255 again', () => {
    const markup = filterMarkup([{ kind: 'duotone', from: srgb('000000'), to: srgb('FFC000') }]);
    const [slope] = transfer(markup, 'R');
    // The bug this test exists for: a slope of 1/255 renders the whole image
    // black, and every attribute in the filter still looks well formed.
    expect(slope).not.toBeCloseTo(1 / 255, 6);
    expect(slope).toBeCloseTo(1, 6);
  });

  it('matches a:clrChange at the source colour, not at level zero', () => {
    const markup = filterMarkup([
      { kind: 'clrChange', from: srgb('F1C40F'), to: srgb('000000'), useAlpha: true },
    ]);
    const table = /<feFuncR type="discrete" tableValues="([^"]+)"/.exec(markup)?.[1] ?? '';
    const ones = table.split(' ').flatMap((value, at) => (value === '1' ? [at] : []));
    // 0xF1 is 241; rounding a 0..1 channel instead lands on 0 or 1.
    expect(ones).toStrictEqual([0xf1]);
  });
});

/**
 * A shape's effects are a filter in points. Chromium rasterises a rotated filter at the
 * resolution of its user space: the same glow written in EMU took nine seconds a shape on
 * a 120-pixel thumbnail and thirty milliseconds in points (ADR 0054).
 */
describe('effects, in point space', () => {
  const GLOW_EMU = 76200;
  const GLOW =
    `<a:effectLst><a:glow rad="${String(GLOW_EMU)}"><a:srgbClr val="4472C4"><a:alpha val="40000"/>` +
    '</a:srgbClr></a:glow></a:effectLst>';
  const RECT = { x: 4572000, y: 2286000, cx: 3048000, cy: 2286000 };

  function glowed(rot: number): string {
    return renderSlide(
      buildChain({ shapes: [sp({ rect: RECT, rot, fill: SOLID('4472C4'), line: GLOW })] }).slide,
      SIZE,
      { idPrefix: 't' },
    );
  }

  it('writes the filter in points, inside a scale pair around the paths', () => {
    const markup = glowed(15);
    // The dilate is what `paint` asked for, 12700 times smaller: 76200 EMU is 6 pt.
    const graph = effectFilter(
      [{ kind: 'glow', rad: GLOW_EMU, color: { space: 'srgb', hex: '4472C4', transforms: [] } }],
      { x: 0, y: 0, w: RECT.cx, h: RECT.cy },
      () => ({ css: '#4472C4', alpha: 0.4 }),
    );
    const dilate = graph.primitives.find((p) => p.op === 'morphology');
    if (dilate?.op !== 'morphology') throw new Error('no dilate');
    expect(dilate.radius).toBeLessThan(GLOW_EMU);
    expect(markup).toContain(`operator="dilate" radius="${num(dilate.radius / 12700)}"`);
    expect(markup).toContain(
      `<filter id="t-1" filterUnits="userSpaceOnUse" x="${num(-graph.margin.left / 12700)}"`,
    );
    const region = /<filter [^>]*width="([0-9.]+)"/.exec(markup);
    expect(Number(region?.[1])).toBeCloseTo(
      (RECT.cx + graph.margin.left + graph.margin.right) / 12700,
      3,
    );
    const into = /transform="scale\(([0-9.]+)\)" filter="url\(#t-1\)"/.exec(markup);
    const outOf = /<g transform="scale\(([0-9.e-]+)\)"><path /.exec(markup);
    expect(into?.[1]).toBe('12700');
    expect(Number(into?.[1]) * Number(outOf?.[1])).toBeCloseTo(1, 12);
  });

  it('leaves a shape without effects alone', () => {
    const markup = renderSlide(
      buildChain({ shapes: [sp({ rect: RECT, rot: 15, fill: SOLID('4472C4') })] }).slide,
      SIZE,
      { idPrefix: 't' },
    );
    expect(markup).not.toContain('<filter');
    expect(markup).not.toContain('scale(12700)');
  });

  it('draws the glow beside the shape, and the shape where it was', async () => {
    const width = 960;
    const height = 540;
    const image = new Image();
    const href = URL.createObjectURL(new Blob([glowed(15)], { type: 'image/svg+xml' }));
    image.src = href;
    await image.decode();
    const canvas = document.createElement('canvas');
    canvas.width = width;
    canvas.height = height;
    const context = canvas.getContext('2d', { willReadFrequently: true });
    if (context === null) throw new Error('no 2d context');
    context.fillStyle = '#FFFFFF';
    context.fillRect(0, 0, width, height);
    const started = performance.now();
    context.drawImage(image, 0, 0, width, height);
    const pixel = (x: number, y: number): number[] => [...context.getImageData(x, y, 1, 1).data];
    URL.revokeObjectURL(href);
    // The centre is the fill; three points above the turned top edge is glow; the corner is not.
    const centre = pixel(480, 270);
    expect(Math.abs((centre[0] ?? 0) - 0x44) + Math.abs((centre[2] ?? 0) - 0xc4)).toBeLessThan(4);
    const above = pixel(480, 174);
    expect(above[2]).toBeGreaterThan((above[0] ?? 0) + 20);
    expect(pixel(2, 2).slice(0, 3)).toEqual([255, 255, 255]);
    expect(performance.now() - started).toBeLessThan(3000);
  });
});
