import { describe, expect, it } from 'vitest';
import { arcEnd, type ArcParameters } from './arc.js';
import { GeometryError } from '../errors.js';
import { evaluateGuides, type ShapeSize } from '../formula/evaluate.js';
import {
  DEFAULT_PRECISION,
  pathData,
  pathScale,
  resolveGeometry,
  resolvePath,
  type PathSegment,
  type ResolvedPath,
} from './path.js';
import { getPreset, presetNames } from '../presets/index.js';
import {
  custGeom,
  type Geometry,
  type Point,
  type PresetCommand,
  type PresetPath,
  type PresetPoint,
  type PresetShape,
} from '../types.js';

/**
 * Two things are being checked here and they are not the same thing.
 *
 * The first is that a path comes out where the shape says it should. The
 * presets are full of guides that state, in shape coordinates, a point the path
 * reaches in its own coordinates - `flowChartDisplay`'s text rectangle is the
 * two ends of its flat top edge, `flowChartMagneticDisk`'s is the bottom of its
 * top ellipse - so path space can be checked against arithmetic that knows
 * nothing about it.
 *
 * The second is that `a:prstGeom` and `a:custGeom` really are one thing. Every
 * preset is rewritten into an eight-times path space, which is a legal
 * `custGeom` expressible in `fmla` alone, and asked to emit the same string.
 */

function preset(name: string): PresetShape {
  const shape = getPreset(name);
  if (shape === undefined) throw new Error(`${name} is not a preset`);
  return shape;
}

function value(guides: ReadonlyMap<string, number>, name: string): number {
  const found = guides.get(name);
  if (found === undefined) throw new Error(`guide ${name} was never defined`);
  return found;
}

function pathAt(
  resolved: { readonly paths: readonly ResolvedPath[] },
  index: number,
): ResolvedPath {
  const found = resolved.paths[index];
  if (found === undefined) throw new Error(`no path at ${index}`);
  return found;
}

function segmentAt(path: ResolvedPath, index: number): PathSegment {
  const found = path.segments[index];
  if (found === undefined) throw new Error(`no segment at ${index}`);
  return found;
}

/** The point a segment leaves the pen at, or null for a `close`. */
function endOf(segment: PathSegment): Point | null {
  return segment.kind === 'close' ? null : segment.to;
}

const SIZES: readonly ShapeSize[] = [
  { w: 200, h: 100 },
  { w: 100, h: 100 },
  { w: 320, h: 90 },
  { w: 96, h: 512 },
];

describe('pathData', () => {
  const move: PathSegment = { kind: 'move', to: { x: 1, y: 2 } };

  it('emits one letter per command', () => {
    expect(
      pathData([
        move,
        { kind: 'line', to: { x: 3, y: 4 } },
        { kind: 'quad', c1: { x: 5, y: 6 }, to: { x: 7, y: 8 } },
        { kind: 'cubic', c1: { x: 9, y: 10 }, c2: { x: 11, y: 12 }, to: { x: 13, y: 14 } },
        {
          kind: 'arc',
          from: { x: 13, y: 14 },
          to: { x: 15, y: 16 },
          rx: 17,
          ry: 18,
          largeArc: true,
          sweep: false,
        },
        { kind: 'close' },
      ]),
    ).toBe('M1 2L3 4Q5 6 7 8C9 10 11 12 13 14A17 18 0 1 0 15 16Z');
  });

  it('rounds to the requested precision and drops trailing zeros', () => {
    const segments: PathSegment[] = [{ kind: 'move', to: { x: 1 / 3, y: 2 } }];
    expect(pathData(segments)).toBe('M0.3333 2');
    expect(pathData(segments, 1)).toBe('M0.3 2');
    expect(pathData(segments, 0)).toBe('M0 2');
    expect(DEFAULT_PRECISION).toBe(4);
  });

  it('never writes a negative zero', () => {
    // `-0` is legal SVG but it is noise in a diff, and it makes two paths that
    // are the same path compare unequal.
    expect(pathData([{ kind: 'move', to: { x: -0, y: -0.000001 } }])).toBe('M0 0');
  });

  it('is deterministic enough to compare two paths for equality', () => {
    const a = resolveGeometry(preset('triangle'), { w: 200, h: 100 });
    const b = resolveGeometry(preset('triangle'), { w: 200, h: 100 });
    expect(pathAt(a, 0).d).toBe(pathAt(b, 0).d);
  });
});

describe('path space', () => {
  const square: PresetPath = {
    w: 0,
    h: 0,
    fill: 'norm',
    stroke: true,
    extrusionOk: true,
    commands: [
      { kind: 'moveTo', to: { x: '0', y: '0' } },
      { kind: 'lnTo', to: { x: '10', y: '0' } },
      { kind: 'lnTo', to: { x: '10', y: '10' } },
      { kind: 'close' },
    ],
  };
  const guides = new Map<string, number>();
  const size: ShapeSize = { w: 40, h: 80 };

  it('treats zero as "no path space", not as a divisor', () => {
    expect(pathScale({ ...square, w: 0, h: 0 }, size)).toStrictEqual({ x: 1, y: 1 });
    const unscaled = resolvePath(square, guides, size);
    expect(unscaled.d).toBe('M0 0L10 0L10 10Z');
    expect(unscaled.finite).toBe(true);
  });

  it('scales each axis independently', () => {
    expect(pathScale({ ...square, w: 10, h: 0 }, size)).toStrictEqual({ x: 4, y: 1 });
    expect(pathScale({ ...square, w: 0, h: 10 }, size)).toStrictEqual({ x: 1, y: 8 });
    expect(resolvePath({ ...square, w: 10, h: 0 }, guides, size).d).toBe('M0 0L40 0L40 10Z');
    expect(resolvePath({ ...square, w: 0, h: 10 }, guides, size).d).toBe('M0 0L10 0L10 80Z');
  });

  it('a path space equal to the shape size is the same as none at all', () => {
    expect(resolvePath({ ...square, w: size.w, h: size.h }, guides, size).d).toBe(
      resolvePath(square, guides, size).d,
    );
  });

  it('is the identity when the coordinates are scaled with the space', () => {
    const doubled: PresetPath = {
      ...square,
      w: 2 * size.w,
      h: 2 * size.h,
      commands: square.commands.map((command) =>
        command.kind === 'lnTo' || command.kind === 'moveTo'
          ? {
              ...command,
              to: { x: String(Number(command.to.x) * 2), y: String(Number(command.to.y) * 2) },
            }
          : command,
      ),
    };
    expect(resolvePath(doubled, guides, size).d).toBe(
      resolvePath({ ...square, w: size.w, h: size.h }, guides, size).d,
    );
  });

  it('scales an arc radius per axis and leaves its angles alone', () => {
    const arc: PresetPath = {
      w: 100,
      h: 100,
      fill: 'norm',
      stroke: true,
      extrusionOk: true,
      commands: [
        { kind: 'moveTo', to: { x: '50', y: '0' } },
        { kind: 'arcTo', wR: '50', hR: '50', stAng: '16200000', swAng: '10800000' },
      ],
    };
    const resolved = resolvePath(arc, guides, { w: 400, h: 100 });
    const segment = segmentAt(resolved, 1);
    if (segment.kind !== 'arc') throw new Error('expected an arc');
    expect(segment.rx).toBeCloseTo(200, 9);
    expect(segment.ry).toBeCloseTo(50, 9);
  });

  it('counts the presets that declare a path space', () => {
    let scaled = 0;
    let unscaled = 0;
    const spaces = new Set<string>();
    for (const name of presetNames()) {
      for (const path of preset(name).pathLst) {
        if (path.w === 0 && path.h === 0) unscaled += 1;
        else {
          scaled += 1;
          spaces.add(`${path.w}x${path.h}`);
        }
      }
    }
    expect([scaled, unscaled, scaled + unscaled]).toStrictEqual([50, 270, 320]);
    expect([...spaces].sort()).toStrictEqual([
      '10x10',
      '1x1',
      '20x20',
      '21600x21600',
      '2x2',
      '43200x43200',
      '5x5',
      '6x6',
      '8x8',
    ]);
  });

  it('finds no positional operand in a scaled path that is not a literal', () => {
    // The reason the "what does a guide mean in path space" question is a
    // decision rather than a finding: it is unreachable from the corpus.
    const literal = /^[-+]?\d+$/;
    const positional = new Set<string>();
    const angular = new Set<string>();
    for (const name of presetNames()) {
      for (const path of preset(name).pathLst) {
        if (path.w === 0 && path.h === 0) continue;
        for (const command of path.commands) {
          switch (command.kind) {
            case 'moveTo':
            case 'lnTo':
              for (const t of [command.to.x, command.to.y]) if (!literal.test(t)) positional.add(t);
              break;
            case 'quadBezTo':
              for (const t of [command.c1.x, command.c1.y, command.to.x, command.to.y])
                if (!literal.test(t)) positional.add(t);
              break;
            case 'cubicBezTo':
              for (const t of [
                command.c1.x,
                command.c1.y,
                command.c2.x,
                command.c2.y,
                command.to.x,
                command.to.y,
              ])
                if (!literal.test(t)) positional.add(t);
              break;
            case 'arcTo':
              for (const t of [command.wR, command.hR]) if (!literal.test(t)) positional.add(t);
              for (const t of [command.stAng, command.swAng]) if (!literal.test(t)) angular.add(t);
              break;
            case 'close':
              break;
          }
        }
      }
    }
    expect([...positional]).toStrictEqual([]);
    expect([...angular].sort()).toStrictEqual(['3cd4', 'cd2', 'cd4']);
  });
});

describe('the presets state the path-space rule themselves', () => {
  /**
   * The centre of a half-turn arc, without asking the module under test.
   *
   * A 180 degree sweep puts the two endpoints on opposite ends of a diameter,
   * so their midpoint is the centre whatever the parameterisation.
   */
  function centreOfHalfTurn(segment: PathSegment): Point {
    if (segment.kind !== 'arc') throw new Error('expected an arc');
    return { x: (segment.from.x + segment.to.x) / 2, y: (segment.from.y + segment.to.y) / 2 };
  }

  function residual(point: Point, centre: Point, rx: number, ry: number): number {
    const dx = (point.x - centre.x) / rx;
    const dy = (point.y - centre.y) / ry;
    return Math.abs(dx * dx + dy * dy - 1);
  }

  it('flowChartDisplay: the text rectangle is the two ends of the flat top edge', () => {
    // Path space 6 x 6, top edge from x=1 to x=5. The guides say w/6 and 5w/6.
    for (const size of SIZES) {
      const resolved = resolveGeometry(preset('flowChartDisplay'), size);
      const path = pathAt(resolved, 0);
      const left = segmentAt(path, 1);
      const right = segmentAt(path, 2);
      if (left.kind !== 'line' || right.kind !== 'line') throw new Error('expected lines');
      expect(left.to.x).toBeCloseTo(value(resolved.guides, 'wd6'), 9);
      expect(right.to.x).toBeCloseTo(value(resolved.guides, 'x2'), 9);
      expect(resolved.textRect?.l).toBeCloseTo(left.to.x, 9);
      expect(resolved.textRect?.r).toBeCloseTo(right.to.x, 9);
    }
  });

  it('flowChartDisplay: the worked example in the README', () => {
    const resolved = resolveGeometry(preset('flowChartDisplay'), { w: 120, h: 60 });
    expect(resolved.paths.length).toBe(1);
    expect(pathAt(resolved, 0).d).toBe('M0 30L20 0L100 0A20 30 0 0 1 100 60L20 60Z');
    expect(pathAt(resolved, 0).fill).toBe('norm');
    expect(resolved.textRect).toStrictEqual({ l: 20, t: 0, r: 100, b: 60 });
  });

  it('flowChartMagneticDisk: the text rectangle sits on the bottom of the top ellipse', () => {
    // Path space 6 x 6. `y3 = 5h/6` is where `L 6,5` lands; `hd3` is the full
    // height of the arc whose vertical radius is one unit of six.
    for (const size of SIZES) {
      const resolved = resolveGeometry(preset('flowChartMagneticDisk'), size);
      const path = pathAt(resolved, 0);
      const cap = segmentAt(path, 1);
      const side = segmentAt(path, 2);
      if (cap.kind !== 'arc' || side.kind !== 'line') throw new Error('expected an arc and a line');
      expect(side.to.y).toBeCloseTo(value(resolved.guides, 'y3'), 9);
      expect(side.to.x).toBeCloseTo(size.w, 9);
      expect(cap.ry * 2).toBeCloseTo(value(resolved.guides, 'hd3'), 9);
      expect(resolved.textRect?.t).toBeCloseTo(cap.ry * 2, 9);
      expect(centreOfHalfTurn(cap).y).toBeCloseTo(cap.ry, 9);
    }
  });

  it('flowChartPunchedTape: the text rectangle sits below the scallops', () => {
    // Path space 20 x 20, scallop radius 2, `L 20,18`.
    for (const size of SIZES) {
      const resolved = resolveGeometry(preset('flowChartPunchedTape'), size);
      const path = pathAt(resolved, 0);
      const scallop = segmentAt(path, 1);
      if (scallop.kind !== 'arc') throw new Error('expected an arc');
      expect(scallop.ry * 2).toBeCloseTo(value(resolved.guides, 'hd5'), 9);
      expect(resolved.textRect?.t).toBeCloseTo(scallop.ry * 2, 9);
      const side = path.segments.find((s) => s.kind === 'line');
      expect(side?.kind).toBe('line');
      if (side?.kind === 'line') {
        expect(side.to.y).toBeCloseTo(value(resolved.guides, 'y2'), 9);
      }
    }
  });

  it('flowChartTerminator: all four text-rectangle corners lie on the end caps', () => {
    // The strongest of the four, because the numbers can only have come from
    // the arc: `il = w * 1018/21600` and `it = h * 3163/21600` put the corner
    // at 45 degrees round the cap, where the offsets are the radii over
    // root two. 1018/3475 = 0.29295 against 1 - 1/sqrt(2) = 0.29289.
    for (const size of SIZES) {
      const resolved = resolveGeometry(preset('flowChartTerminator'), size);
      const path = pathAt(resolved, 0);
      const rect = resolved.textRect;
      if (rect === null) throw new Error('flowChartTerminator has a text rectangle');

      const right = segmentAt(path, 2);
      const left = segmentAt(path, 4);
      if (right.kind !== 'arc' || left.kind !== 'arc') throw new Error('expected two arcs');

      for (const [cap, x] of [
        [right, rect.r],
        [left, rect.l],
      ] as const) {
        const centre = centreOfHalfTurn(cap);
        expect(centre.y).toBeCloseTo(size.h / 2, 9);
        expect(cap.ry).toBeCloseTo(size.h / 2, 9);
        for (const y of [rect.t, rect.b]) {
          expect(residual({ x, y }, centre, cap.rx, cap.ry)).toBeLessThan(2e-4);
        }
      }
    }
  });
});

describe('cloud settles where an arc is resolved', () => {
  /**
   * The alternative reading, written out so the test can show what it costs:
   * scale the radii into shape space first and resolve the arc there.
   */
  function ringGapTheOtherWay(name: string, size: ShapeSize): number {
    const shape = preset(name);
    const path = shape.pathLst[0];
    if (path === undefined) throw new Error('no path');
    const guides = evaluateGuides(shape, size);
    const sx = size.w / path.w;
    const sy = size.h / path.h;
    const at = (token: string): number => {
      const literal = Number(token);
      return Number.isNaN(literal) ? value(guides, token) : literal;
    };
    let cursor: Point = { x: 0, y: 0 };
    let first: Point | null = null;
    for (const command of path.commands) {
      if (command.kind === 'moveTo') {
        cursor = { x: at(command.to.x) * sx, y: at(command.to.y) * sy };
        first ??= cursor;
      } else if (command.kind === 'arcTo') {
        const arc: ArcParameters = {
          wR: at(command.wR) * sx,
          hR: at(command.hR) * sy,
          stAng: at(command.stAng),
          swAng: at(command.swAng),
        };
        cursor = arcEnd(cursor, arc);
      }
    }
    if (first === null) throw new Error('no move');
    return Math.hypot(cursor.x - first.x, cursor.y - first.y) / Math.max(size.w, size.h);
  }

  function ringGap(name: string, size: ShapeSize): number {
    const path = pathAt(resolveGeometry(preset(name), size), 0);
    const start = endOf(segmentAt(path, 0));
    let last: Point | null = null;
    for (const segment of path.segments) last = endOf(segment) ?? last;
    if (start === null || last === null) throw new Error('no endpoints');
    return Math.hypot(last.x - start.x, last.y - start.y) / Math.max(size.w, size.h);
  }

  it('closes its ring of eleven arcs at every aspect ratio', () => {
    // 0.32% is the residue of the integer angles the file was authored with,
    // and it is the whole of the error at a square size.
    for (const name of ['cloud', 'cloudCallout']) {
      for (const size of [
        { w: 100, h: 100 },
        { w: 200, h: 100 },
        { w: 400, h: 100 },
        { w: 100, h: 800 },
      ]) {
        expect(ringGap(name, size)).toBeLessThan(0.004);
      }
    }
  });

  it('and the other order tears it open the moment the shape is not square', () => {
    expect(ringGapTheOtherWay('cloud', { w: 100, h: 100 })).toBeCloseTo(
      ringGap('cloud', { w: 100, h: 100 }),
      12,
    );
    for (const [size, floor] of [
      [{ w: 200, h: 100 }, 0.05],
      [{ w: 400, h: 100 }, 0.03],
      [{ w: 100, h: 800 }, 0.04],
    ] as const) {
      expect(ringGapTheOtherWay('cloud', size)).toBeGreaterThan(floor);
      expect(ringGap('cloud', size)).toBeLessThan(0.004);
    }
  });
});

/**
 * Rewrite a geometry so every path is drawn in a space `factor` times the shape
 * size - a legal `custGeom`, expressible in `fmla` alone.
 *
 * This is what a "convert to freeform" would have to be able to do without
 * changing the picture. Every positional operand gets a guide that multiplies
 * it, the angles are left alone because they belong to no coordinate system,
 * and `rect` and `cxnLst` are untouched because they are in shape space and
 * always were.
 */
function toPathSpace(shape: Geometry, size: ShapeSize, factor: number): Geometry {
  const gdLst = [...shape.gdLst];
  const cache = new Map<string, string>();
  const scaled = (token: string): string => {
    const seen = cache.get(token);
    if (seen !== undefined) return seen;
    const name = `zzPathSpace${cache.size}`;
    cache.set(token, name);
    gdLst.push({ name, fmla: ['*/', token, String(factor), '1'] });
    return name;
  };
  const point = (p: PresetPoint): PresetPoint => ({ x: scaled(p.x), y: scaled(p.y) });
  const rewrite = (command: PresetCommand): PresetCommand => {
    switch (command.kind) {
      case 'moveTo':
        return { kind: 'moveTo', to: point(command.to) };
      case 'lnTo':
        return { kind: 'lnTo', to: point(command.to) };
      case 'quadBezTo':
        return { kind: 'quadBezTo', c1: point(command.c1), to: point(command.to) };
      case 'cubicBezTo':
        return {
          kind: 'cubicBezTo',
          c1: point(command.c1),
          c2: point(command.c2),
          to: point(command.to),
        };
      case 'arcTo':
        return {
          kind: 'arcTo',
          wR: scaled(command.wR),
          hR: scaled(command.hR),
          stAng: command.stAng,
          swAng: command.swAng,
        };
      case 'close':
        return { kind: 'close' };
    }
  };

  const pathLst = shape.pathLst.map((path) => ({
    ...path,
    w: (path.w === 0 ? size.w : path.w) * factor,
    h: (path.h === 0 ? size.h : path.h) * factor,
    commands: path.commands.map(rewrite),
  }));

  return custGeom({
    avLst: shape.avLst,
    gdLst,
    ahLst: shape.ahLst,
    cxnLst: shape.cxnLst,
    rect: shape.rect,
    pathLst,
  });
}

describe('custGeom is a prstGeom written inline', () => {
  it('resolves a geometry that has no name', () => {
    const inline = custGeom({
      gdLst: [{ name: 'half', fmla: ['*/', 'w', '1', '2'] }],
      rect: { l: 'l', t: 't', r: 'half', b: 'b' },
      pathLst: [
        {
          w: 0,
          h: 0,
          fill: 'norm',
          stroke: true,
          extrusionOk: true,
          commands: [
            { kind: 'moveTo', to: { x: 'l', y: 't' } },
            { kind: 'lnTo', to: { x: 'half', y: 'b' } },
            { kind: 'close' },
          ],
        },
      ],
    });
    const resolved = resolveGeometry(inline, { w: 200, h: 100 });
    expect(resolved.name).toBeNull();
    expect(pathAt(resolved, 0).d).toBe('M0 0L100 100Z');
    expect(resolved.textRect).toStrictEqual({ l: 0, t: 0, r: 100, b: 100 });
  });

  it('fills in the children an a:custGeom is allowed to omit', () => {
    const empty = custGeom();
    expect(empty).toStrictEqual({
      name: null,
      avLst: [],
      gdLst: [],
      ahLst: [],
      cxnLst: [],
      rect: null,
      pathLst: [],
    });
    const resolved = resolveGeometry(empty, { w: 10, h: 10 });
    expect(resolved.paths).toStrictEqual([]);
    expect(resolved.textRect).toBeNull();
    expect(resolved.connectionSites).toStrictEqual([]);
  });

  it('says which shape an operand failed in, and says nothing when there is none', () => {
    const bad: PresetPath[] = [
      {
        w: 0,
        h: 0,
        fill: 'norm',
        stroke: true,
        extrusionOk: true,
        commands: [{ kind: 'moveTo', to: { x: 'nope', y: 't' } }],
      },
    ];
    const anonymous = (): unknown => resolveGeometry(custGeom({ pathLst: bad }), { w: 10, h: 10 });
    expect(anonymous).toThrow(GeometryError);
    expect(anonymous).toThrow(/^operand "nope"/);

    const named = (): unknown =>
      resolveGeometry(custGeom({ name: 'myShape', pathLst: bad }), { w: 10, h: 10 });
    expect(named).toThrow(/^myShape: operand "nope"/);
  });

  it('takes a preset and a custGeom down the same code path', () => {
    for (const name of presetNames()) {
      const shape = preset(name);
      const inline = custGeom({ ...shape, name: null });
      for (const size of SIZES) {
        const fromPreset = resolveGeometry(shape, size);
        const fromInline = resolveGeometry(inline, size);
        expect(fromInline.name).toBeNull();
        expect(fromInline.paths.map((p) => p.d)).toStrictEqual(fromPreset.paths.map((p) => p.d));
        expect(fromInline.textRect).toStrictEqual(fromPreset.textRect);
        expect(fromInline.connectionSites).toStrictEqual(fromPreset.connectionSites);
      }
    }
  });
});

describe('the conformance test: every preset rewritten into a path space', () => {
  const FACTOR = 8;
  const FIRST = SIZES[0] as ShapeSize;

  it('generates guide names that collide with nothing', () => {
    for (const name of presetNames()) {
      const shape = preset(name);
      const guides = evaluateGuides(shape, FIRST);
      for (const gd of toPathSpace(shape, FIRST, FACTOR).gdLst) {
        if (!gd.name.startsWith('zzPathSpace')) continue;
        expect(guides.has(gd.name)).toBe(false);
      }
    }
  });

  it('emits the same path data as the preset it came from', () => {
    let compared = 0;
    for (const name of presetNames()) {
      const shape = preset(name);
      for (const size of SIZES) {
        const direct = resolveGeometry(shape, size);
        const inline = resolveGeometry(toPathSpace(shape, size, FACTOR), size);
        expect(inline.paths.length).toBe(direct.paths.length);
        direct.paths.forEach((path, index) => {
          expect(`${name}[${index}] ${pathAt(inline, index).d}`).toBe(
            `${name}[${index}] ${path.d}`,
          );
          compared += 1;
        });
      }
    }
    expect(compared).toBe(320 * SIZES.length);
  });

  it('and puts every segment in the same place, not merely the rounded string', () => {
    let worst = 0;
    for (const name of presetNames()) {
      const shape = preset(name);
      for (const size of SIZES) {
        const direct = resolveGeometry(shape, size);
        const inline = resolveGeometry(toPathSpace(shape, size, FACTOR), size);
        direct.paths.forEach((path, index) => {
          const other = pathAt(inline, index);
          expect(other.segments.length).toBe(path.segments.length);
          path.segments.forEach((segment, at) => {
            const mirror = other.segments[at];
            if (mirror === undefined) throw new Error('length was already checked');
            expect(mirror.kind).toBe(segment.kind);
            const a = endOf(segment);
            const b = endOf(mirror);
            if (a === null || b === null) return;
            worst = Math.max(worst, Math.abs(a.x - b.x), Math.abs(a.y - b.y));
          });
        });
      }
    }
    expect(worst).toBeLessThan(1e-9);
  });

  it('leaves the text rectangle and the connection sites in shape space', () => {
    const size = { w: 320, h: 90 };
    for (const name of presetNames()) {
      const shape = preset(name);
      const direct = resolveGeometry(shape, size);
      const inline = resolveGeometry(toPathSpace(shape, size, FACTOR), size);
      expect(inline.textRect).toStrictEqual(direct.textRect);
      expect(inline.connectionSites).toStrictEqual(direct.connectionSites);
    }
  });
});

describe('close', () => {
  const size: ShapeSize = { w: 100, h: 100 };
  const guides = new Map<string, number>();

  it('returns the pen to the start of the subpath and carries on drawing', () => {
    const path: PresetPath = {
      w: 0,
      h: 0,
      fill: 'norm',
      stroke: true,
      extrusionOk: true,
      commands: [
        { kind: 'moveTo', to: { x: '5', y: '5' } },
        { kind: 'lnTo', to: { x: '40', y: '5' } },
        { kind: 'close' },
        { kind: 'arcTo', wR: '10', hR: '10', stAng: '0', swAng: '5400000' },
      ],
    };
    const resolved = resolvePath(path, guides, size);
    const arc = segmentAt(resolved, 3);
    if (arc.kind !== 'arc') throw new Error('expected an arc');
    // The arc begins where the closed subpath began, not where the line ended.
    expect(arc.from).toStrictEqual({ x: 5, y: 5 });
    expect(resolved.d.startsWith('M5 5L40 5ZA10 10 0 0 1')).toBe(true);
  });

  it('is not a path terminator, and the corpus says so 35 times over', () => {
    let withoutClose = 0;
    let mostSubpaths = 0;
    const midPath: string[] = [];
    for (const name of presetNames()) {
      for (const path of preset(name).pathLst) {
        const closes = path.commands.filter((c) => c.kind === 'close').length;
        if (closes === 0) withoutClose += 1;
        mostSubpaths = Math.max(
          mostSubpaths,
          path.commands.filter((c) => c.kind === 'moveTo').length,
        );
        path.commands.forEach((command, at) => {
          if (command.kind === 'close' && at !== path.commands.length - 1) midPath.push(name);
        });
      }
    }
    expect([withoutClose, new Set(midPath).size, mostSubpaths]).toStrictEqual([51, 35, 11]);
  });

  it('is a case only a hand-authored custGeom can reach, which is why the test above is synthetic', () => {
    // Every DrawingML command but `arcTo` names its own destination, so where a
    // `close` leaves the pen is only observable if an `arcTo` follows it. That
    // never happens in the presets - so the corpus cannot check this, and the
    // synthetic path above is the only witness there is.
    let closeThenArc = 0;
    for (const name of presetNames()) {
      for (const path of preset(name).pathLst) {
        path.commands.forEach((command, at) => {
          if (command.kind === 'close' && path.commands[at + 1]?.kind === 'arcTo') {
            closeThenArc += 1;
          }
        });
      }
    }
    expect(closeThenArc).toBe(0);
  });
});

describe('a path that cannot be drawn says so instead of emitting NaN', () => {
  function unreachable(size: ShapeSize): string[] {
    const found: string[] = [];
    for (const name of presetNames()) {
      const resolved = resolveGeometry(preset(name), size);
      if (resolved.paths.some((path) => !path.finite)) found.push(name);
    }
    return found;
  }

  it('is exactly the shapes that reach a non-finite guide, and no others', () => {
    expect(unreachable({ w: 0, h: 100 })).toStrictEqual([
      'circularArrow',
      'curvedDownArrow',
      'curvedLeftArrow',
      'curvedRightArrow',
      'curvedUpArrow',
      'ellipseRibbon',
      'ellipseRibbon2',
      'halfFrame',
      'leftCircularArrow',
      'leftRightCircularArrow',
      'moon',
    ]);
    expect(unreachable({ w: 100, h: 0 }).length).toBe(16);
    expect(unreachable({ w: 0, h: 0 }).length).toBe(19);
    for (const size of SIZES) expect(unreachable(size)).toStrictEqual([]);
  });

  it('empties the d string but keeps the numbers in the segments', () => {
    const resolved = resolveGeometry(preset('moon'), { w: 0, h: 100 });
    const broken = resolved.paths.filter((path) => !path.finite);
    expect(broken.length).toBeGreaterThan(0);
    for (const path of broken) {
      expect(path.d).toBe('');
      expect(path.segments.length).toBeGreaterThan(0);
    }
  });

  it('never writes a NaN or an Infinity into a d string', () => {
    for (const size of [...SIZES, { w: 0, h: 100 }, { w: 100, h: 0 }, { w: 0, h: 0 }]) {
      for (const name of presetNames()) {
        for (const path of resolveGeometry(preset(name), size).paths) {
          expect(path.d.includes('NaN')).toBe(false);
          expect(path.d.includes('Infinity')).toBe(false);
        }
      }
    }
  });
});

describe('every preset, at four sizes', () => {
  it('emits one path per a:path, each starting with a move', () => {
    let paths = 0;
    let segments = 0;
    for (const name of presetNames()) {
      const shape = preset(name);
      for (const size of SIZES) {
        const resolved = resolveGeometry(shape, size);
        expect(resolved.paths.length).toBe(shape.pathLst.length);
        resolved.paths.forEach((path, index) => {
          expect(segmentAt(path, 0).kind).toBe('move');
          expect(path.d.startsWith('M')).toBe(true);
          expect(path.fill).toBe(shape.pathLst[index]?.fill);
          expect(path.stroke).toBe(shape.pathLst[index]?.stroke);
          paths += 1;
          segments += path.segments.length;
        });
      }
    }
    expect(paths).toBe(320 * SIZES.length);
    expect(segments).toBeGreaterThan(320 * SIZES.length);
  });

  it('leaves no gap between one segment and the next', () => {
    for (const name of presetNames()) {
      for (const size of SIZES) {
        for (const path of resolveGeometry(preset(name), size).paths) {
          let cursor: Point | null = null;
          let subpathStart: Point | null = null;
          for (const segment of path.segments) {
            if (segment.kind === 'arc' && cursor !== null) {
              expect(Math.hypot(segment.from.x - cursor.x, segment.from.y - cursor.y)).toBeLessThan(
                1e-9,
              );
            }
            if (segment.kind === 'close') cursor = subpathStart;
            else {
              cursor = segment.to;
              if (segment.kind === 'move') subpathStart = segment.to;
            }
          }
        }
      }
    }
  });

  it('cannot have its paths merged into one, for 60 of the 187', () => {
    // The measurement behind one <path> per a:path. Concatenating these shapes'
    // paths loses a distinction their own definitions draw.
    let multiPath = 0;
    let mixed = 0;
    let strokeOnly = 0;
    for (const name of presetNames()) {
      const shape = preset(name);
      if (shape.pathLst.length > 1) multiPath += 1;
      const fills = new Set(shape.pathLst.map((p) => p.fill));
      const strokes = new Set(shape.pathLst.map((p) => p.stroke));
      if (fills.size > 1 || strokes.size > 1) mixed += 1;
      strokeOnly += shape.pathLst.filter((p) => p.fill === 'none' && p.stroke).length;
    }
    expect([multiPath, mixed, strokeOnly]).toStrictEqual([63, 60, 95]);
  });

  it('smileyFace really does draw an unfilled mouth on a filled head', () => {
    const resolved = resolveGeometry(preset('smileyFace'), { w: 200, h: 200 });
    expect(resolved.paths.map((path) => `${path.fill}/${path.stroke ? 'stroke' : 'no'}`)).toContain(
      'none/stroke',
    );
    expect(resolved.paths.map((path) => path.fill)).toContain('norm');
  });
});

describe('resolveGeometry also answers the two questions next to the path', () => {
  it('leaves the text rectangle null for the five shapes that have none', () => {
    const none: string[] = [];
    for (const name of presetNames()) {
      if (resolveGeometry(preset(name), { w: 100, h: 100 }).textRect === null) none.push(name);
    }
    expect(none.sort()).toStrictEqual(['chartPlus', 'chartStar', 'chartX', 'line', 'lineInv']);
  });

  it('keeps a connection site angle in the units the file wrote it in', () => {
    const resolved = resolveGeometry(preset('triangle'), { w: 200, h: 100 });
    expect(resolved.connectionSites.length).toBeGreaterThan(0);
    const top = resolved.connectionSites[0];
    expect(top?.ang).toBe(16200000);
    expect(top?.pos).toStrictEqual({ x: 100, y: 0 });
  });

  it('lets an adjust value move the path', () => {
    const shape = preset('triangle');
    const size = { w: 200, h: 100 };
    const left = pathAt(resolveGeometry(shape, size, { adjust: { adj: 0 } }), 0);
    const right = pathAt(resolveGeometry(shape, size, { adjust: { adj: 100000 } }), 0);
    expect(left.d).not.toBe(right.d);
    expect(left.d).toBe(pathAt(resolveGeometry(shape, size, { adjust: { adj: 0 } }), 0).d);
  });

  it('honours a precision override all the way down', () => {
    // flowChartDisplay is drawn in sixths, so no size makes its coordinates
    // come out whole and the rounding is always doing something.
    const size = { w: 100, h: 100 };
    const coarse = pathAt(resolveGeometry(preset('flowChartDisplay'), size, { precision: 1 }), 0);
    const fine = pathAt(resolveGeometry(preset('flowChartDisplay'), size, { precision: 6 }), 0);
    expect(coarse.d).toContain('16.7');
    expect(fine.d).toContain('16.666667');
    expect(coarse.d.length).toBeLessThan(fine.d.length);
    // Precision changes the string and nothing else.
    expect(coarse.segments).toStrictEqual(fine.segments);
  });
});
