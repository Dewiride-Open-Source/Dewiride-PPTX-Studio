import { describe, expect, it } from 'vitest';
import { GeometryError } from '../errors.js';
import { PresetBucket } from '../decode.js';
import {
  BUCKETS,
  BUCKET_MEMBERS,
  COLLAPSED_WHITESPACE,
  OVER_LONG_FORMULAS,
  PRESET_SOURCE,
  PRESET_TOTALS,
  getPreset,
  presetNames,
} from './index.js';
import type { PresetShape } from '../types.js';

/**
 * What this suite is actually checking.
 *
 * The generated buckets are written by one program and read by another, and
 * the two share no code - the transcoder is in `tools/geometry-codegen` and
 * cannot import this package, for the reason `read-xsd.ts` sets out. So the
 * encoder losing a field, or transposing two, would not be caught by either
 * side on its own.
 *
 * Three things close that gap, in increasing order of strength:
 *
 * 1. **The totals.** `PRESET_TOTALS` is counted by the generator from its
 *    *parse* of the XML, before anything is encoded. Recomputing them here from
 *    the *decode* compares the two ends of the pipeline through a number that
 *    neither end can fudge.
 * 2. **Four shapes written out by hand** from POI's file, covering an arc, a
 *    polar handle, a shape with no adjust values at all, a shape with four
 *    paths, and the two anomalies. A transposition survives a count; it does
 *    not survive these.
 * 3. **A full re-encode**, in `tools/geometry-codegen/codegen.test.ts`, which
 *    decodes every committed bucket and encodes it again character for
 *    character. That is the one that would catch a field silently dropped from
 *    all 187 shapes at once.
 */

function everyShape(): PresetShape[] {
  return presetNames().map((name) => {
    const shape = getPreset(name);
    if (shape === undefined) throw new Error(`${name} is named but does not decode`);
    return shape;
  });
}

describe('the preset roster', () => {
  it('is the 187 ST_ShapeType names, and each appears once', () => {
    const names = presetNames();
    expect(names).toHaveLength(187);
    expect(new Set(names).size).toBe(187);
    expect([...names].sort()).toStrictEqual(names);
  });

  it('partitions those names across the six buckets, with no overlap', () => {
    const members = Object.values(BUCKET_MEMBERS).flat();
    expect(members).toHaveLength(187);
    expect(new Set(members).size).toBe(187);
    expect([...members].sort()).toStrictEqual(presetNames());
  });

  it('holds the bucket sizes the generator reported', () => {
    const sizes = Object.fromEntries(
      Object.entries(BUCKET_MEMBERS).map(([name, list]) => [name, list.length]),
    );
    expect(sizes).toStrictEqual({
      basic: 43,
      arrows: 31,
      callouts: 16,
      flowchart: 29,
      misc: 47,
      stars: 21,
    });
  });

  it('records which POI file it came from', () => {
    expect(PRESET_SOURCE).toStrictEqual({
      file: 'presetShapeDefinitions.xml',
      sha256: '4a762444d8d85876881c02a5b1dedf6f73006fcd8acb7b4e393435615b37c780',
      bytes: 538970,
      // The root element is misspelled in the source and reproduced verbatim.
      root: 'presetShapeDefinitons',
    });
  });
});

describe('decoding against the counts taken before encoding', () => {
  it('recovers every guide, handle, site, rectangle, path and command', () => {
    const shapes = everyShape();
    const sum = (pick: (shape: PresetShape) => number): number =>
      shapes.reduce((total, shape) => total + pick(shape), 0);

    expect({
      shapes: shapes.length,
      avGuides: sum((s) => s.avLst.length),
      gdGuides: sum((s) => s.gdLst.length),
      handles: sum((s) => s.ahLst.length),
      sites: sum((s) => s.cxnLst.length),
      textRects: shapes.filter((s) => s.rect !== null).length,
      paths: sum((s) => s.pathLst.length),
      commands: sum((s) => s.pathLst.reduce((n, path) => n + path.commands.length, 0)),
    }).toStrictEqual({ ...PRESET_TOTALS });
  });

  it('leaves exactly five shapes without a text rectangle', () => {
    const without = everyShape()
      .filter((shape) => shape.rect === null)
      .map((shape) => shape.name);
    expect(without).toStrictEqual(['chartPlus', 'chartStar', 'chartX', 'line', 'lineInv']);
  });

  it('never emits an empty guide name or an operator-only formula', () => {
    for (const shape of everyShape()) {
      for (const guide of [...shape.avLst, ...shape.gdLst]) {
        expect(guide.name).not.toBe('');
        expect(guide.fmla.length).toBeGreaterThanOrEqual(2);
      }
    }
  });
});

describe('shapes written out by hand from the source file', () => {
  it('decodes triangle, the simplest shape with an adjust handle', () => {
    expect(getPreset('triangle')).toStrictEqual({
      name: 'triangle',
      avLst: [{ name: 'adj', fmla: ['val', '50000'] }],
      gdLst: [
        { name: 'a', fmla: ['pin', '0', 'adj', '100000'] },
        { name: 'x1', fmla: ['*/', 'w', 'a', '200000'] },
        { name: 'x2', fmla: ['*/', 'w', 'a', '100000'] },
        { name: 'x3', fmla: ['+-', 'x1', 'wd2', '0'] },
      ],
      // Horizontal only: the Y half of the handle is absent, which is not the
      // same as a Y range of zero.
      ahLst: [
        {
          kind: 'xy',
          gdRefX: 'adj',
          minX: '0',
          maxX: '100000',
          gdRefY: null,
          minY: null,
          maxY: null,
          pos: { x: 'x2', y: 't' },
        },
      ],
      cxnLst: [
        { ang: '3cd4', pos: { x: 'x2', y: 't' } },
        { ang: 'cd2', pos: { x: 'x1', y: 'vc' } },
        { ang: 'cd4', pos: { x: 'l', y: 'b' } },
        { ang: 'cd4', pos: { x: 'x2', y: 'b' } },
        { ang: 'cd4', pos: { x: 'r', y: 'b' } },
        { ang: '0', pos: { x: 'x3', y: 'vc' } },
      ],
      rect: { l: 'x1', t: 'vc', r: 'x3', b: 'b' },
      pathLst: [
        {
          w: 0,
          h: 0,
          fill: 'norm',
          stroke: true,
          extrusionOk: true,
          commands: [
            { kind: 'moveTo', to: { x: 'l', y: 'b' } },
            { kind: 'lnTo', to: { x: 'x2', y: 't' } },
            { kind: 'lnTo', to: { x: 'r', y: 'b' } },
            { kind: 'close' },
          ],
        },
      ],
    });
  });

  it('decodes pie: an arcTo, two polar handles, and a transposed text rectangle', () => {
    const pie = getPreset('pie');
    expect(pie?.avLst).toStrictEqual([
      { name: 'adj1', fmla: ['val', '0'] },
      { name: 'adj2', fmla: ['val', '16200000'] },
    ]);
    // Polar handles that adjust an angle and not a radius: the R half is absent.
    expect(pie?.ahLst).toStrictEqual([
      {
        kind: 'polar',
        gdRefR: null,
        minR: null,
        maxR: null,
        gdRefAng: 'adj1',
        minAng: '0',
        maxAng: '21599999',
        pos: { x: 'x1', y: 'y1' },
      },
      {
        kind: 'polar',
        gdRefR: null,
        minR: null,
        maxR: null,
        gdRefAng: 'adj2',
        minAng: '0',
        maxAng: '21599999',
        pos: { x: 'x2', y: 'y2' },
      },
    ]);
    // arcTo has no destination point. Its end is computed from the radii and
    // the angles, which is what makes 2.3 more than an SVG arc segment.
    expect(pie?.pathLst).toStrictEqual([
      {
        w: 0,
        h: 0,
        fill: 'norm',
        stroke: true,
        extrusionOk: true,
        commands: [
          { kind: 'moveTo', to: { x: 'x1', y: 'y1' } },
          { kind: 'arcTo', wR: 'wd2', hR: 'hd2', stAng: 'stAng', swAng: 'swAng' },
          { kind: 'lnTo', to: { x: 'hc', y: 'vc' } },
          { kind: 'close' },
        ],
      },
    ]);

    // POI writes `l="il" t="ir" r="it" b="ib"`, and the guides are named for
    // the sides they compute: il/ir are horizontal, it/ib vertical. So `t` and
    // `r` are transposed here, and pie is the only one of the 29 shapes using
    // this guide set that has them that way - the other 28 read il/it/ir/ib.
    //
    // Reproduced rather than corrected. This transcoder does not hold opinions
    // about its source, and a silent fix would make the difference between our
    // geometry and POI's invisible. If a later POI corrects it, this assertion
    // fails and whoever regenerates gets to decide deliberately.
    expect(pie?.rect).toStrictEqual({ l: 'il', t: 'ir', r: 'it', b: 'ib' });
  });

  it('decodes cornerTabs: no adjust values, no handles, four separate paths', () => {
    const shape = getPreset('cornerTabs');
    expect(shape?.avLst).toStrictEqual([]);
    expect(shape?.ahLst).toStrictEqual([]);
    expect(shape?.gdLst).toStrictEqual([
      { name: 'md', fmla: ['mod', 'w', 'h', '0'] },
      { name: 'dx', fmla: ['*/', '1', 'md', '20'] },
      { name: 'y1', fmla: ['+-', '0', 'b', 'dx'] },
      { name: 'x1', fmla: ['+-', '0', 'r', 'dx'] },
    ]);
    expect(shape?.cxnLst).toHaveLength(12);
    expect(shape?.rect).toStrictEqual({ l: 'dx', t: 'dx', r: 'x1', b: 'y1' });
    expect(shape?.pathLst).toHaveLength(4);
    expect(shape?.pathLst.map((path) => path.commands.length)).toStrictEqual([4, 4, 4, 4]);
    expect(shape?.pathLst[0]?.commands[0]).toStrictEqual({
      kind: 'moveTo',
      to: { x: 'l', y: 't' },
    });
  });

  it('keeps path fill, stroke and extrusion flags that are not the default', () => {
    const bevel = getPreset('bevel');
    expect(bevel?.pathLst).toHaveLength(6);
    // Five unstroked faces shaded by fill mode, then an unfilled outline that
    // is the only stroked path. Those lighten/darken modes are the whole reason
    // a bevel reads as three-dimensional, and `fill="none"` on the last path is
    // not `norm` - a renderer that treats a missing fill and an explicit none
    // as the same thing paints over the shading.
    expect(bevel?.pathLst.map((path) => [path.fill, path.stroke, path.extrusionOk])).toStrictEqual([
      ['norm', false, false],
      ['lightenLess', false, false],
      ['darkenLess', false, false],
      ['lighten', false, false],
      ['darken', false, false],
      ['none', true, false],
    ]);
  });
});

describe('what the source file gets wrong, carried through', () => {
  it('keeps the four-operand +- formulas rather than truncating them', () => {
    // ECMA's `+-` is x + y - z. These eight carry a fourth operand, always a
    // spare 0, and always in the circular-arrow family. 2.2 decides what an
    // evaluator does with it; the transcoder does not get to decide for it.
    expect(OVER_LONG_FORMULAS.map((entry) => `${entry.shape}/${entry.guide}`)).toStrictEqual([
      'circularArrow/xB',
      'circularArrow/yB',
      'leftCircularArrow/xB',
      'leftCircularArrow/yB',
      'leftRightCircularArrow/xB',
      'leftRightCircularArrow/yB',
      'leftRightCircularArrow/xJ',
      'leftRightCircularArrow/yJ',
    ]);
    expect(getPreset('circularArrow')?.gdLst.find((gd) => gd.name === 'xB')?.fmla).toStrictEqual([
      '+-',
      'xH',
      '0',
      'dxB',
      '0',
    ]);
  });

  it('splits the six double-spaced formulas on whitespace runs, not on one space', () => {
    expect(COLLAPSED_WHITESPACE.map((entry) => `${entry.shape}/${entry.guide}`)).toStrictEqual([
      'heptagon/svc',
      'leftArrow/x1',
      'pentagon/svc',
      'star5/svc',
      'star7/svc',
      'upArrow/y1',
    ]);
    // A single-space split would give this an empty operand and four of them.
    expect(getPreset('heptagon')?.gdLst.find((gd) => gd.name === 'svc')?.fmla).toStrictEqual([
      '*/',
      'vc',
      'vf',
      '100000',
    ]);
  });
});

describe('the lookup', () => {
  it('returns undefined for a name that is not a preset', () => {
    expect(getPreset('notAShape')).toBeUndefined();
    expect(getPreset('')).toBeUndefined();
  });

  it('decodes a shape once and hands back the same object after that', () => {
    expect(getPreset('roundRect')).toBe(getPreset('roundRect'));
  });

  it('answers has() without decoding', () => {
    expect(BUCKETS.flowchart.has('flowChartOr')).toBe(true);
    expect(BUCKETS.flowchart.has('roundRect')).toBe(false);
    expect(BUCKETS.flowchart.get('roundRect')).toBeUndefined();
  });
});

describe('the decoder refuses data it does not understand', () => {
  it('throws a typed error, never a bare TypeError', () => {
    const bucket = new PresetBucket('bad||||||- - - - -,W 1 2');
    expect(() => bucket.get('bad')).toThrowError(GeometryError);
    try {
      bucket.get('bad');
    } catch (error) {
      expect(error).toBeInstanceOf(GeometryError);
      expect((error as GeometryError).code).toBe('PRESET_DECODE');
      expect((error as GeometryError).preset).toBe('bad');
    }
  });

  it('rejects a shape with the wrong number of fields', () => {
    expect(() => new PresetBucket('short|||').get('short')).toThrowError(/expected 6 fields/);
  });

  it('rejects an entry with no name separator at all', () => {
    expect(() => new PresetBucket('nameless')).toThrowError(GeometryError);
  });

  it('rejects a path fill that is not an ST_PathFillMode', () => {
    const bucket = new PresetBucket('bad||||||- - chartreuse - -');
    expect(() => bucket.get('bad')).toThrowError(/not an ST_PathFillMode/);
  });
});
