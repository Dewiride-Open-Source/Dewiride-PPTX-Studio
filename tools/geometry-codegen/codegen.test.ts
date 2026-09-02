import { describe, expect, it } from 'vitest';
import type { PresetShape } from '../../packages/geometry/src/types.ts';
import { ARROWS, ARROWS_DATA } from '../../packages/geometry/src/presets/arrows.gen.ts';
import { BASIC, BASIC_DATA } from '../../packages/geometry/src/presets/basic.gen.ts';
import { CALLOUTS, CALLOUTS_DATA } from '../../packages/geometry/src/presets/callouts.gen.ts';
import { FLOWCHART, FLOWCHART_DATA } from '../../packages/geometry/src/presets/flowchart.gen.ts';
import { MISC, MISC_DATA } from '../../packages/geometry/src/presets/misc.gen.ts';
import { STARS, STARS_DATA } from '../../packages/geometry/src/presets/stars.gen.ts';
import { encodeBucket, encodeShape } from './encode.ts';
import { PresetReadError, readPresets } from './read-presets.ts';
import { BUCKET_NAMES, bucketOf, ST_SHAPE_TYPE } from './shape-types.ts';

/**
 * The transcoder's own suite, and the one check that spans both halves.
 *
 * The input file is not in this repository, so nothing here reads it - which
 * rules out the obvious test and forces a better one. `re-encodes every
 * committed bucket` below decodes all 187 shapes with the *package's* decoder
 * and encodes them again with the *tool's* encoder, then compares against the
 * committed text character for character. Any field the encoder drops, any pair
 * it transposes, any operand it loses to a separator, changes that string.
 *
 * It is the inverse-function property from 0.6, applied to a different pair of
 * functions: `encode(decode(x)) === x` for every shape we ship.
 */

const BUCKETS = [
  ['arrows', ARROWS, ARROWS_DATA],
  ['basic', BASIC, BASIC_DATA],
  ['callouts', CALLOUTS, CALLOUTS_DATA],
  ['flowchart', FLOWCHART, FLOWCHART_DATA],
  ['misc', MISC, MISC_DATA],
  ['stars', STARS, STARS_DATA],
] as const;

describe('encode after decode is the identity, over everything we ship', () => {
  for (const [name, bucket, data] of BUCKETS) {
    it(`re-encodes the ${name} bucket to the committed text`, () => {
      const shapes = bucket.names.map((shape) => {
        const decoded = bucket.get(shape);
        if (decoded === undefined) throw new Error(`${shape} is named but does not decode`);
        return decoded;
      });
      expect(encodeBucket(shapes)).toBe(data);
    });
  }

  it('covers all 187 between them', () => {
    const total = BUCKETS.reduce((sum, [, bucket]) => sum + bucket.names.length, 0);
    expect(total).toBe(187);
  });
});

describe('the ST_ShapeType list this repository wrote down', () => {
  it('has 187 unique names, sorted', () => {
    expect(ST_SHAPE_TYPE).toHaveLength(187);
    expect(new Set(ST_SHAPE_TYPE).size).toBe(187);
    expect([...ST_SHAPE_TYPE].sort()).toStrictEqual([...ST_SHAPE_TYPE]);
  });

  it('agrees with the names in the generated buckets', () => {
    const shipped = BUCKETS.flatMap(([, bucket]) => [...bucket.names]).sort();
    // The point of the independent list: if this ever fails, one of the two
    // sources is wrong and the fix is to find out which - not to paste either
    // over the other.
    expect(shipped).toStrictEqual([...ST_SHAPE_TYPE]);
  });
});

describe('bucket assignment', () => {
  it('gives every name exactly one bucket', () => {
    const counts = new Map(BUCKET_NAMES.map((name) => [name, 0]));
    for (const name of ST_SHAPE_TYPE) {
      const bucket = bucketOf(name);
      counts.set(bucket, (counts.get(bucket) ?? 0) + 1);
    }
    expect(Object.fromEntries(counts)).toStrictEqual({
      basic: 43,
      arrows: 31,
      callouts: 16,
      flowchart: 29,
      misc: 47,
      stars: 21,
    });
  });

  it('files the six *ArrowCallout shapes under arrows, as PowerPoint does', () => {
    // Rule order is what decides this, and it is the one place where the name
    // alone would say the opposite.
    expect(bucketOf('rightArrowCallout')).toBe('arrows');
    expect(bucketOf('upDownArrowCallout')).toBe('arrows');
    expect(bucketOf('quadArrowCallout')).toBe('arrows');
    expect(bucketOf('cloudCallout')).toBe('callouts');
    expect(bucketOf('wedgeRectCallout')).toBe('callouts');
  });

  it('puts ribbons, scrolls and waves with the stars, and connectors in misc', () => {
    expect(bucketOf('leftRightRibbon')).toBe('stars');
    expect(bucketOf('verticalScroll')).toBe('stars');
    expect(bucketOf('doubleWave')).toBe('stars');
    expect(bucketOf('star32')).toBe('stars');
    expect(bucketOf('irregularSeal2')).toBe('stars');
    expect(bucketOf('bentConnector3')).toBe('misc');
    expect(bucketOf('actionButtonSound')).toBe('misc');
    expect(bucketOf('mathNotEqual')).toBe('misc');
  });
});

const SHAPE = (body: string): string =>
  `<?xml version="1.0" encoding="utf-8"?>\n<presetShapeDefinitons>${body}</presetShapeDefinitons>`;

const MINIMAL = SHAPE(
  '<demo><pathLst xmlns="http://schemas.openxmlformats.org/drawingml/2006/main">' +
    '<path><moveTo><pt x="l" y="t"/></moveTo><close/></path></pathLst></demo>',
);

describe('the reader refuses everything outside the profile', () => {
  it('reads the minimal shape it does accept', () => {
    const { shapes, anomalies } = readPresets(MINIMAL);
    expect(anomalies).toStrictEqual([]);
    expect(shapes).toHaveLength(1);
    expect(shapes[0]?.name).toBe('demo');
    expect(shapes[0]?.rect).toBeNull();
    expect(shapes[0]?.pathLst[0]?.commands).toStrictEqual([
      { kind: 'moveTo', to: { x: 'l', y: 't' } },
      { kind: 'close' },
    ]);
  });

  it('rejects a comment, a DOCTYPE and a CDATA section alike', () => {
    for (const intruder of ['<!-- hello -->', '<!DOCTYPE x>', '<![CDATA[x]]>']) {
      expect(() => readPresets(SHAPE(intruder))).toThrowError(PresetReadError);
    }
  });

  it('rejects an element it has never seen', () => {
    expect(() => readPresets(SHAPE('<demo><surprise/></demo>'))).toThrowError(
      /<surprise> is outside the profile/,
    );
  });

  it('rejects an unexpected attribute on an element it knows', () => {
    expect(() =>
      readPresets(
        SHAPE('<demo><cxnLst><cxn ang="0" tilt="3"><pos x="l" y="t"/></cxn></cxnLst></demo>'),
      ),
    ).toThrowError(/unexpected attribute "tilt"/);
  });

  it('rejects a formula operator that is not one of the seventeen', () => {
    expect(() =>
      readPresets(SHAPE('<demo><gdLst><gd name="a" fmla="frobnicate 1 2"/></gdLst></demo>')),
    ).toThrowError(/unknown formula operator "frobnicate"/);
  });

  it('rejects text content, which no preset definition has', () => {
    expect(() => readPresets(SHAPE('<demo>text</demo>'))).toThrowError(/has text content/);
  });

  it('rejects a second definition of the same shape', () => {
    expect(() => readPresets(SHAPE('<demo/><demo/>'))).toThrowError(/defined twice/);
  });

  it('rejects a root element that is spelled correctly', () => {
    // Deliberate: POI's root is misspelled `presetShapeDefinitons`, and a file
    // whose root is spelled properly is a different file than the one this was
    // written against.
    const correct = '<?xml version="1.0"?>\n<presetShapeDefinitions/>';
    expect(() => readPresets(correct)).toThrowError(/the root is <presetShapeDefinitions>/);
  });

  it('reports a wrong operand count without refusing the file', () => {
    const { anomalies } = readPresets(
      SHAPE('<demo><gdLst><gd name="a" fmla="val 1 2"/></gdLst></demo>'),
    );
    expect(anomalies).toStrictEqual([
      { kind: 'over-long-formula', shape: 'demo', guide: 'a', fmla: 'val 1 2' },
    ]);
  });

  it('reports a double space and still splits the operands correctly', () => {
    const { shapes, anomalies } = readPresets(
      SHAPE('<demo><gdLst><gd name="a" fmla="*/ w  h 2"/></gdLst></demo>'),
    );
    expect(anomalies[0]?.kind).toBe('collapsed-whitespace');
    expect(shapes[0]?.gdLst[0]?.fmla).toStrictEqual(['*/', 'w', 'h', '2']);
  });
});

describe('the encoder', () => {
  const shape: PresetShape = {
    name: 'demo',
    avLst: [{ name: 'adj', fmla: ['val', '50000'] }],
    gdLst: [{ name: 'g', fmla: ['?:', 'a', 'b', 'c'] }],
    ahLst: [
      {
        kind: 'xy',
        gdRefX: 'adj',
        minX: '0',
        maxX: '1',
        gdRefY: null,
        minY: null,
        maxY: null,
        pos: { x: 'x', y: 'y' },
      },
      {
        kind: 'polar',
        gdRefR: null,
        minR: null,
        maxR: null,
        gdRefAng: 'adj',
        minAng: '0',
        maxAng: '2',
        pos: { x: 'p', y: 'q' },
      },
    ],
    cxnLst: [{ ang: '3cd4', pos: { x: 'hc', y: 't' } }],
    rect: { l: 'l', t: 't', r: 'r', b: 'b' },
    pathLst: [
      {
        w: 10,
        h: 20,
        fill: 'darken',
        stroke: false,
        extrusionOk: false,
        commands: [
          { kind: 'moveTo', to: { x: 'l', y: 't' } },
          { kind: 'quadBezTo', c1: { x: 'a', y: 'b' }, to: { x: 'c', y: 'd' } },
          {
            kind: 'cubicBezTo',
            c1: { x: 'a', y: 'b' },
            c2: { x: 'c', y: 'd' },
            to: { x: 'e', y: 'f' },
          },
          { kind: 'arcTo', wR: 'wd2', hR: 'hd2', stAng: '0', swAng: 'cd4' },
          { kind: 'close' },
        ],
      },
      {
        w: 0,
        h: 0,
        fill: 'norm',
        stroke: true,
        extrusionOk: true,
        commands: [{ kind: 'lnTo', to: { x: 'r', y: 'b' } }],
      },
    ],
  };

  it('writes every branch of the grammar exactly', () => {
    expect(encodeShape(shape)).toBe(
      'demo|adj val 50000|g ?: a b c|X adj 0 1 - - - x y,P - - - adj 0 2 p q|3cd4 hc t|l t r b|' +
        '10 20 darken 0 0,M l t,Q a b c d,C a b c d e f,A wd2 hd2 0 cd4,Z!- - - - -,L r b',
    );
  });

  it('refuses a token carrying a separator rather than writing a corrupt bucket', () => {
    const bad: PresetShape = { ...shape, cxnLst: [{ ang: 'a,b', pos: { x: 'l', y: 't' } }] };
    expect(() => encodeShape(bad)).toThrowError(/reserves/);
  });
});
