import { describe, expect, it } from 'vitest';
import { builtinGuideNames, builtinGuides } from './builtins.js';
import { FULL_CIRCLE } from './formula.js';
import { getPreset, presetNames } from './presets/index.js';

/**
 * The built-in table is derived from the naming rule. This suite checks it
 * against the only other source there is: the 187 presets, which reference
 * these names and would throw on a missing one.
 *
 * The list below is read off the committed buckets and written out literally,
 * so the two are independent in the way that matters - the table is built by
 * applying a rule to a set of divisors, and the list is an observation of what
 * the data asks for. If they ever disagree, one of them is wrong and the fix is
 * to work out which, not to regenerate one from the other.
 */
const REQUIRED_BY_THE_PRESETS = [
  // the frame
  'l',
  't',
  'r',
  'b',
  'w',
  'h',
  'hc',
  'vc',
  'ss',
  // width fractions
  'wd2',
  'wd3',
  'wd4',
  'wd5',
  'wd6',
  'wd8',
  'wd10',
  'wd12',
  'wd32',
  // height fractions
  'hd2',
  'hd3',
  'hd4',
  'hd5',
  'hd6',
  'hd8',
  'hd10',
  // short-side fractions
  'ssd2',
  'ssd6',
  'ssd8',
  'ssd16',
  'ssd32',
  // angles
  'cd2',
  'cd3',
  'cd4',
  '3cd4',
] as const;

/**
 * Recompute that list from the buckets, so the literal above cannot silently
 * fall out of date. This walks every operand position in every preset and
 * collects the names that no `avLst` or `gdLst` in that shape declares.
 */
function namesThePresetsCannotResolveAlone(): Set<string> {
  const free = new Set<string>();
  const isLiteral = (t: string) => /^[-+]?\d+$/.test(t);

  for (const name of presetNames()) {
    const shape = getPreset(name);
    if (shape === undefined) throw new Error(`${name} does not decode`);

    const declared = new Set([
      ...shape.avLst.map((g) => g.name),
      ...shape.gdLst.map((g) => g.name),
    ]);
    const note = (token: string | null) => {
      if (token !== null && !isLiteral(token) && !declared.has(token)) free.add(token);
    };

    for (const gd of [...shape.avLst, ...shape.gdLst]) for (const t of gd.fmla.slice(1)) note(t);
    for (const ah of shape.ahLst) {
      const refs =
        ah.kind === 'xy'
          ? [ah.gdRefX, ah.minX, ah.maxX, ah.gdRefY, ah.minY, ah.maxY]
          : [ah.gdRefR, ah.minR, ah.maxR, ah.gdRefAng, ah.minAng, ah.maxAng];
      for (const t of refs) note(t);
      note(ah.pos.x);
      note(ah.pos.y);
    }
    for (const cxn of shape.cxnLst) {
      note(cxn.ang);
      note(cxn.pos.x);
      note(cxn.pos.y);
    }
    if (shape.rect !== null)
      for (const t of [shape.rect.l, shape.rect.t, shape.rect.r, shape.rect.b]) note(t);
    for (const path of shape.pathLst) {
      for (const cmd of path.commands) {
        switch (cmd.kind) {
          case 'moveTo':
          case 'lnTo':
            note(cmd.to.x);
            note(cmd.to.y);
            break;
          case 'quadBezTo':
            note(cmd.c1.x);
            note(cmd.c1.y);
            note(cmd.to.x);
            note(cmd.to.y);
            break;
          case 'cubicBezTo':
            note(cmd.c1.x);
            note(cmd.c1.y);
            note(cmd.c2.x);
            note(cmd.c2.y);
            note(cmd.to.x);
            note(cmd.to.y);
            break;
          case 'arcTo':
            note(cmd.wR);
            note(cmd.hR);
            note(cmd.stAng);
            note(cmd.swAng);
            break;
          case 'close':
            break;
        }
      }
    }
  }
  return free;
}

describe('the built-in guides the presets require', () => {
  it('is a set of 34 names, and the literal list above is still accurate', () => {
    const observed = namesThePresetsCannotResolveAlone();
    expect(observed.size).toBe(34);
    expect([...observed].sort()).toStrictEqual([...REQUIRED_BY_THE_PRESETS].sort());
  });

  it('is entirely covered by the table', () => {
    const table = new Set(builtinGuideNames());
    const missing = [...REQUIRED_BY_THE_PRESETS].filter((name) => !table.has(name));
    expect(missing).toStrictEqual([]);
  });

  it('leaves no preset with an unresolvable operand, which is the same claim from the other side', () => {
    const table = new Set(builtinGuideNames());
    expect([...namesThePresetsCannotResolveAlone()].filter((n) => !table.has(n))).toStrictEqual([]);
  });
});

describe('the values', () => {
  const g = builtinGuides(200, 100);

  it('places the frame in the shape’s own space, with the origin at its top left', () => {
    expect(g.get('l')).toBe(0);
    expect(g.get('t')).toBe(0);
    expect(g.get('r')).toBe(200);
    expect(g.get('b')).toBe(100);
    expect(g.get('w')).toBe(200);
    expect(g.get('h')).toBe(100);
    expect(g.get('hc')).toBe(100);
    expect(g.get('vc')).toBe(50);
  });

  it('takes ss from the shorter side and ls from the longer', () => {
    expect(g.get('ss')).toBe(100);
    expect(g.get('ls')).toBe(200);
    const tall = builtinGuides(100, 200);
    expect(tall.get('ss')).toBe(100);
    expect(tall.get('ls')).toBe(200);
  });

  it('divides the width, the height and the shorter side by the number in the name', () => {
    expect(g.get('wd2')).toBe(100);
    expect(g.get('wd4')).toBe(50);
    expect(g.get('wd32')).toBe(200 / 32);
    expect(g.get('hd2')).toBe(50);
    expect(g.get('hd10')).toBe(10);
    expect(g.get('ssd2')).toBe(50);
    expect(g.get('ssd6')).toBeCloseTo(100 / 6, 12);
    expect(g.get('ssd32')).toBe(100 / 32);
  });

  it('divides a full circle by the number in the name, and multiplies by the prefix', () => {
    expect(g.get('cd2')).toBe(10800000);
    expect(g.get('cd3')).toBe(7200000);
    expect(g.get('cd4')).toBe(5400000);
    expect(g.get('cd8')).toBe(2700000);
    expect(g.get('3cd4')).toBe(16200000);
    expect(g.get('3cd8')).toBe(8100000);
    expect(g.get('5cd8')).toBe(13500000);
    expect(g.get('7cd8')).toBe(18900000);

    // The same claim without the constants, so a change to FULL_CIRCLE cannot
    // leave these agreeing with a stale number.
    expect(g.get('cd4')).toBe(FULL_CIRCLE / 4);
    expect(g.get('3cd4')).toBe((3 * FULL_CIRCLE) / 4);
  });

  it('keeps the angles independent of the shape size', () => {
    const tiny = builtinGuides(1, 1);
    const huge = builtinGuides(9144000, 6858000);
    for (const name of ['cd2', 'cd3', 'cd4', 'cd8', '3cd4', '3cd8', '5cd8', '7cd8']) {
      expect(tiny.get(name)).toBe(huge.get(name));
    }
  });
});

describe('the name list', () => {
  it('is sorted by code unit, like presetNames, and for the same reason', () => {
    const names = builtinGuideNames();
    expect([...names].sort()).toStrictEqual(names);
    expect(new Set(names).size).toBe(names.length);
  });

  it('does not include the adjust values, which each preset declares for itself', () => {
    const table = new Set(builtinGuideNames());
    for (const adjust of ['adj', 'adj1', 'adj2', 'adj8', 'hf', 'vf']) {
      expect(table.has(adjust)).toBe(false);
    }
  });
});
