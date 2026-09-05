import { describe, expect, it } from 'vitest';
import bases from '../../../../corpus/ground-truth/color-bases.json' with { type: 'json' };
import fixture from '../../../../corpus/ground-truth/color-transforms.json' with { type: 'json' };
import { applyTransforms } from '../apply.js';
import { PaintError } from '../errors.js';
import { parseAngle, parsePercentage, parseSrgbValue } from '../parse.js';
import { PRESET_COLORS } from './preset-colors.js';
import { mapSchemeName, resolveColor, toCss, toHexColor, type ColorContext } from './resolve.js';
import { SYSTEM_COLORS } from './sys-colors.js';
import { fromLinear, hslToRgb, parseHex, rgbToHsl, toByte, toHex, toLinear } from './transfer.js';
import type { ClrMap, ClrScheme, Color, ColorTransform, SchemeSlot } from '../types.js';

/**
 * Two fixtures, 473 swatches, one question: does this package paint what
 * Microsoft PowerPoint paints?
 *
 * `color-transforms.json` is sub-phase 0.7-C - 214 swatches that settled which
 * colour space each transform computes in. `color-bases.json` is 2.6 - 259 more
 * that settled the six bases, alpha, the colour map, the percentage grammar,
 * and the thing 0.7-C could not see: that every transform hands the next one a
 * *clamped* colour.
 *
 * Both were read back two independent ways - PowerPoint's object model and the
 * centre pixel of its own bitmap export - and the two agree on all 464 opaque
 * swatches across the two experiments. Where they could differ the pixels win,
 * because they are what a user sees; the fixtures record the pixels.
 *
 * The tests below are not only "our number equals their number". Several assert
 * that a *wrong* model would be visibly wrong, with the number it would produce,
 * because a fixture nobody can fail is a fixture nobody maintains.
 */

/* -------------------------------------------------------------------------- */
/* building a Color out of a fixture row                                      */
/* -------------------------------------------------------------------------- */

interface FixtureTransform {
  name: string;
  val?: number;
}

function transforms(list: readonly FixtureTransform[]): ColorTransform[] {
  return list.map((t) =>
    t.val === undefined
      ? ({ op: t.name } as ColorTransform)
      : ({ op: t.name, val: t.val } as ColorTransform),
  );
}

/**
 * The theme both experiments used, rebuilt the way the deck actually wrote it.
 *
 * `dk1` and `lt1` are `a:sysClr` with a `@lastClr`, not `a:srgbClr`, because
 * that is what PowerPoint writes and what the swatch deck copied. It makes the
 * fixture exercise system-colour resolution inside theme resolution rather than
 * only at the top level.
 */
const SCHEME: ClrScheme = {
  dk1: { space: 'sys', name: 'windowText', lastClr: fixture.clrScheme.dk1, transforms: [] },
  lt1: { space: 'sys', name: 'window', lastClr: fixture.clrScheme.lt1, transforms: [] },
  dk2: { space: 'srgb', hex: fixture.clrScheme.dk2, transforms: [] },
  lt2: { space: 'srgb', hex: fixture.clrScheme.lt2, transforms: [] },
  accent1: { space: 'srgb', hex: fixture.clrScheme.accent1, transforms: [] },
  accent2: { space: 'srgb', hex: fixture.clrScheme.accent2, transforms: [] },
  accent3: { space: 'srgb', hex: fixture.clrScheme.accent3, transforms: [] },
  accent4: { space: 'srgb', hex: fixture.clrScheme.accent4, transforms: [] },
  accent5: { space: 'srgb', hex: fixture.clrScheme.accent5, transforms: [] },
  accent6: { space: 'srgb', hex: fixture.clrScheme.accent6, transforms: [] },
  hlink: { space: 'srgb', hex: fixture.clrScheme.hlink, transforms: [] },
  folHlink: { space: 'srgb', hex: fixture.clrScheme.folHlink, transforms: [] },
};

const IDENTITY_MAP: ClrMap = {
  bg1: 'lt1',
  tx1: 'dk1',
  bg2: 'lt2',
  tx2: 'dk2',
  accent1: 'accent1',
  accent2: 'accent2',
  accent3: 'accent3',
  accent4: 'accent4',
  accent5: 'accent5',
  accent6: 'accent6',
  hlink: 'hlink',
  folHlink: 'folHlink',
};

const CROSSED_MAP = bases.crossedClrMap as ClrMap;

const CTX: ColorContext = { scheme: SCHEME, map: IDENTITY_MAP };

type FixtureBase =
  | { kind: 'srgb'; value: string }
  | { kind: 'scheme'; value: string }
  | { kind: 'scrgb'; r: number; g: number; b: number }
  | { kind: 'hsl'; hue: number; sat: number; lum: number }
  | { kind: 'sys'; value: string; lastClr?: string }
  | { kind: 'prst'; value: string }
  | { kind: 'raw'; xml: string; value: string };

function toColor(base: FixtureBase, list: readonly FixtureTransform[]): Color | null {
  const t = transforms(list);
  switch (base.kind) {
    case 'srgb':
      return { space: 'srgb', hex: base.value, transforms: t };
    case 'scheme':
      return { space: 'scheme', name: base.value as never, transforms: t };
    case 'scrgb':
      return { space: 'scrgb', r: base.r, g: base.g, b: base.b, transforms: t };
    case 'hsl':
      return { space: 'hsl', hue: base.hue, sat: base.sat, lum: base.lum, transforms: t };
    case 'sys':
      return { space: 'sys', name: base.value, lastClr: base.lastClr ?? null, transforms: t };
    case 'prst':
      return { space: 'prst', name: base.value, transforms: t };
    case 'raw':
      return null; // a grammar probe; no typed base can express it
  }
}

/** Every swatch a `Color` can be built for, with the answer PowerPoint gave. */
interface Row {
  id: string;
  group: string;
  description: string;
  color: Color;
  expected: string;
  transparency: number | null;
}

function rowsOf(
  swatches: readonly {
    id: string;
    group: string;
    description: string;
    base: unknown;
    transforms: FixtureTransform[];
    expected: string;
    transparency?: number | null;
    fromRepairedDeck?: boolean;
  }[],
): Row[] {
  const out: Row[] = [];
  for (const s of swatches) {
    if (s.fromRepairedDeck === true) continue;
    const color = toColor(s.base as FixtureBase, s.transforms);
    if (color === null) continue;
    out.push({
      id: s.id,
      group: s.group,
      description: s.description,
      color,
      expected: s.expected,
      transparency: s.transparency ?? null,
    });
  }
  return out;
}

const C1 = rowsOf(fixture.swatches);
const C2 = rowsOf(bases.swatches);

/**
 * What PowerPoint's bitmap export shows for a translucent swatch over the white
 * slide behind it.
 *
 * Two things worth noting, and neither is this package's business.
 *
 * The compositing is plain sRGB `a*c + (1-a)*bg` - PowerPoint does *not*
 * composite in linear light, even though it blends `tint` and `shade` there.
 *
 * And its rounding here is not `toByte`'s. All three channels of the 50% swatch
 * land on an exact tie - 161.5, 184.5, 225.5 - and the rasteriser takes the
 * lower byte on all three, where the colour engine's own ties go both ways. So
 * this helper rounds halves down rather than reusing `toByte`, because it is
 * modelling a different piece of PowerPoint.
 */
function overWhite(r: number, g: number, b: number, a: number): string {
  const byte = (c: number): number => Math.min(255, Math.max(0, Math.ceil(c * 255 - 0.5)));
  // And the compositing happens on the **quantised** colour, not on the real
  // number the transform chain produced. The `lumMod`/`lumOff` swatch is what
  // shows it: its green resolves to a value a hair above `AA`, and compositing
  // the real number gives `D5` where PowerPoint gives `D4` - which is what
  // compositing `AA` gives. The colour becomes eight bits per channel first.
  return [toByte(r) / 255, toByte(g) / 255, toByte(b) / 255]
    .map((c) =>
      byte(a * c + (1 - a))
        .toString(16)
        .padStart(2, '0'),
    )
    .join('')
    .toUpperCase();
}

/* -------------------------------------------------------------------------- */

describe('the fixtures are what they claim to be', () => {
  it('carries both experiments, and no swatch from a deck PowerPoint repaired', () => {
    expect(fixture.swatches.length).toBe(214);
    expect(bases.swatches.length).toBe(259);
    // Two swatches wrote an out-of-range `a:hslClr/@hue`. PowerPoint refused the
    // package and only opened it after a repair, which reset the hue to zero.
    // A repaired deck measures the repair, so those two are excluded here and
    // the exclusion is asserted rather than assumed.
    const repaired = bases.swatches.filter((s) => s.fromRepairedDeck);
    expect(repaired.map((s) => s.group)).toEqual(['hsl-range', 'hsl-range']);
    expect(C1.length).toBe(214);
    expect(C2.length).toBe(254); // 259 less the 2 repaired and the 3 grammar probes
  });

  it('was read back two ways that agree', () => {
    const opaque = [...fixture.swatches, ...bases.swatches].filter(
      (s) => (('transparency' in s ? s.transparency : 0) ?? 0) === 0,
    );
    const disagree = opaque.filter((s) => s.objectModel !== null && s.objectModel !== s.expected);
    expect(disagree).toEqual([]);
    expect(opaque.length).toBeGreaterThan(460);
  });
});

/* -------------------------------------------------------------------------- */

describe('0.7-C: every transform, in the space PowerPoint uses', () => {
  it('reproduces all 214 swatches exactly', () => {
    const wrong: string[] = [];
    for (const row of C1) {
      const got = toHexColor(resolveColor(row.color, CTX));
      if (got !== row.expected) wrong.push(`${row.description} -> ${got} want ${row.expected}`);
    }
    expect(wrong).toEqual([]);
  });

  it('and a renderer that blended in sRGB instead of linear light would be 73/255 out', () => {
    // The measurement restated as the size of the mistake it prevents, computed
    // over the same 102 swatches rather than asserted from one hand-picked pair.
    const blends = C1.filter((r) => r.group.startsWith('tint') || r.group.startsWith('shade'));
    expect(blends.length).toBe(102);
    let worst = 0;
    let where = '';
    for (const row of blends) {
      const base = row.color as Extract<Color, { space: 'srgb' }>;
      const rgb = parseHex(base.hex) ?? [0, 0, 0];
      const t = base.transforms[0] as Extract<ColorTransform, { val: number }>;
      const p = t.val / 100000;
      const naive = rgb.map((c) => (t.op === 'tint' ? c * p + (1 - p) : c * p)) as [
        number,
        number,
        number,
      ];
      const got = toHex(naive[0], naive[1], naive[2]);
      const d = Math.max(
        ...[0, 2, 4].map((i) =>
          Math.abs(parseInt(got.slice(i, i + 2), 16) - parseInt(row.expected.slice(i, i + 2), 16)),
        ),
      );
      if (d > worst) {
        worst = d;
        where = `${row.description} naive=${got} actual=${row.expected}`;
      }
    }
    expect(where).toBe('#FF0000 tint=75000 naive=FF4040 actual=FF8989');
    expect(worst).toBe(73);
  });
});

/* -------------------------------------------------------------------------- */

describe('2.6: the six bases', () => {
  const groups = (...names: string[]): Row[] => C2.filter((r) => names.includes(r.group));

  it('a:scrgbClr is linear light, and reading it as sRGB percentages is 61/255 wrong', () => {
    const wrong: string[] = [];
    for (const row of groups('scrgb', 'scrgb-grey', 'scrgb-range')) {
      const got = toHexColor(resolveColor(row.color, CTX));
      if (got !== row.expected) wrong.push(`${row.description} -> ${got} want ${row.expected}`);
    }
    expect(wrong).toEqual([]);
    // The headline: half of full scale is BC, not 80.
    const half = resolveColor({ space: 'scrgb', r: 50000, g: 50000, b: 50000, transforms: [] });
    expect(toHexColor(half)).toBe('BCBCBC');
    expect(toByte(0.5)).toBe(127); // what a percentage reading would paint
  });

  it('a:hslClr uses the same bi-hexcone the lum and sat transforms do', () => {
    const wrong: string[] = [];
    for (const row of groups('hsl-hue', 'hsl-sat', 'hsl-lum')) {
      const got = toHexColor(resolveColor(row.color, CTX));
      if (got !== row.expected) wrong.push(`${row.description} -> ${got} want ${row.expected}`);
    }
    // Two of these land on an exact tie; see the tie table in `transfer.ts`.
    expect(wrong).toEqual([]);
  });

  it('a:prstClr resolves all 147 names, and none of them is a CSS colour name', () => {
    const rows = groups('prst', 'prst-grey');
    expect(rows.length).toBe(147);
    for (const row of rows) {
      expect(toHexColor(resolveColor(row.color, CTX))).toBe(row.expected);
    }
    // The table in the package and the table PowerPoint painted are the same
    // table. This is the only thing stopping the two from drifting.
    expect(Object.keys(PRESET_COLORS).length).toBe(147);
    // Abbreviated, every one of them. `color: dkBlue` is not a colour to a
    // browser, it is nothing at all - which is why the table has to exist.
    expect(PRESET_COLORS['dkBlue']).toBe('00008B');
    expect(PRESET_COLORS['ltGoldenrodYellow']).toBe('FAFAD2');
    expect(PRESET_COLORS['medOrchid']).toBe('BA55D3');
    // Both spellings of the seven grey names are in the enumeration, and agree.
    expect(PRESET_COLORS['dkGrey']).toBe(PRESET_COLORS['dkGray']);
    expect(PRESET_COLORS['ltSlateGrey']).toBe(PRESET_COLORS['ltSlateGray']);
    // The classic trap: HTML's grey, not X11's BEBEBE.
    expect(PRESET_COLORS['gray']).toBe('808080');
  });

  it('a:sysClr resolves the thirty names, and prefers @lastClr - which PowerPoint does not', () => {
    for (const row of groups('sys')) {
      expect(toHexColor(resolveColor(row.color, CTX))).toBe(row.expected);
    }
    expect(Object.keys(SYSTEM_COLORS).length).toBe(30);

    // The deliberate divergence, asserted so it cannot be changed by accident.
    // PowerPoint painted these three the *machine's* colour and ignored the
    // attribute. A browser has no machine to ask, so we take the author's cached
    // value instead - which is the colour the deck had where it was written.
    const lastClr = groups('sys-lastclr');
    expect(lastClr.length).toBe(3);
    for (const row of lastClr) {
      const base = row.color as Extract<Color, { space: 'sys' }>;
      expect(toHexColor(resolveColor(row.color, CTX))).toBe(base.lastClr);
      expect(toHexColor(resolveColor(row.color, CTX))).not.toBe(row.expected);
    }
  });

  it('falls back to the table when there is no @lastClr, and throws when there is neither', () => {
    const noLast: Color = { space: 'sys', name: 'btnFace', lastClr: null, transforms: [] };
    expect(toHexColor(resolveColor(noLast))).toBe(SYSTEM_COLORS['btnFace']);
    const unknown: Color = { space: 'sys', name: 'teapot', lastClr: null, transforms: [] };
    expect(() => resolveColor(unknown)).toThrowError(PaintError);
    try {
      resolveColor(unknown);
    } catch (error) {
      expect((error as PaintError).code).toBe('COLOR_SYS_UNKNOWN');
    }
  });
});

/* -------------------------------------------------------------------------- */

describe('2.6: every transform hands the next one a clamped colour', () => {
  it('reproduces the chains that prove it', () => {
    const wrong: string[] = [];
    for (const row of C2.filter((r) => r.group === 'clamp')) {
      const got = toHexColor(resolveColor(row.color, CTX));
      if (got !== row.expected) wrong.push(`${row.description} -> ${got} want ${row.expected}`);
    }
    expect(wrong).toEqual([]);
  });

  it('and carrying the state across the boundary would be 94/255 out', () => {
    // Lightness 0.5176 + 0.6 is 1.1176, which paints white. Taking 0.6 back off
    // *white* is mid-grey. A model that carries the un-clamped lightness returns
    // the colour it started with, and 0.7-C could not see the difference because
    // no chain in it drove a value out of range twice.
    const there = applyTransforms({ r: 68 / 255, g: 114 / 255, b: 196 / 255, a: 1 }, [
      { op: 'lumOff', val: 60000 },
    ]);
    expect(toHexColor(there)).toBe('FFFFFF');
    const andBack = applyTransforms(there, [{ op: 'lumOff', val: -60000 }]);
    expect(toHexColor(andBack)).toBe('666666');
    expect(toHexColor(andBack)).not.toBe('4472C4');
  });

  it('but the arithmetic inside one transform is not clamped', () => {
    // Saturation of 1.04 and 1.56 are different colours. A model that pins
    // saturation at 1 gives one answer for both and is 19/255 wrong on the
    // second - which is the finding 0.7-C made and this one does not undo.
    const base = { r: 68 / 255, g: 114 / 255, b: 196 / 255, a: 1 };
    expect(toHexColor(applyTransforms(base, [{ op: 'satMod', val: 200000 }]))).toBe('0460FF');
    expect(toHexColor(applyTransforms(base, [{ op: 'satMod', val: 300000 }]))).toBe('004EFF');
  });

  it('and neither a gamma round trip nor an alpha interrupts anything', () => {
    const base = { r: 68 / 255, g: 114 / 255, b: 196 / 255, a: 1 };
    const plain = applyTransforms(base, [
      { op: 'satMod', val: 300000 },
      { op: 'satMod', val: 33333 },
    ]);
    const viaGamma = applyTransforms(base, [
      { op: 'satMod', val: 300000 },
      { op: 'gamma' },
      { op: 'invGamma' },
      { op: 'satMod', val: 33333 },
    ]);
    const viaAlpha = applyTransforms(base, [
      { op: 'satMod', val: 300000 },
      { op: 'alpha', val: 100000 },
      { op: 'satMod', val: 33333 },
    ]);
    expect(toHexColor(plain)).toBe('556FAA');
    expect(toHexColor(viaGamma)).toBe('556FAA');
    expect(toHexColor(viaAlpha)).toBe('556FAA');
  });
});

/* -------------------------------------------------------------------------- */

describe('2.6: alpha', () => {
  it('is 1 - Transparency, and composites over the slide exactly as PowerPoint did', () => {
    const rows = C2.filter((r) => r.group.startsWith('alpha'));
    expect(rows.length).toBe(10);
    for (const row of rows) {
      const got = resolveColor(row.color, CTX);
      expect(got.a).toBeCloseTo(1 - (row.transparency ?? 0), 6);
      // The slide behind these swatches is `bg1`, which is white. If our colour
      // and our alpha are both right, the composite is what the export shows.
      expect(overWhite(got.r, got.g, got.b, got.a)).toBe(row.expected);
    }
  });

  it('does not disturb the colour, wherever it appears in the chain', () => {
    const base = { r: 68 / 255, g: 114 / 255, b: 196 / 255, a: 1 };
    const before = applyTransforms(base, [
      { op: 'alpha', val: 50000 },
      { op: 'lumMod', val: 60000 },
      { op: 'lumOff', val: 40000 },
    ]);
    const middle = applyTransforms(base, [
      { op: 'lumMod', val: 60000 },
      { op: 'alpha', val: 50000 },
      { op: 'lumOff', val: 40000 },
    ]);
    expect(toHexColor(before)).toBe('8FAADC');
    expect(toHexColor(middle)).toBe('8FAADC');
    expect(before.a).toBe(0.5);
    expect(middle.a).toBe(0.5);
  });
});

/* -------------------------------------------------------------------------- */

describe('2.6: the colour map', () => {
  it('sends the twelve mapped names through it and the four slot names round it', () => {
    const ctx: ColorContext = { scheme: SCHEME, map: CROSSED_MAP };
    const wrong: string[] = [];
    for (const row of C2.filter((r) => r.group === 'clrmap')) {
      const got = toHexColor(resolveColor(row.color, ctx));
      if (got !== row.expected) wrong.push(`${row.description} -> ${got} want ${row.expected}`);
    }
    expect(wrong).toEqual([]);
  });

  it('states the rule directly, because on an identity map both readings agree', () => {
    // Which is why every deck in the corpus is silent about it, and why this
    // needed a master written on purpose.
    expect(mapSchemeName('dk1', CROSSED_MAP)).toBe('dk1');
    expect(mapSchemeName('lt2', CROSSED_MAP)).toBe('lt2');
    expect(mapSchemeName('bg1', CROSSED_MAP)).toBe('dk1');
    expect(mapSchemeName('tx1', CROSSED_MAP)).toBe('lt1');
    expect(mapSchemeName('accent1', CROSSED_MAP)).toBe('accent2');
    expect(mapSchemeName('phClr', CROSSED_MAP)).toBeNull();
    // With no map at all, the four bg/tx names still mean something.
    expect(mapSchemeName('bg1', undefined)).toBe('lt1');
    expect(mapSchemeName('tx2', undefined)).toBe('dk2');
  });

  it('resolves a theme slot that is itself a colour with transforms', () => {
    const scheme: ClrScheme = {
      ...SCHEME,
      accent1: { space: 'srgb', hex: '4472C4', transforms: [{ op: 'shade', val: 50000 }] },
    };
    const got = resolveColor({ space: 'scheme', name: 'accent1', transforms: [] }, { scheme });
    expect(toHexColor(got)).toBe(
      toHexColor(
        applyTransforms({ r: 68 / 255, g: 114 / 255, b: 196 / 255, a: 1 }, [
          { op: 'shade', val: 50000 },
        ]),
      ),
    );
  });

  it('refuses a theme that resolves a slot to itself', () => {
    const scheme: ClrScheme = {
      ...SCHEME,
      dk1: { space: 'scheme', name: 'tx1', transforms: [] },
    };
    try {
      resolveColor({ space: 'scheme', name: 'tx1', transforms: [] }, { scheme, map: IDENTITY_MAP });
      expect.unreachable('a cycle should throw');
    } catch (error) {
      expect((error as PaintError).code).toBe('COLOR_SCHEME_CYCLE');
    }
  });
});

/* -------------------------------------------------------------------------- */

describe('2.6: the percentage grammar', () => {
  it('accepts both spellings, because PowerPoint does', () => {
    // Measured: `<a:lumMod val="60%"/><a:lumOff val="40%"/>` in a Transitional
    // slide opened with no repair and painted the same colour, to the byte, as
    // the `60000`/`40000` control alongside it.
    expect(parsePercentage('60000')).toBe(60000);
    expect(parsePercentage('60%')).toBe(60000);
    expect(parsePercentage('60.5%')).toBe(60500);
    expect(parsePercentage('-12.5%')).toBe(-12500);
    expect(parsePercentage('.5%')).toBe(500);
    expect(parsePercentage('-40000')).toBe(-40000);
  });

  it('and the two spellings resolve to the same colour the deck showed', () => {
    const literal: Color = {
      space: 'srgb',
      hex: '4472C4',
      transforms: [
        { op: 'lumMod', val: parsePercentage('60%') },
        { op: 'lumOff', val: parsePercentage('40%') },
      ],
    };
    expect(toHexColor(resolveColor(literal, CTX))).toBe('8FAADC');
    const control = C2.find(
      (r) => r.group === 'percent-control' && r.description.includes('lumMod'),
    );
    expect(control?.expected).toBe('8FAADC');
  });

  it('does not let parseInt eat the per-cent sign', () => {
    // The whole reason this function exists. `parseInt("150%")` is 150, so a
    // 150% transform silently becomes 0.15%.
    expect(parsePercentage('150%')).toBe(150000);
    expect(parseInt('150%', 10)).toBe(150);
  });

  it('rejects what is not a number, with a code', () => {
    for (const bad of ['', 'x', '40 %', '40%%', '4e5', '0x10']) {
      try {
        parsePercentage(bad);
        expect.unreachable(`"${bad}" should not parse`);
      } catch (error) {
        expect((error as PaintError).code).toBe('COLOR_NUMBER');
      }
    }
    expect(parseAngle('-3600000')).toBe(-3600000);
    expect(() => parseAngle('60%')).toThrowError(PaintError);
    expect(parseSrgbValue('#4472c4')).toBe('4472C4');
    expect(() => parseSrgbValue('4472C')).toThrowError(PaintError);
  });
});

/* -------------------------------------------------------------------------- */

describe('the transfer functions', () => {
  it('are exact inverses, and a:gamma is one of them', () => {
    for (let i = 0; i <= 255; i++) {
      const c = i / 255;
      expect(fromLinear(toLinear(c))).toBeCloseTo(c, 12);
    }
    // The format exposes the curve directly, which is what makes the tint and
    // shade result more than a curve fit.
    const base = { r: 68 / 255, g: 114 / 255, b: 196 / 255, a: 1 };
    expect(toHexColor(applyTransforms(base, [{ op: 'gamma' }]))).toBe('8DB2E3');
    expect(toHexColor(applyTransforms(base, [{ op: 'invGamma' }]))).toBe('0F2B8D');
    expect(toHexColor(applyTransforms(base, [{ op: 'gamma' }, { op: 'invGamma' }]))).toBe('4472C4');
  });

  it('round-trips HSL, including values outside 0..1', () => {
    const there = rgbToHsl(68 / 255, 114 / 255, 196 / 255);
    const [r, g, b] = hslToRgb(there);
    expect(toHex(r, g, b)).toBe('4472C4');
    // The out-of-range case is the one that matters: `hslToRgb` has to carry a
    // saturation of 1.56 through the arithmetic rather than reject it.
    const wide = hslToRgb({ h: there.h, s: there.s * 3, l: there.l });
    expect(Math.max(...wide)).toBeGreaterThan(1);
    expect(Math.min(...wide)).toBeLessThan(0);
  });

  it('breaks a tie towards odd, which fits six of the seven ties measured', () => {
    // The table in `transfer.ts`. Recorded here because it is a fit rather than
    // a rule, and the one it misses is named rather than hidden.
    expect(toByte(127.5 / 255)).toBe(127);
    expect(toByte(153.5 / 255)).toBe(153);
    expect(toByte(93.5 / 255)).toBe(93);
    expect(toByte(110.5 / 255)).toBe(111);
    expect(toByte(134.5 / 255)).toBe(135);
    // The counterexample. PowerPoint paints 130 here; we paint 129.
    expect(toByte(129.5 / 255)).toBe(129);
  });

  it('clamps rather than wrapping, in both directions', () => {
    expect(toByte(-5)).toBe(0);
    expect(toByte(5)).toBe(255);
    expect(toByte(Number.NaN)).toBe(0);
  });
});

/* -------------------------------------------------------------------------- */

describe('what a caller has to supply, and what happens when they do not', () => {
  it('names the missing piece rather than painting black', () => {
    const cases: [Color, string][] = [
      [{ space: 'scheme', name: 'accent1', transforms: [] }, 'COLOR_NO_SCHEME'],
      [{ space: 'scheme', name: 'phClr', transforms: [] }, 'COLOR_NO_PHCLR'],
      [{ space: 'prst', name: 'burntSienna', transforms: [] }, 'COLOR_PRST_UNKNOWN'],
      [{ space: 'srgb', hex: 'GGGGGG', transforms: [] }, 'COLOR_HEX'],
      [{ space: 'srgb', hex: '4472C4', transforms: [{ op: 'blur' } as never] }, 'COLOR_TRANSFORM'],
    ];
    for (const [color, code] of cases) {
      try {
        resolveColor(color);
        expect.unreachable(`${code} should have been thrown`);
      } catch (error) {
        expect(error).toBeInstanceOf(PaintError);
        expect((error as PaintError).code).toBe(code);
      }
    }
  });

  it('resolves phClr from the style invocation', () => {
    const phClr = { r: 1, g: 0, b: 0, a: 1 };
    const got = resolveColor(
      { space: 'scheme', name: 'phClr', transforms: [{ op: 'shade', val: 50000 }] },
      { phClr },
    );
    expect(toHexColor(got)).toBe('BC0000');
  });

  it('emits CSS a browser will take', () => {
    expect(toCss({ r: 68 / 255, g: 114 / 255, b: 196 / 255, a: 1 })).toBe('#4472C4');
    expect(toCss({ r: 68 / 255, g: 114 / 255, b: 196 / 255, a: 0.5 })).toBe(
      'rgba(68, 114, 196, 0.5)',
    );
  });
});

/* -------------------------------------------------------------------------- */

describe('the whole of both fixtures, in one number', () => {
  it('is exact on every swatch but the one tie', () => {
    const wrong: string[] = [];
    const ctx = (row: Row): ColorContext =>
      row.group === 'clrmap' ? { scheme: SCHEME, map: CROSSED_MAP } : CTX;
    let checked = 0;
    for (const row of [...C1, ...C2]) {
      // The three `sys-lastclr` rows are the deliberate divergence, asserted
      // above; the alpha rows are composites, asserted above.
      if (row.group === 'sys-lastclr' || row.group.startsWith('alpha')) continue;
      checked += 1;
      const got = toHexColor(resolveColor(row.color, ctx(row)));
      if (got !== row.expected)
        wrong.push(`${row.group} ${row.description}: ${got} != ${row.expected}`);
    }
    expect(checked).toBe(455);
    expect(wrong).toEqual(['edge #4472C4 satOff=-50000: 818387 != 828387']);
  });
});

/** Compile-time proof that the slot list and the type have not drifted apart. */
const _slots: SchemeSlot[] = [
  'dk1',
  'lt1',
  'dk2',
  'lt2',
  'accent1',
  'accent2',
  'accent3',
  'accent4',
  'accent5',
  'accent6',
  'hlink',
  'folHlink',
];
void _slots;
