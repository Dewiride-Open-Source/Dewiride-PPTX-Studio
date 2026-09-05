/**
 * Sub-phase 0.7's findings, as assertions.
 *
 * The experiments themselves need PowerPoint and cannot run in CI. Their
 * *answers* can, and a fixture nothing reads is a fixture that rots: the point
 * of these tests is that the committed measurements stay honest and that the
 * conclusions drawn from them are executable rather than prose.
 *
 * Phase 2.6 will implement colour transforms in `@pptx-studio/paint`. When it
 * does, the model it implements has to reproduce this file. If someone
 * "simplifies" the linearisation to a gamma constant, this suite goes red with
 * the exact swatch that proves it wrong.
 */

import { readFileSync } from 'node:fs';
import { join } from 'node:path';

import { REPO_ROOT } from '../repo/root.ts';
import { describe, expect, it } from 'vitest';

import { buildProbeFont } from './fonts/embedding/build-font.ts';
import {
  BLEND_MODELS,
  HSL_MODELS,
  parseHex,
  type Channel,
  type HslOp,
} from './paint/colour/transforms/models.ts';
import {
  eotFontData,
  flagNames,
  readEot,
  writeEot,
  EOT_FLAG,
  EOT_VERSION_2_2,
} from './fonts/format/eot.ts';
import { NAME_ID, readNames, readOs2, readSfnt } from './fonts/format/sfnt.ts';

const root = REPO_ROOT;
const corpus = join(root, 'corpus', 'ground-truth');

const read = (...parts: string[]): Uint8Array =>
  new Uint8Array(readFileSync(join(corpus, ...parts)));
const readJson = <T>(name: string): T => JSON.parse(readFileSync(join(corpus, name), 'utf8')) as T;

/* -------------------------------------------------------------------------- */
/* C - colour transforms                                                      */
/* -------------------------------------------------------------------------- */

interface Swatch {
  id: string;
  group: string;
  description: string;
  base: { kind: 'srgb' | 'scheme'; value: string };
  transforms: { name: string; val?: number }[];
  expected: string;
  objectModel: string;
}
interface ColorFixture {
  clrScheme: Record<string, string>;
  swatches: Swatch[];
}

/**
 * PowerPoint takes the lower value when a channel lands exactly on x.5.
 * Measured: half-down fits 171 of 171, `Math.round` fits 168. See
 * docs/adr/phase-0-foundation/0007-ground-truth.md §C.
 */
const toByte = (c: Channel): number =>
  Math.max(0, Math.min(255, Math.ceil(Math.max(0, Math.min(1, c)) * 255 - 0.5)));

const toHex = (rgb: readonly [Channel, Channel, Channel]): string =>
  rgb.map((c) => toByte(c).toString(16).padStart(2, '0').toUpperCase()).join('');

const CLR_MAP: Record<string, string> = { bg1: 'lt1', tx1: 'dk1', bg2: 'lt2', tx2: 'dk2' };

describe('experiment C - DrawingML colour transforms', () => {
  const fixture = readJson<ColorFixture>('color-transforms.json');

  const baseRgb = (swatch: Swatch): [Channel, Channel, Channel] => {
    if (swatch.base.kind === 'srgb') return parseHex(swatch.base.value);
    const slot = CLR_MAP[swatch.base.value] ?? swatch.base.value;
    const value = fixture.clrScheme[slot];
    if (value === undefined) throw new Error(`no scheme colour ${slot}`);
    return parseHex(value);
  };

  it('is the measurement it says it is', () => {
    expect(fixture.swatches).toHaveLength(214);
    // Every swatch was read back two independent ways - PowerPoint's object
    // model and its own bitmap export. That they agreed everywhere is the
    // reason the fixture can be trusted at all.
    const disagreeing = fixture.swatches.filter((s) => s.expected !== s.objectModel);
    expect(disagreeing).toEqual([]);
  });

  it('leaves a colour with no transform alone', () => {
    for (const s of fixture.swatches.filter((x) => x.group === 'identity')) {
      expect(s.expected, s.description).toBe(s.base.value);
    }
  });

  it('resolves scheme colours through the clrMap', () => {
    for (const s of fixture.swatches.filter((x) => x.group === 'scheme')) {
      expect(toHex(baseRgb(s)), s.description).toBe(s.expected);
    }
  });

  const blendSwatches = (name: 'tint' | 'shade'): Swatch[] =>
    fixture.swatches.filter((s) => s.transforms.length === 1 && s.transforms[0]!.name === name);

  for (const name of ['tint', 'shade'] as const) {
    it(`computes ${name} in linearised sRGB, exactly`, () => {
      const model = BLEND_MODELS.find((m) => m.id === 'linear-srgb')!;
      const f = name === 'tint' ? model.tint : model.shade;
      const swatches = blendSwatches(name);
      expect(swatches.length).toBeGreaterThan(40);
      for (const s of swatches) {
        const p = (s.transforms[0]!.val ?? 0) / 100000;
        const rgb = baseRgb(s);
        expect(toHex([f(rgb[0], p), f(rgb[1], p), f(rgb[2], p)]), s.description).toBe(s.expected);
      }
    });

    it(`is NOT ${name} under a gamma 2.3 power law - the swatch that proves it`, () => {
      // Guards against "simplifying" the transfer function to LibreOffice's
      // constant. If this ever stops failing, the model above stopped being
      // discriminating and the test above stopped meaning anything.
      const gamma = BLEND_MODELS.find((m) => m.id === 'gamma2.3')!;
      const f = name === 'tint' ? gamma.tint : gamma.shade;
      const wrong = blendSwatches(name).filter((s) => {
        const p = (s.transforms[0]!.val ?? 0) / 100000;
        const rgb = baseRgb(s);
        return toHex([f(rgb[0], p), f(rgb[1], p), f(rgb[2], p)]) !== s.expected;
      });
      expect(wrong.length).toBeGreaterThan(0);
    });
  }

  const HSL_NAMES = new Set([
    'lumMod',
    'lumOff',
    'lum',
    'satMod',
    'satOff',
    'sat',
    'hueMod',
    'hueOff',
    'hue',
  ]);

  it('computes lum/sat/hue on the sRGB values, with no linearisation', () => {
    const model = HSL_MODELS.find((m) => m.id === 'hsl-srgb')!;
    const swatches = fixture.swatches.filter(
      (s) => s.transforms.length > 0 && s.transforms.every((t) => HSL_NAMES.has(t.name)),
    );
    expect(swatches.length).toBeGreaterThan(60);
    for (const s of swatches) {
      const ops: HslOp[] = s.transforms.map((t) => {
        const isHue = t.name.startsWith('hue');
        const p = isHue && t.name !== 'hueMod' ? (t.val ?? 0) / 60000 : (t.val ?? 0) / 100000;
        return { kind: t.name, p } as HslOp;
      });
      expect(toHex(model.apply(baseRgb(s), ops)), s.description).toBe(s.expected);
    }
  });

  it('does not clamp saturation before converting back', () => {
    // satMod 200% and 300% on one base give *different* colours. A model that
    // clamped S to 1 would give the same answer for both, and this is the pair
    // that catches it.
    const at = (val: number): Swatch | undefined =>
      fixture.swatches.find(
        (s) =>
          s.base.value === '4472C4' &&
          s.transforms.length === 1 &&
          s.transforms[0]!.name === 'satMod' &&
          s.transforms[0]!.val === val,
      );
    const two = at(200000);
    const three = at(300000);
    expect(two).toBeDefined();
    expect(three).toBeDefined();
    expect(two!.expected).not.toBe(three!.expected);
  });

  it('applies transforms in document order, not a canonical one', () => {
    const order = fixture.swatches.filter((s) => s.group === 'order');
    expect(order.length).toBeGreaterThanOrEqual(8);
    // Each pair is the same two transforms the other way round.
    for (let i = 0; i < order.length; i += 2) {
      const a = order[i]!;
      const b = order[i + 1]!;
      expect(a.expected, `${a.description} vs ${b.description}`).not.toBe(b.expected);
    }
  });

  const linear = (c: Channel): Channel =>
    c <= 0.04045 ? c / 12.92 : Math.pow((c + 0.055) / 1.055, 2.4);
  const encode = (c: Channel): Channel =>
    c <= 0.0031308 ? c * 12.92 : 1.055 * Math.pow(c, 1 / 2.4) - 0.055;

  const single = (name: string): Swatch[] =>
    fixture.swatches.filter((s) => s.transforms.length === 1 && s.transforms[0]!.name === name);

  it('a:inv complements in linearised sRGB, not in sRGB', () => {
    for (const s of single('inv')) {
      const rgb = baseRgb(s);
      expect(
        toHex([0, 1, 2].map((i) => encode(1 - linear(rgb[i]!))) as [Channel, Channel, Channel]),
        s.description,
      ).toBe(s.expected);
    }
  });

  it('a:gamma is the sRGB encode and a:invGamma its exact inverse', () => {
    for (const s of single('gamma')) {
      const rgb = baseRgb(s);
      expect(toHex(rgb.map(encode) as [Channel, Channel, Channel]), s.description).toBe(s.expected);
    }
    for (const s of single('invGamma')) {
      const rgb = baseRgb(s);
      expect(toHex(rgb.map(linear) as [Channel, Channel, Channel]), s.description).toBe(s.expected);
    }
  });

  it('a:gray is Rec.709 luma of the sRGB values, with no linearisation', () => {
    for (const s of single('gray')) {
      const [r, g, b] = baseRgb(s);
      const y = 0.2126 * r + 0.7152 * g + 0.0722 * b;
      expect(toHex([y, y, y]), s.description).toBe(s.expected);
    }
  });

  it('per-channel Mod and Off work in linearised sRGB', () => {
    const channels = fixture.swatches.filter(
      (s) => s.group === 'channelMod' || s.group === 'channelOff',
    );
    expect(channels).toHaveLength(6);
    for (const s of channels) {
      const rgb = [...baseRgb(s)] as [Channel, Channel, Channel];
      const t = s.transforms[0]!;
      const i = t.name.startsWith('red') ? 0 : t.name.startsWith('green') ? 1 : 2;
      const p = (t.val ?? 0) / 100000;
      const channel = linear(rgb[i]);
      rgb[i] = encode(t.name.endsWith('Mod') ? channel * p : channel + p);
      expect(toHex(rgb), s.description).toBe(s.expected);
    }
  });
});

/* -------------------------------------------------------------------------- */
/* A - what PowerPoint writes                                                 */
/* -------------------------------------------------------------------------- */

interface EotFixture {
  decks: {
    deck: string;
    contentTypeDefault: string | null;
    embedTrueTypeFonts: string | null;
    fonts: {
      part: string;
      flags: string;
      flagNames: string[];
      version: string;
      familyName: string;
      styleName: string;
      versionName: string;
      fullName: string;
      italic: number;
      charset: number;
      payloadBeginsWithSfntSignature: boolean;
    }[];
  }[];
}

describe('experiment A - what PowerPoint writes into .fntdata', () => {
  const fixture = readJson<EotFixture>('eot-headers.json');
  const fonts = fixture.decks.flatMap((d) => d.fonts);

  it('measured four decks and 76 embedded font parts', () => {
    expect(fixture.decks).toHaveLength(4);
    expect(fonts).toHaveLength(76);
  });

  it('compresses every single one with MicroType Express', () => {
    // The finding that reshapes Phase 8. If this ever goes green-with-zero,
    // somebody regenerated the fixture from decks that prove something else.
    expect(fonts.length).toBeGreaterThan(0);
    for (const font of fonts) {
      expect(font.flagNames, font.part).toContain('TTCOMPRESSED');
      expect(font.payloadBeginsWithSfntSignature, font.part).toBe(false);
    }
  });

  it('writes EOT version 2.2, never version 1', () => {
    for (const font of fonts) expect(font.version, font.part).toBe('0x00020002');
  });

  it('subsetting is a setting; compression is not', () => {
    const subset = fonts.filter((f) => f.flagNames.includes('SUBSET'));
    expect(subset.length).toBeGreaterThan(0);
    expect(subset.length).toBeLessThan(fonts.length);
    // ... and the ones that are not subset are still compressed.
    for (const font of fonts.filter((f) => !f.flagNames.includes('SUBSET'))) {
      expect(font.flagNames, font.part).toContain('TTCOMPRESSED');
    }
  });

  // NUL is written as an escape, never as a literal: a literal NUL is invisible
  // in an editor and does not survive the first person who tidies this file.
  it('NUL-terminates FamilyName and StyleName but not VersionName or FullName', () => {
    for (const font of fonts) {
      expect(font.familyName.endsWith(' '), `${font.part} family`).toBe(true);
      expect(font.styleName.endsWith(' '), `${font.part} style`).toBe(true);
      expect(font.versionName.endsWith(' '), `${font.part} version`).toBe(false);
      expect(font.fullName.endsWith(' '), `${font.part} full`).toBe(false);
    }
  });

  it('writes 0xFF for italic, not 0x01', () => {
    const italics = fonts.filter((f) => f.italic !== 0);
    expect(italics.length).toBeGreaterThan(0);
    for (const font of italics) expect(font.italic, font.part).toBe(255);
  });

  it('always emits the fntdata content-type Default and embedTrueTypeFonts', () => {
    for (const deck of fixture.decks) {
      expect(deck.contentTypeDefault, deck.deck).toBe(
        '<Default Extension="fntdata" ContentType="application/x-fontdata"/>',
      );
      expect(deck.embedTrueTypeFonts, deck.deck).toBe('1');
    }
  });
});

/* -------------------------------------------------------------------------- */
/* B - the probe font and our EOT writer                                      */
/* -------------------------------------------------------------------------- */

describe('experiment B - the probe font and the EOT writer', () => {
  it('rebuilds the committed probe font byte for byte', () => {
    // Deterministic by construction: no timestamps in `head`, no ordering that
    // depends on anything but the tags. If this drifts, every measurement taken
    // against the committed font is measuring a different font.
    expect(buildProbeFont('ProbeAlpha').bytes).toEqual(read('fonts', 'probe.ttf'));
  });

  it('builds a font that reads back as what it claims to be', () => {
    const font = readSfnt(read('fonts', 'probe.ttf'));
    const names = readNames(font);
    expect(names.get(NAME_ID.family)).toBe('ProbeAlpha');
    expect(names.get(NAME_ID.fullName)).toBe('ProbeAlpha');
    expect(readOs2(font)?.fsType).toBe(0);
    expect(font.tables.map((t) => t.tag)).toEqual([
      'OS/2',
      'cmap',
      'glyf',
      'head',
      'hhea',
      'hmtx',
      'loca',
      'maxp',
      'name',
      'post',
    ]);
  });

  it('round-trips an SFNT through the EOT writer and reader, byte for byte', () => {
    const font = buildProbeFont('ProbeAlpha');
    for (const version of [undefined, EOT_VERSION_2_2]) {
      const eot = writeEot(font.bytes, {
        familyName: font.familyName,
        styleName: font.styleName,
        versionName: font.versionName,
        fullName: font.fullName,
        panose: font.panose,
        charset: 1,
        italic: false,
        weight: font.weight,
        fsType: font.fsType,
        unicodeRange: font.unicodeRange,
        codePageRange: font.codePageRange,
        checkSumAdjustment: font.checkSumAdjustment,
        version,
      });
      const header = readEot(eot);
      expect(flagNames(header.flags)).toEqual([]);
      expect(header.fsType).toBe(font.fsType);
      expect(eotFontData(eot, header)).toEqual(font.bytes);
    }
  });

  it('refuses a FullName that does not begin with the FamilyName', () => {
    const font = buildProbeFont('ProbeAlpha');
    expect(() =>
      writeEot(font.bytes, {
        familyName: 'ProbeAlpha',
        styleName: 'Regular',
        versionName: 'Version 1.000',
        fullName: 'Something Else',
        panose: font.panose,
        charset: 1,
        italic: false,
        weight: 400,
        fsType: 0,
        unicodeRange: font.unicodeRange,
        codePageRange: font.codePageRange,
        checkSumAdjustment: font.checkSumAdjustment,
      }),
    ).toThrow(/must begin with FamilyName/);
  });

  it('the committed EOTs are the two shapes PowerPoint rejected', () => {
    const v1 = readEot(read('fonts', 'probe-v1.eot'));
    expect(v1.version).toBe(0x00010000);
    expect(v1.flags & EOT_FLAG.TTCOMPRESSED).toBe(0);
    expect(v1.familyName).toBe('ProbeAlpha');

    // Byte-for-byte PowerPoint's own header shape, uncompressed payload aside.
    const v22 = readEot(read('fonts', 'probe-v2_2.eot'));
    expect(v22.version).toBe(EOT_VERSION_2_2);
    expect(v22.flags).toBe(0);
    expect(v22.familyName).toBe('ProbeCharlie ');
    expect(v22.charset).toBe(0);
  });
});
