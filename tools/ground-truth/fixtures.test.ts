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
      expect(font.familyName.endsWith('\u0000'), `${font.part} family`).toBe(true);
      expect(font.styleName.endsWith('\u0000'), `${font.part} style`).toBe(true);
      expect(font.versionName.endsWith('\u0000'), `${font.part} version`).toBe(false);
      expect(font.fullName.endsWith('\u0000'), `${font.part} full`).toBe(false);
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
    expect(v22.familyName).toBe('ProbeCharlie\u0000');
    expect(v22.charset).toBe(0);
  });
});

/* -------------------------------------------------------------------------- */
/* F2 - what the export does at another width                                 */
/* -------------------------------------------------------------------------- */

interface ZoomProbe {
  readonly id: string;
  readonly family: string;
  readonly widthPt?: number;
  readonly periodPt?: number;
  readonly read: { readonly kind: string; readonly y0?: number; readonly y1?: number };
}
interface ZoomScore {
  readonly model: string;
  readonly fits: number;
  readonly of: number;
}
interface ZoomFixture {
  readonly widths: readonly number[];
  readonly slide: { readonly w: number };
  readonly tolerancePx: number;
  readonly boxTolerancePx: number;
  readonly findings: {
    readonly markerOnHairline: string;
    readonly patternScaledFrom: number;
    readonly patternCoverageNominalFrom: number | null;
    readonly borderEdge: string;
    readonly frameStretched: boolean;
    readonly textLinear: boolean;
    readonly textDroppedBelowPx: number;
    readonly gradientIdenticalFrom: number | null;
    readonly exportCeiling: unknown;
  };
  readonly candidates: Readonly<Record<string, readonly ZoomScore[]>>;
  readonly probes: readonly ZoomProbe[];
  readonly measured: Readonly<
    Record<
      string,
      Readonly<
        Record<string, { boxH?: number; periodPx?: number | null; innerEdgePx?: number | null }>
      >
    >
  >;
}

describe('experiment F2 - what the export does at another width', () => {
  const zoom = readJson<ZoomFixture>('zoom.json');
  const cases = (family: string) =>
    zoom.probes
      .filter((probe) => probe.family === family)
      .flatMap((probe) =>
        zoom.widths.map((width) => ({
          probe,
          width,
          scale: width / zoom.slide.w,
          seen: zoom.measured[probe.id]![String(width)]!,
        })),
      );

  it('names one perfect reading per family, and every rival misses', () => {
    for (const [family, scores] of Object.entries(zoom.candidates)) {
      const perfect = scores.filter((score) => score.fits === score.of);
      expect(perfect, family).toHaveLength(1);
      expect(scores.length, family).toBeGreaterThan(1);
    }
  });

  it('sizes a triangle head from a 10-pt vector pen up to 2 pt and from whole pixels above, 49 of 49', () => {
    const rule = (widthPt: number, scale: number): number =>
      widthPt <= 2 ? 5 * Math.max(1, 2 * scale) : 5 * Math.max(1, Math.round(widthPt * scale));
    const all = cases('marker');
    expect(all).toHaveLength(49);
    for (const c of all) {
      const predicted = rule(c.probe.widthPt ?? 0, c.scale);
      expect(
        Math.abs((c.seen.boxH ?? 0) - predicted),
        `${c.probe.id}@${String(c.width)}`,
      ).toBeLessThanOrEqual(zoom.boxTolerancePx);
    }
    // Five nominal widths says nothing on a hairline; the export drew a 10-pt head.
    expect(
      cases('marker').filter((c) => c.probe.widthPt === 0 && c.width === 3840)[0]?.seen.boxH,
    ).toBeGreaterThan(30);
    // A rounded 2-pt pen at 1.25 px/pt says 15 px; the export drew 12.
    expect(
      cases('marker').filter((c) => c.probe.widthPt === 0 && c.width === 1200)[0]?.seen.boxH,
    ).toBe(12);
    expect(zoom.findings.markerOnHairline).toBe('M8');
  });

  it('keeps the 6-pt pattern tile from 960 up and loses it below', () => {
    const horz = cases('pattern').filter((c) => c.probe.id === 'pattern-horz');
    for (const c of horz) {
      if (c.width >= 960) {
        expect(Math.abs((c.seen.periodPx ?? 0) - 6 * c.scale), String(c.width)).toBeLessThanOrEqual(
          zoom.tolerancePx,
        );
      }
    }
    expect(horz.find((c) => c.width === 480)?.seen.periodPx).toBeNull();
    expect(horz.find((c) => c.width === 240)?.seen.periodPx).toBe(3);
    expect(zoom.findings.patternScaledFrom).toBe(960);
    expect(zoom.findings.patternCoverageNominalFrom).toBe(1920);
  });

  it('keeps a picture border wholly outside, snapped to the pixel grid half up, 8 of 8', () => {
    const borders = cases('border').filter((c) => (c.probe.widthPt ?? 1) * c.scale >= 1);
    expect(borders).toHaveLength(8);
    for (const c of borders) {
      const edge = (((c.probe.read.y0 ?? 0) + (c.probe.read.y1 ?? 0)) / 2) * c.scale;
      const predicted = Math.round(edge) - edge;
      expect(
        Math.abs((c.seen.innerEdgePx ?? 0) - predicted),
        `${c.probe.id}@${String(c.width)}`,
      ).toBeLessThanOrEqual(zoom.tolerancePx);
    }
    expect(zoom.findings.borderEdge).toBe('O4');
  });

  it('records what does not scale: the frame stretches, text scales, nothing is dropped', () => {
    expect(zoom.findings.frameStretched).toBe(true);
    expect(zoom.findings.textLinear).toBe(true);
    expect(zoom.findings.textDroppedBelowPx).toBe(0);
    expect(zoom.findings.gradientIdenticalFrom).toBe(480);
    expect(zoom.findings.exportCeiling).toBeNull();
  });
});

interface SnapScore {
  readonly model: string;
  readonly fits: number;
  readonly of: number;
}
interface SnapProbe {
  readonly id: string;
  readonly family: string;
  readonly shape: 'band' | 'from' | 'to';
  readonly centrePt: number;
  readonly widthPt?: number;
  readonly edge?: string;
  readonly read: { readonly axis: 'x' | 'y'; readonly at: number; readonly half: number };
  readonly line?: { readonly from: number; readonly to: number; readonly t: number };
}
interface SnapMeasure {
  readonly from: number;
  readonly cover: readonly number[];
}
interface SnapAtWidth {
  readonly height: number;
  readonly dpi: number;
  readonly eighth: boolean;
  readonly wholeHeight: boolean;
  readonly wholeDpi: boolean;
  readonly cases: number;
  readonly rule: Readonly<Record<string, readonly [number, number]>>;
  readonly ruleMisses: readonly string[];
  readonly mappings: Readonly<Record<string, Readonly<Record<string, readonly [number, number]>>>>;
  readonly partialRows: number;
  readonly quarterRows: number;
}
interface SnapFixture {
  readonly slide: { readonly w: number; readonly h: number };
  readonly widths: readonly number[];
  readonly gridWidths: readonly number[];
  readonly tolerance: number;
  readonly cases: number;
  readonly excluded: readonly string[];
  readonly findings: Readonly<Record<string, unknown>> & {
    readonly partialRows: Readonly<Record<string, number>>;
    readonly holdsAt: readonly number[];
    readonly failsAt: readonly number[];
  };
  readonly candidates: Readonly<Record<string, readonly SnapScore[]>>;
  readonly atWidth: Readonly<Record<string, SnapAtWidth>>;
  readonly probes: readonly SnapProbe[];
  readonly measured: Readonly<Record<string, Readonly<Record<string, SnapMeasure>>>>;
}

describe('experiment F3 - where the export puts an edge on the device grid', () => {
  const snap = readJson<SnapFixture>('snap.json');
  const halfUp = (v: number): number => Math.floor(v + 0.5);
  /** The pen in whole pixels, never under one; a half went up at 240, 480, 960 and 1920 and down at 1200. */
  const pen = (w: number, scale: number): number =>
    Math.max(1, w - Math.floor(w) === 0.5 && scale === 1.25 ? Math.floor(w) : halfUp(w));
  /** Whether the pen sits on pixel centres: when the width rounded half up is odd. */
  const odd = (w: number): boolean => Math.max(1, halfUp(w)) % 2 === 1;
  /** An odd pen sits on pixel centres: the rounded coordinate, half a pixel on. */
  const snapped = (v: number, w: number): number => halfUp(v) + (odd(w) ? 0.5 : 0);
  /** Device pixels per point along a read axis: the rounded height's down rows. */
  const scaleAt = (width: number, axis: 'x' | 'y'): number =>
    axis === 'x'
      ? width / snap.slide.w
      : Math.round((width * snap.slide.h) / snap.slide.w) / snap.slide.h;
  const coverOf = (top: number, bottom: number, from: number, to: number): number[] => {
    const out: number[] = [];
    for (let r = from; r <= to; r++) {
      out.push(Math.min(1, Math.max(0, Math.min(bottom, r + 1) - Math.max(top, r))));
    }
    return out;
  };
  interface Case {
    readonly probe: SnapProbe;
    readonly width: number;
    readonly scale: number;
    readonly from: number;
    readonly to: number;
    readonly c: number;
    readonly w: number;
    readonly seen: number[];
  }
  const caseOf = (probe: SnapProbe, width: number): Case => {
    const scale = scaleAt(width, probe.read.axis);
    const from = Math.round((probe.read.at - probe.read.half) * scale);
    const to = Math.round((probe.read.at + probe.read.half) * scale);
    const measure = snap.measured[probe.id]![String(width)]!;
    const seen = new Array<number>(to - from + 1).fill(0);
    measure.cover.forEach((v, i) => {
      seen[measure.from - from + i] = v;
    });
    return {
      probe,
      width,
      scale,
      from,
      to,
      c: probe.centrePt * scale,
      w: (probe.widthPt ?? 0) * scale,
      seen,
    };
  };
  /** A family's cases on the widths the rule is read from. */
  const cases = (family: string, widths: readonly number[] = snap.gridWidths): Case[] =>
    snap.probes
      .filter((probe) => probe.family === family)
      .flatMap((probe) => widths.map((width) => caseOf(probe, width)));
  const worst = (a: readonly number[], b: readonly number[]): number =>
    Math.max(...a.map((v, i) => Math.abs(v - (b[i] ?? 0))));
  const fits = (all: readonly Case[], predict: (c: Case) => [number, number]): number =>
    all.filter((k) => {
      const [top, bottom] = predict(k);
      return worst(k.seen, coverOf(top, bottom, k.from, k.to)) <= snap.tolerance;
    }).length;
  const band = (k: Case): [number, number] => {
    const n = pen(k.w, k.scale);
    const centre = snapped(k.c, k.w);
    return [centre - n / 2, centre + n / 2];
  };
  const edge = (k: Case, at: number): [number, number] =>
    k.probe.shape === 'from' ? [at, Infinity] : [-Infinity, at];
  /** A picture border: the pen centred half the true width outside the rounded frame edge. */
  const border = (k: Case): [number, number] => {
    const n = pen(k.w, k.scale);
    const centre = halfUp(k.c) - k.w / 2;
    return [centre - n / 2, centre + n / 2];
  };
  /** An algn="in" pen: a device pixel outside the rounded frame edge, half for an odd pen, the rest inside. */
  const inset = (k: Case): [number, number] => {
    const top = halfUp(k.c) - 1 + (odd(k.w) ? 0.5 : 0) + Math.max(0, 1 - k.w);
    return [top, top + pen(k.w, k.scale)];
  };
  /** The family's rule, for the families whose rule has no bias of its own. */
  const rule = (k: Case): [number, number] => {
    switch (k.probe.family) {
      case 'fill':
      case 'picture':
        return edge(k, halfUp(k.c));
      case 'border':
        return border(k);
      case 'inset':
        return inset(k);
      default:
        return band(k);
    }
  };

  /** A losing model's score in the fixture, summed over the families named. */
  const scored = (model: string, ...families: string[]): number =>
    families.reduce(
      (sum, family) => sum + (snap.candidates[family]!.find((s) => s.model === model)?.fits ?? 0),
      0,
    );

  it('names one perfect reading per family on the grid widths, and every rival misses', () => {
    expect(snap.cases).toBe(3388);
    expect(snap.excluded).toHaveLength(0);
    expect(snap.gridWidths).toEqual([240, 480, 960, 1200, 1920, 3840]);
    for (const [family, scores] of Object.entries(snap.candidates)) {
      const perfect = scores.filter((score) => score.fits === score.of);
      expect(perfect, family).toHaveLength(1);
      expect(snap.findings[family], family).toBe(perfect[0]!.model);
      expect(scores.length, family).toBeGreaterThan(1);
    }
    expect(snap.findings.holdsAt).toEqual(snap.gridWidths);
  });

  it('snaps every axis-aligned stroke: the centre rounds half up and an odd pen sits on pixel centres, 852 of 852', () => {
    const all = ['stroke', 'outline', 'triangle', 'rotated', 'tie'].flatMap((f) => cases(f));
    expect(all).toHaveLength(852);
    expect(fits(all, band)).toBe(852);
    // The reading anyone writes first - the true width, antialiased where it lies - is what a
    // browser draws, and it misses three cases in four.
    expect(fits(all, (k) => [k.c - k.w / 2, k.c + k.w / 2])).toBeLessThan(all.length / 4);
    // Rounding the pen's top edge without moving its centre misses every odd pen past the half.
    const topRounded = fits(all, (k) => {
      const n = Math.max(1, halfUp(k.w));
      return [halfUp(k.c - n / 2), halfUp(k.c - n / 2) + n];
    });
    expect(topRounded).toBe(scored('SR', 'stroke', 'outline', 'triangle', 'rotated', 'tie'));
    // The rows the export left partial: two for each 2.5-pixel pen at 1200, and nothing else.
    const partial = cases('stroke').reduce(
      (sum, k) => sum + k.seen.filter((v) => v > 0.1 && v < 0.9).length,
      0,
    );
    expect(partial).toBe(16);
    expect(snap.findings.partialRows['stroke']).toBe(partial);
  });

  it('rounds an exact half pixel of width up at 240, 960 and 1920 and down at 1200, 44 of 44', () => {
    const ties = cases('tie').filter((k) => k.w - Math.floor(k.w) === 0.5 && k.w > 1);
    const at = (width: number): Case[] => ties.filter((k) => k.width === width);
    expect(at(240)).toHaveLength(6);
    expect(at(960)).toHaveLength(12);
    expect(at(1200)).toHaveLength(16);
    expect(at(1920)).toHaveLength(10);
    // No integer-EMU width is a half pixel at 3840, and none of the tie widths is at 480.
    expect(at(3840)).toHaveLength(0);
    expect(at(480)).toHaveLength(0);
    expect(fits(ties, band)).toBe(44);
    const always = (round: (w: number) => number): number =>
      fits(ties, (k) => {
        const n = Math.max(1, round(k.w));
        const centre = snapped(k.c, k.w);
        return [centre - n / 2, centre + n / 2];
      });
    // Up everywhere misses the 16 at 1200; down everywhere misses the other 28.
    expect(always(halfUp)).toBe(28);
    expect(always((w) => Math.floor(w))).toBe(16);
    const halfEven = (v: number): number =>
      v - Math.floor(v) === 0.5
        ? Math.floor(v) % 2 === 0
          ? Math.floor(v)
          : Math.floor(v) + 1
        : halfUp(v);
    expect(always(halfEven)).toBe(22);
    // The fixture scores the same three readings over the whole family.
    expect(scored('SG', 'tie')).toBe(cases('tie').length - 16);
    expect(scored('SD2', 'tie')).toBe(cases('tie').length - 28);
    expect(scored('SE2', 'tie')).toBe(cases('tie').length - 22);
  });

  it('rounds a fill edge and a picture edge half up, 192 of 192', () => {
    const fills = cases('fill').concat(cases('picture'));
    expect(fills).toHaveLength(192);
    expect(fits(fills, (k) => edge(k, halfUp(k.c)))).toBe(192);
    expect(fits(fills, (k) => edge(k, k.c))).toBe(scored('E0', 'fill', 'picture'));
    expect(fits(fills, (k) => edge(k, Math.floor(k.c)))).toBe(scored('EF', 'fill', 'picture'));
  });

  it('takes a flat end onto its pen, and a one-pixel pen a quarter pixel past it, 144 of 144', () => {
    const ends = cases('end');
    expect(ends).toHaveLength(144);
    expect(
      fits(ends, (k) => {
        const reach = pen(k.w, k.scale) === 1 ? (k.probe.shape === 'from' ? -0.25 : 0.25) : 0;
        return edge(k, snapped(k.c, k.w) + reach);
      }),
    ).toBe(144);
    expect(fits(ends, (k) => edge(k, halfUp(k.c)))).toBe(scored('ER', 'end'));
    expect(fits(ends, (k) => edge(k, snapped(k.c, k.w)))).toBe(scored('EP', 'end'));
  });

  it('keeps a picture border half its true width outside the rounded frame edge, 48 of 48', () => {
    const borders = cases('border');
    expect(borders).toHaveLength(48);
    expect(fits(borders, border)).toBe(48);
    // F2's O4 read this border at one width; a pen wholly outside the rounded edge misses 16.
    expect(fits(borders, (k) => [halfUp(k.c) - pen(k.w, k.scale), halfUp(k.c)])).toBe(32);
  });

  it('puts an algn="in" pen a device pixel outside the rounded frame edge and the rest inside, 48 of 48', () => {
    const insets = cases('inset');
    expect(insets).toHaveLength(48);
    expect(fits(insets, inset)).toBe(48);
    // A 1-pt pen at 100 % straddles its frame edge, half a row each side, where a centred one is crisp.
    const straddling = insets.filter((k) => k.width === 960 && k.probe.widthPt === 1);
    expect(straddling).toHaveLength(4);
    for (const k of straddling) {
      expect(k.seen.filter((v) => Math.abs(v - 0.5) < snap.tolerance)).toHaveLength(2);
    }
    // The alignment ignored - the stroke rule centred on the frame edge - is right at 200 % only.
    expect(fits(insets, band)).toBe(scored('IX', 'inset'));
    expect(fits(insets, (k) => [halfUp(k.c), halfUp(k.c) + pen(k.w, k.scale)])).toBe(
      scored('IR', 'inset'),
    );
    expect(fits(insets, (k) => [k.c, k.c + k.w])).toBe(scored('I0', 'inset'));
  });

  it('snaps a slanted line at its endpoints and antialiases between them, 24 of 24', () => {
    const slants = cases('slant');
    expect(slants).toHaveLength(24);
    expect(
      fits(slants, (k) => {
        const line = k.probe.line!;
        const from = snapped(line.from * k.scale, k.w);
        const to = snapped(line.to * k.scale, k.w);
        const centre = from + (to - from) * line.t;
        const n = pen(k.w, k.scale);
        return [centre - n / 2, centre + n / 2];
      }),
    ).toBe(24);
    expect(fits(slants, band)).toBe(scored('SP', 'slant'));
  });

  it('draws a curved outline where the stroke rule puts it, a quarter pixel down at the top', () => {
    const bias = (k: Case): number => (k.probe.read.axis === 'y' ? 0.25 : 0.125);
    const ellipses = cases('ellipse');
    expect(ellipses).toHaveLength(72);
    expect(
      fits(ellipses, (k) => {
        const [top, bottom] = band(k);
        return k.probe.edge === 'bottom' ? [top, bottom] : [top + bias(k), bottom + bias(k)];
      }),
    ).toBe(72);
    expect(fits(ellipses, band)).toBe(scored('SP', 'ellipse'));
    const rounded = cases('roundRect');
    expect(rounded).toHaveLength(72);
    expect(
      fits(rounded, (k) => {
        const [top, bottom] = band(k);
        return pen(k.w, k.scale) > 1 ? [top + bias(k), bottom] : [top, bottom];
      }),
    ).toBe(72);
    expect(fits(rounded, band)).toBe(scored('SP', 'roundRect'));
    expect(snap.findings['ellipse']).toBe('SQ');
    expect(snap.findings['roundRect']).toBe('SN');
  });

  it('holds at 1040 and 1120 but for the exact halves, which round the other way', () => {
    for (const width of [1040, 1120]) {
      const here = snap.atWidth[String(width)]!;
      expect(here.wholeDpi && here.wholeHeight && !here.eighth, String(width)).toBe(true);
    }
    // A 6.5-pixel pen at 1040 is drawn six wide on a pixel centre, the 1-pt L step straddles
    // its row, and the turned rectangle's pen is a row off.
    expect(snap.atWidth['1040']!.ruleMisses).toEqual([
      'outline-L-1-0_25-step',
      'rotated-1-0',
      'tie-6-0',
      'tie-6-0_25',
    ]);
    const six = caseOf(
      snap.probes.find((p) => p.id === 'tie-6-0')!,
      1040,
    );
    expect(fits([six], (k) => [halfUp(k.c) - 2.5, halfUp(k.c) + 3.5])).toBe(1);
    // A 3.5-pixel pen at 1120 is drawn four wide on a pixel centre, a frame edge on an exact
    // half rounds down, and the 2-pt border is crisp one pen outside, a sixth of a row from BT.
    expect(snap.atWidth['1120']!.ruleMisses).toEqual([
      ...['0', '0_25', '0_5', '0_75'].map((o) => `stroke-h-3-${o}`),
      ...['0', '0_25', '0_5', '0_75'].map((o) => `stroke-v-3-${o}`),
      'outline-2-0_5-bottom',
      'outline-L-1-0_25-step',
      'outline-L-2-0_75-step',
      ...['0', '0_25', '0_5', '0_75'].map((o) => `border-2-${o}`),
    ]);
    const three = caseOf(
      snap.probes.find((p) => p.id === 'stroke-h-3-0')!,
      1120,
    );
    expect(fits([three], (k) => [halfUp(k.c) - 1.5, halfUp(k.c) + 2.5])).toBe(1);
    const bottom = caseOf(
      snap.probes.find((p) => p.id === 'outline-2-0_5-bottom')!,
      1120,
    );
    expect(bottom.c - Math.floor(bottom.c)).toBeCloseTo(0.5, 9);
    expect(fits([bottom], (k) => [Math.floor(k.c) - 1, Math.floor(k.c) + 1])).toBe(1);
    const wide = caseOf(
      snap.probes.find((p) => p.id === 'border-2-0')!,
      1120,
    );
    expect(fits([wide], (k) => [halfUp(k.c) - 2, halfUp(k.c)])).toBe(1);
  });

  it('keeps the columns and moves the rows by the stretch, in quarter pixels, at a fractional height', () => {
    const plain = ['stroke', 'tie', 'fill', 'border', 'inset'];
    for (const width of [120, 1000, 1320]) {
      const here = snap.atWidth[String(width)]!;
      expect(here.wholeHeight, String(width)).toBe(false);
      expect(here.height).toBe(Math.round((width * snap.slide.h) / snap.slide.w));
      // Every partial row is a quarter, as a 2x2 supersampler leaves them.
      expect(here.quarterRows / here.partialRows, String(width)).toBeGreaterThan(0.9);
      // The columns hold: every vertical stroke and every left and right edge.
      const columns = ['stroke', 'fill', 'picture']
        .flatMap((f) => cases(f, [width]))
        .filter((k) => k.probe.read.axis === 'x');
      expect(fits(columns, rule), String(width)).toBe(columns.length);
      // The rows: the rule at the width's own scale, then moved by the stretch - exactly, or
      // rounded up to the next quarter pixel.
      const uniform = width / snap.slide.w;
      const stretch = here.height / ((width * snap.slide.h) / snap.slide.w) - 1;
      const rows = plain.flatMap((f) => cases(f, [width])).filter((k) => k.probe.read.axis === 'y');
      const moved =
        (move: (v: number) => number) =>
        (k: Case): [number, number] => {
          const [top, bottom] = rule({
            ...k,
            scale: uniform,
            c: k.probe.centrePt * uniform,
            w: (k.probe.widthPt ?? 0) * uniform,
          });
          const carry = (v: number): number => (Number.isFinite(v) ? move(v) : v);
          return [carry(top), carry(bottom)];
        };
      const exact = fits(
        rows,
        moved((v) => v * (1 + stretch)),
      );
      const quarter = fits(
        rows,
        moved((v) => v + Math.ceil(v * stretch * 4) / 4),
      );
      const recorded = (name: string): number =>
        Object.entries(here.mappings[name]!)
          .filter(([family]) => plain.includes(family))
          .reduce((sum, [, [f]]) => sum + f, 0);
      expect(exact, String(width)).toBe(recorded('uniform'));
      expect(quarter, String(width)).toBe(recorded('quarterStep'));
      expect(quarter, String(width)).toBeGreaterThan(exact);
      // At 1320 that is every one of them; at 1000 all but the 12-pt pen, 12.5 pixels there and
      // an exact half at a scale that is not a power of two; at 120 the strokes gain a quarter row.
      if (width === 1320) expect(quarter).toBe(rows.length);
      if (width === 1000) {
        expect(quarter).toBe(rows.length - 2);
        expect(rows.filter((k) => k.probe.widthPt === 12).map((k) => k.probe.id)).toEqual([
          'tie-12-0',
          'tie-12-0_25',
        ]);
      }
      if (width === 120) expect(quarter).toBeLessThan(rows.length - 2);
    }
  });

  it('snaps nothing on either axis at a fractional dots per inch, and no resampling tried fits', () => {
    for (const width of [1008, 1100, 1184]) {
      const here = snap.atWidth[String(width)]!;
      expect(here.wholeDpi, String(width)).toBe(false);
      expect(here.rule['stroke']![0], String(width)).toBeLessThanOrEqual(12);
      expect(here.rule['tie']![0], String(width)).toBe(0);
      expect(here.rule['fill']![0], String(width)).toBeLessThanOrEqual(4);
      // Coverage is not in quarters: the rows were resampled, not supersampled.
      expect(here.quarterRows / here.partialRows, String(width)).toBeLessThan(0.2);
      for (const [name, byFamily] of Object.entries(here.mappings)) {
        const fitsAll = Object.values(byFamily).reduce((sum, [f]) => sum + f, 0);
        const ofAll = Object.values(byFamily).reduce((sum, [, n]) => sum + n, 0);
        expect(fitsAll / ofAll, `${name}@${String(width)}`).toBeLessThan(0.6);
      }
    }
    expect(snap.findings.failsAt).toEqual([120, 1000, 1008, 1040, 1100, 1120, 1184, 1320]);
  });
});
