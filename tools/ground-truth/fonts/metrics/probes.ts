/**
 * Experiment T13 - what a font-table reader must do to agree with the browser.
 *
 * `cli render` measures text in Node, where there is no canvas, so it reads the
 * font's own tables. Every one of those reads has more than one plausible
 * source, and a real face cannot say which is right because its tables agree
 * with each other. These probes disagree on purpose. ADR 0042.
 */

import { buildFont, type Baselines, type FontSpec } from '../../lib/truetype.ts';

export interface Probe {
  readonly id: string;
  /** The one question this font exists to answer. */
  readonly asks: string;
  readonly spec: Partial<FontSpec>;
}

/**
 * Three vertical metrics, three different values.
 *
 * `hhea` 800/-200, `usWin` 900/300 and `sTypo` 700/-100 are pairwise distinct,
 * so `fontBoundingBoxAscent` at a known size names its source outright.
 */
export const METRIC_SPLIT: Partial<FontSpec> = {
  familyName: 'PptxStudio Split',
  hheaAscender: 800,
  hheaDescender: -200,
  winAscent: 900,
  winDescent: 300,
  typoAscender: 700,
  typoDescender: -100,
};

/**
 * The split on a 2048 em, so no metric lands on a whole pixel at `BOX_PX`.
 *
 * usWin is 928.7109375 over 220.21484375, one fraction each side of the half,
 * so floor, ceil and round half up predict different pairs; no `BASE`, so the
 * ideographic fallback is read off that fractional descent too.
 */
export const METRIC_SPLIT_2048: Partial<FontSpec> = {
  familyName: 'PptxStudio Split 2048',
  unitsPerEm: 2048,
  hheaAscender: 1638,
  hheaDescender: -410,
  winAscent: 1902,
  winDescent: 451,
  typoAscender: 1434,
  typoDescender: -205,
};

/**
 * The split on a 2000 em, where usWin's ascent is exactly 928.5 px at `BOX_PX`.
 *
 * An even integer part, so half up says 929 and half to even says 928; the
 * usWin descent and both hhea metrics are whole pixels.
 */
export const METRIC_SPLIT_2000: Partial<FontSpec> = {
  familyName: 'PptxStudio Split 2000',
  unitsPerEm: 2000,
  hheaAscender: 1602,
  hheaDescender: -398,
  winAscent: 1857,
  winDescent: 400,
  typoAscender: 1401,
  typoDescender: -201,
};

/**
 * The split on a 2560 em, where every usWin and hhea side sits on the half at `BOX_PX`.
 *
 * As 32-bit floats 2336/2560 and 544/2560 round up while 672/2560 and 2464/2560
 * round down, so a browser holding the em ratio in single precision reports
 * 913/262 through usWin and 962/213 through hhea; half up says 913/263 and 963/213.
 */
export const METRIC_SPLIT_2560: Partial<FontSpec> = {
  familyName: 'PptxStudio Split 2560',
  unitsPerEm: 2560,
  hheaAscender: 2464,
  hheaDescender: -544,
  winAscent: 2336,
  winDescent: 672,
  typoAscender: 1792,
  typoDescender: -256,
};

/** The font has A..H, so the kerned pair is AB and AC is the unkerned control. */
const PAIR = [{ left: 'A', right: 'B', adjust: -200 }] as const;

/**
 * U+20000, the first character of CJK Extension B.
 *
 * Above the BMP on purpose: a format 4 cmap cannot address it at all, so a
 * reader that takes the first subtable it understands cannot find it.
 */
export const ASTRAL = 0x20000;

/**
 * U+4E00, the one character `packages/text`'s own baseline probe measures.
 *
 * The two engines have to be asked the same question in the same script, so a
 * font that answers about the ideographic baseline must map it. ADR 0039.
 */
export const IDEOGRAPH = 0x4e00;

/**
 * `BASE` coordinates distinct from every vertical metric and from each other.
 *
 * 450, 410 and 380 match no ascent (900, 800, 700), no descent (300, 200, 100),
 * no `icfb` or `hang` coordinate, four fifths of no ascent, and not zero. Three
 * scripts rather than two because the answer names which one was looked up.
 */
const DEFAULT_BASELINES: Baselines = { hang: 620, icfb: -530, ideo: -450, romn: 0 };
const LATIN_BASELINES: Baselines = { hang: 630, icfb: -535, ideo: -410, romn: 0 };
const HAN_BASELINES: Baselines = { hang: 660, icfb: -560, ideo: -380, romn: 0 };
const NO_IDEO_BASELINES: Baselines = { hang: 610, icfb: -520, romn: 0 };

/** A coordinate a fifth of an em below the em, which no cell could hold. */
const OUTSIZE_BASELINES: Baselines = { hang: 640, icfb: -570, ideo: -1200, romn: 0 };

export const PROBES: readonly Probe[] = [
  {
    id: 'plain',
    asks: 'is the advance linear in size, or does the browser hint it',
    spec: { familyName: 'PptxStudio Plain' },
  },
  {
    id: 'split',
    asks: 'which table the font bounding box comes from',
    spec: METRIC_SPLIT,
  },
  {
    id: 'split-usetypo',
    asks: 'whether fsSelection bit 7 moves the box to the typo metrics',
    spec: { ...METRIC_SPLIT, familyName: 'PptxStudio Split UseTypo', useTypoMetrics: true },
  },
  {
    id: 'kern-only',
    asks: 'whether a legacy kern table is honoured at all',
    spec: { familyName: 'PptxStudio Kern', kern: PAIR },
  },
  {
    id: 'gpos-only',
    asks: 'whether GPOS PairPos is honoured at all',
    spec: { familyName: 'PptxStudio Gpos', gpos: PAIR },
  },
  {
    id: 'gpos-extension',
    asks: 'whether a PairPos behind a type 9 Extension lookup still kerns',
    spec: { familyName: 'PptxStudio Gpos Extension', gpos: PAIR, gposExtension: true },
  },
  {
    id: 'astral',
    asks: 'which cmap subtable a reader must prefer when a font has two',
    spec: { familyName: 'PptxStudio Astral', astral: ASTRAL },
  },
  {
    id: 'kern-and-gpos',
    asks: 'which one wins when both are present and they disagree',
    spec: {
      familyName: 'PptxStudio Both',
      kern: PAIR,
      gpos: [{ left: 'A', right: 'B', adjust: -100 }],
    },
  },
  {
    id: 'base-table',
    asks: 'whether the ideographic baseline is a BASE coordinate, and of which script',
    spec: {
      ...METRIC_SPLIT,
      familyName: 'PptxStudio Base',
      base: {
        DFLT: DEFAULT_BASELINES,
        latn: LATIN_BASELINES,
        hani: HAN_BASELINES,
        kana: HAN_BASELINES,
      },
      ideograph: IDEOGRAPH,
    },
  },
  {
    id: 'base-no-dflt',
    asks: 'whether the lookup falls through to another script when DFLT is absent',
    spec: {
      ...METRIC_SPLIT,
      familyName: 'PptxStudio Base NoDflt',
      base: { latn: LATIN_BASELINES, hani: HAN_BASELINES, kana: HAN_BASELINES },
      ideograph: IDEOGRAPH,
    },
  },
  {
    id: 'base-outsize',
    asks: 'whether a coordinate outside the em is reported as written or clamped',
    spec: {
      ...METRIC_SPLIT,
      familyName: 'PptxStudio Base Outsize',
      base: { DFLT: OUTSIZE_BASELINES, latn: OUTSIZE_BASELINES },
      ideograph: IDEOGRAPH,
    },
  },
  {
    id: 'base-no-ideo',
    asks: 'what a BASE table carrying no ideo coordinate falls back to',
    spec: {
      ...METRIC_SPLIT,
      familyName: 'PptxStudio Base NoIdeo',
      base: {
        DFLT: NO_IDEO_BASELINES,
        latn: NO_IDEO_BASELINES,
        hani: NO_IDEO_BASELINES,
        kana: NO_IDEO_BASELINES,
      },
      ideograph: IDEOGRAPH,
    },
  },
  {
    id: 'split-2048',
    asks: 'which whole pixel the face box rounds to, above and below the half',
    spec: METRIC_SPLIT_2048,
  },
  {
    id: 'split-2000',
    asks: 'whether a half rounds up or to even',
    spec: METRIC_SPLIT_2000,
  },
  {
    id: 'split-2560',
    asks: 'whether the em ratio is held in single precision before the half rounds',
    spec: METRIC_SPLIT_2560,
  },
];

/** The strings measured in every probe font. */
export const STRINGS: readonly string[] = ['A', 'B', 'AB', 'AC', 'ABCDEFGH', 'A B'];

/**
 * The sizes every string is measured at.
 *
 * Small sizes are where hinting would show: if the advance is linear at 8px it
 * is linear everywhere, and if it is not, this is where it breaks.
 */
export const SIZES: readonly number[] = [8, 12, 16, 32, 100, 1000];

/** The size the face box is read at, matching `FACE_BOX_PX` in the product. */
export const BOX_PX = 1000;

/**
 * The sizes a baseline is scored at.
 *
 * Every BASE coordinate is a whole number of pixels at both. Only the 2048 and
 * 2560 em probes' fallback descents are not, which is what they are there to show.
 */
export const BASELINE_SIZES: readonly number[] = [100, 1000];

export interface ProbeFont {
  readonly id: string;
  readonly asks: string;
  readonly family: string;
  readonly spec: FontSpec;
  readonly base64: string;
}

export function probeFonts(): readonly ProbeFont[] {
  return PROBES.map((probe) => {
    const built = buildFont(probe.spec);
    return {
      id: probe.id,
      asks: probe.asks,
      family: built.spec.familyName,
      spec: built.spec,
      base64: Buffer.from(built.bytes).toString('base64'),
    };
  });
}
