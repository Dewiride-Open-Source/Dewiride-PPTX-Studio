/**
 * Experiment T13 - what a font-table reader must do to agree with the browser.
 *
 * `cli render` measures text in Node, where there is no canvas, so it reads the
 * font's own tables. Every one of those reads has more than one plausible
 * source, and a real face cannot say which is right because its tables agree
 * with each other. These probes disagree on purpose. ADR 0042.
 */

import { buildFont, type FontSpec } from '../../lib/truetype.ts';

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

/** The font has A..H, so the kerned pair is AB and AC is the unkerned control. */
const PAIR = [{ left: 'A', right: 'B', adjust: -200 }] as const;

/**
 * U+20000, the first character of CJK Extension B.
 *
 * Above the BMP on purpose: a format 4 cmap cannot address it at all, so a
 * reader that takes the first subtable it understands cannot find it.
 */
export const ASTRAL = 0x20000;

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
