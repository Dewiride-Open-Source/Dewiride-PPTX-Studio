/**
 * Experiment T14 - the font-table reader against Chromium, on faces nobody built.
 *
 * T13 asks with seven fonts built to disagree with themselves, which says
 * nothing about class-based GPOS kerning, GSUB ligatures or an `hmtx` table
 * with thousands of entries. These samples ask with whatever faces a machine
 * carries. ADR 0042, open questions 2 and 3.
 */

export interface Sample {
  readonly id: string;
  readonly script: string;
  /** Whether a disagreement on this sample fails the job. */
  readonly gated: boolean;
  /** Whether the browser runs a shaper here that the reader has no answer for. */
  readonly shaped: boolean;
  /** The one question this sample exists to answer. */
  readonly asks: string;
  readonly text: string;
}

/** Fifteen pairs a Latin text face carries class-based GPOS adjustments for. */
const KERN_PAIRS = 'AV AW AY LT LV LY PA TA To Va Vo Wa We Ya Yo';

/** Long enough that a per-glyph truncation error would accumulate into sight. */
const LONG = Array.from({ length: 30 }, () => KERN_PAIRS).join(' ');

/**
 * The strings measured on every face.
 *
 * No `f` or `F` in any gated sample: `liga` is on by default in `measureText`
 * and `fi`/`ff`/`fl`/`ffi`/`ffl` are the only substitutions a Latin text face
 * applies unasked, so the reader would be scored against a shaper.
 */
export const SAMPLES: readonly Sample[] = [
  {
    id: 'latin-short',
    script: 'latin',
    gated: true,
    shaped: false,
    asks: 'the per-glyph error, with nothing accumulated',
    text: 'HEAD BADGE',
  },
  {
    id: 'latin-kern-pairs',
    script: 'latin',
    gated: true,
    shaped: false,
    asks: 'class-based GPOS kerning, which no built font could exercise',
    text: KERN_PAIRS,
  },
  {
    id: 'latin-long',
    script: 'latin',
    gated: true,
    shaped: false,
    asks: 'whether truncating each advance drifts over hundreds of glyphs',
    text: LONG,
  },
  {
    id: 'latin-spread',
    script: 'latin',
    gated: true,
    shaped: false,
    asks: 'a cmap addressed across its range, not in one segment',
    text: 'AB Æ × ÷ ¿ ¡ « » – — ‘ ’ “ ” • … € ™ ° ±',
  },
  {
    id: 'cyrillic',
    script: 'cyrillic',
    gated: true,
    shaped: false,
    asks: 'a second alphabet, and a second cmap segment block',
    text: 'Съешь же ещё этих мягких французских булок да выпей чаю',
  },
  {
    id: 'greek',
    script: 'greek',
    gated: true,
    shaped: false,
    asks: 'a third alphabet',
    text: 'Ξεσκεπάζω την ψυχοφθόρα βδελυγμία',
  },
  {
    id: 'hebrew',
    script: 'hebrew',
    gated: true,
    shaped: false,
    asks: 'the control: right-to-left without joining, which separates direction from shaping',
    text: 'דג סקרן שט בים מאוכזב ולפתע מצא חברה',
  },
  {
    id: 'cjk',
    script: 'cjk',
    gated: true,
    shaped: false,
    asks: 'advances that do not depend on context, and a format 12 cmap',
    text: '永東京都漢字測試日本語中文',
  },
  {
    id: 'latin-ligatures',
    script: 'latin',
    gated: false,
    shaped: false,
    asks: 'what GSUB liga costs a reader that does not read GSUB',
    text: 'office difficult waffle fluffy',
  },
  {
    id: 'arabic',
    script: 'arabic',
    gated: false,
    shaped: true,
    asks: 'how wide an unshaped sum measures where a shaper would join',
    text: 'العربية لغة جميلة وواسعة',
  },
  {
    id: 'devanagari',
    script: 'devanagari',
    gated: false,
    shaped: true,
    asks: 'a second shaped script, so the answer is not one script per anecdote',
    text: 'हिन्दी एक सुंदर भाषा है',
  },
];

/** The six sizes and the box size of T13, so the two experiments sit in one table. */
export const SIZES: readonly number[] = [8, 12, 16, 32, 100, 1000];

/** The size the face box is read at, matching `FACE_BOX_PX` in the product. */
export const BOX_PX = 1000;

/**
 * The relative disagreement a gated comparison may carry.
 *
 * T13 has the reader reproducing Chromium exactly, so a residual on a real face
 * is a feature the reader misses rather than noise; this is a fifteenth of the
 * 0.149% median by which that browser disagrees with PowerPoint (T2), and
 * 0.004 px on a ten-glyph string at 8 px - far under the half pixel that moves
 * a line break.
 */
export const AGREEMENT = 1e-4;

/** Equality, for the rival scoreboard: T13 scored its readings at 1e-12 px. */
export const EXACT_PX = 1e-9;

/**
 * The five readings of a string's width, each scored against `measureText`.
 *
 * `shipped` is driven through `createFontMeasurer` rather than reimplemented,
 * so what is scored is the code an `npm install` gets.
 */
export const READINGS = [
  'shipped',
  'exact',
  'advanceRounded',
  'unkerned',
  'wholeStringTruncated',
] as const;

export type Reading = (typeof READINGS)[number];

/** The reading the product implements, and the one that has to win. */
export const SHIPPED: Reading = 'shipped';

/**
 * The four readings of the face box.
 *
 * T13's built fonts all have `unitsPerEm` 1000 and integer metrics, so the four
 * tie there; a real face with `unitsPerEm` 2048 puts the box at 928.22265625
 * and separates them.
 */
export const BOX_VARIANTS = ['exact', 'round', 'floor', 'ceil'] as const;

export type BoxVariant = (typeof BOX_VARIANTS)[number];

export function applyBoxVariant(variant: BoxVariant, value: number): number {
  if (variant === 'exact') return value;
  if (variant === 'round') return Math.round(value);
  if (variant === 'floor') return Math.floor(value);
  return Math.ceil(value);
}

/** Every character any sample uses, for picking one the face certainly has. */
export function sampleCharacters(): readonly number[] {
  const seen = new Set<number>();
  for (const sample of SAMPLES) {
    for (const character of sample.text) {
      const code = character.codePointAt(0);
      if (code !== undefined) seen.add(code);
    }
  }
  return [...seen].sort((a, b) => a - b);
}
