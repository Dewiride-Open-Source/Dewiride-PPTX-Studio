/**
 * Experiment T14 - the font-table reader against Chromium, on faces nobody built.
 *
 * T13 asks with fourteen fonts built to disagree with themselves, which says
 * nothing about class-based GPOS kerning, GSUB ligatures or an `hmtx` table
 * with thousands of entries. These samples ask with whatever faces a machine
 * carries. ADR 0042, open questions 2 and 3.
 */

export interface Sample {
  readonly id: string;
  readonly script: string;
  /**
   * Whether a disagreement on this sample fails the job.
   *
   * A sample that is not gated is recorded instead, and has to measure outside
   * `widthTolerance` on some face: what it prices is a feature of the browser
   * the reader does not implement.
   */
  readonly gated: boolean;
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
    asks: 'the per-glyph error, with nothing accumulated',
    text: 'HEAD BADGE',
  },
  {
    id: 'latin-kern-pairs',
    script: 'latin',
    gated: true,
    asks: 'class-based GPOS kerning, which no built font could exercise',
    text: KERN_PAIRS,
  },
  {
    id: 'latin-long',
    script: 'latin',
    gated: true,
    asks: 'whether truncating each advance drifts over hundreds of glyphs',
    text: LONG,
  },
  {
    id: 'latin-spread',
    script: 'latin',
    gated: true,
    asks: 'a cmap addressed across its range, not in one segment',
    text: 'AB Æ × ÷ ¿ ¡ « » – — ‘ ’ “ ” • … € ™ ° ±',
  },
  {
    id: 'cyrillic',
    script: 'cyrillic',
    gated: true,
    asks: 'a second alphabet, and a second cmap segment block',
    text: 'Съешь же ещё этих мягких французских булок да выпей чаю',
  },
  {
    id: 'greek',
    script: 'greek',
    gated: true,
    asks: 'a third alphabet',
    text: 'Ξεσκεπάζω την ψυχοφθόρα βδελυγμία',
  },
  {
    id: 'hebrew',
    script: 'hebrew',
    gated: true,
    asks: 'the control: right-to-left without joining, which separates direction from shaping',
    text: 'דג סקרן שט בים מאוכזב ולפתע מצא חברה',
  },
  {
    id: 'cjk',
    script: 'cjk',
    gated: true,
    asks: 'advances that do not depend on context, and a format 12 cmap',
    text: '永東京都漢字測試日本語中文',
  },
  {
    id: 'latin-ligatures',
    script: 'latin',
    gated: false,
    asks: 'what GSUB liga costs a reader that does not read GSUB',
    text: 'office difficult waffle fluffy',
  },
  {
    id: 'arabic',
    script: 'arabic',
    gated: false,
    asks: 'how wide an unshaped sum measures where a shaper would join',
    text: 'العربية لغة جميلة وواسعة',
  },
  {
    id: 'devanagari',
    script: 'devanagari',
    gated: false,
    asks: 'a second shaped script, so the answer is not one script per anecdote',
    text: 'हिन्दी एक सुंदर भाषा है',
  },
];

/** The six sizes and the box size of T13, so the two experiments sit in one table. */
export const SIZES: readonly number[] = [8, 12, 16, 32, 100, 1000];

/** The size the face box is read at, matching `FACE_BOX_PX` in the product. */
export const BOX_PX = 1000;

/**
 * The pixel Chromium quantises each glyph advance to on a Linux runner.
 *
 * FreeMono - one advance for every glyph - is out by exactly `glyphs` times
 * `round(0.6px) - 0.6px` at every size, which is per-glyph rounding rather than
 * one rounding of a total. `advanceQuantum` refuses a run that says otherwise.
 */
export const ADVANCE_QUANTUM = 1;

/** The reader's own step: T13's per-glyph advance truncation and kern rounding. */
export const READER_STEP = 1 / 65536 + 1 / 131072;

/**
 * How far a whole-pixel browser and the reader may sit apart.
 *
 * The browser quantises the advance and the kern adjustment separately, so a
 * run carries one rounding per glyph and one more per applied kern; a single
 * joint rounding is refuted, because it could not cost the 0.648 px a glyph
 * T14 measured on Lato Thin. Inclusive at the half quantum. ADR 0045.
 */
export function widthTolerance(glyphs: number, kerns = 0): number {
  return (glyphs + kerns) * (ADVANCE_QUANTUM / 2) + glyphs * READER_STEP;
}

/** The face box is one measurement, so it carries one glyph's quantisation. */
export const BOX_TOLERANCE = ADVANCE_QUANTUM / 2;

/**
 * The glyphs a sample is drawn with, which is its code points.
 *
 * True of every gated sample by construction, since none carries an `f` or a
 * joining script; a recorded sample is drawn with fewer, which makes its
 * tolerance generous and is why a recorded sample is never gated.
 */
export function glyphsOf(text: string): number {
  return [...text].length;
}

/** Equality, for the column that shows what the quantisation hides: T13 scored at 1e-12 px. */
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

/** The one rival a whole-pixel browser can still separate: it is not a rounding. */
export const UNKERNED: Reading = 'unkerned';

/** One ascent and descent in font units, the descent positive below the baseline. */
export interface BoxPair {
  readonly ascent: number;
  readonly descent: number;
}

/** `FaceMetrics.candidates`, as the artifact writes it: an absent table is null. */
export interface BoxCandidates {
  readonly hhea: BoxPair;
  readonly usWin: BoxPair | null;
  readonly sTypo: BoxPair | null;
  readonly useTypoMetrics: boolean;
}

const DIRECTWRITE_BOX = 'usWin, or sTypo when fsSelection bit 7 is set';
const FREETYPE_BOX = 'hhea, or sTypo when fsSelection bit 7 is set';

/**
 * Which table the browser read the face box out of, as rival readings.
 *
 * A reading that needs a table this face has no bytes for answers `undefined`
 * and is scored over the faces that carry it. T13 settled DirectWrite at 28/28
 * and T14 settled FreeType at 79/79, on different tables. ADR 0045.
 */
export const BOX_READINGS: Readonly<
  Record<string, (candidates: BoxCandidates) => BoxPair | undefined>
> = {
  'hhea.ascender/descender': (c) => c.hhea,
  'OS/2.usWinAscent/usWinDescent': (c) => c.usWin ?? undefined,
  'OS/2.sTypoAscender/sTypoDescender': (c) => c.sTypo ?? undefined,
  [DIRECTWRITE_BOX]: (c) => {
    // What `metricsOf` answers, including its fall back to hhea with no OS/2.
    if (c.usWin === null || c.sTypo === null) return c.hhea;
    return c.useTypoMetrics ? c.sTypo : c.usWin;
  },
  [FREETYPE_BOX]: (c) => (c.useTypoMetrics && c.sTypo !== null ? c.sTypo : c.hhea),
};

/**
 * The reading `metricsOf` implements here, which no rival may fit more faces than.
 *
 * DirectWrite and FreeType read different tables, so the shipped reading is a
 * property of the runner; macOS takes DirectWrite's by assumption, and a run
 * there scores that assumption. ADR 0045, ADR 0052.
 */
export function shippedBoxFor(platform: string): string {
  return platform === 'linux' ? FREETYPE_BOX : DIRECTWRITE_BOX;
}

/** The rasteriser a runner reads fonts through, as the summary should name it. */
export function rasteriserOf(platform: string): string {
  if (platform === 'linux') return 'FreeType';
  if (platform === 'win32') return 'DirectWrite';
  return `${platform === 'darwin' ? 'CoreText' : platform}, scored against the DirectWrite reading it is assumed to share (ADR 0045)`;
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
