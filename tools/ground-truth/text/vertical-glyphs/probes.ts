/**
 * Experiment T10 - the probe table for how a glyph is turned in vertical text.
 *
 * T6 answered where the block of a turned frame sits. This answers what happens
 * to the glyphs inside it: which of them stand upright, which face PowerPoint
 * asks GDI for, and where in its cell an upright one is drawn. See ADR 0039.
 */

/** `ST_TextVerticalType`, the five this experiment turns text with. */
export type VertType = 'horz' | 'vert' | 'vert270' | 'eaVert' | 'mongolianVert';

export const VERT_TYPES: readonly VertType[] = [
  'horz',
  'vert',
  'vert270',
  'eaVert',
  'mongolianVert',
];

/**
 * The faces PowerPoint resolved to an `@`-prefixed variant of themselves.
 *
 * Meiryo, Yu Mincho and BIZ UDGothic were dropped after the drawing stream came
 * back naming a substitute: a probe that measures a face nobody asked for is not
 * a reading of the face in the file.
 */
export const FACES: readonly string[] = [
  'Yu Gothic',
  'MS Gothic',
  'SimSun',
  'Microsoft YaHei',
  'Malgun Gothic',
];

/** The two faces the geometry is measured at a second size, to prove it scales. */
export const SCALED_FACES: readonly string[] = ['Yu Gothic', 'MS Gothic'];

/**
 * U+4E00, one horizontal stroke.
 *
 * The whole orientation question in one glyph: upright its ink is wide and flat,
 * turned it is tall and narrow, and nothing in between is possible.
 */
export const BAR = '一';

/** U+30C8, asymmetric on both axes, so it separates upright from upside down. */
export const KATAKANA = 'ト';

/** One glyph whose cell is not a full em, to see what the cross axis follows. */
export const LATIN = 'L';

export interface Probe {
  readonly id: string;
  readonly family: string;
  readonly vars: Readonly<Record<string, string | number>>;
  readonly vert: VertType;
  readonly face: string;
  /** Points. */
  readonly sizePt: number;
  readonly text: string;
  /** The slide is exported as a bitmap, because the glyph's ink is the answer. */
  readonly raster: boolean;
}

/** The frame every probe is drawn in: square, zero insets, top left. */
export const FRAME_PT = { widthPt: 220, heightPt: 250 };

/** Points per slide, matching `sheet-pptx`. */
export const SLIDE_PT = { widthPt: 960, heightPt: 540 };

/** Probes that need the bitmap sit six to a slide; the rest sit eight. */
export const RASTER_PER_SLIDE = 6;
export const STREAM_PER_SLIDE = 8;

/**
 * The mixed strings, whose substrings are all distinct.
 *
 * A record names no shape, so two probes that drew the same characters could not
 * be told apart; every stretch here appears once in the whole deck.
 */
export const MIXED: readonly { readonly id: string; readonly text: string }[] = [
  { id: 'ea-only', text: '縦書' },
  { id: 'latin-only', text: 'Kite' },
  { id: 'ea-then-latin', text: '日月Sun' },
  { id: 'latin-then-ea', text: 'Moon星空' },
  { id: 'ea-latin-ea', text: '雨Wind雪' },
  { id: 'fullwidth-latin', text: 'ＡＢ' },
  { id: 'halfwidth-kana', text: 'ｱｲ' },
  { id: 'ea-punctuation', text: '、。' },
];

function probe(spec: Omit<Probe, 'id'> & { readonly id: string }): Probe {
  return spec;
}

/**
 * Every probe, raster ones first so the bitmaps land on the earliest slides.
 *
 * The families:
 *
 * - `cell` asks where an upright glyph sits, at two sizes on two faces so the
 *   answer has to be a fraction of the em rather than a constant.
 * - `glyph` repeats it on an asymmetric character, which is what separates
 *   upright from upside down.
 * - `advance` puts two glyphs in a row: the cell pitch is their spacing.
 * - `latin` asks whether a Latin run turns with the line even where the East
 *   Asian run beside it does not.
 * - `pitch` puts two paragraphs in the frame, so the line pitch is read from the
 *   two pens rather than from a bounding box.
 * - `face` is the drawing stream alone: which face each stretch was drawn in.
 */
export function allProbes(): readonly Probe[] {
  const raster: Probe[] = [];

  for (const face of FACES) {
    for (const vert of VERT_TYPES) {
      raster.push(
        probe({
          id: `cell-${face}-96-${vert}`,
          family: 'cell',
          vars: { face, size: 96, vert },
          vert,
          face,
          sizePt: 96,
          text: BAR,
          raster: true,
        }),
      );
    }
  }
  for (const face of SCALED_FACES) {
    for (const vert of VERT_TYPES) {
      raster.push(
        probe({
          id: `cell-${face}-48-${vert}`,
          family: 'cell',
          vars: { face, size: 48, vert },
          vert,
          face,
          sizePt: 48,
          text: BAR,
          raster: true,
        }),
      );
    }
  }
  for (const vert of VERT_TYPES) {
    raster.push(
      probe({
        id: `glyph-${vert}`,
        family: 'glyph',
        vars: { face: 'Yu Gothic', size: 96, vert },
        vert,
        face: 'Yu Gothic',
        sizePt: 96,
        text: KATAKANA,
        raster: true,
      }),
      probe({
        id: `advance-${vert}`,
        family: 'advance',
        vars: { face: 'Yu Gothic', size: 96, vert },
        vert,
        face: 'Yu Gothic',
        sizePt: 96,
        text: BAR + BAR,
        raster: true,
      }),
      probe({
        id: `latin-${vert}`,
        family: 'latin',
        vars: { face: 'Yu Gothic', size: 96, vert },
        vert,
        face: 'Yu Gothic',
        sizePt: 96,
        text: LATIN,
        raster: true,
      }),
    );
  }

  const stream: Probe[] = [];
  for (const face of FACES) {
    stream.push(
      probe({
        id: `pitch-${face}`,
        family: 'pitch',
        vars: { face, size: 96, vert: 'eaVert' },
        vert: 'eaVert',
        face,
        sizePt: 96,
        // Two paragraphs, written as two lines by `build-deck`.
        text: BAR,
        raster: false,
      }),
    );
  }
  for (const vert of ['horz', 'vert', 'eaVert', 'mongolianVert'] as const) {
    for (const mixed of MIXED) {
      stream.push(
        probe({
          id: `face-${vert}-${mixed.id}`,
          family: 'face',
          vars: { face: 'Yu Gothic', size: 24, vert, mixed: mixed.id },
          vert,
          face: 'Yu Gothic',
          sizePt: 24,
          text: mixed.text,
          raster: false,
        }),
      );
    }
  }

  return [...raster, ...stream];
}
