/**
 * Experiment T11 - the probe table for the two WordArt columns.
 *
 * T10 answered how a glyph is turned. This answers the two numbers the WordArt
 * directions need and no other direction has: the width of the column each
 * character is stacked in, and the pitch it is stacked at. See ADR 0040.
 */

/** The two `ST_TextVerticalType` values that stack one character per cell. */
export type WordArtVert = 'wordArtVert' | 'wordArtVertRtl';

export const WORD_ART_TYPES: readonly WordArtVert[] = ['wordArtVert', 'wordArtVertRtl'];

/** `horz` is carried as the control the cell is measured against. */
export type VertType = WordArtVert | 'horz';

/**
 * The faces the column is fitted over.
 *
 * Spread in the font box is what makes the fit falsifiable, so these run from
 * MS Gothic, whose box is exactly the em, to Segoe UI, whose box is a third
 * larger again.
 */
export const FACES: readonly string[] = [
  'Arial',
  'Courier New',
  'Times New Roman',
  'Verdana',
  'Segoe UI',
  'MS Gothic',
  'Yu Gothic',
  'SimSun',
  'Malgun Gothic',
];

/** The faces that draw both scripts, which is what separates face from script. */
export const EA_FACES: readonly string[] = ['MS Gothic', 'Yu Gothic', 'SimSun', 'Malgun Gothic'];

/** Three sizes, so a reading has to be a fraction of the em rather than a constant. */
export const SIZES_PT: readonly number[] = [12, 18, 28];

/** Four Latin characters of visibly unequal advance, so a per-glyph cell would show. */
export const LATIN = 'Wxyz';

/** Five full-width characters, whose advance is one em on every face here. */
export const CJK = '日本語の文';

/** One string in both scripts, so the pitch is read per glyph rather than per run. */
export const MIXED = '月Wi空';

/** U+4E00, one horizontal stroke: upright its ink is wide and flat, turned it is not. */
export const BAR = '一';

/** A Latin glyph with an unmistakable foot, to see which way a Latin cell points. */
export const UPRIGHT_LATIN = 'L';

/** Long enough that the column has to break before the frame's bottom edge. */
export const OVERLONG = 'Wxyzabcdefghijklmnop';

export interface Probe {
  readonly id: string;
  readonly family: string;
  readonly vars: Readonly<Record<string, string | number>>;
  readonly vert: VertType;
  readonly face: string;
  /** Points. */
  readonly sizePt: number;
  readonly text: string;
  /** Written as that many paragraphs of the same text, to space two columns. */
  readonly paragraphs: number;
  /** `a:bodyPr/@wrap`; only the break family lets a column end. */
  /** `a:bodyPr/@wrap`; only the break family lets a column end. */
  readonly wrap: 'none' | 'square';
  /** A second run in the same paragraph, to see which run's cell the column takes. */
  readonly secondRun?: { readonly text: string; readonly sizePt: number } | undefined;
  /** The slide is exported as a bitmap, because the glyph's ink is the answer. */
  readonly raster: boolean;
}

/** The frame every probe is drawn in: zero insets, anchored top left. */
export const FRAME_PT = { widthPt: 220, heightPt: 240 };

/** Points per slide, matching `sheet-pptx`. */
export const SLIDE_PT = { widthPt: 960, heightPt: 540 };

/** Probes that need the bitmap sit six to a slide; the rest sit eight. */
export const RASTER_PER_SLIDE = 6;
export const STREAM_PER_SLIDE = 8;

function probe(spec: Probe): Probe {
  return spec;
}

function stream(spec: Omit<Probe, 'paragraphs' | 'wrap' | 'raster'>): Probe {
  return probe({ ...spec, paragraphs: 1, wrap: 'none', raster: false });
}

/** Which way the ink of a stacked glyph points, which only the bitmap can say. */
function uprightProbes(): readonly Probe[] {
  const out: Probe[] = [];
  for (const vert of ['horz', ...WORD_ART_TYPES] as const) {
    for (const [face, text] of [
      ['Yu Gothic', BAR],
      ['Arial', UPRIGHT_LATIN],
    ] as const) {
      out.push(
        probe({
          id: `upright-${vert}-${face.replace(/\s+/g, '')}`,
          family: 'upright',
          vars: { vert, face, size: 96 },
          vert,
          face,
          sizePt: 96,
          text,
          paragraphs: 1,
          wrap: 'none',
          raster: true,
        }),
      );
    }
  }
  return out;
}

/**
 * A row of horizontal runs across the slide.
 *
 * Each puts its pen on its own left edge, so the set is one line through the
 * origin and its residual says whether the stream's scale is fitted or guessed.
 */
function scaleProbes(): readonly Probe[] {
  return Array.from({ length: STREAM_PER_SLIDE }, (_, at) =>
    stream({
      id: `scale-${String(at)}`,
      family: 'scale',
      vars: { face: 'Arial', size: 18, script: 'latin', vert: 'horz' },
      vert: 'horz',
      face: 'Arial',
      sizePt: 18,
      text: LATIN,
    }),
  );
}

/** The fit: every face at every size, so the cell has to be a fraction of the em. */
function cellProbes(): readonly Probe[] {
  const out: Probe[] = [];
  for (const face of FACES) {
    for (const sizePt of SIZES_PT) {
      out.push(
        stream({
          id: `cell-${face.replace(/\s+/g, '')}-${String(sizePt)}`,
          family: 'cell',
          vars: { face, size: sizePt, script: 'latin', vert: 'wordArtVert' },
          vert: 'wordArtVert',
          face,
          sizePt,
          text: LATIN,
        }),
      );
    }
  }
  return out;
}

/** The same fit in CJK, which is what tells a property of the face from one of the script. */
function scriptProbes(): readonly Probe[] {
  const out: Probe[] = [];
  for (const face of EA_FACES) {
    for (const sizePt of SIZES_PT) {
      out.push(
        stream({
          id: `script-${face.replace(/\s+/g, '')}-${String(sizePt)}`,
          family: 'script',
          vars: { face, size: sizePt, script: 'cjk', vert: 'wordArtVert' },
          vert: 'wordArtVert',
          face,
          sizePt,
          text: CJK,
        }),
      );
    }
  }
  for (const face of EA_FACES) {
    out.push(
      stream({
        id: `mixed-${face.replace(/\s+/g, '')}`,
        family: 'mixed',
        vars: { face, size: 28, script: 'mixed', vert: 'wordArtVert' },
        vert: 'wordArtVert',
        face,
        sizePt: 28,
        text: MIXED,
      }),
    );
  }
  return out;
}

/**
 * Two paragraphs, an overflowing one, and a repeat of the Latin cell.
 *
 * The pair gives the column pitch and the edge the columns stack from, which is
 * the only thing the two WordArt types disagree about.
 */
function columnProbes(): readonly Probe[] {
  const out: Probe[] = [];
  for (const vert of WORD_ART_TYPES) {
    for (const face of ['Arial', 'MS Gothic']) {
      out.push(
        probe({
          id: `columns-${vert}-${face.replace(/\s+/g, '')}`,
          family: 'columns',
          vars: { face, size: 18, vert },
          vert,
          face,
          sizePt: 18,
          text: face === 'Arial' ? LATIN : CJK,
          paragraphs: 2,
          wrap: 'none',
          raster: false,
        }),
      );
    }
    out.push(
      probe({
        id: `break-${vert}`,
        family: 'break',
        vars: { face: 'Arial', size: 18, vert },
        vert,
        face: 'Arial',
        sizePt: 18,
        text: OVERLONG,
        paragraphs: 1,
        wrap: 'square',
        raster: false,
      }),
      probe({
        id: `rtl-${vert}`,
        family: 'cell',
        vars: { face: 'Arial', size: 18, script: 'latin', vert },
        vert,
        face: 'Arial',
        sizePt: 18,
        text: LATIN,
        paragraphs: 1,
        wrap: 'none',
        raster: false,
      }),
    );
  }
  return out;
}

/**
 * One paragraph of two sizes, which is the only thing that separates a cell
 * belonging to the run from a cell belonging to the column.
 */
function sizeProbes(): readonly Probe[] {
  return ['Arial', 'MS Gothic'].map((face) =>
    probe({
      id: `sizes-${face.replace(/\s+/g, '')}`,
      family: 'sizes',
      vars: { face, size: 12, script: 'latin', vert: 'wordArtVert' },
      vert: 'wordArtVert',
      face,
      sizePt: 12,
      text: 'Wx',
      secondRun: { text: 'yz', sizePt: 28 },
      paragraphs: 1,
      wrap: 'none',
      raster: false,
    }),
  );
}

/** Every probe, raster ones first so the bitmaps land on the earliest slides. */
export function allProbes(): readonly Probe[] {
  return [
    ...uprightProbes(),
    ...scaleProbes(),
    ...cellProbes(),
    ...scriptProbes(),
    ...columnProbes(),
    ...sizeProbes(),
  ];
}
