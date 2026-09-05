/**
 * What a bullet is, where it goes, and what number it carries.
 *
 * Measured against PowerPoint in experiment T5 -
 * `corpus/ground-truth/bullets.json` and
 * `docs/adr/phase-3-text/0031-bullets-fields-and-script-runs.md`. Four of the rules here are
 * ones a careful implementation gets wrong, and each is wrong in a way that
 * looks right on the deck in front of you:
 *
 * - **`a:buFont` does not reach an autonumber.** It re-faces a character bullet
 *   and is inert on a number, which is drawn in the first run's face whatever
 *   the file says. PowerPoint's own UI writes `buFont="+mj-lt"` when it converts
 *   a paragraph to a numbered list, and that attribute changes nothing.
 * - **`a:buClr` and `a:buSzPct` *do* reach one.** So the three decorations do
 *   not travel together, and "buFont is decorative" generalises to the wrong
 *   answer for the other two.
 * - **The numbers are not in the file.** A six-item numbered list PowerPoint
 *   authored itself carries no `startAt` anywhere; every number on every slide
 *   is computed at layout time by the rule in `numberParagraphs`.
 * - **A symbol bullet is mapped through the ANSI codepage**, not the code point.
 *   `char="•"` in Wingdings draws U+F095, a filled circle; adding 0xF000 to
 *   U+2022 gives U+F022, an envelope.
 */

import { autonumberTypeface, formatAutonumber, START_AT_MAX, START_AT_MIN } from './autonumber.js';
import { TextError } from '../errors.js';

/* -------------------------------------------------------------------------- */
/* the bullet a paragraph declares                                            */
/* -------------------------------------------------------------------------- */

/**
 * The exclusive group in `CT_TextParagraphProperties`: one of four, or nothing.
 *
 * `none` is `a:buNone` and `inherit` is the absence of all four, and they are
 * different: the first suppresses a bullet the level above declared, the second
 * takes it.
 */
export type BulletKind = 'none' | 'char' | 'autonum' | 'blip';

/** `a:buSzPct` / `a:buSzPts` / `a:buSzTx`, which are exclusive of each other. */
export type BulletSize =
  | { readonly kind: 'percent'; readonly value: number }
  | { readonly kind: 'points'; readonly value: number }
  | { readonly kind: 'text' };

/** `a:buClr` / `a:buClrTx`. The colour itself is the model's, not this package's. */
export type BulletColor<TColor> =
  { readonly kind: 'color'; readonly value: TColor } | { readonly kind: 'text' };

/** `a:buFont` / `a:buFontTx`. */
export type BulletFont =
  { readonly kind: 'typeface'; readonly value: string } | { readonly kind: 'text' };

/**
 * A paragraph's bullet, after the cascade has merged it.
 *
 * Every field is separately optional because T5 measured the four groups to
 * merge **per property**, exactly as 3.1 measured the rest of `a:pPr` to: a
 * level declaring only `a:buFont` re-faces the character it inherits and keeps
 * everything else, and a level declaring only `a:buSzTx` cancels an inherited
 * `a:buSzPct` and keeps the character, the font and the colour. Measured on
 * eleven cases twice over - once through a shape's own `a:lstStyle` and once
 * through three hops of the master chain.
 */
export interface ResolvedBullet<TColor> {
  readonly kind: BulletKind;
  /** `a:buChar/@char`, as written. The private-use mapping happens at draw time. */
  readonly char?: string | undefined;
  /** `a:buAutoNum/@type`. */
  readonly scheme?: string | undefined;
  /** `a:buAutoNum/@startAt`. Absent is measured to behave exactly as `1`. */
  readonly startAt?: number | undefined;
  /** The relationship id of an `a:buBlip`'s image. */
  readonly blip?: string | undefined;
  readonly font?: BulletFont | undefined;
  readonly size?: BulletSize | undefined;
  readonly color?: BulletColor<TColor> | undefined;
}

/* -------------------------------------------------------------------------- */
/* the numbering pass                                                         */
/* -------------------------------------------------------------------------- */

/** What `numberParagraphs` needs to know about one paragraph. */
export interface NumberedParagraph {
  /** `a:pPr/@lvl`, 0 through 8. */
  readonly level: number;
  /** The resolved `a:buAutoNum/@type`, or `null` for any other bullet kind. */
  readonly scheme: string | null;
  /** The resolved `a:buAutoNum/@startAt`; absent behaves as 1. */
  readonly startAt: number;
}

/**
 * The number each paragraph's bullet carries, or `null` where it has none.
 *
 * Measured on fourteen packages, one interruption each, and it is the only rule
 * of the six candidates that fits all fifteen readings. The four it beats are
 * all designs somebody would write:
 *
 * | reading                                                    | fits |
 * | ---------------------------------------------------------- | ---- |
 * | **measured**                                               | 15   |
 * | a change of `startAt` does not break the run               | 13   |
 * | an unnumbered paragraph consumes a number                  | 13   |
 * | any intervening paragraph breaks the run                   | 12   |
 * | one counter per level that never restarts                  | 9    |
 * | one counter per text body                                  | 7    |
 *
 * The two that decide it:
 *
 * - **A deeper level does not break the run.** `1, 2, [1, 2], 3, 4` - the outer
 *   list resumes at three after a nested one, so the walk skips deeper
 *   paragraphs rather than stopping at them.
 * - **A shallower one does.** `1, [1], 2, [1]` - the inner list restarts every
 *   time it is re-entered, because the outer paragraph between the two inner
 *   ones ends the inner run.
 *
 * And one that is easy to miss: a paragraph whose `startAt` differs from its
 * predecessor's starts a **new** run, so `7, 1, 2, 3` is what a `startAt="7"`
 * followed by three plain paragraphs renders - not `7, 8, 9, 10`. Repeat the
 * `startAt` on every paragraph and it *is* `7, 8, 9, 10`.
 */
export function numberParagraphs(
  paragraphs: readonly NumberedParagraph[],
): readonly (number | null)[] {
  return paragraphs.map((paragraph, index) => {
    if (paragraph.scheme === null) return null;
    let before = 0;
    for (let back = index - 1; back >= 0; back--) {
      const other = paragraphs[back];
      if (other === undefined) break;
      if (other.level > paragraph.level) continue;
      if (other.level < paragraph.level) break;
      if (other.scheme !== paragraph.scheme || other.startAt !== paragraph.startAt) break;
      before++;
    }
    const value = paragraph.startAt + before;
    // A list long enough to run past `ST_TextBulletStartAtNum` is one PowerPoint
    // would itself wrap; `formatAutonumber` refuses rather than guessing, so the
    // ceiling is enforced here where the caller can see why.
    if (value > START_AT_MAX) {
      throw new TextError(
        'TEXT_AUTONUMBER_VALUE',
        `paragraph ${String(index)} would be numbered ${String(value)}, past ST_TextBulletStartAtNum`,
        String(value),
      );
    }
    return value;
  });
}

/* -------------------------------------------------------------------------- */
/* the symbol-font mapping                                                    */
/* -------------------------------------------------------------------------- */

/**
 * The twenty-seven code points where Windows-1252 is not Latin-1.
 *
 * This table is the whole of the private-use finding. A symbol bullet is mapped
 * by **ANSI codepage byte**, not by code point, and U+2022 BULLET - by a
 * distance the most common bullet character there is - is byte 0x95 in
 * Windows-1252 and nothing at all in Latin-1.
 */
const CP1252_HIGH: Readonly<Record<number, number>> = {
  0x20ac: 0x80,
  0x201a: 0x82,
  0x0192: 0x83,
  0x201e: 0x84,
  0x2026: 0x85,
  0x2020: 0x86,
  0x2021: 0x87,
  0x02c6: 0x88,
  0x2030: 0x89,
  0x0160: 0x8a,
  0x2039: 0x8b,
  0x0152: 0x8c,
  0x017d: 0x8e,
  0x2018: 0x91,
  0x2019: 0x92,
  0x201c: 0x93,
  0x201d: 0x94,
  0x2022: 0x95,
  0x2013: 0x96,
  0x2014: 0x97,
  0x02dc: 0x98,
  0x2122: 0x99,
  0x0161: 0x9a,
  0x203a: 0x9b,
  0x0153: 0x9c,
  0x017e: 0x9e,
  0x0178: 0x9f,
};

/**
 * The symbol faces this experiment measured, and no more.
 *
 * A list rather than a rule, deliberately. What makes a face a symbol face is a
 * (3,0) symbol subtable in its `cmap`, and reading that needs the font itself -
 * which is 8.1's SFNT reader, not this package. Until then a face outside this
 * list is treated as an ordinary one, which is the behaviour that leaves the
 * character alone rather than mangling it.
 */
export const SYMBOL_TYPEFACES: readonly string[] = [
  'Wingdings',
  'Wingdings 2',
  'Wingdings 3',
  'Webdings',
  'Symbol',
];

const SYMBOL_SET: ReadonlySet<string> = new Set(SYMBOL_TYPEFACES);

/** Whether `typeface` is one of the symbol faces T5 measured. */
export function isSymbolTypeface(typeface: string): boolean {
  return SYMBOL_SET.has(typeface);
}

/**
 * The character a symbol face actually draws for `char`.
 *
 * Four readings were scored over 21 probes and only this one fits all of them:
 *
 * | reading                                            | fits |
 * | -------------------------------------------------- | ---- |
 * | **the ANSI codepage byte, plus 0xF000**            | 21   |
 * | the low byte, plus 0xF000                          | 20   |
 * | only U+0020..U+00FF, plus 0xF000                   | 19   |
 * | every face maps, not only the symbol ones          | 17   |
 * | nothing is ever mapped                             | 13   |
 *
 * The two that separate them are `char="•"` in Wingdings, which draws U+F095
 * rather than U+F022, and `char="§"` in **Arial**, which draws a section sign
 * and not a Wingdings glyph.
 */
export function symbolBulletChar(char: string, typeface: string | undefined): string {
  if (typeface === undefined || !isSymbolTypeface(typeface)) return char;
  const code = char.codePointAt(0);
  if (code === undefined) return char;
  const byte = code >= 0x20 && code <= 0xff ? code : (CP1252_HIGH[code] ?? code & 0xff);
  return String.fromCodePoint(0xf000 + byte);
}

/* -------------------------------------------------------------------------- */
/* what to draw                                                               */
/* -------------------------------------------------------------------------- */

/** The run properties a bullet takes its face, size and colour from. */
export interface BulletContext<TColor> {
  /** The **first** run's typeface. Measured: not the largest run's, not the paragraph's. */
  readonly typeface: string;
  /** The first run's `@sz`, hundredths of a point. */
  readonly sz: number;
  readonly color: TColor;
}

/** A bullet resolved to the thing a renderer draws. */
export type DrawableBullet<TColor> =
  | { readonly kind: 'none' }
  | {
      readonly kind: 'text';
      /** The characters to draw, already mapped for a symbol face. */
      readonly text: string;
      readonly typeface: string;
      /** Hundredths of a point. */
      readonly sz: number;
      readonly color: TColor;
    }
  | {
      readonly kind: 'picture';
      readonly blip: string;
      /** The height to scale the picture to, in hundredths of a point. */
      readonly sz: number;
    };

/**
 * The size a bullet is drawn at, in hundredths of a point.
 *
 * `a:buSzPct` is a percentage of the **first run's** size - not of the
 * paragraph's `a:defRPr`, and not of the largest run's, both of which were
 * probed with sizes that disagree. `a:buSzPts` is absolute. `a:buSzTx` and an
 * absent size are both the run's own size.
 */
function bulletSize(size: BulletSize | undefined, runSz: number): number {
  if (size === undefined || size.kind === 'text') return runSz;
  if (size.kind === 'points') return size.value;
  return Math.round((runSz * size.value) / 100000);
}

/**
 * What to draw for one paragraph's bullet.
 *
 * `number` is what `numberParagraphs` returned for this paragraph, and is
 * required for an autonumber precisely because it cannot be derived from the
 * paragraph alone - it is a property of the paragraph's position in its run.
 */
export function drawableBullet<TColor>(
  bullet: ResolvedBullet<TColor>,
  context: BulletContext<TColor>,
  number?: number,
): DrawableBullet<TColor> {
  if (bullet.kind === 'none') return { kind: 'none' };

  const sz = bulletSize(bullet.size, context.sz);
  const color =
    bullet.color === undefined || bullet.color.kind === 'text' ? context.color : bullet.color.value;

  if (bullet.kind === 'blip') {
    if (bullet.blip === undefined) {
      throw new TextError('TEXT_BULLET', 'a picture bullet with no a:buBlip relationship');
    }
    return { kind: 'picture', blip: bullet.blip, sz };
  }

  if (bullet.kind === 'autonum') {
    if (bullet.scheme === undefined) {
      throw new TextError('TEXT_BULLET', 'an autonumber bullet with no a:buAutoNum/@type');
    }
    if (number === undefined) {
      throw new TextError(
        'TEXT_BULLET',
        `the ${bullet.scheme} bullet needs its number; call numberParagraphs first`,
        bullet.scheme,
      );
    }
    // `a:buFont` is measured inert here. The two Wingdings schemes are the sole
    // exception, and they are an exception about the *glyphs*, not the file:
    // their code points exist in no text face.
    const typeface = autonumberTypeface(bullet.scheme) ?? context.typeface;
    return { kind: 'text', text: formatAutonumber(bullet.scheme, number), typeface, sz, color };
  }

  if (bullet.char === undefined) {
    throw new TextError('TEXT_BULLET', 'a character bullet with no a:buChar/@char');
  }
  const typeface =
    bullet.font === undefined || bullet.font.kind === 'text' ? context.typeface : bullet.font.value;
  return { kind: 'text', text: symbolBulletChar(bullet.char, typeface), typeface, sz, color };
}

/* -------------------------------------------------------------------------- */
/* where it goes                                                              */
/* -------------------------------------------------------------------------- */

/**
 * The height a picture bullet is scaled to, as a fraction of the font size.
 *
 * One number and not three: a square source at 12, 24 and 48 points advances
 * 8.4, 16.8 and 33.6. The width follows the source's aspect ratio, so a 64x16
 * picture is four times as wide as a square one and a 16x64 picture a quarter -
 * a renderer that letterboxes into a square box puts the text in the wrong place
 * on both.
 */
export const BLIP_BULLET_HEIGHT = 0.7;

/** Where the bullet is drawn and where the paragraph's text starts, in points. */
export interface BulletLayout {
  /** The bullet's left edge, from the text frame's left inset. */
  readonly bulletLeftPt: number;
  /** The first line's text left edge. */
  readonly firstLineLeftPt: number;
  /** Every line after the first. */
  readonly wrappedLeftPt: number;
}

/**
 * Where a bulleted paragraph's bullet and text sit, in points from the inset.
 *
 * Scored over 108 probes sweeping `marL` across nine values and `indent` across
 * six, with two bullets whose advances straddle the interesting region. Two
 * things a plausible implementation gets wrong:
 *
 * - **A positive `indent` does not move the bullet.** Only a hanging indent
 *   does, so the term is `min(0, indent)`. Eighteen probes at `indent="12"` put
 *   the text exactly where `indent="0"` does.
 * - **The bullet is clamped to the frame.** A hanging indent deeper than `marL`
 *   would put it left of the text box; PowerPoint pins it at zero and the text
 *   follows, which is why `marL="0" indent="-36"` starts the text at 36 and not
 *   at the bullet's own width.
 *
 * The `-indent` term is the third: a hanging indent wider than the bullet holds
 * the first line out to it, so the text lines up with the wrapped lines below
 * only when `marL` is at least the hanging width.
 *
 * A wrapped line is at `marL` exactly - no bullet, no clamp - which is a
 * separate reading from the first line's and was scored separately.
 */
export function bulletLayout(
  marLPt: number,
  indentPt: number,
  bulletAdvancePt: number,
): BulletLayout {
  for (const [name, value] of [
    ['marL', marLPt],
    ['indent', indentPt],
    ['the bullet advance', bulletAdvancePt],
  ] as const) {
    if (!Number.isFinite(value)) {
      throw new TextError(
        'TEXT_BULLET',
        `${name} is ${String(value)}, not a length`,
        String(value),
      );
    }
  }
  if (bulletAdvancePt < 0) {
    throw new TextError(
      'TEXT_BULLET',
      `the bullet advance is ${String(bulletAdvancePt)}, which is negative`,
      String(bulletAdvancePt),
    );
  }
  const bulletLeftPt = Math.max(0, marLPt + Math.min(0, indentPt));
  return {
    bulletLeftPt,
    firstLineLeftPt: Math.max(marLPt, -indentPt, bulletLeftPt + bulletAdvancePt),
    wrappedLeftPt: marLPt,
  };
}

/**
 * The width a picture bullet occupies, in points.
 *
 * `aspect` is the source image's width divided by its height.
 */
export function blipBulletWidth(sz: number, aspect: number): number {
  if (!Number.isFinite(aspect) || aspect <= 0) {
    throw new TextError(
      'TEXT_BULLET',
      `a picture bullet aspect ratio of ${String(aspect)} is not usable`,
      String(aspect),
    );
  }
  return (BLIP_BULLET_HEIGHT * sz * aspect) / 100;
}

/** Whether `startAt` is inside `ST_TextBulletStartAtNum`, which PowerPoint enforces. */
export function isStartAtInRange(startAt: number): boolean {
  return Number.isInteger(startAt) && startAt >= START_AT_MIN && startAt <= START_AT_MAX;
}
