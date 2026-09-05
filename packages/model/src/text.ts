/**
 * The text model, and the ten sources a run's size can come from.
 *
 * Everything here is `T | undefined` for the same reason the rest of the model
 * is: a run whose `a:rPr` declares no size is not an 18-point run, it is a run
 * that has not said, and the difference is what the whole cascade is about.
 *
 * ## What the cascade turned out to be
 *
 * Measured in 3.1 against PowerPoint, over a ladder of nine packages each
 * declaring one source fewer than the last. Scored over all 40320 permutations
 * of the eight sources the plan names: **none of them fits** while a placeholder
 * is assumed to read every source, and exactly one order fits once the two it
 * ignores are dropped. `corpus/ground-truth/text-cascade.json`, and
 * `docs/adr/phase-3-text/0027-the-text-cascade.md`.
 *
 * The walk, nearest first:
 *
 * 1. the run's own `a:rPr`
 * 2. the paragraph's `a:pPr/a:defRPr`
 * 3. the shape's `a:lstStyle`, at the paragraph's level
 * 4. each ancestor placeholder's `a:lstStyle`, up the same chain 2.9 measured
 *    for geometry - `@idx` alone to the layout, the folded `@type` to the master
 *
 * and then **one of two termini**, which is where every implementation this
 * project has read goes wrong:
 *
 * - a shape whose chain ends at a type with a bucket reads the master's
 *   `p:txStyles[bucket]` **and stops**
 * - a shape with no bucket reads `p:defaultTextStyle` **and stops**
 *
 * A body placeholder ignores a `p:defaultTextStyle` that is right there in the
 * package; a slide-number placeholder reads it. They are not two rungs of one
 * ladder.
 *
 * ## Three things the plan calls sources and PowerPoint does not read
 *
 * `a:objectDefaults/a:spDef/a:lstStyle`, `a:defPPr`, and `p:otherStyle`. The
 * last is the surprising one: ECMA calls it "the text style for all other text",
 * and nothing measurable reads it - not a shape that is not a placeholder on the
 * slide, the layout or the master, and not the `dt`/`ftr`/`sldNum` placeholders
 * the plan assigned to it, which fall through to `p:defaultTextStyle` instead.
 */

import type { Color, Effect, Fill, Line } from '@pptx-studio/paint';
import type { XElement } from '@pptx-studio/xml';

/* -------------------------------------------------------------------------- */
/* scalars                                                                    */
/* -------------------------------------------------------------------------- */

/** `ST_TextAlignType`. */
export type TextAlign = 'l' | 'ctr' | 'r' | 'just' | 'justLow' | 'dist' | 'thaiDist';

/** `ST_TextFontAlignType`. */
export type FontAlign = 'auto' | 't' | 'ctr' | 'base' | 'b';

/** `ST_TextUnderlineType`, all seventeen, as written. */
export type Underline =
  | 'none'
  | 'words'
  | 'sng'
  | 'dbl'
  | 'heavy'
  | 'dotted'
  | 'dottedHeavy'
  | 'dash'
  | 'dashHeavy'
  | 'dashLong'
  | 'dashLongHeavy'
  | 'dotDash'
  | 'dotDashHeavy'
  | 'dotDotDash'
  | 'dotDotDashHeavy'
  | 'wavy'
  | 'wavyHeavy'
  | 'wavyDbl';

/** `ST_TextStrikeType`. */
export type Strike = 'noStrike' | 'sngStrike' | 'dblStrike';

/** `ST_TextCapsType`. */
export type Caps = 'none' | 'small' | 'all';

/**
 * `a:lnSpc`, `a:spcBef` and `a:spcAft`, which are one of two quite different
 * things and say which in their child element's name.
 *
 * `a:spcPct` is hundred-thousandths of the line height - 150000 is 150%.
 * `a:spcPts` is hundredths of a point - 3000 is 30pt. Conflating them collapses
 * every slide that uses the second, and `ST_Percentage` accepts *both* the
 * `150000` and the `150%` spellings, so the parse is not `parseInt` either:
 * `parseInt("150%")` is 150, which is 0.15% line spacing. Measured in 3.1, and
 * in 2.6 for the colour transforms that share the type.
 */
export type Spacing =
  | { readonly kind: 'percent'; readonly value: number }
  | { readonly kind: 'points'; readonly value: number };

/**
 * One of the four typeface slots on a run.
 *
 * `@typeface` may be a real face name or a theme reference - `+mj-lt`,
 * `+mn-lt`, `+mj-ea`, `+mn-ea`, `+mj-cs`, `+mn-cs` - and the reference is kept
 * as written rather than resolved here, because which theme it resolves against
 * is a property of the sheet and not of the run.
 */
export interface Typeface {
  readonly typeface: string;
  /** `@panose`, twenty hex characters. */
  readonly panose: string | undefined;
  /** `@pitchFamily`, `(family << 4) | pitch`. */
  readonly pitchFamily: number | undefined;
  /** `@charset`, **signed**: Shift-JIS is -128, not 128. */
  readonly charset: number | undefined;
}

/**
 * Which of a theme's two font collections a `+mj-`/`+mn-` reference names.
 *
 * `script` is spelled the way `FontCollection` spells it rather than the way the
 * reference does - `lt` in the markup is the `a:latin` element, and a resolver
 * that carried the reference's spelling would need a second map between two
 * three-member sets to index the thing it just parsed.
 */
export type ThemeFontRef = {
  readonly collection: 'major' | 'minor';
  readonly script: 'latin' | 'ea' | 'cs';
};

const THEME_FONT = /^\+(mj|mn)-(lt|ea|cs)$/;

/** `+mn-lt` and its five siblings, or `null` for a real typeface name. */
export function themeFontRef(typeface: string): ThemeFontRef | null {
  const match = THEME_FONT.exec(typeface);
  if (match === null) return null;
  const script = match[2];
  return {
    collection: match[1] === 'mj' ? 'major' : 'minor',
    script: script === 'lt' ? 'latin' : (script as 'ea' | 'cs'),
  };
}

/* -------------------------------------------------------------------------- */
/* run properties                                                             */
/* -------------------------------------------------------------------------- */

/**
 * `CT_TextCharacterProperties` - `a:rPr`, `a:defRPr` and `a:endParaRPr`.
 *
 * Every field is optional in the file and optional here. The cascade merges
 * these **per property**: measured in 3.1 with five levels declaring one
 * property each, which produced a 24-point bold italic underlined struck run
 * where a nearest-level-takes-all reading gives an 18-point struck one.
 */
export interface RunProps {
  /** `@sz`, hundredths of a point. 2400 is 24pt. */
  readonly sz: number | undefined;
  readonly b: boolean | undefined;
  readonly i: boolean | undefined;
  readonly u: Underline | undefined;
  readonly strike: Strike | undefined;
  readonly cap: Caps | undefined;
  /** `@spc`, hundredths of a point. May be negative. */
  readonly spc: number | undefined;
  /** `@kern`, hundredths of a point: the size at or above which to kern. */
  readonly kern: number | undefined;
  /** `@baseline`, hundred-thousandths. 30000 is superscript. */
  readonly baseline: number | undefined;
  readonly noProof: boolean | undefined;
  readonly lang: string | undefined;
  readonly altLang: string | undefined;
  readonly latin: Typeface | undefined;
  readonly ea: Typeface | undefined;
  readonly cs: Typeface | undefined;
  readonly sym: Typeface | undefined;
  /** The text's own fill: `a:solidFill` and its siblings inside `a:rPr`. */
  readonly fill: Fill | undefined;
  /** `a:ln` inside `a:rPr`: the outline of the glyphs. */
  readonly line: Line | undefined;
  readonly effects: readonly Effect[] | undefined;
  /** `a:highlight`, which is a colour and not a fill. */
  readonly highlight: Color | undefined;
  /** The element this was read from. Edits go here, never to the fields above. */
  readonly node: XElement;
}

/* -------------------------------------------------------------------------- */
/* paragraph properties                                                       */
/* -------------------------------------------------------------------------- */

/* -------------------------------------------------------------------------- */
/* bullets                                                                    */
/* -------------------------------------------------------------------------- */

/**
 * The exclusive group in `CT_TextParagraphProperties`: one of four, or nothing.
 *
 * `undefined` on `ParaProps` is the fifth case and the one that matters -
 * a level that declares no bullet element at all **inherits** one, where
 * `'none'` suppresses it.
 */
export type BulletKind = 'none' | 'char' | 'autonum' | 'blip';

/** `a:buFont` / `a:buFontTx`, which are exclusive of each other. */
export type BulletFont =
  | { readonly kind: 'typeface'; readonly value: Typeface }
  /** `a:buFontTx`: follow the text, and cancel anything inherited. */
  | { readonly kind: 'text' };

/** `a:buSzPct` / `a:buSzPts` / `a:buSzTx`. */
export type BulletSize =
  /** `a:buSzPct/@val`, thousandths of a percent of the first run's size. */
  | { readonly kind: 'percent'; readonly value: number }
  /** `a:buSzPts/@val`, hundredths of a point, absolute. */
  | { readonly kind: 'points'; readonly value: number }
  | { readonly kind: 'text' };

/** `a:buClr` / `a:buClrTx`. */
export type BulletColor =
  { readonly kind: 'color'; readonly value: Color } | { readonly kind: 'text' };

/** `a:buAutoNum`. */
export interface BulletAutoNum {
  /** `@type`, one of the 41 `ST_TextAutonumberScheme` values. */
  readonly type: string;
  /**
   * `@startAt`. Absent is measured to behave exactly as `1` - but the two are
   * kept apart, because a paragraph whose `startAt` differs from its
   * predecessor's begins a new run, and "absent" has to compare equal to "1"
   * for that rule rather than to nothing.
   */
  readonly startAt: number | undefined;
}

/**
 * `CT_TextParagraphProperties` - `a:pPr` and every `a:lvlNpPr`.
 *
 * `@lvl` is deliberately absent. It is the level a paragraph *is at*, not a
 * property that inherits: an `a:lvl3pPr` carrying `lvl="7"` would be nonsense,
 * and a cascade that inherited it would change which level the next lookup used
 * halfway through resolving one paragraph.
 */
export interface ParaProps {
  /** `@marL`, EMU. Measured to default to **0**, not to the schema's 347663. */
  readonly marL: number | undefined;
  readonly marR: number | undefined;
  /** `@indent`, EMU, usually negative. Defaults to 0, not to -342900. */
  readonly indent: number | undefined;
  readonly algn: TextAlign | undefined;
  /** `@defTabSz`, EMU. */
  readonly defTabSz: number | undefined;
  readonly rtl: boolean | undefined;
  readonly eaLnBrk: boolean | undefined;
  readonly fontAlgn: FontAlign | undefined;
  /** `@latinLnBrk`. ECMA says the default is true; PowerPoint behaves as false. */
  readonly latinLnBrk: boolean | undefined;
  readonly hangingPunct: boolean | undefined;
  readonly lnSpc: Spacing | undefined;
  readonly spcBef: Spacing | undefined;
  readonly spcAft: Spacing | undefined;
  /** `a:defRPr`: the run properties every run in this paragraph starts from. */
  readonly defRPr: RunProps | undefined;
  /**
   * Which of `a:buNone`, `a:buChar`, `a:buAutoNum` and `a:buBlip` this level
   * states, or `undefined` where it states none and inherits instead.
   */
  readonly buKind: BulletKind | undefined;
  /** `a:buChar/@char`, exactly as written. The symbol mapping happens at draw time. */
  readonly buChar: string | undefined;
  readonly buAutoNum: BulletAutoNum | undefined;
  /** `a:buBlip/a:blip/@r:embed`. */
  readonly buBlip: string | undefined;
  /**
   * The three decorations, each its own slot.
   *
   * They merge independently of the kind and of each other, which is measured
   * rather than assumed: a level declaring only `a:buFont` re-faces the
   * character it inherits and keeps its size and colour, and a level declaring
   * only `a:buSzTx` cancels an inherited `a:buSzPct` and keeps the rest.
   * Measured on eleven cases twice - once through a shape's own `a:lstStyle`
   * and once through three hops of the master chain.
   */
  readonly buFont: BulletFont | undefined;
  readonly buSize: BulletSize | undefined;
  readonly buColor: BulletColor | undefined;
  readonly node: XElement;
}

/**
 * `CT_TextListStyle`: nine levels, and an `a:defPPr` nothing reads.
 *
 * `levels[0]` is `a:lvl1pPr`, which a paragraph at `lvl="0"` uses. Measured in
 * 3.1: the levels are independent all the way up the chain - a layout
 * placeholder declaring only `a:lvl2pPr` changes the second level and leaves the
 * first and third to the master - so this is nine slots and not one object.
 *
 * `a:defPPr` is parsed and kept because it is in the file, and it is **not** a
 * source: measured on a shape's own `a:lstStyle` and on `p:defaultTextStyle`,
 * declaring only an `a:defPPr` changed nothing either time.
 */
export interface ListStyle {
  /** Nine entries, `a:lvl1pPr` through `a:lvl9pPr`, absent ones `undefined`. */
  readonly levels: readonly (ParaProps | undefined)[];
  /** `a:defPPr`. Preserved, never resolved against. */
  readonly defPPr: ParaProps | undefined;
  readonly node: XElement;
}

/** The number of levels a `CT_TextListStyle` holds, and `@lvl`'s range plus one. */
export const LEVELS = 9;

/* -------------------------------------------------------------------------- */
/* the text body                                                              */
/* -------------------------------------------------------------------------- */

/** `a:r`: a run of text with its own properties. */
export interface TextRun {
  readonly kind: 'run';
  readonly props: RunProps | undefined;
  readonly text: string;
  readonly node: XElement;
}

/**
 * `a:br`: a hard line break.
 *
 * Not a paragraph break. It starts no new bullet, resets no first-line indent
 * and applies neither `spcBef` nor `spcAft` - which is 3.6's problem, and is
 * recorded here so that nothing later mistakes it for an `a:p`.
 */
export interface TextBreak {
  readonly kind: 'br';
  readonly props: RunProps | undefined;
  readonly node: XElement;
}

/** `a:fld`: a field, with the text PowerPoint last cached for it. */
export interface TextField {
  readonly kind: 'field';
  /** `@id`, a required `ST_Guid`. Regenerating it is a repair risk. */
  readonly id: string;
  /** `@type`, one of the fourteen reserved names or a custom one. */
  readonly fieldType: string | undefined;
  readonly props: RunProps | undefined;
  readonly text: string;
  readonly node: XElement;
}

export type TextContent = TextRun | TextBreak | TextField;

/** `a:p`. */
export interface Paragraph {
  readonly props: ParaProps | undefined;
  /**
   * `a:pPr/@lvl`, 0 through 8, defaulting to 0.
   *
   * Read off the paragraph rather than inherited, and clamped rather than
   * thrown on: `ST_TextIndentLevelType` bounds it at 8, and a file outside the
   * bound is one PowerPoint still opens.
   */
  readonly level: number;
  readonly content: readonly TextContent[];
  /** `a:endParaRPr`: the properties of the (empty) run after the last one. */
  readonly endParaRPr: RunProps | undefined;
  readonly node: XElement;
}

/** `ST_TextAnchoringType`. PowerPoint lays `just` and `dist` out as `b`. */
export type TextAnchor = 't' | 'ctr' | 'b' | 'just' | 'dist';

/** `ST_TextVerticalType`. */
export type VerticalText =
  'horz' | 'vert' | 'vert270' | 'wordArtVert' | 'eaVert' | 'mongolianVert' | 'wordArtVertRtl';

/** `ST_TextWrappingType`. */
export type TextWrap = 'none' | 'square';

/** `ST_TextVertOverflowType`. */
export type VertOverflow = 'overflow' | 'ellipsis' | 'clip';

/** `ST_TextHorzOverflowType`. */
export type HorzOverflow = 'overflow' | 'clip';

/** The one autofit child an `a:bodyPr` may carry. */
export type Autofit =
  | { readonly kind: 'none' }
  | { readonly kind: 'shape' }
  | {
      readonly kind: 'normal';
      /** `@fontScale`, thousandths of a percent. */
      readonly fontScale: number | undefined;
      /** `@lnSpcReduction`, thousandths of a percent. */
      readonly lnSpcReduction: number | undefined;
    };

/**
 * `a:bodyPr`, as the file states it.
 *
 * Lengths are EMU and `rot` is 60000ths of a degree, both as written. Every
 * attribute inherits through the placeholder chain on its own - measured in 3.6
 * at all three levels - so absence has to survive parsing.
 */
export interface BodyProps {
  readonly anchor: TextAnchor | undefined;
  readonly anchorCtr: boolean | undefined;
  readonly lIns: number | undefined;
  readonly tIns: number | undefined;
  readonly rIns: number | undefined;
  readonly bIns: number | undefined;
  readonly vert: VerticalText | undefined;
  readonly wrap: TextWrap | undefined;
  readonly vertOverflow: VertOverflow | undefined;
  readonly horzOverflow: HorzOverflow | undefined;
  readonly rot: number | undefined;
  readonly upright: boolean | undefined;
  readonly numCol: number | undefined;
  readonly spcCol: number | undefined;
  readonly rtlCol: boolean | undefined;
  readonly spcFirstLastPara: boolean | undefined;
  readonly compatLnSpc: boolean | undefined;
  readonly fromWordArt: boolean | undefined;
  readonly forceAA: boolean | undefined;
  readonly autofit: Autofit | undefined;
  readonly node: XElement;
}

/** `p:txBody` / `a:txBody`. */
export interface TextBody {
  readonly bodyPr: BodyProps | undefined;
  readonly lstStyle: ListStyle | undefined;
  readonly paragraphs: readonly Paragraph[];
  readonly node: XElement;
}

/* -------------------------------------------------------------------------- */
/* the master's text styles                                                   */
/* -------------------------------------------------------------------------- */

/**
 * The bucket of `p:txStyles` a shape reads, or `null` for one that reads none.
 *
 * `other` is in the union because the element exists, not because anything
 * reaches it - see the module comment.
 */
export type TextStyleBucket = 'title' | 'body' | 'other';

/** `p:txStyles` on a master. */
export interface TextStyles {
  readonly title: ListStyle | undefined;
  readonly body: ListStyle | undefined;
  readonly other: ListStyle | undefined;
  readonly node: XElement;
}
