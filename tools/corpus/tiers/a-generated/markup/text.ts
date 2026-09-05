/**
 * Text markup for the Tier A probes.
 *
 * Six decks - `a02` and `a07` through `a11` - are about what happens inside a
 * `p:txBody`, and between them they need every child of `a:pPr` and `a:rPr`.
 * Writing that inline in six files would be six chances to get the sequence
 * wrong, and the sequence is the part PowerPoint refuses over.
 *
 * So the orders below are copied out of `packages/xml/src/edit/schema-order.gen.ts`,
 * which is generated from the ECMA-376 Transitional XSDs, and every builder
 * here emits its children in that order and no other:
 *
 * ```
 * pPr  = lnSpc, spcBef, spcAft, buClr|buClrTx, buSzPct|buSzPts|buSzTx,
 *        buFont|buFontTx, buAutoNum|buBlip|buChar|buNone, tabLst, defRPr, extLst
 * rPr  = ln, <fill>, effectDag|effectLst, highlight, uLn|uLnTx, uFill|uFillTx,
 *        latin, ea, cs, sym, hlinkClick, hlinkMouseOver, rtl, extLst
 * p    = pPr, (br|fld|r)*, endParaRPr
 * fld  = rPr, pPr, t
 * bodyPr = prstTxWarp, noAutofit|normAutofit|spAutoFit, scene3d, flatTx|sp3d, extLst
 * lstStyle = defPPr, lvl1pPr … lvl9pPr, extLst
 * ```
 *
 * Two of those are worth saying out loud because they are counter-intuitive
 * enough to be written wrong from memory. `a:rtl` is the **last** child of
 * `a:rPr` before `a:extLst`, not an attribute and not near the top - that was
 * experiment E11's question. And `a:fld` puts `a:rPr` before `a:pPr`, which is
 * the opposite of the containment order everywhere else in DrawingML.
 */

import { escapeAttribute, escapeXml, NS_R } from './chassis.ts';

// ------------------------------------------------------------ run properties

export interface RunProps {
  /** `@lang`. What `Intl.DateTimeFormat` is keyed on when a field renders. */
  readonly lang?: string;
  readonly altLang?: string;
  /** `@sz` in hundredths of a point: 1800 is 18pt. */
  readonly sz?: number;
  readonly b?: boolean;
  readonly i?: boolean;
  /** `ST_TextUnderlineType`. */
  readonly u?: string;
  /** `noStrike`, `sngStrike` or `dblStrike`. */
  readonly strike?: string;
  readonly kern?: number;
  /** `none`, `small` or `all`. */
  readonly cap?: string;
  /** `@spc`, letter spacing in hundredths of a point. Negative is legal. */
  readonly spc?: number;
  /** `@baseline`, an `ST_Percentage`: superscript is positive. */
  readonly baseline?: number;
  readonly normalizeH?: boolean;
  readonly noProof?: boolean;
  readonly dirty?: boolean;
  readonly smtClean?: boolean;
  // --- children, emitted in the order above ---
  readonly fill?: string;
  readonly highlight?: string;
  /** `a:latin/@typeface`. */
  readonly latin?: string;
  /** `a:ea/@typeface` - the East Asian face for a mixed run. */
  readonly ea?: string;
  /** `a:cs/@typeface` - the complex-script face. */
  readonly cs?: string;
  /** `a:sym/@typeface` - the symbol face, which is what Wingdings text needs. */
  readonly sym?: string;
  /** Relationship id for `a:hlinkClick`. */
  readonly hlink?: string;
  /** `a:rtl`, the element. Right-to-left paragraph direction lives on `a:pPr`. */
  readonly rtl?: boolean;
}

function onOff(name: string, value: boolean | undefined): string {
  return value === undefined ? '' : ` ${name}="${value ? '1' : '0'}"`;
}

function attr(name: string, value: string | number | undefined): string {
  return value === undefined ? '' : ` ${name}="${String(value)}"`;
}

function font(tag: string, typeface: string | undefined): string {
  return typeface === undefined ? '' : `<a:${tag} typeface="${escapeAttribute(typeface)}"/>`;
}

/** `a:rPr`, `a:defRPr` or `a:endParaRPr` - the same content model under three names. */
export function rPr(tag: 'rPr' | 'defRPr' | 'endParaRPr', props: RunProps = {}): string {
  const attributes =
    attr('lang', props.lang) +
    attr('altLang', props.altLang) +
    attr('sz', props.sz) +
    onOff('b', props.b) +
    onOff('i', props.i) +
    attr('u', props.u) +
    attr('strike', props.strike) +
    attr('kern', props.kern) +
    attr('cap', props.cap) +
    attr('spc', props.spc) +
    onOff('normalizeH', props.normalizeH) +
    attr('baseline', props.baseline) +
    onOff('noProof', props.noProof) +
    onOff('dirty', props.dirty) +
    onOff('smtClean', props.smtClean);

  const children =
    (props.fill ?? '') +
    (props.highlight === undefined ? '' : `<a:highlight>${props.highlight}</a:highlight>`) +
    font('latin', props.latin) +
    font('ea', props.ea) +
    font('cs', props.cs) +
    font('sym', props.sym) +
    (props.hlink === undefined ? '' : `<a:hlinkClick xmlns:r="${NS_R}" r:id="${props.hlink}"/>`) +
    (props.rtl === undefined ? '' : `<a:rtl val="${props.rtl ? '1' : '0'}"/>`);

  return children === ''
    ? `<a:${tag}${attributes}/>`
    : `<a:${tag}${attributes}>${children}</a:${tag}>`;
}

// ------------------------------------------------------ paragraph properties

export interface ParaProps {
  /** `@marL`. The schema default is 347663, and applying it eagerly is a bug. */
  readonly marL?: number;
  readonly marR?: number;
  /** `@lvl`, zero-based: `lvl="1"` selects `lvl2pPr`. */
  readonly lvl?: number;
  /** `@indent`. The schema default is -342900; same caveat as `marL`. */
  readonly indent?: number;
  /** `l`, `ctr`, `r`, `just`, `justLow`, `dist` or `thaiDist`. */
  readonly algn?: string;
  readonly defTabSz?: number;
  /** `@rtl` - the paragraph's own direction, distinct from `a:rtl` on a run. */
  readonly rtl?: boolean;
  readonly eaLnBrk?: boolean;
  /** `auto`, `t`, `ctr`, `base` or `b`. */
  readonly fontAlgn?: string;
  /**
   * `@latinLnBrk`. ECMA says the default is true; Word and PowerPoint behave as
   * if it is false, which is why a probe deck states it either way explicitly.
   */
  readonly latinLnBrk?: boolean;
  readonly hangingPunct?: boolean;
  // --- children, in order ---
  /** `a:lnSpc` contents: one `a:spcPct` or `a:spcPts`. */
  readonly lnSpc?: string;
  readonly spcBef?: string;
  readonly spcAft?: string;
  /** `a:buClr` contents, or the literal `<a:buClrTx/>`. */
  readonly buClr?: string;
  /** `a:buSzPct`/`a:buSzPts`/`a:buSzTx` markup, already built. */
  readonly buSz?: string;
  /** `a:buFont` markup, or the literal `<a:buFontTx/>`. */
  readonly buFont?: string;
  /** One of `a:buAutoNum`, `a:buBlip`, `a:buChar`, `a:buNone`. */
  readonly bullet?: string;
  /** `a:tabLst` markup. */
  readonly tabLst?: string;
  /** `a:defRPr` markup - build it with `rPr('defRPr', …)`. */
  readonly defRPr?: string;
}

/** `a:pPr`, `a:defPPr`, or one of `a:lvl1pPr` … `a:lvl9pPr`. */
export function pPr(tag: string, props: ParaProps = {}): string {
  const attributes =
    attr('marL', props.marL) +
    attr('marR', props.marR) +
    attr('lvl', props.lvl) +
    attr('indent', props.indent) +
    attr('algn', props.algn) +
    attr('defTabSz', props.defTabSz) +
    onOff('rtl', props.rtl) +
    onOff('eaLnBrk', props.eaLnBrk) +
    attr('fontAlgn', props.fontAlgn) +
    onOff('latinLnBrk', props.latinLnBrk) +
    onOff('hangingPunct', props.hangingPunct);

  const children =
    (props.lnSpc === undefined ? '' : `<a:lnSpc>${props.lnSpc}</a:lnSpc>`) +
    (props.spcBef === undefined ? '' : `<a:spcBef>${props.spcBef}</a:spcBef>`) +
    (props.spcAft === undefined ? '' : `<a:spcAft>${props.spcAft}</a:spcAft>`) +
    (props.buClr ?? '') +
    (props.buSz ?? '') +
    (props.buFont ?? '') +
    (props.bullet ?? '') +
    (props.tabLst ?? '') +
    (props.defRPr ?? '');

  return children === ''
    ? `<a:${tag}${attributes}/>`
    : `<a:${tag}${attributes}>${children}</a:${tag}>`;
}

/**
 * `a:spcPct`.
 *
 * `ST_Percentage` has two lexical forms and both are legal: `150000` and
 * `"150%"`. `parseInt("150%")` yields 150, i.e. 0.15% line spacing, and the
 * slide collapses to a single line - so `a07` writes both forms deliberately
 * and this helper is how it says which one it meant.
 */
export const spcPct = (value: number | string): string => `<a:spcPct val="${String(value)}"/>`;

export const spcPts = (hundredthsOfAPoint: number): string =>
  `<a:spcPts val="${String(hundredthsOfAPoint)}"/>`;

// --------------------------------------------------------------- paragraphs

export interface Run {
  readonly text: string;
  readonly props?: RunProps;
}

/** `a:r`. `CT_RegularTextRun` is `rPr` then `t`, and `a:t` is never empty here. */
export function run(text: string, props: RunProps = {}): string {
  return `<a:r>${rPr('rPr', props)}<a:t>${escapeXml(text)}</a:t></a:r>`;
}

/**
 * `a:br` - a hard line break.
 *
 * Not a paragraph break. It starts no new bullet number, applies no `spcBef` or
 * `spcAft`, and does not reset the first-line indent. Rendering it as a
 * paragraph is a visible bug in every bulleted list it appears in.
 */
export const br = (props: RunProps = {}): string => `<a:br>${rPr('rPr', props)}</a:br>`;

/**
 * `a:fld`.
 *
 * `@id` is an `ST_Guid` and is required. It is generated once when the field is
 * created and persists unchanged; regenerating it on export is a documented
 * route to a repair prompt, so every GUID in this corpus is a literal written
 * down in the deck that uses it.
 *
 * The `a:t` is the **cached** text - what the authoring application last
 * rendered. A viewer that cannot compute the field falls back to it, so a field
 * with no cached text renders as nothing at all.
 */
export function field(spec: {
  readonly id: string;
  readonly type: string;
  readonly text: string;
  readonly props?: RunProps;
  readonly paraProps?: string;
}): string {
  return (
    `<a:fld id="${spec.id}" type="${escapeAttribute(spec.type)}">` +
    rPr('rPr', spec.props ?? {}) +
    (spec.paraProps ?? '') +
    `<a:t>${escapeXml(spec.text)}</a:t>` +
    '</a:fld>'
  );
}

/** `a:p`. Pass `content` already built, so a paragraph can mix runs, breaks and fields. */
export function para(spec: {
  readonly props?: ParaProps;
  readonly content?: string;
  /** `a:endParaRPr`, which is what sets the height of an empty paragraph. */
  readonly endProps?: RunProps;
}): string {
  return (
    '<a:p>' +
    (spec.props === undefined ? '' : pPr('pPr', spec.props)) +
    (spec.content ?? '') +
    (spec.endProps === undefined ? '' : rPr('endParaRPr', spec.endProps)) +
    '</a:p>'
  );
}

/**
 * One paragraph, one run, nothing else - the common case in a probe deck.
 *
 * Named `textLine` rather than `line` because `shapes.ts` already exports a
 * `line` that builds an `a:ln`, and a deck that imports both should not have to
 * think about which one it got.
 */
export const textLine = (text: string, props: RunProps = {}, paraProps?: ParaProps): string =>
  para({ ...(paraProps === undefined ? {} : { props: paraProps }), content: run(text, props) });

// --------------------------------------------------------------- body and list

export interface BodyProps {
  readonly rot?: number;
  /**
   * `@spcFirstLastPara`, default **false**: the first paragraph's `spcBef` and
   * the last's `spcAft` are discarded. Honouring them anyway shifts every
   * centred and bottom-anchored text box.
   */
  readonly spcFirstLastPara?: boolean;
  /** `overflow`, `ellipsis` or `clip`. */
  readonly vertOverflow?: string;
  readonly horzOverflow?: string;
  /** `horz`, `vert`, `vert270`, `wordArtVert`, `eaVert`, `mongolianVert`, `wordArtVertRtl`. */
  readonly vert?: string;
  /** `none` or `square`. */
  readonly wrap?: string;
  /** Defaults are asymmetric: 91440 / 45720 / 91440 / 45720. */
  readonly lIns?: number;
  readonly tIns?: number;
  readonly rIns?: number;
  readonly bIns?: number;
  readonly numCol?: number;
  readonly spcCol?: number;
  readonly rtlCol?: boolean;
  /** `t`, `ctr`, `b`, `just` or `dist`. */
  readonly anchor?: string;
  readonly anchorCtr?: boolean;
  readonly upright?: boolean;
  readonly compatLnSpc?: boolean;
  /** `a:noAutofit`, `a:normAutofit` or `a:spAutoFit` markup, already built. */
  readonly autofit?: string;
}

/** `a:bodyPr`. */
export function bodyPr(props: BodyProps = {}): string {
  const attributes =
    attr('rot', props.rot) +
    onOff('spcFirstLastPara', props.spcFirstLastPara) +
    attr('vertOverflow', props.vertOverflow) +
    attr('horzOverflow', props.horzOverflow) +
    attr('vert', props.vert) +
    attr('wrap', props.wrap) +
    attr('lIns', props.lIns) +
    attr('tIns', props.tIns) +
    attr('rIns', props.rIns) +
    attr('bIns', props.bIns) +
    attr('numCol', props.numCol) +
    attr('spcCol', props.spcCol) +
    onOff('rtlCol', props.rtlCol) +
    attr('anchor', props.anchor) +
    onOff('anchorCtr', props.anchorCtr) +
    onOff('upright', props.upright) +
    onOff('compatLnSpc', props.compatLnSpc);
  return props.autofit === undefined
    ? `<a:bodyPr${attributes}/>`
    : `<a:bodyPr${attributes}>${props.autofit}</a:bodyPr>`;
}

/**
 * `a:normAutofit`.
 *
 * `@fontScale` and `@lnSpcReduction` are `ST_TextFontScalePercentOrPercentString`
 * and `ST_TextSpacingPercentOrPercentString`, so both accept the two lexical
 * forms. PowerPoint only ever writes values from its own discrete ladder and
 * re-snaps anything else on open, which is visible as text jumping - so `a11`
 * carries the ladder steps and nothing between them.
 */
export function normAutofit(fontScale?: number | string, lnSpcReduction?: number | string): string {
  return (
    '<a:normAutofit' +
    (fontScale === undefined ? '' : ` fontScale="${String(fontScale)}"`) +
    (lnSpcReduction === undefined ? '' : ` lnSpcReduction="${String(lnSpcReduction)}"`) +
    '/>'
  );
}

/** A whole `p:txBody`. `CT_TextBody` is bodyPr, lstStyle, then one or more `a:p`. */
export function txBody(spec: {
  readonly bodyPr?: string;
  readonly lstStyle?: string;
  readonly paras: string;
}): string {
  return (
    '<p:txBody>' +
    (spec.bodyPr ?? '<a:bodyPr/>') +
    (spec.lstStyle ?? '<a:lstStyle/>') +
    spec.paras +
    '</p:txBody>'
  );
}

/**
 * The same body, in the DrawingML namespace.
 *
 * `p:txBody` and `a:txBody` are the **same complex type** - `CT_TextBody`,
 * bodyPr, lstStyle, one or more `a:p` - declared twice under two namespaces.
 * A shape holds the PresentationML one; a table cell and a chart's rich text
 * hold the DrawingML one. Everything inside is identical, so the only thing
 * that ever goes wrong here is the wrapper, and a table whose cells say
 * `p:txBody` is a package PowerPoint refuses.
 */
export function aTxBody(spec: {
  readonly bodyPr?: string;
  readonly lstStyle?: string;
  readonly paras: string;
}): string {
  return (
    '<a:txBody>' +
    (spec.bodyPr ?? '<a:bodyPr/>') +
    (spec.lstStyle ?? '<a:lstStyle/>') +
    spec.paras +
    '</a:txBody>'
  );
}

/** `a:lstStyle` from a sparse level map. Levels are emitted 1..9 in order. */
export function lstStyle(levels: Readonly<Record<number, ParaProps>>, defPPr?: ParaProps): string {
  const body =
    (defPPr === undefined ? '' : pPr('defPPr', defPPr)) +
    [1, 2, 3, 4, 5, 6, 7, 8, 9]
      .map((n) => {
        const level = levels[n];
        return level === undefined ? '' : pPr(`lvl${String(n)}pPr`, level);
      })
      .join('');
  return body === '' ? '<a:lstStyle/>' : `<a:lstStyle>${body}</a:lstStyle>`;
}

// -------------------------------------------------------------------- bullets

/** `a:buChar`. `@char` is one character; `buFont` decides which glyph it is. */
export const buChar = (char: string): string => `<a:buChar char="${escapeAttribute(char)}"/>`;

export function buAutoNum(type: string, startAt?: number): string {
  return (
    `<a:buAutoNum type="${type}"` +
    (startAt === undefined ? '' : ` startAt="${String(startAt)}"`) +
    '/>'
  );
}

/** `a:buBlip`, whose only child is `a:blip`. The image is a slide relationship. */
export const buBlip = (relId: string): string =>
  `<a:buBlip><a:blip r:embed="${relId}"/></a:buBlip>`;

export const buFont = (typeface: string, pitchFamily?: number, charset?: number): string =>
  `<a:buFont typeface="${escapeAttribute(typeface)}"` +
  (pitchFamily === undefined ? '' : ` pitchFamily="${String(pitchFamily)}"`) +
  (charset === undefined ? '' : ` charset="${String(charset)}"`) +
  '/>';

export const buNone = '<a:buNone/>';
export const buSzPct = (value: number): string => `<a:buSzPct val="${String(value)}"/>`;
