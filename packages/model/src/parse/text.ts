/**
 * `p:txBody`, `a:lstStyle` and `p:txStyles` into the text model.
 *
 * Nothing here resolves anything and nothing here defaults anything the file
 * could have said. That is the same rule the rest of the parser follows, and in
 * this one corner it is load-bearing rather than tidy: `@marL` and `@indent`
 * have schema defaults of 347663 and -342900, and a parser that materialises
 * them gives every paragraph a 27-point hanging indent it never asked for *and*
 * blocks the value it should have inherited, because a default that has been
 * written down cannot be told from a declaration. Measured in 3.1 - PowerPoint
 * resolves both to 0 when nothing declares them.
 *
 * The one place a value is supplied is `Paragraph.level`, which is 0 when
 * `@lvl` is absent because that is what the attribute means rather than what the
 * schema says, and which is not inherited by anything.
 */

import { parsePercentage } from '@pptx-studio/paint';
import { attributeValue, childElements, firstChild, textContent } from '@pptx-studio/xml';
import type { XElement } from '@pptx-studio/xml';

import { ModelError } from '../errors.js';
import { parseBodyPropsChild } from './body.js';
import { parseColorChild, parseEffects, parseFill, parseLine } from './paint.js';
import {
  type BulletAutoNum,
  type BulletColor,
  type BulletFont,
  type BulletKind,
  type BulletSize,
  LEVELS,
  type Caps,
  type FontAlign,
  type ListStyle,
  type Paragraph,
  type ParaProps,
  type RunProps,
  type Spacing,
  type Strike,
  type TextAlign,
  type TextBody,
  type TextContent,
  type TextStyles,
  type Typeface,
  type Underline,
} from '../text.js';

/* -------------------------------------------------------------------------- */
/* attribute readers                                                          */
/* -------------------------------------------------------------------------- */

/**
 * `ST_OnOff`, which has four lexical forms and not two.
 *
 * `1`/`true` and `0`/`false`. Anything else is a file that says something we do
 * not understand, and guessing is how an explicit `b="0"` becomes bold.
 */
function boolOf(element: XElement, name: string, partName: string): boolean | undefined {
  const raw = attributeValue(element, name);
  if (raw === undefined) return undefined;
  if (raw === '1' || raw === 'true') return true;
  if (raw === '0' || raw === 'false') return false;
  throw new ModelError(
    'MODEL_TEXT_ATTR',
    `${element.qname}/@${name} is "${raw}", which is not a ST_OnOff`,
    partName,
    name,
  );
}

function intOf(element: XElement, name: string, partName: string): number | undefined {
  const raw = attributeValue(element, name);
  if (raw === undefined) return undefined;
  const value = Number(raw);
  if (!Number.isInteger(value)) {
    throw new ModelError(
      'MODEL_TEXT_ATTR',
      `${element.qname}/@${name} is "${raw}", which is not an integer`,
      partName,
      name,
    );
  }
  return value;
}

/**
 * An attribute whose value must be one of a fixed set.
 *
 * A value outside the set throws rather than falling back to the first member.
 * A silent fallback here paints the wrong thing and gives nothing to trace it
 * by, and `ST_TextAlignType` in particular has a member - `just` - whose
 * substitute would be plausible on most slides and wrong on the rest.
 */
function enumOf<T extends string>(
  element: XElement,
  name: string,
  allowed: readonly T[],
  partName: string,
): T | undefined {
  const raw = attributeValue(element, name);
  if (raw === undefined) return undefined;
  if (!(allowed as readonly string[]).includes(raw)) {
    throw new ModelError(
      'MODEL_TEXT_ATTR',
      `${element.qname}/@${name} is "${raw}", which is not one of ${allowed.join(', ')}`,
      partName,
      name,
    );
  }
  return raw as T;
}

const ALIGN: readonly TextAlign[] = ['l', 'ctr', 'r', 'just', 'justLow', 'dist', 'thaiDist'];
const FONT_ALIGN: readonly FontAlign[] = ['auto', 't', 'ctr', 'base', 'b'];
const STRIKE: readonly Strike[] = ['noStrike', 'sngStrike', 'dblStrike'];
const CAPS: readonly Caps[] = ['none', 'small', 'all'];
const UNDERLINE: readonly Underline[] = [
  'none',
  'words',
  'sng',
  'dbl',
  'heavy',
  'dotted',
  'dottedHeavy',
  'dash',
  'dashHeavy',
  'dashLong',
  'dashLongHeavy',
  'dotDash',
  'dotDashHeavy',
  'dotDotDash',
  'dotDotDashHeavy',
  'wavy',
  'wavyHeavy',
  'wavyDbl',
];

/* -------------------------------------------------------------------------- */
/* typefaces                                                                  */
/* -------------------------------------------------------------------------- */

/**
 * `a:latin` and its three siblings.
 *
 * `@charset` is **signed**: Shift-JIS is written `-128`, and reading it as an
 * unsigned byte turns the one value that identifies Japanese text into 128,
 * which is nothing. `@panose` is kept as the twenty hex characters it is,
 * because it is written back verbatim and nothing here interprets it.
 */
function typefaceOf(element: XElement, partName: string): Typeface {
  const typeface = attributeValue(element, 'typeface');
  if (typeface === undefined) {
    throw new ModelError(
      'MODEL_TEXT_ATTR',
      `${element.qname} has no @typeface`,
      partName,
      element.qname,
    );
  }
  return {
    typeface,
    panose: attributeValue(element, 'panose'),
    pitchFamily: intOf(element, 'pitchFamily', partName),
    charset: intOf(element, 'charset', partName),
  };
}

/**
 * A named typeface child.
 *
 * Split from `typefaceOf` because `a:buFont` carries the same attributes on the
 * element itself rather than on a child of a known name, and reading it through
 * a lookup by qname would need the caller to pass its own element as its own
 * parent.
 */
function parseTypeface(parent: XElement, qname: string, partName: string): Typeface | undefined {
  const element = firstChild(parent, qname);
  return element === undefined ? undefined : typefaceOf(element, partName);
}

/* -------------------------------------------------------------------------- */
/* run properties                                                             */
/* -------------------------------------------------------------------------- */

/** `a:rPr`, `a:defRPr` or `a:endParaRPr` - one type under three names. */
export function parseRunProps(element: XElement, partName: string): RunProps {
  return {
    sz: intOf(element, 'sz', partName),
    b: boolOf(element, 'b', partName),
    i: boolOf(element, 'i', partName),
    u: enumOf(element, 'u', UNDERLINE, partName),
    strike: enumOf(element, 'strike', STRIKE, partName),
    cap: enumOf(element, 'cap', CAPS, partName),
    spc: intOf(element, 'spc', partName),
    kern: intOf(element, 'kern', partName),
    baseline: intOf(element, 'baseline', partName),
    noProof: boolOf(element, 'noProof', partName),
    lang: attributeValue(element, 'lang'),
    altLang: attributeValue(element, 'altLang'),
    latin: parseTypeface(element, 'a:latin', partName),
    ea: parseTypeface(element, 'a:ea', partName),
    cs: parseTypeface(element, 'a:cs', partName),
    sym: parseTypeface(element, 'a:sym', partName),
    fill: parseFill(element, partName),
    line: parseLine(element, partName),
    effects: parseEffects(element),
    highlight: parseColorChild(firstChild(element, 'a:highlight')) ?? undefined,
    node: element,
  };
}

function parseRunPropsChild(
  parent: XElement,
  qname: string,
  partName: string,
): RunProps | undefined {
  const element = firstChild(parent, qname);
  return element === undefined ? undefined : parseRunProps(element, partName);
}

/* -------------------------------------------------------------------------- */
/* spacing                                                                    */
/* -------------------------------------------------------------------------- */

/**
 * `a:lnSpc`, `a:spcBef` or `a:spcAft`.
 *
 * The child's name is the unit. `a:spcPct` is a `ST_TextSpacingPercentOrPercentString`
 * and so accepts both the `150000` and the `150%` spellings - `parsePercentage`
 * handles the second, measured in 2.6 for colour and again in 3.1 here.
 * `a:spcPts` is hundredths of a point and has only one spelling.
 */
function parseSpacing(parent: XElement, qname: string, partName: string): Spacing | undefined {
  const element = firstChild(parent, qname);
  if (element === undefined) return undefined;
  const pct = firstChild(element, 'a:spcPct');
  if (pct !== undefined) {
    const raw = attributeValue(pct, 'val');
    if (raw === undefined) {
      throw new ModelError('MODEL_TEXT_ATTR', `${qname}/a:spcPct has no @val`, partName, qname);
    }
    return { kind: 'percent', value: parsePercentage(raw) };
  }
  const pts = firstChild(element, 'a:spcPts');
  if (pts !== undefined) {
    const value = intOf(pts, 'val', partName);
    if (value === undefined) {
      throw new ModelError('MODEL_TEXT_ATTR', `${qname}/a:spcPts has no @val`, partName, qname);
    }
    return { kind: 'points', value };
  }
  // `<a:lnSpc/>` with neither child is legal markup that states nothing, and
  // that is different from stating a spacing of zero.
  return undefined;
}

/* -------------------------------------------------------------------------- */
/* paragraph properties                                                       */
/* -------------------------------------------------------------------------- */

/** `a:pPr`, `a:defPPr` or any `a:lvlNpPr` - one type under eleven names. */
export function parseParaProps(element: XElement, partName: string): ParaProps {
  return {
    marL: intOf(element, 'marL', partName),
    marR: intOf(element, 'marR', partName),
    indent: intOf(element, 'indent', partName),
    algn: enumOf(element, 'algn', ALIGN, partName),
    defTabSz: intOf(element, 'defTabSz', partName),
    rtl: boolOf(element, 'rtl', partName),
    eaLnBrk: boolOf(element, 'eaLnBrk', partName),
    fontAlgn: enumOf(element, 'fontAlgn', FONT_ALIGN, partName),
    latinLnBrk: boolOf(element, 'latinLnBrk', partName),
    hangingPunct: boolOf(element, 'hangingPunct', partName),
    lnSpc: parseSpacing(element, 'a:lnSpc', partName),
    spcBef: parseSpacing(element, 'a:spcBef', partName),
    spcAft: parseSpacing(element, 'a:spcAft', partName),
    defRPr: parseRunPropsChild(element, 'a:defRPr', partName),
    ...parseBullet(element, partName),
    node: element,
  };
}

/** The four bullet elements, of which at most one may appear. */
const BULLET_KINDS: readonly (readonly [string, BulletKind])[] = [
  ['a:buNone', 'none'],
  ['a:buChar', 'char'],
  ['a:buAutoNum', 'autonum'],
  ['a:buBlip', 'blip'],
];

/**
 * The bullet, as four independent slots plus the exclusive kind.
 *
 * Four rather than one because T5 measured them to merge separately: a level
 * that declares only `a:buFont` re-faces an inherited `a:buChar` and keeps
 * everything else. Folding them into a single value would make that impossible
 * to express, and the fold is what a reader writes first because the schema puts
 * the four kinds in one exclusive group.
 *
 * `a:buFontTx`, `a:buSzTx` and `a:buClrTx` are parsed as *values* rather than as
 * absences, because they are: PowerPoint writes nothing at all to mean "follow
 * the text", so the explicit element only ever appears in order to cancel
 * something a level above declared.
 */
function parseBullet(
  element: XElement,
  partName: string,
): Pick<ParaProps, 'buKind' | 'buChar' | 'buAutoNum' | 'buBlip' | 'buFont' | 'buSize' | 'buColor'> {
  let buKind: BulletKind | undefined;
  for (const [name, kind] of BULLET_KINDS) {
    if (firstChild(element, name) !== undefined) {
      if (buKind !== undefined) {
        throw new ModelError(
          'MODEL_TEXT_BULLET',
          `a:pPr declares both ${buKind} and ${kind} bullets, which are exclusive`,
          partName,
          'a:pPr',
        );
      }
      buKind = kind;
    }
  }

  const charElement = firstChild(element, 'a:buChar');
  const autoNum = firstChild(element, 'a:buAutoNum');
  const blip = firstChild(element, 'a:buBlip');
  const font = firstChild(element, 'a:buFont');
  const szPct = firstChild(element, 'a:buSzPct');
  const szPts = firstChild(element, 'a:buSzPts');
  const clr = firstChild(element, 'a:buClr');

  let buAutoNum: BulletAutoNum | undefined;
  if (autoNum !== undefined) {
    const type = attributeValue(autoNum, 'type');
    if (type === undefined) {
      throw new ModelError(
        'MODEL_TEXT_BULLET',
        'a:buAutoNum has no @type, which the schema requires',
        partName,
        'a:buAutoNum',
      );
    }
    buAutoNum = { type, startAt: intOf(autoNum, 'startAt', partName) };
  }

  let buBlip: string | undefined;
  if (blip !== undefined) {
    const embed = firstChild(blip, 'a:blip');
    buBlip = embed === undefined ? undefined : attributeValue(embed, 'r:embed');
    if (buBlip === undefined) {
      throw new ModelError(
        'MODEL_TEXT_BULLET',
        'a:buBlip has no a:blip/@r:embed to resolve',
        partName,
        'a:buBlip',
      );
    }
  }

  let buFont: BulletFont | undefined;
  if (firstChild(element, 'a:buFontTx') !== undefined) buFont = { kind: 'text' };
  else if (font !== undefined) buFont = { kind: 'typeface', value: typefaceOf(font, partName) };

  let buSize: BulletSize | undefined;
  if (firstChild(element, 'a:buSzTx') !== undefined) buSize = { kind: 'text' };
  else if (szPct !== undefined) {
    buSize = { kind: 'percent', value: requiredPercentage(szPct, partName, 'a:buSzPct') };
  } else if (szPts !== undefined) {
    buSize = { kind: 'points', value: requiredInt(szPts, partName, 'a:buSzPts') };
  }

  let buColor: BulletColor | undefined;
  if (firstChild(element, 'a:buClrTx') !== undefined) buColor = { kind: 'text' };
  else if (clr !== undefined) {
    const value = parseColorChild(clr);
    if (value === undefined || value === null) {
      throw new ModelError('MODEL_TEXT_BULLET', 'a:buClr holds no colour', partName, 'a:buClr');
    }
    buColor = { kind: 'color', value };
  }

  return {
    buKind,
    buChar: charElement === undefined ? undefined : (attributeValue(charElement, 'char') ?? ''),
    buAutoNum,
    buBlip,
    buFont,
    buSize,
    buColor,
  };
}

/** `@val` on a `a:buSzPct`, which the schema makes required. */
function requiredPercentage(element: XElement, partName: string, what: string): number {
  const raw = attributeValue(element, 'val');
  if (raw === undefined) {
    throw new ModelError('MODEL_TEXT_BULLET', `${what} has no @val`, partName, what);
  }
  return parsePercentage(raw);
}

/** `@val` on a `a:buSzPts`, in hundredths of a point. */
function requiredInt(element: XElement, partName: string, what: string): number {
  const value = intOf(element, 'val', partName);
  if (value === undefined) {
    throw new ModelError('MODEL_TEXT_BULLET', `${what} has no @val`, partName, what);
  }
  return value;
}

/* -------------------------------------------------------------------------- */
/* list styles                                                                */
/* -------------------------------------------------------------------------- */

const LEVEL_NAMES: readonly string[] = Array.from(
  { length: LEVELS },
  (_, i) => `a:lvl${String(i + 1)}pPr`,
);

/**
 * `a:lstStyle`, or a `p:txStyles` bucket, which is the same type unwrapped.
 *
 * The nine levels are independent: measured in 3.1, a layout placeholder
 * declaring only `a:lvl2pPr` changes the second level and leaves the first and
 * third to the master. So a level nobody declared stays `undefined` rather than
 * falling back to level one, and the resolver asks per level.
 */
export function parseListStyle(element: XElement, partName: string): ListStyle {
  const levels: (ParaProps | undefined)[] = [];
  for (const name of LEVEL_NAMES) {
    const level = firstChild(element, name);
    levels.push(level === undefined ? undefined : parseParaProps(level, partName));
  }
  const defPPr = firstChild(element, 'a:defPPr');
  return {
    levels,
    defPPr: defPPr === undefined ? undefined : parseParaProps(defPPr, partName),
    node: element,
  };
}

function parseListStyleChild(
  parent: XElement,
  qname: string,
  partName: string,
): ListStyle | undefined {
  const element = firstChild(parent, qname);
  return element === undefined ? undefined : parseListStyle(element, partName);
}

/** `p:txStyles` on a master. */
export function parseTextStyles(element: XElement, partName: string): TextStyles {
  return {
    title: parseListStyleChild(element, 'p:titleStyle', partName),
    body: parseListStyleChild(element, 'p:bodyStyle', partName),
    other: parseListStyleChild(element, 'p:otherStyle', partName),
    node: element,
  };
}

/** `p:defaultTextStyle` on `ppt/presentation.xml`, which is a `CT_TextListStyle`. */
export function parseDefaultTextStyle(
  presentation: XElement,
  partName: string,
): ListStyle | undefined {
  return parseListStyleChild(presentation, 'p:defaultTextStyle', partName);
}

/* -------------------------------------------------------------------------- */
/* paragraphs and runs                                                        */
/* -------------------------------------------------------------------------- */

/**
 * The paragraph's level, from `a:pPr/@lvl`.
 *
 * `ST_TextIndentLevelType` bounds it at 0..8 and PowerPoint still opens a file
 * outside the bound, so this clamps rather than throws - refusing to render a
 * deck PowerPoint renders is a worse answer than rendering it the way
 * PowerPoint does. A non-integer is a different matter and does throw: that is
 * a file saying something nobody can act on.
 */
function levelOf(pPr: XElement | undefined, partName: string): number {
  if (pPr === undefined) return 0;
  const lvl = intOf(pPr, 'lvl', partName);
  if (lvl === undefined) return 0;
  return Math.max(0, Math.min(LEVELS - 1, lvl));
}

function parseContent(element: XElement, partName: string): TextContent | null {
  if (element.qname === 'a:r') {
    const t = firstChild(element, 'a:t');
    return {
      kind: 'run',
      props: parseRunPropsChild(element, 'a:rPr', partName),
      text: t === undefined ? '' : textContent(t),
      node: element,
    };
  }
  if (element.qname === 'a:br') {
    return {
      kind: 'br',
      props: parseRunPropsChild(element, 'a:rPr', partName),
      node: element,
    };
  }
  if (element.qname === 'a:fld') {
    const id = attributeValue(element, 'id');
    if (id === undefined) {
      // `@id` is a required `ST_Guid` and PowerPoint may repair a file that
      // lost one. Inventing a replacement would be writing markup we did not
      // read, so this refuses instead.
      throw new ModelError('MODEL_TEXT_FIELD', 'a:fld has no @id', partName, 'a:fld');
    }
    const t = firstChild(element, 'a:t');
    return {
      kind: 'field',
      id,
      fieldType: attributeValue(element, 'type'),
      props: parseRunPropsChild(element, 'a:rPr', partName),
      text: t === undefined ? '' : textContent(t),
      node: element,
    };
  }
  return null;
}

function parseParagraph(element: XElement, partName: string): Paragraph {
  const pPr = firstChild(element, 'a:pPr');
  const content: TextContent[] = [];
  for (const child of childElements(element)) {
    const parsed = parseContent(child, partName);
    if (parsed !== null) content.push(parsed);
  }
  return {
    props: pPr === undefined ? undefined : parseParaProps(pPr, partName),
    level: levelOf(pPr, partName),
    content,
    endParaRPr: parseRunPropsChild(element, 'a:endParaRPr', partName),
    node: element,
  };
}

/** A `p:txBody` or an `a:txBody`. */
export function parseTextBody(element: XElement, partName: string): TextBody {
  return {
    bodyPr: parseBodyPropsChild(element, partName),
    lstStyle: parseListStyleChild(element, 'a:lstStyle', partName),
    paragraphs: childElements(element)
      .filter((child) => child.qname === 'a:p')
      .map((child) => parseParagraph(child, partName)),
    node: element,
  };
}

/** The `p:txBody` of a shape, when it has one. */
export function parseTextBodyChild(parent: XElement, partName: string): TextBody | undefined {
  const element = firstChild(parent, 'p:txBody') ?? firstChild(parent, 'a:txBody');
  return element === undefined ? undefined : parseTextBody(element, partName);
}
