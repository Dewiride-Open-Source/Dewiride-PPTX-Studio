/**
 * `a:bodyPr`: the text frame, as the file states it.
 *
 * Every field is `T | undefined` for the reason the rest of the model is - a
 * frame that states no anchor is not a top-anchored frame, and 3.6 measured that
 * the difference is the whole of the placeholder cascade.
 */

import { attributeValue, firstChild } from '@pptx-studio/xml';
import type { XElement } from '@pptx-studio/xml';

import { ModelError } from '../errors.js';
import type {
  Autofit,
  BodyProps,
  HorzOverflow,
  TextAnchor,
  TextWrap,
  VertOverflow,
  VerticalText,
} from '../text.js';

const ANCHORS: readonly TextAnchor[] = ['t', 'ctr', 'b', 'just', 'dist'];
const VERTICALS: readonly VerticalText[] = [
  'horz',
  'vert',
  'vert270',
  'wordArtVert',
  'eaVert',
  'mongolianVert',
  'wordArtVertRtl',
];
const WRAPS: readonly TextWrap[] = ['none', 'square'];
const VERT_OVERFLOWS: readonly VertOverflow[] = ['overflow', 'ellipsis', 'clip'];
const HORZ_OVERFLOWS: readonly HorzOverflow[] = ['overflow', 'clip'];

function attrError(element: XElement, name: string, raw: string, why: string, part: string): never {
  throw new ModelError(
    'MODEL_TEXT_ATTR',
    `${element.qname}/@${name} is "${raw}", which is not ${why}`,
    part,
    name,
  );
}

function boolOf(element: XElement, name: string, part: string): boolean | undefined {
  const raw = attributeValue(element, name);
  if (raw === undefined) return undefined;
  if (raw === '1' || raw === 'true') return true;
  if (raw === '0' || raw === 'false') return false;
  return attrError(element, name, raw, 'a ST_OnOff', part);
}

function intOf(element: XElement, name: string, part: string): number | undefined {
  const raw = attributeValue(element, name);
  if (raw === undefined) return undefined;
  const value = Number(raw);
  if (!Number.isInteger(value)) return attrError(element, name, raw, 'an integer', part);
  return value;
}

function enumOf<T extends string>(
  element: XElement,
  name: string,
  allowed: readonly T[],
  part: string,
): T | undefined {
  const raw = attributeValue(element, name);
  if (raw === undefined) return undefined;
  if (!(allowed as readonly string[]).includes(raw)) {
    return attrError(element, name, raw, `one of ${allowed.join(', ')}`, part);
  }
  return raw as T;
}

/**
 * `a:noAutofit` / `a:normAutofit` / `a:spAutoFit`, whichever is present.
 *
 * The scale attributes are `ST_TextFontScalePercentOrPercentString`, which is
 * the same both-spellings grammar as every other percentage here; they are kept
 * as thousandths so `@pptx-studio/text` reads them the way it was measured.
 */
function parseAutofit(element: XElement, part: string): Autofit | undefined {
  if (firstChild(element, 'a:noAutofit') !== undefined) return { kind: 'none' };
  if (firstChild(element, 'a:spAutoFit') !== undefined) return { kind: 'shape' };
  const norm = firstChild(element, 'a:normAutofit');
  if (norm === undefined) return undefined;
  return {
    kind: 'normal',
    fontScale: percentOf(norm, 'fontScale', part),
    lnSpcReduction: percentOf(norm, 'lnSpcReduction', part),
  };
}

/** Thousandths of a percent, accepting both the `92500` and the `92.5%` spellings. */
function percentOf(element: XElement, name: string, part: string): number | undefined {
  const raw = attributeValue(element, name);
  if (raw === undefined) return undefined;
  const value = raw.endsWith('%') ? Number(raw.slice(0, -1)) * 1000 : Number(raw);
  if (!Number.isFinite(value)) return attrError(element, name, raw, 'a ST_Percentage', part);
  return value;
}

/** Parse an `a:bodyPr`. Lengths stay in EMU and angles in 60000ths of a degree. */
export function parseBodyProps(element: XElement, partName: string): BodyProps {
  return {
    anchor: enumOf(element, 'anchor', ANCHORS, partName),
    anchorCtr: boolOf(element, 'anchorCtr', partName),
    lIns: intOf(element, 'lIns', partName),
    tIns: intOf(element, 'tIns', partName),
    rIns: intOf(element, 'rIns', partName),
    bIns: intOf(element, 'bIns', partName),
    vert: enumOf(element, 'vert', VERTICALS, partName),
    wrap: enumOf(element, 'wrap', WRAPS, partName),
    vertOverflow: enumOf(element, 'vertOverflow', VERT_OVERFLOWS, partName),
    horzOverflow: enumOf(element, 'horzOverflow', HORZ_OVERFLOWS, partName),
    rot: intOf(element, 'rot', partName),
    upright: boolOf(element, 'upright', partName),
    numCol: intOf(element, 'numCol', partName),
    spcCol: intOf(element, 'spcCol', partName),
    rtlCol: boolOf(element, 'rtlCol', partName),
    spcFirstLastPara: boolOf(element, 'spcFirstLastPara', partName),
    compatLnSpc: boolOf(element, 'compatLnSpc', partName),
    fromWordArt: boolOf(element, 'fromWordArt', partName),
    forceAA: boolOf(element, 'forceAA', partName),
    autofit: parseAutofit(element, partName),
    node: element,
  };
}

/** The `a:bodyPr` of a `p:txBody`, when it has one. */
export function parseBodyPropsChild(parent: XElement, partName: string): BodyProps | undefined {
  const element = firstChild(parent, 'a:bodyPr');
  return element === undefined ? undefined : parseBodyProps(element, partName);
}
