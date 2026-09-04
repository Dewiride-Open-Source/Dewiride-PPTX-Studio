/**
 * Typed failures, as everywhere else in this repository.
 *
 * Two kinds of thing go wrong here and they want different treatment.
 *
 * A **broken package** - a slide that names no layout, a layout whose master is
 * missing, a master with no theme - is not a programming mistake. PowerPoint
 * repairs all three rather than refusing them (measured in 2.9, and after the
 * repair the binding is simply gone), and a viewer that throws on a file
 * PowerPoint opens is a viewer nobody can use. So those are *not* errors here:
 * the chain records what it found and `parent` is `null`. There is a code for
 * each all the same, because `validate` refuses to *write* one, and because a
 * caller that asks for the theme of a master that has none needs an answer with
 * a name.
 *
 * A **caller mistake** - resolving `phClr` outside a style invocation, asking
 * for a sheet that is not in the document - is an error and throws.
 */
export type ModelErrorCode =
  /** A part the chain needs is not in the package. */
  | 'MODEL_PART_MISSING'
  /** A part is present but its root element is not the one its rel promised. */
  | 'MODEL_PART_KIND'
  /** `ppt/presentation.xml` is missing, or the root rels do not name it. */
  | 'MODEL_NO_PRESENTATION'
  /** A slide has no `slideLayout` relationship, or a layout no `slideMaster`. */
  | 'MODEL_NO_PARENT'
  /** A master has no `theme` relationship. */
  | 'MODEL_NO_THEME'
  /** A `p:clrMap` or `a:overrideClrMapping` missing one of its twelve attributes. */
  | 'MODEL_CLRMAP_INCOMPLETE'
  /** A `p:clrMap` attribute naming something that is not one of the twelve slots. */
  | 'MODEL_CLRMAP_SLOT'
  /** A `p:ph/@idx` that is not a non-negative integer below 2^32. */
  | 'MODEL_PLACEHOLDER_IDX'
  /** An `a:off`/`a:ext`/`@rot` attribute that is not an integer. */
  | 'MODEL_XFRM_NUMBER'
  /** A style-matrix reference whose `@idx` does not parse. */
  | 'MODEL_STYLE_IDX'
  /** An `a:fontRef/@idx` outside `major`, `minor` and `none`. */
  | 'MODEL_FONT_COLLECTION'
  /** A theme with no `a:fmtScheme`, asked for a style-matrix entry. */
  | 'MODEL_NO_STYLE_MATRIX'
  /** A `a:fillStyleLst`/`a:lnStyleLst`/`a:bgFillStyleLst` with no entries at all. */
  | 'MODEL_STYLE_LIST_EMPTY'
  /** A sheet chain that returns to a sheet it has already visited. */
  | 'MODEL_SHEET_CYCLE'
  /** A sheet that does not belong to the document it was asked about. */
  | 'MODEL_FOREIGN_SHEET'
  /** `a:prstGeom` with no `@prst`, or an `a:custGeom` command missing a point. */
  | 'MODEL_GEOMETRY';

export class ModelError extends Error {
  override readonly name = 'ModelError';
  readonly code: ModelErrorCode;
  /** The part the failure is about, when there is one. */
  readonly partName: string | null;
  /** Whatever detail names the failure: an attribute value, a rel id, a slot. */
  readonly detail: string | null;

  constructor(
    code: ModelErrorCode,
    message: string,
    partName: string | null = null,
    detail: string | null = null,
  ) {
    super(message);
    this.code = code;
    this.partName = partName;
    this.detail = detail;
  }
}

export function isModelError(value: unknown): value is ModelError {
  return value instanceof ModelError;
}

export const MODEL_ERROR_CODES: readonly ModelErrorCode[] = [
  'MODEL_PART_MISSING',
  'MODEL_PART_KIND',
  'MODEL_NO_PRESENTATION',
  'MODEL_NO_PARENT',
  'MODEL_NO_THEME',
  'MODEL_CLRMAP_INCOMPLETE',
  'MODEL_CLRMAP_SLOT',
  'MODEL_PLACEHOLDER_IDX',
  'MODEL_XFRM_NUMBER',
  'MODEL_GEOMETRY',
  'MODEL_STYLE_IDX',
  'MODEL_FONT_COLLECTION',
  'MODEL_NO_STYLE_MATRIX',
  'MODEL_STYLE_LIST_EMPTY',
  'MODEL_SHEET_CYCLE',
  'MODEL_FOREIGN_SHEET',
];
