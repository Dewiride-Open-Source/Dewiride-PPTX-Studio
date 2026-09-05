/**
 * The document model, and the one property that makes the whole product
 * possible: **`undefined` means "this sheet did not say".**
 *
 * Every visual property below is `T | undefined`, and absence is a fact the
 * file states rather than a gap to be filled in. PowerPoint writes a slide
 * placeholder with no `a:xfrm` at all - measured in 2.9 on its own output,
 * where a title that was typed into but not moved still has none, and one
 * nudged by a single point suddenly has the full resolved rectangle baked in -
 * so a parser that defaults geometry at parse time destroys the only record of
 * which shapes are still bound to their layout. That record is what Change
 * Layout reads, what the inspector's provenance chips read, and what tells a
 * theme swap which shapes to invalidate.
 *
 * So nothing here is resolved. `resolve.ts` walks the chain, `style.ts` reaches
 * the theme, and both report *where* the answer came from.
 */

import type { ClrMap, ClrScheme, Color, Effect, Fill, Line } from '@pptx-studio/paint';
import type { ShapeGeometry } from './parse/geometry.js';
import type { TextBody, TextStyles } from './text.js';
import type { XElement } from '@pptx-studio/xml';

/* -------------------------------------------------------------------------- */
/* geometry                                                                   */
/* -------------------------------------------------------------------------- */

/** `a:off` and `a:ext` together, in EMU, plus the three transform attributes. */
export interface Xfrm {
  readonly x: number;
  readonly y: number;
  readonly cx: number;
  readonly cy: number;
  /** `@rot`, sixtieths of a degree. Zero when the attribute was absent. */
  readonly rot: number;
  readonly flipH: boolean;
  readonly flipV: boolean;
  /** `a:chOff`/`a:chExt`, present only on a group's `a:xfrm`. */
  readonly child: {
    readonly x: number;
    readonly y: number;
    readonly cx: number;
    readonly cy: number;
  } | null;
}

/* -------------------------------------------------------------------------- */
/* placeholders                                                               */
/* -------------------------------------------------------------------------- */

/**
 * `ST_PlaceholderType`. All nineteen, as written.
 *
 * Not all nineteen are legal everywhere. A slide **master** accepts only
 * `title`, `body`, `dt`, `ftr`, `sldNum` and `hdr`: a master carrying
 * `ctrTitle`, `subTitle`, `obj` or `pic` is repaired on open, measured in 2.9.
 * That is not a curiosity - it is the reason the layout-to-master hop has to
 * fold `ctrTitle` into `title` and every content type into `body`, because
 * otherwise the stock "Title and Content" layout could inherit nothing from the
 * stock master.
 */
export type PlaceholderType =
  | 'title'
  | 'body'
  | 'ctrTitle'
  | 'subTitle'
  | 'obj'
  | 'chart'
  | 'tbl'
  | 'clipArt'
  | 'dgm'
  | 'media'
  | 'sldImg'
  | 'pic'
  | 'sldNum'
  | 'hdr'
  | 'ftr'
  | 'dt';

export const PLACEHOLDER_TYPES: readonly PlaceholderType[] = [
  'title',
  'body',
  'ctrTitle',
  'subTitle',
  'obj',
  'chart',
  'tbl',
  'clipArt',
  'dgm',
  'media',
  'sldImg',
  'pic',
  'sldNum',
  'hdr',
  'ftr',
  'dt',
];

/** The six a slide master may carry. Measured; see `PlaceholderType`. */
export const MASTER_PLACEHOLDER_TYPES: readonly PlaceholderType[] = [
  'title',
  'body',
  'dt',
  'ftr',
  'sldNum',
  'hdr',
];

/** `@sz`, a sizing hint. Recorded, and not part of matching - measured. */
export type PlaceholderSize = 'full' | 'half' | 'quarter';

export interface Placeholder {
  /**
   * `@type` as written, or `null` when the attribute was absent.
   *
   * `null` is kept rather than defaulted, because "the file said `obj`" and
   * "the file said nothing" are different facts and only the second one is safe
   * to rewrite. `normalizePlaceholder` supplies the default when a comparison
   * needs one.
   */
  readonly type: PlaceholderType | null;
  /** `@idx` as written, or `null` when absent. Defaults to 0 when compared. */
  readonly idx: number | null;
  readonly size: PlaceholderSize | null;
  /** `@orient`. Recorded, and not part of matching. */
  readonly orient: 'horz' | 'vert' | null;
  readonly hasCustomPrompt: boolean;
}

/** A placeholder with its defaults supplied: what the matcher compares. */
export interface NormalPlaceholder {
  readonly type: PlaceholderType;
  readonly idx: number;
}

/* -------------------------------------------------------------------------- */
/* shapes                                                                     */
/* -------------------------------------------------------------------------- */

/** One of the six things a `p:spTree` may contain. */
export type ShapeKind = 'sp' | 'pic' | 'grpSp' | 'graphicFrame' | 'cxnSp' | 'contentPart';

/** `a:lnRef`, `a:fillRef` and `a:effectRef`: an index into the style matrix. */
export interface StyleRef {
  /**
   * `@idx`. Zero means none; 1..999 index the main list; 1001 and up index the
   * background list, offset by 1000. Measured on both `a:fillRef` and
   * `p:bgRef`, which turn out to be the same rule over the same six entries.
   */
  readonly idx: number;
  /** The colour `phClr` takes inside the entry this reference names. */
  readonly color: Color | null;
}

/** `a:fontRef`. `@idx` is a collection name, never a number. */
export interface FontRef {
  readonly idx: 'major' | 'minor' | 'none';
  readonly color: Color | null;
}

/** `p:style`: four references into the theme's `a:fmtScheme` and `a:fontScheme`. */
export interface ShapeStyle {
  readonly lnRef: StyleRef;
  readonly fillRef: StyleRef;
  readonly effectRef: StyleRef;
  readonly fontRef: FontRef;
}

/**
 * A shape, exactly as its part wrote it.
 *
 * Everything optional here is optional *in the file*, and the resolver is the
 * only thing entitled to fill any of it in.
 */
export interface Shape {
  readonly kind: ShapeKind;
  /** `p:cNvPr/@id`. Unique within a part, and may repeat across parts. */
  readonly cNvPrId: number;
  readonly name: string;
  /** `p:cNvPr/@descr`, the alt text. */
  readonly descr: string | null;
  readonly hidden: boolean;
  readonly placeholder: Placeholder | null;
  readonly xfrm: Xfrm | undefined;
  readonly fill: Fill | undefined;
  readonly line: Line | undefined;
  readonly effects: readonly Effect[] | undefined;
  readonly style: ShapeStyle | undefined;
  /** `a:prstGeom/@prst`, or `null` for a `custGeom` or no geometry at all. */
  readonly prstGeom: string | undefined;
  /**
   * `a:prstGeom` or `a:custGeom`, parsed far enough to draw.
   *
   * `undefined` when the shape declares neither, which is the ordinary state of
   * a slide placeholder: it inherits its outline from its layout exactly as it
   * inherits its rectangle, and `resolve` is what asks.
   */
  readonly geometry: ShapeGeometry | undefined;
  /**
   * `p:txBody`, when the shape has one.
   *
   * Present on every `p:sp` PowerPoint writes, including the ones with no text
   * in them, and absent from a `p:pic` or a `p:graphicFrame`. Its `a:lstStyle`
   * is one level of the text cascade; the paragraphs are the text itself.
   */
  readonly text: TextBody | undefined;
  /** Children, for a `grpSp`. Empty for everything else. */
  readonly children: readonly Shape[];
  /** The element this was read from. Edits go here, never to the fields above. */
  readonly node: XElement;
}

/* -------------------------------------------------------------------------- */
/* the background                                                             */
/* -------------------------------------------------------------------------- */

/**
 * `p:bg`, which is one of two quite different things.
 *
 * `p:bgPr` states a fill. `p:bgRef` names an entry of the theme's style matrix
 * and the colour to invoke it with - the same mechanism as `a:fillRef` on a
 * shape, over the same table, with the same 1000-offset. Both were measured
 * against the same six theme entries in 2.9 and agreed on all six.
 */
export type Background =
  | { readonly kind: 'fill'; readonly fill: Fill; readonly effects: readonly Effect[] | undefined }
  | { readonly kind: 'ref'; readonly ref: StyleRef };

/* -------------------------------------------------------------------------- */
/* the theme                                                                  */
/* -------------------------------------------------------------------------- */

/** `a:fmtScheme`: four lists of style entries, held as the elements they were. */
export interface FormatScheme {
  readonly name: string | null;
  /** `a:fillStyleLst`. Each entry is a whole `EG_FillProperties` element. */
  readonly fillStyles: readonly Fill[];
  /** `a:lnStyleLst`. Each entry is a whole `a:ln`. */
  readonly lineStyles: readonly Line[];
  /** `a:effectStyleLst`. Each entry's `a:effectLst`, flattened. */
  readonly effectStyles: readonly (readonly Effect[])[];
  /** `a:bgFillStyleLst`, which `@idx` 1001 and up reach. */
  readonly bgFillStyles: readonly Fill[];
}

/**
 * `a:fontScheme`: the two collections `+mj-` and `+mn-` name.
 *
 * Three scripts each, because `a:rPr` has four typeface slots and three of them
 * take a theme reference - `+mn-ea` on an `a:ea` is how a CJK run follows the
 * theme. The `a:font` script-tag list under each collection is 3.7's business
 * and is not modelled here.
 */
export interface FontCollection {
  /** `a:latin/@typeface`. An empty string means the theme states none. */
  readonly latin: string | null;
  readonly ea: string | null;
  readonly cs: string | null;
}

export interface FontScheme {
  readonly name: string | null;
  readonly major: FontCollection;
  readonly minor: FontCollection;
}

export interface Theme {
  readonly partName: string;
  readonly name: string | null;
  readonly scheme: ClrScheme;
  readonly fonts: FontScheme;
  readonly format: FormatScheme | null;
  readonly node: XElement;
}

/* -------------------------------------------------------------------------- */
/* sheets                                                                     */
/* -------------------------------------------------------------------------- */

export type SheetKind = 'slide' | 'layout' | 'master';

/**
 * `p:clrMapOvr`, which has exactly two forms and one surprise.
 *
 * `a:masterClrMapping` means "use the map already in force". The surprise,
 * measured in 2.9, is what "already in force" means for a slide: it is the
 * **layout's** map, not the master's. A layout carrying an
 * `a:overrideClrMapping` changes the colours of every slide bound to it that
 * says `masterClrMapping`, and a slide's own override beats its layout's. The
 * element's name says master and it means parent.
 */
export type ColorMapOverride =
  { readonly kind: 'inherit' } | { readonly kind: 'override'; readonly map: ClrMap };

/**
 * A slide, a layout or a master, with its parent already resolved.
 *
 * The parent link comes from **this part's own `.rels`** and from nowhere else.
 * Not from `p:sldLayoutIdLst`, not from `p:sldMasterIdLst`, and not from the
 * part's number in its folder: a two-master deck saved by PowerPoint numbers its
 * layouts `slideLayout1..22` in one flat folder, and which master owns which is
 * recorded only in the masters' relationship parts. The list elements exist to
 * give each sheet an id and an order in the UI, not to bind anything.
 */
export interface Sheet {
  readonly kind: SheetKind;
  readonly partName: string;
  /** `p:cSld/@name`. */
  readonly name: string | null;
  /** `p:sldLayout/@type`, one of `ST_SlideLayoutType`. Layouts only. */
  readonly layoutType: string | null;
  /** `p:sldLayout/@matchingName`. Layouts only, and only when written. */
  readonly matchingName: string | null;
  readonly shapes: readonly Shape[];
  readonly background: Background | undefined;
  /** `p:clrMap`, on a master. `undefined` everywhere else. */
  readonly clrMap: ClrMap | undefined;
  /**
   * `p:txStyles`, on a master. `undefined` everywhere else, and meaningfully so.
   *
   * A master that declares none does not get a partial one: PowerPoint
   * substitutes its whole built-in set, measured in 3.1. So the difference
   * between `undefined` and a `p:txStyles` whose buckets say nothing is the
   * difference between a 28-point body placeholder and an 18-point one.
   */
  readonly txStyles: TextStyles | undefined;
  /** `p:clrMapOvr`, on a layout or a slide. */
  readonly clrMapOvr: ColorMapOverride | undefined;
  /** `@showMasterSp`, defaulting to true. Slides and layouts. */
  readonly showMasterShapes: boolean;
  /** The layout of a slide, the master of a layout, `null` for a master and for
   *  a sheet whose binding is missing - which PowerPoint repairs rather than
   *  refuses, so it has to be representable. */
  readonly parent: Sheet | null;
  /** The theme, on a master. Every other sheet reaches it through `parent`. */
  readonly theme: Theme | null;
  readonly node: XElement;
}

/* -------------------------------------------------------------------------- */
/* resolution                                                                 */
/* -------------------------------------------------------------------------- */

/**
 * Where a resolved value came from.
 *
 * 2.9 declared this union in advance, guessing at the members the text cascade
 * would need. 3.1 measured them, and the guess was wrong in one place:
 * `themeObjDefaults` is gone, because `a:objectDefaults/a:spDef/a:lstStyle` is
 * not a source. A package declaring a size there and nowhere else resolves to
 * the built-in default, so the member named a level that never fires - see
 * `corpus/ground-truth/text-cascade.json`.
 */
export type Origin =
  /** The run's own `a:rPr`. */
  | 'run'
  /** The paragraph's `a:pPr`, or its `a:defRPr`. */
  | 'paragraph'
  /** The shape itself: its `p:spPr`, or its own `a:lstStyle`. */
  | 'shape'
  | 'layoutPh'
  | 'masterPh'
  /** The master's `p:txStyles`, in the bucket the chain's last type selects. */
  | 'txStyles'
  /** `p:defaultTextStyle`, which only a shape with no bucket reaches. */
  | 'defaultTextStyle'
  /** PowerPoint's own text styles, substituted for a master that declares none. */
  | 'builtin'
  | 'theme'
  | 'schemaDefault';

/**
 * A value, where it came from, and whether the thing that owns it said so.
 *
 * `explicit` is not the same as `origin === 'shape'`: a layout placeholder that
 * declares a fill is explicit *about that fill*, and the slide that inherits it
 * is not. The inspector's hollow-versus-filled dot is exactly this bit.
 */
export interface Resolved<T> {
  readonly value: T;
  readonly origin: Origin;
  readonly explicit: boolean;
  /** The sheet the value was found on, when one was involved. */
  readonly sheet: Sheet | null;
  /** The shape the value was found on, when one was involved. */
  readonly shape: Shape | null;
}
