/** EMU in one inch, which is what turns an image resolution into a size. */
export const EMU_PER_INCH = 914400;

/**
 * What a DrawingML colour is, before anything has been decided about it.
 *
 * ## The union is lazy on purpose
 *
 * `EG_ColorChoice` has six members and they are not six spellings of a hex
 * triple. `a:hslClr` names a hue and a lightness; `a:schemeClr` names a slot in
 * a theme this colour has never seen; `a:sysClr` names a value that belongs to
 * the machine. Flattening all six to RGB at parse time is the same mistake
 * `xfrm === undefined` exists to avoid one layer up: it destroys the fact that
 * the deck said `accent1`, and with it the ability to swap a palette (7.6), to
 * tell the inspector where a colour came from (6.5), or to re-resolve when a
 * `clrMapOvr` changes. So a `Color` keeps the space it was written in, and
 * `resolveColor` is the only place it becomes a number.
 *
 * ## Values are in the units the file uses
 *
 * A percentage transform carries hundred-thousandths (`val="40000"` is 40%) and
 * an angle carries sixtieth-of-a-degree units (`val="7200000"` is 120 degrees),
 * because that is what the attribute says and converting early would put a
 * different number in the debugger from the one in the markup. `parsePercentage`
 * and `parseAngle` do the string half, including the `ST_Percentage` trap where
 * the same attribute may legally read `"40000"` or `"40%"`.
 */

/* -------------------------------------------------------------------------- */
/* transforms                                                                 */
/* -------------------------------------------------------------------------- */

/** Transforms whose `val` is a `ST_Percentage` family type: hundred-thousandths. */
export type PercentTransformOp =
  | 'tint'
  | 'shade'
  | 'alpha'
  | 'alphaOff'
  | 'alphaMod'
  | 'hueMod'
  | 'sat'
  | 'satOff'
  | 'satMod'
  | 'lum'
  | 'lumOff'
  | 'lumMod'
  | 'red'
  | 'redOff'
  | 'redMod'
  | 'green'
  | 'greenOff'
  | 'greenMod'
  | 'blue'
  | 'blueOff'
  | 'blueMod';

/** Transforms whose `val` is an angle: sixtieths of a degree. */
export type AngleTransformOp = 'hue' | 'hueOff';

/** Transforms with no attribute at all. */
export type FlagTransformOp = 'comp' | 'inv' | 'gray' | 'gamma' | 'invGamma';

export type ColorTransformOp = PercentTransformOp | AngleTransformOp | FlagTransformOp;

/**
 * One transform, in document order.
 *
 * Order is load-bearing and not a detail: on `#4472C4`, `lumMod 60%` then
 * `lumOff 40%` gives `8FAADC` and the reverse gives `517CC8`; `tint 50%` then
 * `shade 50%` gives `8D93A7` and the reverse gives `BEC2D1`. A renderer that
 * sorts these into a canonical order to make its own life easier is wrong on
 * one of every such pair, and PowerPoint's own "Lighter 40%" gallery entry is
 * the first pair.
 */
export type ColorTransform =
  | { readonly op: PercentTransformOp; readonly val: number }
  | { readonly op: AngleTransformOp; readonly val: number }
  | { readonly op: FlagTransformOp };

/** The 28 members of `EG_ColorTransform`, for a parser to test a tag against. */
export const COLOR_TRANSFORM_OPS: readonly ColorTransformOp[] = [
  'alpha',
  'alphaMod',
  'alphaOff',
  'blue',
  'blueMod',
  'blueOff',
  'comp',
  'gamma',
  'gray',
  'green',
  'greenMod',
  'greenOff',
  'hue',
  'hueMod',
  'hueOff',
  'inv',
  'invGamma',
  'lum',
  'lumMod',
  'lumOff',
  'red',
  'redMod',
  'redOff',
  'sat',
  'satMod',
  'satOff',
  'shade',
  'tint',
];

const ANGLE_OPS = new Set<string>(['hue', 'hueOff']);
const FLAG_OPS = new Set<string>(['comp', 'inv', 'gray', 'gamma', 'invGamma']);

/** Which of the three shapes a transform tag takes. Drives parsing, not resolution. */
export function transformShape(op: string): 'percent' | 'angle' | 'flag' | null {
  if (FLAG_OPS.has(op)) return 'flag';
  if (ANGLE_OPS.has(op)) return 'angle';
  return COLOR_TRANSFORM_OPS.includes(op as ColorTransformOp) ? 'percent' : null;
}

/* -------------------------------------------------------------------------- */
/* the theme                                                                  */
/* -------------------------------------------------------------------------- */

/** The twelve children of `a:clrScheme`, and the only things a `clrMap` maps *to*. */
export type SchemeSlot =
  | 'dk1'
  | 'lt1'
  | 'dk2'
  | 'lt2'
  | 'accent1'
  | 'accent2'
  | 'accent3'
  | 'accent4'
  | 'accent5'
  | 'accent6'
  | 'hlink'
  | 'folHlink';

export const SCHEME_SLOTS: readonly SchemeSlot[] = [
  'dk1',
  'lt1',
  'dk2',
  'lt2',
  'accent1',
  'accent2',
  'accent3',
  'accent4',
  'accent5',
  'accent6',
  'hlink',
  'folHlink',
];

/** The twelve attributes of `p:clrMap`, and the only names a map is consulted for. */
export type MappedColorName =
  | 'bg1'
  | 'tx1'
  | 'bg2'
  | 'tx2'
  | 'accent1'
  | 'accent2'
  | 'accent3'
  | 'accent4'
  | 'accent5'
  | 'accent6'
  | 'hlink'
  | 'folHlink';

export const MAPPED_COLOR_NAMES: readonly MappedColorName[] = [
  'bg1',
  'tx1',
  'bg2',
  'tx2',
  'accent1',
  'accent2',
  'accent3',
  'accent4',
  'accent5',
  'accent6',
  'hlink',
  'folHlink',
];

/**
 * Every value `a:schemeClr/@val` may take: seventeen, not twelve and not
 * sixteen.
 *
 * The four that trip people are `dk1`/`lt1`/`dk2`/`lt2`. They are *slot* names,
 * so a `schemeClr` naming one reaches the theme directly and the `clrMap` is not
 * consulted. `bg1`/`tx1`/`bg2`/`tx2` are *mapped* names for the same four slots
 * and do go through the map, which is how one master can invert a theme without
 * touching the theme. On the usual identity-ish map both spellings agree, which
 * is exactly why the bug survives to production.
 *
 * `phClr` is neither: it is the colour the enclosing `fillRef`/`lnRef`/`effectRef`
 * was invoked with, and it resolves from the context rather than from the theme.
 */
export type SchemeColorName = SchemeSlot | MappedColorName | 'phClr';

/** A parsed `a:clrScheme`. Each slot holds a colour, not a hex string - see below. */
export type ClrScheme = Readonly<Record<SchemeSlot, Color>>;

/** A parsed `p:clrMap` or `p:clrMapOvr`. All twelve attributes are required. */
export type ClrMap = Readonly<Record<MappedColorName, SchemeSlot>>;

/* -------------------------------------------------------------------------- */
/* the colour                                                                 */
/* -------------------------------------------------------------------------- */

/**
 * A colour exactly as the file wrote it: one of six bases, plus its transforms.
 *
 * A theme slot holds one of these rather than a hex string because that is what
 * the markup holds - the standard Office theme writes
 * `<a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>`, and a slot may
 * carry transforms of its own.
 */
export type Color =
  /** `a:srgbClr` - six hex digits. */
  | { readonly space: 'srgb'; readonly hex: string; readonly transforms: readonly ColorTransform[] }
  /** `a:scrgbClr` - three percentages, hundred-thousandths. */
  | {
      readonly space: 'scrgb';
      readonly r: number;
      readonly g: number;
      readonly b: number;
      readonly transforms: readonly ColorTransform[];
    }
  /** `a:hslClr` - hue in sixtieths of a degree, sat and lum in hundred-thousandths. */
  | {
      readonly space: 'hsl';
      readonly hue: number;
      readonly sat: number;
      readonly lum: number;
      readonly transforms: readonly ColorTransform[];
    }
  /** `a:schemeClr` - a name resolved against a theme and a colour map. */
  | {
      readonly space: 'scheme';
      readonly name: SchemeColorName;
      readonly transforms: readonly ColorTransform[];
    }
  /** `a:sysClr` - a machine colour, with the last value the writer saw. */
  | {
      readonly space: 'sys';
      readonly name: string;
      readonly lastClr: string | null;
      readonly transforms: readonly ColorTransform[];
    }
  /** `a:prstClr` - one of the 140 names in `ST_PresetColorVal`. */
  | {
      readonly space: 'prst';
      readonly name: string;
      readonly transforms: readonly ColorTransform[];
    };

/** A resolved colour. Channels are 0..1 and clamped; `a` is opacity, 1 is opaque. */
export interface Rgba {
  readonly r: number;
  readonly g: number;
  readonly b: number;
  readonly a: number;
}
