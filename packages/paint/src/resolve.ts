import { applyTransforms } from './apply.js';
import { PaintError } from './errors.js';
import { PRESET_COLORS } from './preset-colors.js';
import { SYSTEM_COLORS } from './sys-colors.js';
import { clamp01, fromLinear, hslToRgb, parseHex, toByte, toHex } from './transfer.js';
import {
  MAPPED_COLOR_NAMES,
  type ClrMap,
  type ClrScheme,
  type Color,
  type MappedColorName,
  type Rgba,
  type SchemeColorName,
  type SchemeSlot,
} from './types.js';

/**
 * From a `Color` to a number, and the four things that have to be supplied
 * before that is possible.
 *
 * A colour is not self-contained. `accent1` needs a theme *and* a colour map;
 * `phClr` needs the style invocation it sits inside; `windowText` needs a
 * machine. Rather than let any of those default silently to black - which is the
 * failure mode where every themed shape on a slide renders correctly-shaped and
 * completely wrong - each one that is missing throws with its own code.
 */
export interface ColorContext {
  /** The theme's `a:clrScheme`. Required by any `a:schemeClr`. */
  readonly scheme?: ClrScheme;
  /**
   * The `p:clrMap` in force, from the master or overridden by a layout or slide.
   * Required by a `schemeClr` naming one of the twelve mapped names.
   */
  readonly map?: ClrMap;
  /**
   * The colour the enclosing `a:fillRef`/`a:lnRef`/`a:effectRef` was invoked
   * with, which is what `schemeClr val="phClr"` means. Supplied by the style
   * resolver in 2.9; there is nothing sensible to default it to.
   */
  readonly phClr?: Rgba;
  /**
   * Values for `a:sysClr` when the element has no `@lastClr`. Defaults to the
   * table in `sys-colors.ts`, which is one Windows machine's light theme - see
   * the long note there about why `@lastClr` wins here and does not in
   * PowerPoint.
   */
  readonly systemColors?: Readonly<Record<string, string>>;
}

const MAPPED: ReadonlySet<string> = new Set<string>(MAPPED_COLOR_NAMES);

/** Is this one of the twelve names a `p:clrMap` has an attribute for? */
function isMapped(name: SchemeColorName): name is MappedColorName {
  return MAPPED.has(name);
}

/**
 * Which theme slot a `a:schemeClr/@val` names.
 *
 * `dk1`/`lt1`/`dk2`/`lt2` are slot names and reach the theme directly; the
 * twelve mapped names go through the map. Measured in 2.6 against a master whose
 * map was deliberately crossed - `bg1="dk1"`, `accent1="accent2"` - because on
 * the identity map every deck ships with, the two readings give the same answer
 * and a resolver that runs `dk1` through the map looks correct. Under the
 * crossed map `accent1` painted `accent2`'s colour and `bg1` painted `dk1`'s,
 * both as this function says.
 *
 * Returns `null` for `phClr`, which names no slot.
 */
export function mapSchemeName(name: SchemeColorName, map: ClrMap | undefined): SchemeSlot | null {
  if (name === 'phClr') return null;
  if (!isMapped(name)) return name;
  const mapped = map?.[name];
  if (mapped !== undefined) return mapped;
  // `bg1`/`tx1`/`bg2`/`tx2` have no meaning without a map, but the other eight
  // mapped names are identity in every map anyone writes, so falling back to
  // the same name is the useful answer rather than a throw.
  if (name === 'bg1') return 'lt1';
  if (name === 'tx1') return 'dk1';
  if (name === 'bg2') return 'lt2';
  if (name === 'tx2') return 'dk2';
  return name;
}

function baseRgb(color: Color, ctx: ColorContext, seen: Set<SchemeSlot>): Rgba {
  switch (color.space) {
    case 'srgb': {
      const rgb = parseHex(color.hex);
      if (rgb === null) {
        throw new PaintError(
          'COLOR_HEX',
          `srgbClr val="${color.hex}" is not six hex digits`,
          color.hex,
        );
      }
      return { r: rgb[0], g: rgb[1], b: rgb[2], a: 1 };
    }

    case 'scrgb':
      // **Linear light, not sRGB percentages.** Measured 14/14 in 2.6, and it is
      // the single largest divergence in colour: `r="50000"` is `BC`, not `80`.
      // The name is the clue and it is the one nobody follows - "scRGB" is a
      // linear-light space, and reading the attributes as plain percentages is
      // 61 units out of 255 wrong at the midpoint of every channel.
      return {
        r: fromLinear(clamp01(color.r / 100000)),
        g: fromLinear(clamp01(color.g / 100000)),
        b: fromLinear(clamp01(color.b / 100000)),
        a: 1,
      };

    case 'hsl': {
      // The same bi-hexcone the `lumMod` family uses, so a deck can express a
      // colour either way and get the same answer. 16/16 exact except at the two
      // ties `toByte` documents.
      const [r, g, b] = hslToRgb({
        h: color.hue / 60000,
        s: color.sat / 100000,
        l: color.lum / 100000,
      });
      return { r: clamp01(r), g: clamp01(g), b: clamp01(b), a: 1 };
    }

    case 'sys': {
      // `@lastClr` first - see `sys-colors.ts`. This deliberately differs from
      // what PowerPoint does on a machine that has a system palette, because a
      // browser does not have one and the author's cached value is closer to
      // what the deck looked like than the viewer's Windows theme would be.
      const table = ctx.systemColors ?? SYSTEM_COLORS;
      const value = color.lastClr ?? table[color.name];
      if (value === undefined) {
        throw new PaintError(
          'COLOR_SYS_UNKNOWN',
          `sysClr val="${color.name}" has no lastClr and no value in the system colour table`,
          color.name,
        );
      }
      const rgb = parseHex(value);
      if (rgb === null) {
        throw new PaintError('COLOR_HEX', `sysClr resolved to "${value}"`, value);
      }
      return { r: rgb[0], g: rgb[1], b: rgb[2], a: 1 };
    }

    case 'prst': {
      const value = PRESET_COLORS[color.name];
      if (value === undefined) {
        throw new PaintError(
          'COLOR_PRST_UNKNOWN',
          `prstClr val="${color.name}" is not one of the 147 preset colour names`,
          color.name,
        );
      }
      const rgb = parseHex(value)!;
      return { r: rgb[0], g: rgb[1], b: rgb[2], a: 1 };
    }

    case 'scheme': {
      if (color.name === 'phClr') {
        if (ctx.phClr === undefined) {
          throw new PaintError(
            'COLOR_NO_PHCLR',
            'schemeClr val="phClr" outside a style invocation: no placeholder colour in context',
            'phClr',
          );
        }
        return ctx.phClr;
      }
      if (ctx.scheme === undefined) {
        throw new PaintError(
          'COLOR_NO_SCHEME',
          `schemeClr val="${color.name}" needs a clrScheme, and none was supplied`,
          color.name,
        );
      }
      const slot = mapSchemeName(color.name, ctx.map);
      // The cast is deliberate. `ClrScheme` promises all twelve slots, so the
      // type says this cannot be missing; a parser that built one from a theme
      // with eleven children disagrees, and this is where that shows up as a
      // named error instead of `undefined.transforms`.
      const slotColor = slot === null ? undefined : (ctx.scheme[slot] as Color | undefined);
      if (slot === null || slotColor === undefined) {
        throw new PaintError(
          'COLOR_SCHEME_NAME',
          `"${color.name}" is not a scheme colour name, or the clrScheme has no such slot`,
          color.name,
        );
      }
      if (seen.has(slot)) {
        // A theme whose `dk1` is `<a:schemeClr val="tx1"/>` and whose map sends
        // `tx1` back to `dk1`. Legal markup, infinite regress, and something a
        // hand-edited theme can produce.
        throw new PaintError(
          'COLOR_SCHEME_CYCLE',
          `clrScheme slot "${slot}" resolves to itself`,
          slot,
        );
      }
      seen.add(slot);
      // The slot's own colour, with its own transforms, before this colour's.
      return resolveWith(slotColor, ctx, seen);
    }
  }
}

function resolveWith(color: Color, ctx: ColorContext, seen: Set<SchemeSlot>): Rgba {
  return applyTransforms(baseRgb(color, ctx, seen), color.transforms);
}

/**
 * Resolve a colour to numbers: the base, then its transforms in document order.
 *
 * Throws `PaintError` when the context is missing something the colour needs.
 * Nothing here returns black on failure, because a themed deck that renders
 * entirely black is the single hardest colour bug to notice from a screenshot
 * and the easiest to notice from an exception.
 */
export function resolveColor(color: Color, ctx: ColorContext = {}): Rgba {
  return resolveWith(color, ctx, new Set());
}

/** Six upper-case hex digits. Alpha is dropped, because the format has nowhere to put it. */
export function toHexColor(color: Rgba): string {
  return toHex(color.r, color.g, color.b);
}

/**
 * A CSS colour.
 *
 * `#RRGGBB` when opaque, `rgba(...)` when not - rather than `#RRGGBBAA`, which
 * SVG presentation attributes in older renderers do not accept. A renderer that
 * would rather carry opacity separately has `Rgba.a` and should: `fill-opacity`
 * composites the same way and survives a fill being replaced.
 */
export function toCss(color: Rgba): string {
  if (color.a >= 1) return `#${toHexColor(color)}`;
  const r = toByte(color.r);
  const g = toByte(color.g);
  const b = toByte(color.b);
  return `rgba(${String(r)}, ${String(g)}, ${String(b)}, ${String(Math.round(color.a * 1000) / 1000)})`;
}
