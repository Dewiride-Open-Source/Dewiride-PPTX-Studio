/**
 * The text cascade: one walk, two termini, and the origin of every answer.
 *
 * `resolveRun(context, paragraph, run, pick)` returns the first level that
 * declared the property asked for **and which level that was**, exactly as
 * `resolve` does for geometry and fills. The inspector's provenance chips,
 * per-property Reset, and knowing what a theme swap invalidates all fall out of
 * the same field, read by the same function the renderer reads.
 *
 * ## The walk
 *
 * 1. the run's own `a:rPr`
 * 2. the paragraph's `a:pPr/a:defRPr`
 * 3. the shape's `a:lstStyle`, at the paragraph's level
 * 4. each ancestor placeholder's `a:lstStyle`, up the chain 2.9 measured for
 *    geometry: `@idx` alone to the layout, the folded `@type` to the master
 *
 * ## The two termini, which are not one ladder
 *
 * A shape whose chain ends at a type with a bucket - `title` and `ctrTitle` to
 * `p:titleStyle`, the content types to `p:bodyStyle` - reads that bucket and
 * **stops**. A shape with no bucket reads `p:defaultTextStyle` and **stops**.
 *
 * Measured in 3.1, and it is where every implementation this project has read
 * goes wrong. Scored over the eight combinations of the three choices anyone
 * could make, this reading fits 37 of 37 probes and the reading the plan
 * specified - the shape's own `@type` picks the bucket, the header-and-footer
 * trio reads `p:otherStyle`, and so does a shape that is not a placeholder -
 * fits 13. Three of the four disagreements are visible on any deck:
 *
 * - **The `@type` at the end of the chain picks the bucket, not the shape's.**
 *   The first hop matches on `@idx` alone, so a slide `title` at idx 0 really
 *   does land on a layout `body` at idx 0, and when it does it reads
 *   `p:bodyStyle`. Measured both ways round and twice more through a `sldNum`.
 * - **`dt`, `ftr`, `sldNum` and `hdr` reach no bucket at all** and fall through
 *   to `p:defaultTextStyle`, which a body placeholder ignores even when it is
 *   there in the same package.
 * - **Nothing measurable reads `p:otherStyle`** - not those four, not a shape
 *   that is not a placeholder, on the slide, the layout or the master.
 *
 * ## Every level contributes only what it declares
 *
 * The walk is per property, not per level. Measured with five levels declaring
 * one property each: the run came back 24-point bold italic underlined and
 * struck through, where a nearest-level-takes-all reading gives an 18-point run
 * that is only struck through.
 */

import { ModelError } from '../errors.js';
import { inheritanceChain, normalizePlaceholder } from './placeholder.js';
import { themeOf } from './resolve.js';
import { resolveTypeface } from './typeface.js';
import { BUILTIN_TEXT_STYLES, TEXT_FLOOR, type BuiltinLevel } from '../builtin-text-styles.js';
import {
  type BulletAutoNum,
  type BulletColor,
  type BulletFont,
  type BulletKind,
  type BulletSize,
  LEVELS,
  type ListStyle,
  type Paragraph,
  type ParaProps,
  type RunProps,
  type TextContent,
  themeFontRef,
} from '../text.js';
import type { Origin, Resolved, Shape, Sheet } from '../types.js';

/* -------------------------------------------------------------------------- */
/* the bucket                                                                 */
/* -------------------------------------------------------------------------- */

/** `title` and `ctrTitle`. Everything else with a bucket has the body one. */
const TITLE_TYPES: ReadonlySet<string> = new Set(['title', 'ctrTitle']);

/**
 * The four types that reach no bucket.
 *
 * Measured, and against the plan, which assigned all four to `p:otherStyle`. A
 * `sldNum` placeholder in a package declaring `p:otherStyle` at 12 points and
 * `p:defaultTextStyle` at 20 came back at 20.
 */
const NO_BUCKET_TYPES: ReadonlySet<string> = new Set(['dt', 'ftr', 'sldNum', 'hdr']);

/** Which `p:txStyles` bucket a placeholder type reads, or `null` for none. */
export function bucketOf(type: string): 'title' | 'body' | null {
  if (TITLE_TYPES.has(type)) return 'title';
  if (NO_BUCKET_TYPES.has(type)) return null;
  return 'body';
}

/* -------------------------------------------------------------------------- */
/* the context                                                                */
/* -------------------------------------------------------------------------- */

/**
 * Everything the cascade needs that is not on the shape.
 *
 * `defaultTextStyle` comes from `ppt/presentation.xml` and is therefore a
 * property of the package rather than of any sheet, which is why it is passed
 * in rather than walked to. A caller that has a `Document` has it already;
 * `undefined` means the package declared none, which is a state PowerPoint
 * distinguishes and this must too.
 */
export interface TextContext {
  readonly sheet: Sheet;
  readonly shape: Shape;
  readonly defaultTextStyle: ListStyle | undefined;
}

/** One level of the walk: a list style, and what to call it if it answers. */
interface Level {
  readonly style: ListStyle | undefined;
  readonly origin: Origin;
  readonly sheet: Sheet | null;
  readonly shape: Shape | null;
}

/** Where the walk ends: the bucket the chain's last type selects, and whose. */
interface Terminus {
  readonly bucket: 'title' | 'body' | null;
  readonly master: Sheet | null;
}

function terminusOf(context: TextContext): Terminus {
  const chain = inheritanceChain(context.shape, context.sheet);
  // The type at the *end* of the chain, not the shape's own. A slide title that
  // matched a layout body reads p:bodyStyle - measured, both ways round.
  const placeholder = chain.at(-1)?.shape.placeholder ?? null;
  let master: Sheet | null = null;
  for (let sheet: Sheet | null = context.sheet; sheet !== null; sheet = sheet.parent) {
    if (sheet.kind === 'master') {
      master = sheet;
      break;
    }
  }
  return {
    bucket: placeholder === null ? null : bucketOf(normalizePlaceholder(placeholder).type),
    master,
  };
}

/**
 * The list styles a shape reads, nearest first, ending at whichever terminus it
 * reaches.
 *
 * Exported because it *is* the finding: a caller that wants to show the user
 * where a value could have come from - the inspector's Overrides panel, layout
 * compatibility scoring - needs the same list the resolver walks, not a second
 * opinion about it.
 */
export function textLevels(context: TextContext): readonly Level[] {
  const levels: Level[] = inheritanceChain(context.shape, context.sheet).map((link, index) => ({
    style: link.shape.text?.lstStyle,
    origin: index === 0 ? 'shape' : link.origin,
    sheet: link.sheet,
    shape: link.shape,
  }));

  const { bucket, master } = terminusOf(context);
  if (bucket === null) {
    levels.push({
      style: context.defaultTextStyle,
      origin: 'defaultTextStyle',
      sheet: null,
      shape: null,
    });
    return levels;
  }

  // The master's own p:txStyles. When it has none, PowerPoint substitutes its
  // built-in styles wholesale rather than merging them underneath - that is not
  // a `ListStyle`, so `floorOf` applies it and this level is simply absent.
  levels.push({
    style: master?.txStyles?.[bucket],
    origin: 'txStyles',
    sheet: master,
    shape: null,
  });
  return levels;
}

/**
 * The values left when every level has been asked and none of them said.
 *
 * Two quite different things, and which applies turns on one question: does the
 * master declare `p:txStyles` at all? If it does not, PowerPoint substitutes its
 * whole built-in set, and a body placeholder is 28 points with a hanging bullet
 * indent. If it does, there is no per-property backstop underneath it and what
 * remains is the floor - 18 points, no margin, no indent. Measured as two
 * separate probes precisely because the two are easy to conflate and differ on
 * every partially-specified master.
 */
export function floorOf(context: TextContext, level: number): BuiltinLevel {
  const { bucket, master } = terminusOf(context);
  if (master?.txStyles !== undefined) return TEXT_FLOOR;
  return BUILTIN_TEXT_STYLES[bucket ?? 'other'][clampLevel(level)] ?? TEXT_FLOOR;
}

function clampLevel(level: number): number {
  if (!Number.isInteger(level)) {
    throw new ModelError('MODEL_TEXT_LEVEL', `paragraph level ${String(level)} is not an integer`);
  }
  return Math.max(0, Math.min(LEVELS - 1, level));
}

/* -------------------------------------------------------------------------- */
/* resolving                                                                  */
/* -------------------------------------------------------------------------- */

function resolvedAt<T>(value: T, level: Level, explicit: boolean): Resolved<T> {
  return { value, origin: level.origin, explicit, sheet: level.sheet, shape: level.shape };
}

/**
 * A paragraph property, from the paragraph's own `a:pPr` or up the walk.
 *
 * `pick` returns `undefined` for "this level did not say" and anything else for
 * a value, including `null` - the same contract `resolve` uses, and for the same
 * reason: a level may legitimately declare a `null`, and conflating that with
 * silence is how an explicit value gets overwritten by an inherited one.
 */
export function resolveParagraph<T>(
  context: TextContext,
  paragraph: Paragraph,
  pick: (props: ParaProps) => T | undefined,
): Resolved<T> | undefined {
  if (paragraph.props !== undefined) {
    const own = pick(paragraph.props);
    if (own !== undefined) {
      return {
        value: own,
        origin: 'paragraph',
        explicit: true,
        sheet: context.sheet,
        shape: context.shape,
      };
    }
  }
  const level = clampLevel(paragraph.level);
  for (const entry of textLevels(context)) {
    const props = entry.style?.levels[level];
    if (props === undefined) continue;
    const value = pick(props);
    if (value !== undefined) return resolvedAt(value, entry, entry.origin === 'shape');
  }
  return undefined;
}

/**
 * A run property, from the run's own `a:rPr` and then every `a:defRPr` above it.
 *
 * The paragraph's `a:pPr/a:defRPr` is one level, not a special case: it sits
 * between the run and the shape's `a:lstStyle`, which the ladder pinned by
 * declaring a different size at each and reading back which one won.
 */
export function resolveRun<T>(
  context: TextContext,
  paragraph: Paragraph,
  run: TextContent | undefined,
  pick: (props: RunProps) => T | undefined,
): Resolved<T> | undefined {
  if (run?.props !== undefined) {
    const own = pick(run.props);
    if (own !== undefined) {
      return {
        value: own,
        origin: 'run',
        explicit: true,
        sheet: context.sheet,
        shape: context.shape,
      };
    }
  }
  if (paragraph.props?.defRPr !== undefined) {
    const value = pick(paragraph.props.defRPr);
    if (value !== undefined) {
      return {
        value,
        origin: 'paragraph',
        explicit: true,
        sheet: context.sheet,
        shape: context.shape,
      };
    }
  }
  const level = clampLevel(paragraph.level);
  for (const entry of textLevels(context)) {
    const defRPr = entry.style?.levels[level]?.defRPr;
    if (defRPr === undefined) continue;
    const value = pick(defRPr);
    if (value !== undefined) return resolvedAt(value, entry, entry.origin === 'shape');
  }
  return undefined;
}

/* -------------------------------------------------------------------------- */
/* the three properties with a measured floor                                 */
/* -------------------------------------------------------------------------- */

/**
 * The font size of a run, in hundredths of a point, always.
 *
 * The only resolver here that cannot return `undefined`: text has to be drawn at
 * some size, and the size when nothing declared one is measured rather than
 * chosen. `origin` is `builtin` when the master declares no `p:txStyles` and
 * `schemaDefault` when it does and the bucket said nothing - two different
 * numbers on the same slide, and the field says which was used.
 */
export function resolveSize(
  context: TextContext,
  paragraph: Paragraph,
  run: TextContent | undefined,
): Resolved<number> {
  const found = resolveRun(context, paragraph, run, (props) => props.sz);
  if (found !== undefined) return found;
  return floorResolved(context, paragraph, (level) => level.sz);
}

/**
 * The Latin typeface of a run, with any theme reference followed.
 *
 * The second resolver that cannot return `undefined`: a run has to be drawn in
 * some face, and the face when nothing named one is measured rather than
 * chosen - the theme's minor, on all six probes reaching a silent cascade.
 */
export function resolveLatinTypeface(
  context: TextContext,
  paragraph: Paragraph,
  run: TextContent | undefined,
): Resolved<string> {
  const scheme = themeOf(context.sheet)?.fonts ?? null;
  /** A name a source gave, with a theme reference followed where one can be. */
  const faceOf = (typeface: string): string | undefined => {
    if (themeFontRef(typeface) === null) return typeface;
    return scheme === null ? undefined : resolveTypeface(typeface, scheme);
  };

  const named = resolveRun(context, paragraph, run, (props) => props.latin);
  if (named !== undefined) {
    const face = faceOf(named.value.typeface);
    if (face !== undefined) return { ...named, value: face };
  }
  const floor = floorOf(context, paragraph.level);
  const reference = floor.typeface ?? TEXT_FLOOR.typeface;
  const face = reference === null ? undefined : faceOf(reference);
  if (face === undefined) {
    throw new ModelError(
      'MODEL_TEXT_TYPEFACE',
      'no source in the cascade names a typeface, and the theme names none either',
    );
  }
  return {
    value: face,
    origin: floor === TEXT_FLOOR ? 'schemaDefault' : 'builtin',
    explicit: false,
    sheet: null,
    shape: null,
  };
}

/** `@marL` in EMU. Measured to be 0 when nothing declares it, not 347663. */
export function resolveMarginLeft(context: TextContext, paragraph: Paragraph): Resolved<number> {
  const found = resolveParagraph(context, paragraph, (props) => props.marL);
  if (found !== undefined) return found;
  return floorResolved(context, paragraph, (level) => level.marL);
}

/** `@indent` in EMU. Measured to be 0 when nothing declares it, not -342900. */
export function resolveIndent(context: TextContext, paragraph: Paragraph): Resolved<number> {
  const found = resolveParagraph(context, paragraph, (props) => props.indent);
  if (found !== undefined) return found;
  return floorResolved(context, paragraph, (level) => level.indent);
}

/* -------------------------------------------------------------------------- */
/* bullets                                                                    */
/* -------------------------------------------------------------------------- */

/**
 * The bullet, merged the way T5 measured it: five slots, each on its own.
 *
 * The kind is one slot because the schema makes the four elements exclusive, and
 * the three decorations are three more because PowerPoint merges them
 * separately - a level declaring only `a:buFont` re-faces the character it
 * inherits and keeps its size and colour. Measured on eleven cases twice over,
 * once through a shape's own `a:lstStyle` and once through three hops of the
 * master chain, and the two agree.
 *
 * Every one of these returns `undefined` when nothing in the chain declares the
 * slot, which for the kind is the difference between "no bullet" and "nobody
 * said" - `a:buNone` resolves to `'none'` and an empty chain to `undefined`.
 * Nothing here supplies a default, because a paragraph with no bullet anywhere
 * in its chain genuinely has none.
 */
export function resolveBulletKind(
  context: TextContext,
  paragraph: Paragraph,
): Resolved<BulletKind> | undefined {
  return resolveParagraph(context, paragraph, (props) => props.buKind);
}

/** `a:buChar/@char`, as written. The symbol mapping belongs to the renderer. */
export function resolveBulletChar(
  context: TextContext,
  paragraph: Paragraph,
): Resolved<string> | undefined {
  return resolveParagraph(context, paragraph, (props) => props.buChar);
}

/** `a:buAutoNum`, type and start value together, since one is meaningless alone. */
export function resolveBulletAutoNum(
  context: TextContext,
  paragraph: Paragraph,
): Resolved<BulletAutoNum> | undefined {
  return resolveParagraph(context, paragraph, (props) => props.buAutoNum);
}

/** `a:buBlip/a:blip/@r:embed`, which the caller resolves against the part's rels. */
export function resolveBulletBlip(
  context: TextContext,
  paragraph: Paragraph,
): Resolved<string> | undefined {
  return resolveParagraph(context, paragraph, (props) => props.buBlip);
}

/**
 * `a:buFont` or `a:buFontTx`.
 *
 * Worth knowing before using it: this is **inert for an autonumber**. PowerPoint
 * draws a number in the first run's face whatever `a:buFont` says - measured on
 * twenty probes, including one where `buFont` named Wingdings and the run named
 * Courier New - and its own UI writes `buFont="+mj-lt"` on every list it
 * numbers. The resolver still reports it, because the file says it and 1.3 has
 * to write it back; the renderer is where it is ignored.
 */
export function resolveBulletFont(
  context: TextContext,
  paragraph: Paragraph,
): Resolved<BulletFont> | undefined {
  return resolveParagraph(context, paragraph, (props) => props.buFont);
}

/** `a:buSzPct`, `a:buSzPts` or `a:buSzTx`; a percentage is of the first run's size. */
export function resolveBulletSize(
  context: TextContext,
  paragraph: Paragraph,
): Resolved<BulletSize> | undefined {
  return resolveParagraph(context, paragraph, (props) => props.buSize);
}

/** `a:buClr` or `a:buClrTx`. Unlike the font, this one *does* reach an autonumber. */
export function resolveBulletColor(
  context: TextContext,
  paragraph: Paragraph,
): Resolved<BulletColor> | undefined {
  return resolveParagraph(context, paragraph, (props) => props.buColor);
}

function floorResolved(
  context: TextContext,
  paragraph: Paragraph,
  pick: (level: BuiltinLevel) => number,
): Resolved<number> {
  const floor = floorOf(context, paragraph.level);
  return {
    value: pick(floor),
    origin: floor === TEXT_FLOOR ? 'schemaDefault' : 'builtin',
    explicit: false,
    sheet: null,
    shape: null,
  };
}
