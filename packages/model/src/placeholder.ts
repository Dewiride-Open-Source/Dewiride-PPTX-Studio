/**
 * Which placeholder on the parent sheet a placeholder inherits from.
 *
 * ## The two hops do not use the same key, and neither uses the pair
 *
 * Every implementation this project has read matches on `(type, idx)`, usually
 * with a ladder of fallbacks for titles and for the header-and-footer trio. That
 * is not what PowerPoint does. Measured in 2.9 across 46 probes, each a slide
 * placeholder with no geometry of its own whose *rendered position* names the
 * layout placeholder it matched:
 *
 * - **Slide to layout matches on `@idx` alone.** `@type` is not consulted, not
 *   as a tiebreak and not as a fallback. A slide `title` at idx 0 takes a layout
 *   `body` at idx 0; a slide `body` at idx 2 takes a layout `sldNum` at idx 2; a
 *   slide `title` at idx 9 against a layout holding a title at idx 0 matches
 *   **nothing** and renders at the origin with zero size. Scored against the
 *   probes: `idx` alone 34/34, `(type, idx)` 13/34, family-and-idx 24/34, type
 *   alone 6/34.
 * - **Layout to master matches on type alone**, taking the *first* placeholder
 *   of that type in the master's shape tree - `@idx` is not even a tiebreak. A
 *   layout body at idx 5, against a master holding bodies at idx 1 and idx 5,
 *   takes the one at idx 1.
 *
 * The asymmetry is not arbitrary. Every stock PowerPoint layout numbers its
 * date, footer and slide-number placeholders 10, 11 and 12 while every stock
 * master numbers the same three 2, 3 and 4, so a second hop that looked at `idx`
 * would inherit nothing in the templates Office ships. And a master's
 * placeholder vocabulary is only `title`, `body`, `dt`, `ftr`, `sldNum` and
 * `hdr` - PowerPoint repairs any master carrying `ctrTitle`, `subTitle`, `obj`
 * or `pic` - so a second hop that looked at the type as written would find
 * nothing for the `ctrTitle` in layout 1 or the `obj` in layout 2.
 *
 * ## Each hop reads the placeholder of the sheet it is leaving
 *
 * A slide placeholder with no `@type` (so `obj`) matching a layout `ctrTitle`
 * reaches the master's **title**, not its body. So the type used for the second
 * hop is the *layout's*, not the slide's - which is what one would expect from a
 * chain of sheets, and is worth stating because the alternative reading gives a
 * different answer on exactly this case.
 */

import type { NormalPlaceholder, Placeholder, PlaceholderType, Shape, Sheet } from './types.js';

/**
 * The type of a `p:ph` that has none.
 *
 * `obj`, measured: a bare `<p:ph/>` on a slide reads back through PowerPoint's
 * object model as `ppPlaceholderObject`. ECMA's own default for the attribute
 * is `body`, and the two behave identically at the first hop because that hop
 * ignores the type - but they differ at the second, where `obj` folds to `body`
 * and `body` is already there. They agree, and it is worth knowing why rather
 * than by luck.
 */
export const DEFAULT_PLACEHOLDER_TYPE: PlaceholderType = 'obj';

/** The `@idx` of a `p:ph` that has none. */
export const DEFAULT_PLACEHOLDER_IDX = 0;

/** Supply the two defaults, so a comparison has something to compare. */
export function normalizePlaceholder(ph: Placeholder): NormalPlaceholder {
  return {
    type: ph.type ?? DEFAULT_PLACEHOLDER_TYPE,
    idx: ph.idx ?? DEFAULT_PLACEHOLDER_IDX,
  };
}

/** The four types that fold to `title` or stay as themselves on a master. */
const MASTER_TITLE: ReadonlySet<PlaceholderType> = new Set<PlaceholderType>(['title', 'ctrTitle']);

/** The types a master may carry, which are the only targets of the second hop. */
const MASTER_OWN: ReadonlySet<PlaceholderType> = new Set<PlaceholderType>([
  'body',
  'dt',
  'ftr',
  'sldNum',
  'hdr',
]);

/**
 * The type a placeholder presents when it asks a master for its parent.
 *
 * `ctrTitle` becomes `title`; everything a master may not carry - `subTitle`,
 * `obj`, `pic`, `chart`, `tbl`, `clipArt`, `dgm`, `media`, `sldImg` - becomes
 * `body`. Measured on `pic`, `obj`, `subTitle` and `ctrTitle`; the rest follow
 * because a master cannot hold them either and `body` is the only remaining
 * content type.
 */
export function masterPlaceholderType(type: PlaceholderType): PlaceholderType {
  if (MASTER_TITLE.has(type)) return 'title';
  if (MASTER_OWN.has(type)) return type;
  return 'body';
}

/** Every placeholder in a sheet's top-level shape tree, in document order. */
export function placeholders(sheet: Sheet): readonly Shape[] {
  return sheet.shapes.filter((shape) => shape.placeholder !== null);
}

/**
 * Slide to layout: the first placeholder whose `@idx` agrees.
 *
 * "First" matters only for a layout that has two placeholders at one index,
 * which PowerPoint accepts without repairing; the first in the shape tree is
 * the one that wins.
 */
export function matchInLayout(ph: Placeholder, layout: Sheet): Shape | null {
  const want = normalizePlaceholder(ph);
  for (const shape of layout.shapes) {
    if (shape.placeholder === null) continue;
    if (normalizePlaceholder(shape.placeholder).idx === want.idx) return shape;
  }
  return null;
}

/** Layout to master: the first placeholder of the folded type. */
export function matchInMaster(ph: Placeholder, master: Sheet): Shape | null {
  const want = masterPlaceholderType(normalizePlaceholder(ph).type);
  for (const shape of master.shapes) {
    if (shape.placeholder === null) continue;
    if (masterPlaceholderType(normalizePlaceholder(shape.placeholder).type) === want) return shape;
  }
  return null;
}

/** One link of the chain: a shape, the sheet it is on, and how it was reached. */
export interface ChainLink {
  readonly shape: Shape;
  readonly sheet: Sheet;
  /** `shape` for the first link, then `layoutPh`, then `masterPh`. */
  readonly origin: 'shape' | 'layoutPh' | 'masterPh';
}

/**
 * The chain a shape inherits along, nearest first.
 *
 * One entry for a non-placeholder, and for a placeholder that matched nothing.
 * Up to three for a slide placeholder that reached the master. Each hop asks
 * with the placeholder of the sheet it is leaving, which is why the layout's
 * `@type` and not the slide's decides the second hop.
 */
export function inheritanceChain(shape: Shape, sheet: Sheet): readonly ChainLink[] {
  const chain: ChainLink[] = [{ shape, sheet, origin: 'shape' }];
  if (shape.placeholder === null) return chain;

  let current = shape;
  let currentSheet = sheet;
  const seen = new Set<string>([sheet.partName]);

  while (currentSheet.parent !== null && current.placeholder !== null) {
    const parentSheet = currentSheet.parent;
    if (seen.has(parentSheet.partName)) break;
    seen.add(parentSheet.partName);

    const next =
      parentSheet.kind === 'master'
        ? matchInMaster(current.placeholder, parentSheet)
        : matchInLayout(current.placeholder, parentSheet);
    if (next === null) break;

    chain.push({
      shape: next,
      sheet: parentSheet,
      origin: parentSheet.kind === 'master' ? 'masterPh' : 'layoutPh',
    });
    current = next;
    currentSheet = parentSheet;
  }

  return chain;
}

/**
 * A placeholder that reached no parent and states no geometry of its own.
 *
 * PowerPoint renders one at the origin with zero width and height - it does not
 * refuse the file, does not drop the shape, and does not invent a rectangle.
 * Measured on four probes: a title where the layout has none, a body at an index
 * nobody has, and both against a layout with no placeholders at all.
 */
export const ORPHAN_RECT = { x: 0, y: 0, cx: 0, cy: 0 } as const;
