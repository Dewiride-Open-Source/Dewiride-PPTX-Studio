/**
 * The generic resolver: one function, and four features fall out of it.
 *
 * `resolve(shape, sheet, pick)` walks the inheritance chain and returns the
 * first level that declared the thing asked for, **with a record of which level
 * that was**. Carrying the origin costs one field and buys the inspector's
 * provenance chips, per-property Reset, layout-compatibility scoring, and
 * correct invalidation on a theme swap - all reading the same function the
 * renderer reads, so none of them can drift from what is on the screen.
 *
 * What descends is not a guess. Measured in 2.9, on a slide placeholder that
 * declares nothing: an `spPr` solid fill, a `p:style` and an `a:ln` all arrive
 * from the layout placeholder, and all three arrive from the *master*
 * placeholder when the layout placeholder declares nothing either. Geometry is
 * not special; it travels the same chain as everything else.
 */

import type { ClrMap, ClrScheme, ColorContext, Rgba } from '@pptx-studio/paint';

import { ModelError } from '../errors.js';
import { inheritanceChain } from './placeholder.js';
import type { Resolved, Shape, Sheet, Theme, Xfrm } from '../types.js';

/**
 * The first level of the chain that declared something.
 *
 * `pick` must return `undefined` for "this level did not say" and anything else
 * for a value, including `null` - a shape may legitimately declare a `null`, and
 * conflating that with silence is how an explicit `a:noFill` gets overwritten by
 * an inherited colour.
 */
export function resolve<T>(
  shape: Shape,
  sheet: Sheet,
  pick: (shape: Shape) => T | undefined,
): Resolved<T> | undefined {
  const chain = inheritanceChain(shape, sheet);
  for (const [index, link] of chain.entries()) {
    const value = pick(link.shape);
    if (value === undefined) continue;
    return {
      value,
      origin: link.origin,
      explicit: index === 0,
      sheet: link.sheet,
      shape: link.shape,
    };
  }
  return undefined;
}

/**
 * The rectangle a shape occupies, and which sheet stated it.
 *
 * The single most consequential call in the package. A slide placeholder that
 * PowerPoint wrote has **no** `a:xfrm` - not when it was created, not when text
 * was typed into it, only when the user moved or resized it, at which point the
 * whole resolved rectangle is baked in at once. So an absent `xfrm` is the
 * statement "still bound to my layout", and it is what Change Layout reads.
 */
export function resolveXfrm(shape: Shape, sheet: Sheet): Resolved<Xfrm> | undefined {
  return resolve(shape, sheet, (s) => s.xfrm);
}

/* -------------------------------------------------------------------------- */
/* the sheet chain                                                            */
/* -------------------------------------------------------------------------- */

/** Every sheet from this one up to its master, nearest first. */
export function sheetChain(sheet: Sheet): readonly Sheet[] {
  const chain: Sheet[] = [];
  const seen = new Set<string>();
  let current: Sheet | null = sheet;
  while (current !== null) {
    if (seen.has(current.partName)) {
      throw new ModelError(
        'MODEL_SHEET_CYCLE',
        `the sheet chain returns to ${current.partName}`,
        current.partName,
      );
    }
    seen.add(current.partName);
    chain.push(current);
    current = current.parent;
  }
  return chain;
}

/** The master at the top of a sheet's chain, or `null` if the chain is broken. */
export function masterOf(sheet: Sheet): Sheet | null {
  for (const link of sheetChain(sheet)) if (link.kind === 'master') return link;
  return null;
}

/**
 * The theme a sheet resolves against.
 *
 * The master's, reached through this sheet's own chain. **Not** the one
 * `ppt/presentation.xml.rels` names: PowerPoint writes a `theme` relationship
 * there too, it always points at `theme1.xml`, and on a two-master deck every
 * slide on the second master would take the first master's palette. Measured -
 * the two masters resolved `accent1` to two different colours, and each
 * master's background used its own theme's `bgFillStyleLst`.
 *
 * A master must own its theme part outright: two masters pointing at the same
 * one is repaired on open, so a shared theme is a broken file rather than an
 * economy.
 */
export function themeOf(sheet: Sheet): Theme | null {
  return masterOf(sheet)?.theme ?? null;
}

export function schemeOf(sheet: Sheet): ClrScheme | undefined {
  return themeOf(sheet)?.scheme;
}

/**
 * The colour map in force on a sheet.
 *
 * `a:masterClrMapping` says "the map already in force", and the surprise is what
 * that means: for a slide it is the **layout's** map, not the master's. Measured
 * - a layout carrying an `a:overrideClrMapping` changed `accent1` on a slide
 * that declared `masterClrMapping`, and a slide with an override of its own beat
 * its layout's. The element is named for the sheet it usually ends at rather
 * than for the one it asks.
 */
export function colorMapOf(sheet: Sheet): ClrMap | undefined {
  for (const link of sheetChain(sheet)) {
    if (link.kind === 'master') return link.clrMap;
    if (link.clrMapOvr?.kind === 'override') return link.clrMapOvr.map;
  }
  return undefined;
}

/**
 * Everything `resolveColor` needs, for a colour written on this sheet.
 *
 * `phClr` is supplied by whatever style invocation the colour sits inside, and
 * there is nothing sensible to default it to: `style.ts` passes it in, and a
 * `schemeClr val="phClr"` reached any other way throws rather than painting
 * black.
 */
export function colorContextOf(sheet: Sheet, phClr?: Rgba): ColorContext {
  const scheme = schemeOf(sheet);
  const map = colorMapOf(sheet);
  return {
    ...(scheme === undefined ? {} : { scheme }),
    ...(map === undefined ? {} : { map }),
    ...(phClr === undefined ? {} : { phClr }),
  };
}
