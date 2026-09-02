import type { PresetBucket } from '../decode.js';
import type { PresetBucketName, PresetShape } from '../types.js';
import { ARROWS } from './arrows.gen.js';
import { BASIC } from './basic.gen.js';
import { CALLOUTS } from './callouts.gen.js';
import { FLOWCHART } from './flowchart.gen.js';
import { MISC } from './misc.gen.js';
import { STARS } from './stars.gen.js';

/**
 * All 187 presets, and the stable way to reach one.
 *
 * ## The trade this module makes, said out loud
 *
 * Importing it pulls in all six buckets - about 120 KB of encoded text before
 * compression - because a general renderer cannot know which shapes a deck will
 * contain until it opens one. That is the right default and it is also the
 * expensive one.
 *
 * A consumer who does know can import a bucket directly:
 *
 * ```ts
 * import { FLOWCHART } from '@pptx-studio/geometry/presets/flowchart.gen.js';
 * ```
 *
 * and pay for 29 shapes instead of 187. The bucket boundaries are a packaging
 * decision rather than a fact about geometry, so they may be rebalanced; a
 * consumer importing one directly should pin a version. `getPreset` never
 * moves.
 */

export const BUCKETS: Readonly<Record<PresetBucketName, PresetBucket>> = {
  arrows: ARROWS,
  basic: BASIC,
  callouts: CALLOUTS,
  flowchart: FLOWCHART,
  misc: MISC,
  stars: STARS,
};

/**
 * Built on first lookup, not at module load.
 *
 * Six map lookups per miss is not the cost worth avoiding; a top-level
 * `new Map` over 187 names in every bundle that so much as imports this module
 * is. Deferring it also keeps the module free of load-time work, which is what
 * `"sideEffects": false` promises the bundler.
 */
let index: Map<string, PresetBucket> | null = null;

function indexed(): Map<string, PresetBucket> {
  if (index !== null) return index;
  const built = new Map<string, PresetBucket>();
  for (const bucket of Object.values(BUCKETS)) {
    for (const name of bucket.names) built.set(name, bucket);
  }
  index = built;
  return built;
}

/**
 * The preset for an `a:prstGeom/@prst` value, or `undefined` if it is not one
 * of the 187.
 *
 * Undefined rather than a throw, and rather than a silent rectangle: a deck
 * naming a shape type that does not exist is a fact the caller should be able
 * to report, and 2.10's renderer draws nothing for it rather than guessing.
 */
export function getPreset(name: string): PresetShape | undefined {
  return indexed().get(name)?.get(name);
}

/**
 * Every `ST_ShapeType` name that has a definition, sorted by code unit.
 *
 * Not `localeCompare`, deliberately. It is collation, so it depends on the ICU
 * build behind the engine and on the ambient locale - `noSmoking` and
 * `nonIsoscelesTrapezoid` swap places between the two orderings, which is
 * exactly the kind of difference that makes a list stable on one browser and
 * not on another. A code-unit sort is the same everywhere.
 */
export function presetNames(): readonly string[] {
  return [...indexed().keys()].sort();
}

export { ARROWS, BASIC, CALLOUTS, FLOWCHART, MISC, STARS };
export {
  BUCKET_MEMBERS,
  COLLAPSED_WHITESPACE,
  OVER_LONG_FORMULAS,
  PRESET_SOURCE,
  PRESET_TOTALS,
} from './manifest.gen.js';
