/**
 * The committed oracle, opened: PowerPoint's grid for any slide, at 960 wide and at the zoom
 * widths the capture also exported, without a caller knowing which file holds it. ADR 0054.
 */

import { readFileSync } from 'node:fs';
import { join } from 'node:path';

import { repoPath } from '../repo/root.ts';

import { FidelityError } from './errors.ts';
import { decodeGridSet } from './metric/grid.ts';
import type { Grid } from './metric/reduce.ts';

export interface OracleRecord {
  readonly key: string;
  readonly deck: string;
  readonly slide: number;
  readonly cells: number;
  /** Relative to `corpus/ground-truth/`. */
  readonly file: string;
}

export interface ZoomRecord extends OracleRecord {
  /** The export width in pixels, and the cell that keeps the grid 120 wide. */
  readonly width: number;
  readonly cell: number;
}

export interface Oracle {
  readonly powerpoint: string;
  readonly rasterWidth: number;
  readonly slides: number;
  readonly noiseFloor: {
    readonly slidesAffected: number;
    readonly worstMeanBpDrop: number;
    readonly worstMaxD: number;
    readonly worstCells: number;
  };
  readonly jitter: readonly { key: string; meanBp: number; maxD: number; cells: number }[];
  readonly records: readonly OracleRecord[];
  readonly zoom: readonly ZoomRecord[];
}

/** How far PowerPoint disagreed with its own second export of one slide. */
export interface JitterRow {
  readonly key: string;
  readonly meanBp: number;
  readonly maxD: number;
  readonly cells: number;
}

/** What one capture measured: the decks it exported, and their rows. */
export interface OracleCapture {
  readonly powerpoint: string;
  readonly rasterWidth: number;
  readonly captured: ReadonlySet<string>;
  readonly records: readonly OracleRecord[];
  readonly zoom: readonly ZoomRecord[];
  readonly jitter: readonly JitterRow[];
}

/**
 * The committed oracle with one capture's decks replaced: every other deck keeps its rows,
 * and two PowerPoint builds are refused rather than averaged.
 */
export function mergeOracle(existing: Oracle | null, fresh: OracleCapture): Oracle {
  if (existing !== null && existing.powerpoint !== fresh.powerpoint) {
    throw new FidelityError(
      'FID_ORACLE_MISSING',
      `the oracle was captured on PowerPoint ${existing.powerpoint} and this run on ` +
        `${fresh.powerpoint}; two builds are not one oracle`,
    );
  }
  const kept = (row: { readonly deck: string }): boolean => !fresh.captured.has(row.deck);
  const keptRecords = (existing?.records ?? []).filter(kept);
  const keptKeys = new Set(keptRecords.map((row) => row.key));
  const records = [...keptRecords, ...fresh.records];
  const jitter = [
    ...(existing?.jitter ?? []).filter((row) => keptKeys.has(row.key)),
    ...fresh.jitter,
  ];
  return {
    powerpoint: fresh.powerpoint,
    rasterWidth: fresh.rasterWidth,
    slides: records.length,
    noiseFloor: {
      slidesAffected: jitter.length,
      worstMeanBpDrop: jitter.reduce((worst, row) => Math.max(worst, 10000 - row.meanBp), 0),
      worstMaxD: jitter.reduce((worst, row) => Math.max(worst, row.maxD), 0),
      worstCells: jitter.reduce((worst, row) => Math.max(worst, row.cells), 0),
    },
    jitter,
    records,
    zoom: [...(existing?.zoom ?? []).filter(kept), ...fresh.zoom],
  };
}

export interface OpenedOracle {
  readonly oracle: Oracle;
  /** Every slide the oracle holds at 960 wide. */
  readonly keys: ReadonlySet<string>;
  gridFor(key: string): Grid;
  zoomGridFor(key: string, width: number): Grid;
}

/** Where the committed oracle lives. */
export const ORACLE_DIR = repoPath('corpus/ground-truth/render/fidelity');

export function openOracle(dir = ORACLE_DIR): OpenedOracle {
  const oracle = JSON.parse(readFileSync(join(dir, 'oracle.json'), 'utf8')) as Oracle;
  const sets = new Map<string, ReadonlyMap<string, Grid>>();
  const setOf = (file: string): ReadonlyMap<string, Grid> => {
    let set = sets.get(file);
    if (set === undefined) {
      const path = join(dir, '..', '..', file);
      set = decodeGridSet(new Uint8Array(readFileSync(path)), file);
      sets.set(file, set);
    }
    return set;
  };
  const gridOf = (record: OracleRecord | undefined, what: string): Grid => {
    if (record === undefined) {
      throw new FidelityError('FID_ORACLE_MISSING', `no oracle grid for ${what}`, what);
    }
    const grid = setOf(record.file).get(record.key);
    if (grid === undefined) {
      throw new FidelityError('FID_ORACLE_MISSING', `${record.file} holds no ${what}`, what);
    }
    return grid;
  };
  const records = new Map(oracle.records.map((record) => [record.key, record]));
  return {
    oracle,
    keys: new Set(records.keys()),
    gridFor: (key) => gridOf(records.get(key), key),
    zoomGridFor: (key, width) =>
      gridOf(
        oracle.zoom.find((record) => record.key === key && record.width === width),
        `${key} at ${String(width)} px wide`,
      ),
  };
}
