import { mkdtempSync, mkdirSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';

import { describe, expect, it } from 'vitest';

import { FidelityError } from './errors.ts';
import { encodeGridSet } from './metric/grid.ts';
import type { Grid } from './metric/reduce.ts';
import { mergeOracle, openOracle, type Oracle } from './oracle.ts';

/** A one-cell grid whose luma is `value`, so two grids are told apart by one byte. */
function grid(value: number, cell = 8): Grid {
  return {
    width: 1,
    height: 1,
    cell,
    luma: [value],
    chromaWidth: 1,
    chromaHeight: 1,
    cb: [128],
    cr: [128],
  };
}

/** An oracle directory of one deck in two sets, and one zoomed set. */
function oracleDir(): string {
  const root = mkdtempSync(join(tmpdir(), 'oracle-'));
  const dir = join(root, 'render', 'fidelity');
  mkdirSync(join(dir, 'grids'), { recursive: true });
  mkdirSync(join(dir, 'zoom'), { recursive: true });
  writeFileSync(
    join(dir, 'grids', 'd.1.ppt.grids'),
    encodeGridSet([{ key: 'd-01', grid: grid(10) }]),
  );
  writeFileSync(
    join(dir, 'grids', 'd.2.ppt.grids'),
    encodeGridSet([{ key: 'd-02', grid: grid(20) }]),
  );
  writeFileSync(
    join(dir, 'zoom', 'd.240.ppt.grids'),
    encodeGridSet([{ key: 'd-01', grid: grid(30, 2) }]),
  );
  writeFileSync(
    join(dir, 'oracle.json'),
    JSON.stringify({
      powerpoint: 'test',
      rasterWidth: 960,
      slides: 2,
      noiseFloor: { slidesAffected: 0, worstMeanBpDrop: 0, worstMaxD: 0, worstCells: 0 },
      jitter: [],
      records: [
        { key: 'd-01', deck: 'd', slide: 1, cells: 1, file: 'render/fidelity/grids/d.1.ppt.grids' },
        { key: 'd-02', deck: 'd', slide: 2, cells: 1, file: 'render/fidelity/grids/d.2.ppt.grids' },
      ],
      zoom: [
        {
          key: 'd-01',
          deck: 'd',
          slide: 1,
          cells: 1,
          width: 240,
          cell: 2,
          file: 'render/fidelity/zoom/d.240.ppt.grids',
        },
      ],
    }),
  );
  return dir;
}

describe('openOracle', () => {
  it('finds a slide in whichever set the index names', () => {
    const opened = openOracle(oracleDir());
    expect([...opened.keys]).toEqual(['d-01', 'd-02']);
    expect(opened.gridFor('d-01').luma).toEqual([10]);
    expect(opened.gridFor('d-02').luma).toEqual([20]);
  });

  it('finds a zoomed grid by width, and nothing at a width the capture did not export', () => {
    const opened = openOracle(oracleDir());
    const zoomed = opened.zoomGridFor('d-01', 240);
    expect(zoomed.luma).toEqual([30]);
    expect(zoomed.cell).toBe(2);
    expect(() => opened.zoomGridFor('d-01', 3840)).toThrow(FidelityError);
    expect(() => opened.zoomGridFor('d-02', 240)).toThrow(/no oracle grid/);
  });

  it('refuses a slide the index does not hold', () => {
    expect(() => openOracle(oracleDir()).gridFor('d-03')).toThrow(FidelityError);
  });
});

describe('mergeOracle', () => {
  const row = (deck: string, slide: number, file = `render/fidelity/grids/${deck}.ppt.grids`) => ({
    key: `${deck}-${String(slide).padStart(2, '0')}`,
    deck,
    slide,
    cells: 8160,
    file,
  });
  const existing: Oracle = {
    powerpoint: '16.0.1',
    rasterWidth: 960,
    slides: 3,
    noiseFloor: { slidesAffected: 2, worstMeanBpDrop: 9, worstMaxD: 40, worstCells: 12 },
    jitter: [
      { key: 'a01-x-01', meanBp: 9991, maxD: 40, cells: 12 },
      { key: 'a02-y-01', meanBp: 9995, maxD: 20, cells: 3 },
    ],
    records: [row('a01-x', 1), row('a02-y', 1), row('a02-y', 2)],
    zoom: [{ ...row('a02-y', 1, 'render/fidelity/zoom/a02-y.240.ppt.grids'), width: 240, cell: 2 }],
  };

  it('replaces the captured decks and keeps every other row, jitter and zoomed grid as it was', () => {
    const merged = mergeOracle(existing, {
      powerpoint: '16.0.1',
      rasterWidth: 960,
      captured: new Set(['a02-y']),
      records: [row('a02-y', 1, 'render/fidelity/grids/a02-y.1.ppt.grids')],
      zoom: [],
      jitter: [{ key: 'a02-y-01', meanBp: 9990, maxD: 50, cells: 30 }],
    });
    expect(merged.records.map((r) => r.key)).toEqual(['a01-x-01', 'a02-y-01']);
    expect(merged.records[0]?.file).toBe('render/fidelity/grids/a01-x.ppt.grids');
    expect(merged.records[1]?.file).toBe('render/fidelity/grids/a02-y.1.ppt.grids');
    expect(merged.slides).toBe(2);
    // The recaptured deck's zoomed grids and old jitter go with it; the kept deck's stay.
    expect(merged.zoom).toEqual([]);
    expect(merged.jitter.map((j) => j.key)).toEqual(['a01-x-01', 'a02-y-01']);
    expect(merged.noiseFloor).toEqual({
      slidesAffected: 2,
      worstMeanBpDrop: 10,
      worstMaxD: 50,
      worstCells: 30,
    });
  });

  it('is the capture alone when there is nothing to merge into', () => {
    const merged = mergeOracle(null, {
      powerpoint: '16.0.2',
      rasterWidth: 960,
      captured: new Set(['a01-x']),
      records: [row('a01-x', 1)],
      zoom: [],
      jitter: [],
    });
    expect(merged.powerpoint).toBe('16.0.2');
    expect(merged.records).toHaveLength(1);
    expect(merged.noiseFloor.slidesAffected).toBe(0);
  });

  it('refuses to put two PowerPoint builds in one oracle', () => {
    expect(() =>
      mergeOracle(existing, {
        powerpoint: '16.0.2',
        rasterWidth: 960,
        captured: new Set(['a01-x']),
        records: [],
        zoom: [],
        jitter: [],
      }),
    ).toThrow(FidelityError);
  });
});
