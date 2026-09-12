import { mkdtempSync, mkdirSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';

import { describe, expect, it } from 'vitest';

import { FidelityError } from './errors.ts';
import { encodeGridSet } from './metric/grid.ts';
import type { Grid } from './metric/reduce.ts';
import { openOracle } from './oracle.ts';

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
