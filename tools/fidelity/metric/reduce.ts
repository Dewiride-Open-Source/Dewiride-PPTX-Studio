/**
 * A raster into the integer grid the fidelity score is computed over.
 *
 * Two properties are load-bearing and both are why this is written the way it
 * is. It is **pure integer arithmetic**, so two machines that agree on the
 * pixels agree on the grid exactly, with no float summation order to differ
 * over. And it holds **no free variables**, so `page.evaluate` can stringify it
 * and the browser runs the same source Node does - one implementation, not two
 * that agree until they do not. ADR 0035.
 */

/** The reduced planes of one raster, all values in `0..255`. */
export interface Grid {
  /** Luma cells across and down; the raster is `cell` times this. */
  readonly width: number;
  readonly height: number;
  /** The side of one luma cell, in raster pixels. */
  readonly cell: number;
  /** `width * height` luma means, row-major. */
  readonly luma: readonly number[];
  /** Chroma at half resolution in each axis, row-major. */
  readonly chromaWidth: number;
  readonly chromaHeight: number;
  readonly cb: readonly number[];
  readonly cr: readonly number[];
}

/** What `reduceRgba` is handed; `data` is RGBA, four bytes per pixel. */
export interface RasterInput {
  readonly data: ArrayLike<number>;
  readonly width: number;
  readonly height: number;
  readonly cell: number;
}

/**
 * The raster reduced to luma and decimated chroma.
 *
 * BT.601 at 8-bit precision, the same coefficients JPEG uses, chosen because
 * they are integers everyone has written down rather than because this is a
 * JPEG. Chroma is decimated a further 2x2 for the reason chroma always is: the
 * eye resolves it less finely, and it halves what has to be committed.
 *
 * Declared as a `function` with no captured scope so that
 * `page.evaluate(reduceRgba, …)` sends this exact body to the browser.
 */
export function reduceRgba(input: RasterInput): Grid {
  const { data, width, height, cell } = input;
  if (!Number.isInteger(cell) || cell <= 0) throw new Error(`cell ${String(cell)} is not a size`);
  if (width % cell !== 0 || height % cell !== 0) {
    throw new Error(
      `raster ${String(width)}x${String(height)} is not whole cells of ${String(cell)}`,
    );
  }

  const cols = width / cell;
  const rows = height / cell;
  const luma: number[] = new Array<number>(cols * rows);
  const cellCb: number[] = new Array<number>(cols * rows);
  const cellCr: number[] = new Array<number>(cols * rows);
  const half = (cell * cell) >> 1;

  for (let cy = 0; cy < rows; cy++) {
    for (let cx = 0; cx < cols; cx++) {
      let sr = 0;
      let sg = 0;
      let sb = 0;
      for (let y = 0; y < cell; y++) {
        let at = ((cy * cell + y) * width + cx * cell) * 4;
        for (let x = 0; x < cell; x++) {
          sr += data[at] ?? 0;
          sg += data[at + 1] ?? 0;
          sb += data[at + 2] ?? 0;
          at += 4;
        }
      }
      // Rounded means, half up, so a cell one level brighter is one level here.
      const r = Math.floor((sr + half) / (cell * cell));
      const g = Math.floor((sg + half) / (cell * cell));
      const b = Math.floor((sb + half) / (cell * cell));
      const at = cy * cols + cx;
      luma[at] = (77 * r + 150 * g + 29 * b + 128) >> 8;
      const cb = ((-43 * r - 85 * g + 128 * b + 128) >> 8) + 128;
      const cr = ((128 * r - 107 * g - 21 * b + 128) >> 8) + 128;
      cellCb[at] = cb < 0 ? 0 : cb > 255 ? 255 : cb;
      cellCr[at] = cr < 0 ? 0 : cr > 255 ? 255 : cr;
    }
  }

  const chromaWidth = Math.ceil(cols / 2);
  const chromaHeight = Math.ceil(rows / 2);
  const cb: number[] = new Array<number>(chromaWidth * chromaHeight);
  const cr: number[] = new Array<number>(chromaWidth * chromaHeight);
  for (let y = 0; y < chromaHeight; y++) {
    for (let x = 0; x < chromaWidth; x++) {
      let sumCb = 0;
      let sumCr = 0;
      let n = 0;
      for (let dy = 0; dy < 2; dy++) {
        for (let dx = 0; dx < 2; dx++) {
          const sy = y * 2 + dy;
          const sx = x * 2 + dx;
          if (sy >= rows || sx >= cols) continue;
          sumCb += cellCb[sy * cols + sx] ?? 0;
          sumCr += cellCr[sy * cols + sx] ?? 0;
          n += 1;
        }
      }
      // A ragged edge averages the cells that exist, so an odd grid is not a
      // special case and no cell is invented to square it off.
      cb[y * chromaWidth + x] = Math.floor((sumCb + (n >> 1)) / n);
      cr[y * chromaWidth + x] = Math.floor((sumCr + (n >> 1)) / n);
    }
  }

  return { width: cols, height: rows, cell, luma, chromaWidth, chromaHeight, cb, cr };
}
