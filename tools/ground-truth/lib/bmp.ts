/**
 * Decode an uncompressed Windows BMP far enough to sample a pixel.
 *
 * PowerPoint's `Slide.Export` writes both PNG and BMP. We sample the BMP: it
 * needs no decompressor, so there is no library between what PowerPoint painted
 * and what we compare against, and a swatch's interior has no anti-aliasing to
 * worry about. The PNG is exported alongside purely so a human can look at it.
 */

export interface Bitmap {
  readonly width: number;
  readonly height: number;
  /** `[r, g, b]`, 0-255. Origin top-left, whichever way the file stored it. */
  pixel(x: number, y: number): [number, number, number];
}

export function readBmp(bytes: Uint8Array): Bitmap {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  if (bytes[0] !== 0x42 || bytes[1] !== 0x4d) throw new Error('not a BMP: missing "BM"');

  const dataOffset = view.getUint32(0x0a, true);
  const headerSize = view.getUint32(0x0e, true);
  if (headerSize < 40) throw new Error(`unsupported BMP header size ${String(headerSize)}`);

  const width = view.getInt32(0x12, true);
  const rawHeight = view.getInt32(0x16, true);
  const bitCount = view.getUint16(0x1c, true);
  const compression = view.getUint32(0x1e, true);
  if (compression !== 0 && compression !== 3) {
    throw new Error(`unsupported BMP compression ${String(compression)}`);
  }
  if (bitCount !== 24 && bitCount !== 32) {
    throw new Error(`unsupported BMP bit depth ${String(bitCount)}`);
  }

  // A negative height means the rows are stored top-down.
  const topDown = rawHeight < 0;
  const height = Math.abs(rawHeight);
  const bytesPerPixel = bitCount / 8;
  const stride = Math.floor((width * bitCount + 31) / 32) * 4;

  return {
    width,
    height,
    pixel(x: number, y: number): [number, number, number] {
      if (x < 0 || y < 0 || x >= width || y >= height) {
        throw new Error(
          `pixel (${String(x)},${String(y)}) is outside ${String(width)}x${String(height)}`,
        );
      }
      const row = topDown ? y : height - 1 - y;
      const at = dataOffset + row * stride + x * bytesPerPixel;
      // BMP stores BGR.
      return [bytes[at + 2]!, bytes[at + 1]!, bytes[at]!];
    },
  };
}

export function hex(rgb: readonly [number, number, number]): string {
  return rgb.map((c) => c.toString(16).padStart(2, '0').toUpperCase()).join('');
}

/** The smallest rectangle holding every pixel darker than `threshold`, or `null`. */
export interface InkBox {
  readonly left: number;
  readonly top: number;
  readonly right: number;
  readonly bottom: number;
}

/**
 * Where the ink is, in pixels, inclusive on every edge.
 *
 * The probe slides are white on a light grey reference rectangle, so a single
 * darkness threshold separates drawn glyphs from everything else.
 */
export function inkBox(bitmap: Bitmap, threshold = 200): InkBox | null {
  let left = bitmap.width;
  let top = bitmap.height;
  let right = -1;
  let bottom = -1;
  for (let y = 0; y < bitmap.height; y++) {
    for (let x = 0; x < bitmap.width; x++) {
      const [r, g, b] = bitmap.pixel(x, y);
      if (r >= threshold && g >= threshold && b >= threshold) continue;
      if (x < left) left = x;
      if (x > right) right = x;
      if (y < top) top = y;
      if (y > bottom) bottom = y;
    }
  }
  return right < 0 ? null : { left, top, right, bottom };
}

/**
 * Write a 24-bit BMP, bottom-up, with an explicit resolution in the header.
 *
 * A probe image needs a declared DPI, because whether PowerPoint sizes a tile
 * from the image's own resolution or from a fixed 96 is a question only an image
 * that disagrees with 96 can answer. BMP states it in two header fields, so the
 * probe needs no metadata library and no PNG encoder.
 */
export function writeBmp(
  width: number,
  height: number,
  pixel: (x: number, y: number) => number,
  dpi = 96,
): Uint8Array {
  if (width < 1 || height < 1) throw new Error(`a ${String(width)}x${String(height)} BMP is empty`);
  const stride = Math.floor((width * 24 + 31) / 32) * 4;
  const dataSize = stride * height;
  const out = new Uint8Array(54 + dataSize);
  const view = new DataView(out.buffer);

  out[0] = 0x42;
  out[1] = 0x4d;
  view.setUint32(2, out.length, true);
  view.setUint32(10, 54, true);
  view.setUint32(14, 40, true);
  view.setInt32(18, width, true);
  view.setInt32(22, height, true);
  view.setUint16(26, 1, true);
  view.setUint16(28, 24, true);
  view.setUint32(34, dataSize, true);
  // Pixels per metre, which is how BMP spells DPI: 96 dpi is 3780.
  const perMetre = Math.round(dpi / 0.0254);
  view.setInt32(38, perMetre, true);
  view.setInt32(42, perMetre, true);

  for (let y = 0; y < height; y++) {
    // Bottom-up: the last row of the image is the first row in the file.
    const row = 54 + (height - 1 - y) * stride;
    for (let x = 0; x < width; x++) {
      const rgb = pixel(x, y);
      out[row + x * 3] = rgb & 0xff;
      out[row + x * 3 + 1] = (rgb >> 8) & 0xff;
      out[row + x * 3 + 2] = (rgb >> 16) & 0xff;
    }
  }
  return out;
}

/** Every column that holds ink, as a sorted list. For finding a rule under a word. */
export function inkRows(bitmap: Bitmap, threshold = 200): readonly number[] {
  const rows: number[] = [];
  for (let y = 0; y < bitmap.height; y++) {
    for (let x = 0; x < bitmap.width; x++) {
      const [r, g, b] = bitmap.pixel(x, y);
      if (r < threshold || g < threshold || b < threshold) {
        rows.push(y);
        break;
      }
    }
  }
  return rows;
}
