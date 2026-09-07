/**
 * How big a picture is, read from its own header.
 *
 * A tiled `a:blipFill` needs the image's pixel count and its declared
 * resolution, and both are measurable without decoding a single pixel. Reading
 * the header rather than the image is what keeps this synchronous, which is what
 * lets `renderSlide` stay a pure string function. ADR 0036.
 */

import type { ImageSize } from '@pptx-studio/paint';

import { RenderError } from '../errors.js';

/** What PowerPoint assumes when an image declares nothing. Measured in C6. */
const DEFAULT_DPI = 96;

const DPI_PER_METRE = 0.0254;

function big32(bytes: Uint8Array, at: number): number {
  return (
    (((bytes[at] ?? 0) << 24) |
      ((bytes[at + 1] ?? 0) << 16) |
      ((bytes[at + 2] ?? 0) << 8) |
      (bytes[at + 3] ?? 0)) >>>
    0
  );
}

function little32(bytes: Uint8Array, at: number): number {
  return (
    ((bytes[at + 3] ?? 0) << 24) |
    ((bytes[at + 2] ?? 0) << 16) |
    ((bytes[at + 1] ?? 0) << 8) |
    (bytes[at] ?? 0)
  );
}

function little16(bytes: Uint8Array, at: number): number {
  return ((bytes[at + 1] ?? 0) << 8) | (bytes[at] ?? 0);
}

function big16(bytes: Uint8Array, at: number): number {
  return ((bytes[at] ?? 0) << 8) | (bytes[at + 1] ?? 0);
}

function png(bytes: Uint8Array): ImageSize | null {
  if (big32(bytes, 0) !== 0x89504e47) return null;
  const widthPx = big32(bytes, 16);
  const heightPx = big32(bytes, 20);
  let dpi = DEFAULT_DPI;
  // Walk the chunks for `pHYs`, which states pixels per unit and the unit.
  let at = 8;
  while (at + 8 <= bytes.length) {
    const length = big32(bytes, at);
    const type = String.fromCharCode(...bytes.subarray(at + 4, at + 8));
    if (type === 'pHYs' && bytes[at + 16] === 1) {
      dpi = big32(bytes, at + 8) * DPI_PER_METRE;
      break;
    }
    if (type === 'IDAT' || type === 'IEND') break;
    at += 12 + length;
  }
  return { widthPx, heightPx, dpi };
}

function jpeg(bytes: Uint8Array): ImageSize | null {
  if (big16(bytes, 0) !== 0xffd8) return null;
  let dpi = DEFAULT_DPI;
  let at = 2;
  while (at + 4 <= bytes.length) {
    if (bytes[at] !== 0xff) {
      at += 1;
      continue;
    }
    const marker = bytes[at + 1] ?? 0;
    const length = big16(bytes, at + 2);
    // APP0/JFIF states a density and the unit it is in.
    if (marker === 0xe0 && bytes[at + 11] === 1) dpi = big16(bytes, at + 12) || DEFAULT_DPI;
    if (marker === 0xe0 && bytes[at + 11] === 2) dpi = big16(bytes, at + 12) * 2.54;
    const isFrame =
      marker >= 0xc0 && marker <= 0xcf && marker !== 0xc4 && marker !== 0xc8 && marker !== 0xcc;
    if (isFrame) {
      return { heightPx: big16(bytes, at + 5), widthPx: big16(bytes, at + 7), dpi };
    }
    at += 2 + length;
  }
  return null;
}

function gif(bytes: Uint8Array): ImageSize | null {
  if (big32(bytes, 0) !== 0x47494638) return null;
  return { widthPx: little16(bytes, 6), heightPx: little16(bytes, 8), dpi: DEFAULT_DPI };
}

function bmp(bytes: Uint8Array): ImageSize | null {
  if (bytes[0] !== 0x42 || bytes[1] !== 0x4d) return null;
  const perMetre = little32(bytes, 38);
  return {
    widthPx: little32(bytes, 18),
    heightPx: Math.abs(little32(bytes, 22)),
    dpi: perMetre > 0 ? perMetre * DPI_PER_METRE : DEFAULT_DPI,
  };
}

/**
 * The size of a PNG, JPEG, GIF or BMP.
 *
 * Anything else throws: a tile sized from a guess is a wrong slide nobody can
 * trace, and the formats that need decoding - EMF, WMF, TIFF - are Phase 10.
 */
export function imageSize(bytes: Uint8Array): ImageSize {
  const size = png(bytes) ?? jpeg(bytes) ?? gif(bytes) ?? bmp(bytes);
  if (size === null) {
    throw new RenderError('RENDER_IMAGE_FORMAT', 'the image is not a PNG, JPEG, GIF or BMP');
  }
  if (!(size.widthPx > 0) || !(size.heightPx > 0)) {
    throw new RenderError('RENDER_IMAGE_FORMAT', 'the image header declares no pixels');
  }
  return size;
}

const BASE64 = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/';

/**
 * A `data:` URI, encoded here rather than through `btoa`.
 *
 * `btoa` is a DOM API and the core packages run in a Worker with no DOM, and
 * `Buffer` is Node. Twenty lines is cheaper than either dependency.
 */
export function dataUri(bytes: Uint8Array, contentType: string): string {
  let out = '';
  for (let at = 0; at < bytes.length; at += 3) {
    const a = bytes[at] ?? 0;
    const b = bytes[at + 1] ?? 0;
    const c = bytes[at + 2] ?? 0;
    const left = bytes.length - at;
    out += BASE64[a >> 2];
    out += BASE64[((a & 3) << 4) | (b >> 4)];
    out += left > 1 ? BASE64[((b & 15) << 2) | (c >> 6)] : '=';
    out += left > 2 ? BASE64[c & 63] : '=';
  }
  return `data:${contentType};base64,${out}`;
}
