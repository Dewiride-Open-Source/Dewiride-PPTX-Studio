/**
 * Decode the PNG `Slide.Export` writes: indexed, RGB or RGBA, not interlaced.
 *
 * A BMP needs no decompressor, which is why `bmp.ts` is the default; a sweep that exports thousands
 * of slides at 4 px per point cannot afford 24 MB each, so it exports PNG and proves this decoder
 * against a BMP twin of the same slide.
 */

import { inflateSync } from 'node:zlib';

import type { Bitmap } from './bmp.ts';

const SIGNATURE = [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a];
/** Samples per pixel by colour type: RGB, indexed, RGBA. */
const CHANNELS: Readonly<Record<number, number>> = { 2: 3, 3: 1, 6: 4 };

interface Chunks {
  readonly width: number;
  readonly height: number;
  readonly depth: number;
  readonly colour: number;
  readonly palette: Uint8Array | null;
  readonly data: Uint8Array;
}

function chunksOf(bytes: Uint8Array): Chunks {
  if (!SIGNATURE.every((b, i) => bytes[i] === b)) throw new Error('not a PNG: bad signature');
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  let header: Uint8Array | null = null;
  let palette: Uint8Array | null = null;
  const idat: Uint8Array[] = [];
  for (let at = 8; at < bytes.length;) {
    const length = view.getUint32(at);
    const type = new TextDecoder().decode(bytes.subarray(at + 4, at + 8));
    const data = bytes.subarray(at + 8, at + 8 + length);
    if (type === 'IEND') break;
    if (type === 'IHDR') header = data;
    else if (type === 'PLTE') palette = data;
    else if (type === 'IDAT') idat.push(data);
    at += 12 + length;
  }
  if (header === null) throw new Error('PNG has no IHDR');
  if (header[12] !== 0) throw new Error('unsupported interlaced PNG');
  const data = new Uint8Array(idat.reduce((n, d) => n + d.length, 0));
  let offset = 0;
  for (const d of idat) {
    data.set(d, offset);
    offset += d.length;
  }
  const head = new DataView(header.buffer, header.byteOffset, header.byteLength);
  return {
    width: head.getUint32(0),
    height: head.getUint32(4),
    depth: header[8] ?? 0,
    colour: header[9] ?? -1,
    palette,
    data,
  };
}

function paeth(a: number, b: number, c: number): number {
  const p = a + b - c;
  const pa = Math.abs(p - a);
  const pb = Math.abs(p - b);
  const pc = Math.abs(p - c);
  if (pa <= pb && pa <= pc) return a;
  return pb <= pc ? b : c;
}

function predictor(filter: number, a: number, b: number, c: number): number {
  switch (filter) {
    case 0:
      return 0;
    case 1:
      return a;
    case 2:
      return b;
    case 3:
      return (a + b) >> 1;
    case 4:
      return paeth(a, b, c);
    default:
      throw new Error(`unsupported PNG filter ${String(filter)}`);
  }
}

/** A decoded PNG that also answers `0xRRGGBB`, for scans too long to allocate an array per pixel. */
export interface PackedBitmap extends Bitmap {
  rgb(x: number, y: number): number;
}

export function readPng(bytes: Uint8Array): PackedBitmap {
  const { width, height, depth, colour, palette, data } = chunksOf(bytes);
  const channels = CHANNELS[colour];
  if (channels === undefined) throw new Error(`unsupported PNG colour type ${String(colour)}`);
  const depths = colour === 3 ? [1, 2, 4, 8] : [8];
  if (!depths.includes(depth)) {
    throw new Error(`unsupported PNG bit depth ${String(depth)} for colour type ${String(colour)}`);
  }
  if (colour === 3 && palette === null) throw new Error('indexed PNG has no PLTE');
  // Below 8 bits per sample the filters work one byte apart, not one pixel.
  const stride = Math.ceil((width * channels * depth) / 8);
  const step = Math.max(1, (channels * depth) / 8);
  const raw = inflateSync(data);
  if (raw.length !== height * (stride + 1)) {
    throw new Error(
      `PNG data is ${String(raw.length)} bytes, expected ${String(height * (stride + 1))}`,
    );
  }
  const packed =
    palette === null
      ? null
      : Array.from(
          { length: palette.length / 3 },
          (_, i) => (palette[3 * i]! << 16) | (palette[3 * i + 1]! << 8) | palette[3 * i + 2]!,
        );
  const out = new Uint8Array(height * stride);
  let decoded = 0;
  // Rows are unfiltered on demand: each depends on the one above, and scans rarely reach the bottom.
  const ensure = (y: number): void => {
    for (; decoded <= y; decoded++) {
      const filter = raw[decoded * (stride + 1)] ?? -1;
      const line = decoded * (stride + 1) + 1;
      const row = decoded * stride;
      for (let x = 0; x < stride; x++) {
        const a = x >= step ? out[row + x - step]! : 0;
        const b = decoded > 0 ? out[row - stride + x]! : 0;
        const c = x >= step && decoded > 0 ? out[row - stride + x - step]! : 0;
        out[row + x] = (raw[line + x]! + predictor(filter, a, b, c)) & 0xff;
      }
    }
  };
  const rgb = (x: number, y: number): number => {
    if (x < 0 || y < 0 || x >= width || y >= height) {
      throw new Error(
        `pixel (${String(x)},${String(y)}) is outside ${String(width)}x${String(height)}`,
      );
    }
    ensure(y);
    if (packed !== null) {
      const bit = x * depth;
      const byte = out[y * stride + (bit >> 3)]!;
      const value = packed[(byte >> (8 - depth - (bit & 7))) & ((1 << depth) - 1)];
      if (value === undefined)
        throw new Error(`palette index out of range at ${String(x)},${String(y)}`);
      return value;
    }
    const at = y * stride + x * channels;
    return (out[at]! << 16) | (out[at + 1]! << 8) | out[at + 2]!;
  };
  return {
    width,
    height,
    rgb,
    pixel(x: number, y: number): [number, number, number] {
      const v = rgb(x, y);
      return [(v >> 16) & 0xff, (v >> 8) & 0xff, v & 0xff];
    },
  };
}
