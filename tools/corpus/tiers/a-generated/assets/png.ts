/**
 * A tiny deterministic PNG, for the probes that need something to fill with.
 *
 * `tools/bench/png.ts` already writes PNGs and this deliberately does not use
 * it: what a benchmark deck needs is *incompressible noise*, so the ZIP layer
 * does the work a real photograph makes it do. What a fill probe needs is the
 * opposite - an image whose orientation is obvious, so that a wrong `a:tile`
 * flip or an ignored `a:srcRect` crop is visible rather than plausible. Noise
 * is the one image in which a crop bug cannot be seen.
 *
 * Deflate level 0 - stored blocks - so the bytes do not depend on which zlib
 * built them, for the same reason the corpus archives are stored.
 */

import { deflateSync } from 'node:zlib';

const SIGNATURE = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);

let table: Uint32Array | undefined;

function crc32(bytes: Uint8Array): number {
  if (table === undefined) {
    table = new Uint32Array(256);
    for (let n = 0; n < 256; n++) {
      let c = n;
      for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
      table[n] = c >>> 0;
    }
  }
  let crc = 0xffffffff;
  for (const byte of bytes) crc = table[(crc ^ byte) & 0xff]! ^ (crc >>> 8);
  return (crc ^ 0xffffffff) >>> 0;
}

function chunk(type: string, data: Uint8Array): Uint8Array {
  const encoder = new TextEncoder();
  const typeBytes = encoder.encode(type);
  const out = new Uint8Array(12 + data.length);
  const view = new DataView(out.buffer);
  view.setUint32(0, data.length, false);
  out.set(typeBytes, 4);
  out.set(data, 8);
  // The CRC covers the type and the data, never the length.
  const body = new Uint8Array(4 + data.length);
  body.set(typeBytes, 0);
  body.set(data, 4);
  view.setUint32(8 + data.length, crc32(body), false);
  return out;
}

/** A truecolour PNG from a pixel function. */
export function png(
  width: number,
  height: number,
  pixel: (x: number, y: number) => number,
): Uint8Array {
  const stride = width * 3 + 1;
  const raw = new Uint8Array(stride * height);
  for (let y = 0; y < height; y++) {
    const row = y * stride;
    raw[row] = 0; // filter: none
    for (let x = 0; x < width; x++) {
      const rgb = pixel(x, y);
      raw[row + 1 + x * 3] = (rgb >> 16) & 0xff;
      raw[row + 2 + x * 3] = (rgb >> 8) & 0xff;
      raw[row + 3 + x * 3] = rgb & 0xff;
    }
  }

  const ihdr = new Uint8Array(13);
  const header = new DataView(ihdr.buffer);
  header.setUint32(0, width, false);
  header.setUint32(4, height, false);
  ihdr[8] = 8; // bit depth
  ihdr[9] = 2; // colour type 2: truecolour RGB
  ihdr[10] = 0;
  ihdr[11] = 0;
  ihdr[12] = 0;

  const deflated = deflateSync(raw, { level: 0 });
  const parts = [
    SIGNATURE,
    chunk('IHDR', ihdr),
    chunk('IDAT', new Uint8Array(deflated.buffer, deflated.byteOffset, deflated.byteLength)),
    chunk('IEND', new Uint8Array(0)),
  ];
  const total = parts.reduce((sum, part) => sum + part.length, 0);
  const out = new Uint8Array(total);
  let at = 0;
  for (const part of parts) {
    out.set(part, at);
    at += part.length;
  }
  return out;
}

/**
 * 32 x 32, four coloured quadrants with a white L in the top-left one.
 *
 * The L is the part that matters. Four quadrants alone cannot distinguish a
 * horizontal flip from a vertical one plus a rotation; an asymmetric mark can.
 * `a:tile/@flip`, `a:srcRect` and `a:blip`'s colour effects are all probed
 * against this image, and every one of them is a no-op on a symmetric picture.
 */
export function quadrantPng(): Uint8Array {
  const quadrants = [0xc0392b, 0x27ae60, 0x2980b9, 0xf1c40f];
  return png(32, 32, (x, y) => {
    // The mark: a 3px-wide L inside the top-left quadrant.
    if (x >= 4 && x < 7 && y >= 4 && y < 12) return 0xffffff;
    if (x >= 4 && x < 12 && y >= 9 && y < 12) return 0xffffff;
    return quadrants[(y < 16 ? 0 : 2) + (x < 16 ? 0 : 1)] ?? 0;
  });
}
