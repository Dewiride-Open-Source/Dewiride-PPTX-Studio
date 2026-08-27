/**
 * Valid PNGs of a chosen size, made of noise.
 *
 * A benchmark deck has to be 200 MB and a real 200 MB deck is almost entirely
 * media, so the media has to be both **large** and **incompressible**. Anything
 * else measures the wrong thing: a deck padded with zeroes deflates to nothing
 * and the ZIP reader never does the work a real file makes it do.
 *
 * PNG rather than JPEG because a valid PNG is a header, a zlib stream and a
 * trailer, and a valid baseline JPEG is a Huffman-coded entropy stream. The
 * point here is a file PowerPoint will open, not a photograph.
 *
 * The zlib stream is written at level 0 - stored DEFLATE blocks - which is the
 * only reason this is fast. Compressing noise costs real time and produces
 * output very slightly larger than the input; storing it costs a memcpy.
 *
 * Deterministic: the same seed gives the same bytes, so a deck built twice from
 * one recipe hashes the same and `corpus/bench/manifest.json` can pin it.
 */

import { crc32, deflateSync } from 'node:zlib';

const encoder = new TextEncoder();

const PNG_SIGNATURE = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);

/** mulberry32: small, fast, and identical on every platform we might run on. */
export function mulberry32(seed: number): () => number {
  let state = seed >>> 0;
  return () => {
    state = (state + 0x6d2b79f5) >>> 0;
    let t = state;
    t = Math.imul(t ^ (t >>> 15), t | 1);
    t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

function chunk(type: string, data: Uint8Array): Uint8Array {
  const typeBytes = encoder.encode(type);
  const out = new Uint8Array(12 + data.length);
  const view = new DataView(out.buffer);
  view.setUint32(0, data.length, false);
  out.set(typeBytes, 4);
  out.set(data, 8);
  // The CRC covers the type and the data, and not the length. Getting that
  // wrong produces a file every decoder rejects and no decoder explains.
  const body = new Uint8Array(4 + data.length);
  body.set(typeBytes, 0);
  body.set(data, 4);
  view.setUint32(8 + data.length, crc32(body), false);
  return out;
}

/**
 * A truecolour PNG of `size` x `size` pixels filled with noise.
 *
 * Roughly `size * size * 3 + size` bytes on disk. 465 gives about 634 KiB,
 * which is a plausible size for a photograph on a slide.
 */
export function noisePng(size: number, seed: number): Uint8Array {
  const random = mulberry32(seed);

  // One filter byte per scanline (0 = None) followed by RGB triples.
  const stride = size * 3 + 1;
  const raw = new Uint8Array(stride * size);
  for (let y = 0; y < size; y++) {
    const row = y * stride;
    raw[row] = 0;
    for (let x = 1; x < stride; x++) {
      raw[row + x] = (random() * 256) | 0;
    }
  }

  const ihdr = new Uint8Array(13);
  const header = new DataView(ihdr.buffer);
  header.setUint32(0, size, false);
  header.setUint32(4, size, false);
  ihdr[8] = 8; // bit depth
  ihdr[9] = 2; // colour type 2: truecolour RGB
  ihdr[10] = 0; // deflate
  ihdr[11] = 0; // adaptive filtering
  ihdr[12] = 0; // no interlace

  const deflated = deflateSync(raw, { level: 0 });
  const idat = new Uint8Array(deflated.buffer, deflated.byteOffset, deflated.byteLength);

  const parts = [
    PNG_SIGNATURE,
    chunk('IHDR', ihdr),
    chunk('IDAT', idat),
    chunk('IEND', new Uint8Array(0)),
  ];
  const total = parts.reduce((sum, part) => sum + part.length, 0);
  const png = new Uint8Array(total);
  let at = 0;
  for (const part of parts) {
    png.set(part, at);
    at += part.length;
  }
  return png;
}

/** The pixel width whose noise PNG comes closest to `bytes` on disk. */
export function sizeForBytes(bytes: number): number {
  // size*size*3 + size raw, plus ~5 bytes per 65535-byte stored block, plus
  // about 60 bytes of chunk overhead. Solving the quadratic is close enough:
  // the caller wants "roughly this big", never "exactly".
  return Math.max(2, Math.round((Math.sqrt(1 + 12 * bytes) - 1) / 6));
}
