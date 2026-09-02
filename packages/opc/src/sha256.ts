/**
 * SHA-256, synchronous, in about a hundred lines.
 *
 * ## Why not `crypto.subtle.digest`
 *
 * Two reasons, and the second is the one that decided it.
 *
 * It is asynchronous. Everything that wants a digest here - the round-trip
 * comparator, `cli bisect` in 1.5, the media dedupe in 10.8, the FontFace
 * registry name in 8.6 - is otherwise synchronous, and making them all async to
 * reach a hash turns one library call into a colour change across four
 * packages.
 *
 * And it is not always there. `crypto.subtle` is gated on a **secure context**:
 * present on `https:` and on `http://localhost`, absent on a plain `http:`
 * origin. A library that renders a `.pptx` entirely in the browser has no say
 * in where it is served from, and "the digest works on my machine and is
 * `undefined` on the intranet deployment" is a failure mode that shows up
 * exactly once, in production, at somebody else's site.
 *
 * ## Why it is in this package
 *
 * `@pptx-studio/opc` is the only layer-0 package that deals in bytes at all -
 * `@pptx-studio/xml` works in text - and it already exports one digest,
 * {@link crc32}, because the ZIP format requires it. A second one is at home
 * beside it. The alternative is a utility package holding two functions.
 *
 * The two are not interchangeable and the difference matters where they are
 * used. CRC-32 is a 32-bit error-detecting code: a birthday collision is
 * expected within about 2^16 inputs, which is fewer parts than a large corpus
 * has. It is right for "did this entry survive the archive" and wrong for
 * "these two media parts are the same file".
 *
 * FIPS 180-4. Verified against the standard's own two test vectors and the
 * empty-string vector in `sha256.test.ts`.
 */

/** Round constants: the first 32 bits of the fractional parts of the cube roots of the first 64 primes. */
const K = /* @__PURE__ */ new Uint32Array([
  0x428a2f98, 0x71374491, 0xb5c0fbcf, 0xe9b5dba5, 0x3956c25b, 0x59f111f1, 0x923f82a4, 0xab1c5ed5,
  0xd807aa98, 0x12835b01, 0x243185be, 0x550c7dc3, 0x72be5d74, 0x80deb1fe, 0x9bdc06a7, 0xc19bf174,
  0xe49b69c1, 0xefbe4786, 0x0fc19dc6, 0x240ca1cc, 0x2de92c6f, 0x4a7484aa, 0x5cb0a9dc, 0x76f988da,
  0x983e5152, 0xa831c66d, 0xb00327c8, 0xbf597fc7, 0xc6e00bf3, 0xd5a79147, 0x06ca6351, 0x14292967,
  0x27b70a85, 0x2e1b2138, 0x4d2c6dfc, 0x53380d13, 0x650a7354, 0x766a0abb, 0x81c2c92e, 0x92722c85,
  0xa2bfe8a1, 0xa81a664b, 0xc24b8b70, 0xc76c51a3, 0xd192e819, 0xd6990624, 0xf40e3585, 0x106aa070,
  0x19a4c116, 0x1e376c08, 0x2748774c, 0x34b0bcb5, 0x391c0cb3, 0x4ed8aa4a, 0x5b9cca4f, 0x682e6ff3,
  0x748f82ee, 0x78a5636f, 0x84c87814, 0x8cc70208, 0x90befffa, 0xa4506ceb, 0xbef9a3f7, 0xc67178f2,
]);

/** Initial hash value: the first 32 bits of the fractional parts of the square roots of the first 8 primes. */
const H0 = /* @__PURE__ */ new Uint32Array([
  0x6a09e667, 0xbb67ae85, 0x3c6ef372, 0xa54ff53a, 0x510e527f, 0x9b05688c, 0x1f83d9ab, 0x5be0cd19,
]);

/**
 * One 64-byte block, in place.
 *
 * `h` and `w` are supplied by the caller and reused across every block of a
 * message, because allocating a 64-word schedule per block allocates sixteen
 * million times on a gigabyte and that cost is visible.
 */
function compress(h: Uint32Array, w: Uint32Array, block: Uint8Array, at: number): void {
  for (let i = 0; i < 16; i++) {
    const j = at + i * 4;
    w[i] = ((block[j]! << 24) | (block[j + 1]! << 16) | (block[j + 2]! << 8) | block[j + 3]!) >>> 0;
  }
  for (let i = 16; i < 64; i++) {
    const x = w[i - 15]!;
    const y = w[i - 2]!;
    const s0 = ((x >>> 7) | (x << 25)) ^ ((x >>> 18) | (x << 14)) ^ (x >>> 3);
    const s1 = ((y >>> 17) | (y << 15)) ^ ((y >>> 19) | (y << 13)) ^ (y >>> 10);
    w[i] = (w[i - 16]! + s0 + w[i - 7]! + s1) >>> 0;
  }

  let a = h[0]!;
  let b = h[1]!;
  let c = h[2]!;
  let d = h[3]!;
  let e = h[4]!;
  let f = h[5]!;
  let g = h[6]!;
  let hh = h[7]!;

  for (let i = 0; i < 64; i++) {
    const s1 = ((e >>> 6) | (e << 26)) ^ ((e >>> 11) | (e << 21)) ^ ((e >>> 25) | (e << 7));
    const ch = (e & f) ^ (~e & g);
    const t1 = (hh + s1 + ch + K[i]! + w[i]!) >>> 0;
    const s0 = ((a >>> 2) | (a << 30)) ^ ((a >>> 13) | (a << 19)) ^ ((a >>> 22) | (a << 10));
    const maj = (a & b) ^ (a & c) ^ (b & c);
    const t2 = (s0 + maj) >>> 0;
    hh = g;
    g = f;
    f = e;
    e = (d + t1) >>> 0;
    d = c;
    c = b;
    b = a;
    a = (t1 + t2) >>> 0;
  }

  h[0] = (h[0]! + a) >>> 0;
  h[1] = (h[1]! + b) >>> 0;
  h[2] = (h[2]! + c) >>> 0;
  h[3] = (h[3]! + d) >>> 0;
  h[4] = (h[4]! + e) >>> 0;
  h[5] = (h[5]! + f) >>> 0;
  h[6] = (h[6]! + g) >>> 0;
  h[7] = (h[7]! + hh) >>> 0;
}

/**
 * The 32-byte digest of `bytes`.
 *
 * Whole blocks are read straight out of the input; only the final partial block
 * is copied, so digesting a 200 MB part allocates 128 bytes rather than 200 MB.
 */
export function sha256(bytes: Uint8Array): Uint8Array {
  const h = new Uint32Array(H0);
  const w = new Uint32Array(64);

  const whole = bytes.length - (bytes.length % 64);
  for (let at = 0; at < whole; at += 64) compress(h, w, bytes, at);

  // The tail: what is left, a 0x80 byte, zero padding, and the message length
  // in bits as a 64-bit big-endian integer in the last eight bytes. One block
  // if the length fits after the marker, two if it does not.
  const rest = bytes.length - whole;
  const tail = new Uint8Array(rest < 56 ? 64 : 128);
  tail.set(bytes.subarray(whole));
  tail[rest] = 0x80;

  // `length * 8` overflows 32 bits above 512 MB, and the inflation budget
  // allows a part twice that. The high word is `length / 2 ** 29`, which is
  // exact for any length a Uint8Array can hold.
  const high = Math.floor(bytes.length / 0x20000000);
  const low = (bytes.length * 8) >>> 0;
  const end = tail.length;
  tail[end - 8] = (high >>> 24) & 0xff;
  tail[end - 7] = (high >>> 16) & 0xff;
  tail[end - 6] = (high >>> 8) & 0xff;
  tail[end - 5] = high & 0xff;
  tail[end - 4] = (low >>> 24) & 0xff;
  tail[end - 3] = (low >>> 16) & 0xff;
  tail[end - 2] = (low >>> 8) & 0xff;
  tail[end - 1] = low & 0xff;

  for (let at = 0; at < tail.length; at += 64) compress(h, w, tail, at);

  const digest = new Uint8Array(32);
  for (let i = 0; i < 8; i++) {
    const value = h[i]!;
    digest[i * 4] = (value >>> 24) & 0xff;
    digest[i * 4 + 1] = (value >>> 16) & 0xff;
    digest[i * 4 + 2] = (value >>> 8) & 0xff;
    digest[i * 4 + 3] = value & 0xff;
  }
  return digest;
}

const HEX = '0123456789abcdef';

/** The digest of `bytes` as 64 lowercase hex digits - the spelling the corpus manifests use. */
export function sha256Hex(bytes: Uint8Array): string {
  const digest = sha256(bytes);
  let out = '';
  for (const byte of digest) out += HEX[byte >>> 4]! + HEX[byte & 0x0f]!;
  return out;
}

/**
 * The digest of a string, encoded UTF-8 first.
 *
 * `TextEncoder` rather than a hand-rolled encoder: it is in every browser and
 * every Worker, and it is not gated on a secure context the way `crypto.subtle`
 * is.
 */
export function sha256HexOfText(text: string): string {
  return sha256Hex(new TextEncoder().encode(text));
}
