/**
 * CRC-32 (IEEE 802.3, reflected, polynomial 0xEDB88320) - the checksum ZIP
 * stores for every entry.
 *
 * This exists because `fflate` does not export one, and because without it the
 * reader's zip-bomb defence is silently lossy. Inflating into a
 * caller-supplied output buffer caps allocation, but when the DEFLATE stream
 * produces more bytes than the buffer holds, the surplus is **dropped without
 * error** - the reader gets short data and no signal. Verifying the CRC is what
 * turns that silent truncation into `ERR_CRC_MISMATCH`. The cap makes the
 * reader safe; the checksum makes it honest.
 *
 * The writer in 0.3 needs the same function, from the same implementation, so
 * that what we verify on read is what we emit on write.
 */

/** 256-entry lookup table, built once on first use. */
let table: Uint32Array | undefined;

function crcTable(): Uint32Array {
  if (table) return table;
  const next = new Uint32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) {
      c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    }
    next[n] = c >>> 0;
  }
  table = next;
  return next;
}

/**
 * CRC-32 of `data`, as an unsigned 32-bit number.
 *
 * `seed` lets a caller checksum a byte stream in chunks: pass the previous
 * result back in. It is the finalised value, not the internal register, so
 * `crc32(b, crc32(a))` equals `crc32(concat(a, b))`.
 */
export function crc32(data: Uint8Array, seed = 0): number {
  const t = crcTable();
  let c = ~seed >>> 0;
  for (let i = 0; i < data.length; i++) {
    c = (c >>> 8) ^ t[(c ^ data[i]!) & 0xff]!;
  }
  return ~c >>> 0;
}
