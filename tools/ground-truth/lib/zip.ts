/**
 * A minimal read/write ZIP for the ground-truth tools.
 *
 * `@pptx-studio/opc` already has a far better ZIP reader, with decompression
 * budgets and a part-name grammar. This is deliberately *not* that: `tools/`
 * stays self-contained, the same call `schema-codegen` made. Pointing a tool at
 * a package's source drags that package's module resolution into the tools
 * TypeScript project, and pointing it at `dist` makes `pnpm typecheck` depend
 * on a build. Neither is worth it for eighty lines.
 *
 * The corollary: nothing here is hardened. It reads archives this repository
 * produced or PowerPoint produced, on a developer's machine, by hand. Hostile
 * input is `opc`'s job and `opc` does it properly.
 */

import { deflateRawSync, inflateRawSync } from 'node:zlib';

const LOCAL_SIG = 0x04034b50;
const CENTRAL_SIG = 0x02014b50;
const EOCD_SIG = 0x06054b50;

export interface ZipEntry {
  readonly name: string;
  readonly bytes: Uint8Array;
  /**
   * Store this entry rather than letting the writer choose.
   *
   * Set by the corpus generator, and for a reason that outlives the entry: a
   * deflated archive's bytes depend on which zlib built it, so a fixture whose
   * SHA-256 is pinned would need re-pinning after a routine Node upgrade with
   * no source change. Stored bytes are a pure function of the XML above them.
   * `readZip` has always handled both methods, so nothing else changes.
   */
  readonly store?: boolean;
  /**
   * The general-purpose bit flag, written to both headers. Defaults to 0.
   *
   * PowerPoint 16.0.20326 writes `0x0006` on every entry - bits 1 and 2, which
   * for method 8 are the deflate level hint and mean nothing to a decompressor.
   * Bit 11 is the one with consequences: it declares the entry name to be
   * UTF-8 rather than CP437. `a35-zip-shapes` is the only deck that sets
   * either, and `corpus/ground-truth/powerpoint-conventions.json` is where the
   * measurement lives.
   *
   * Bit 3 is not supported: it moves the sizes into a trailing data descriptor,
   * which this writer does not emit, so setting it would produce an archive
   * that lies about itself.
   */
  readonly flags?: number;
  /**
   * Extra-field bytes for the **local** header only, verbatim.
   *
   * The one this repository cares about is Microsoft's `0xA220` growth hint:
   * a header id, a length, the signature `0xA028`, a padding value and that
   * many zero bytes, so an editor can rewrite a part slightly larger in place
   * without moving every entry after it. PowerPoint writes 520 bytes of it on
   * the first entry of a package. Readers ignore what they do not recognise,
   * which is exactly why an unrecognised extra field is worth having in the
   * corpus: it is bytes between the local header and the payload that a naive
   * reader will walk straight through.
   */
  readonly extra?: Uint8Array;
}

/**
 * Microsoft's `0xA220` "growth hint" extra field, at a total length in bytes.
 *
 * Layout: header id `0xA220`, data length, signature `0xA028`, the padding
 * value, then that many zero bytes. `total` counts the four header bytes, so
 * PowerPoint's 520 is 4 + 2 + 2 + 512.
 */
export function growthHint(total = 520): Uint8Array {
  if (total < 8) throw new Error('a growth hint needs at least 8 bytes');
  const field = new Uint8Array(total);
  const view = new DataView(field.buffer);
  view.setUint16(0, 0xa220, true);
  view.setUint16(2, total - 4, true);
  view.setUint16(4, 0xa028, true);
  view.setUint16(6, total - 8, true);
  return field;
}

/** Every entry of an archive, in central-directory order. */
export function readZip(archive: Uint8Array): ZipEntry[] {
  const view = new DataView(archive.buffer, archive.byteOffset, archive.byteLength);

  let eocd = -1;
  for (let i = archive.length - 22; i >= 0; i--) {
    if (view.getUint32(i, true) === EOCD_SIG) {
      eocd = i;
      break;
    }
  }
  if (eocd < 0) throw new Error('not a zip archive: no end-of-central-directory record');

  const count = view.getUint16(eocd + 10, true);
  let offset = view.getUint32(eocd + 16, true);

  const entries: ZipEntry[] = [];
  for (let i = 0; i < count; i++) {
    if (view.getUint32(offset, true) !== CENTRAL_SIG) {
      throw new Error(`central directory entry ${String(i)} has a bad signature`);
    }
    const flags = view.getUint16(offset + 8, true);
    const method = view.getUint16(offset + 10, true);
    const compressedSize = view.getUint32(offset + 20, true);
    const nameLength = view.getUint16(offset + 28, true);
    const extraLength = view.getUint16(offset + 30, true);
    const commentLength = view.getUint16(offset + 32, true);
    const localOffset = view.getUint32(offset + 42, true);
    const name = new TextDecoder().decode(archive.subarray(offset + 46, offset + 46 + nameLength));

    if (view.getUint32(localOffset, true) !== LOCAL_SIG) {
      throw new Error(`local header for ${name} has a bad signature`);
    }
    const localNameLength = view.getUint16(localOffset + 26, true);
    const localExtraLength = view.getUint16(localOffset + 28, true);
    const start = localOffset + 30 + localNameLength + localExtraLength;
    const raw = archive.subarray(start, start + compressedSize);

    entries.push({
      name,
      bytes: method === 0 ? raw : new Uint8Array(inflateRawSync(raw)),
      store: method === 0,
      flags,
      ...(localExtraLength === 0
        ? {}
        : {
            extra: archive.slice(
              localOffset + 30 + localNameLength,
              localOffset + 30 + localNameLength + localExtraLength,
            ),
          }),
    });
    offset += 46 + nameLength + extraLength + commentLength;
  }
  return entries;
}

/** The one entry with this name, or `undefined`. */
export function entry(entries: readonly ZipEntry[], name: string): Uint8Array | undefined {
  return entries.find((e) => e.name === name)?.bytes;
}

/* -------------------------------------------------------------------------- */
/* writing                                                                    */
/* -------------------------------------------------------------------------- */

// A fixed DOS timestamp, so building the same package twice gives the same
// bytes. 1980-01-01 00:00:00 is the earliest the format can express.
const DOS_TIME = 0;
const DOS_DATE = 0x0021;

const EMPTY = new Uint8Array(0);

let crcTable: Uint32Array | undefined;

function crc32(bytes: Uint8Array): number {
  if (crcTable === undefined) {
    crcTable = new Uint32Array(256);
    for (let n = 0; n < 256; n++) {
      let c = n;
      for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
      crcTable[n] = c >>> 0;
    }
  }
  let crc = 0xffffffff;
  for (const byte of bytes) crc = crcTable[(crc ^ byte) & 0xff]! ^ (crc >>> 8);
  return (crc ^ 0xffffffff) >>> 0;
}

/**
 * Write a ZIP32 archive: no directory entries, no ZIP64, entries in the order
 * given. `[Content_Types].xml` must be first - that is an OPC rule, not a ZIP
 * one, and it is the caller's business.
 */
export function writeZip(entries: readonly ZipEntry[]): Uint8Array {
  const locals: Uint8Array[] = [];
  const centrals: Uint8Array[] = [];
  let offset = 0;

  for (const e of entries) {
    const name = new TextEncoder().encode(e.name);
    const crc = crc32(e.bytes);
    // Store when the caller says so, or when deflate would not help; every OPC
    // reader handles both, and a stored entry is easier to look at in a hex
    // editor when something is wrong.
    const deflated =
      e.store === true ? undefined : new Uint8Array(deflateRawSync(e.bytes, { level: 9 }));
    const useDeflate = deflated !== undefined && deflated.length < e.bytes.length;
    const payload = useDeflate && deflated !== undefined ? deflated : e.bytes;
    const method = useDeflate ? 8 : 0;

    const flags = e.flags ?? 0;
    const extra = e.extra ?? EMPTY;

    const local = new Uint8Array(30 + name.length + extra.length + payload.length);
    const lv = new DataView(local.buffer);
    lv.setUint32(0, LOCAL_SIG, true);
    lv.setUint16(4, 20, true); // version needed: 2.0
    lv.setUint16(6, flags, true);
    lv.setUint16(8, method, true);
    lv.setUint16(10, DOS_TIME, true);
    lv.setUint16(12, DOS_DATE, true);
    lv.setUint32(14, crc, true);
    lv.setUint32(18, payload.length, true);
    lv.setUint32(22, e.bytes.length, true);
    lv.setUint16(26, name.length, true);
    lv.setUint16(28, extra.length, true);
    local.set(name, 30);
    local.set(extra, 30 + name.length);
    local.set(payload, 30 + name.length + extra.length);
    locals.push(local);

    const central = new Uint8Array(46 + name.length);
    const cv = new DataView(central.buffer);
    cv.setUint32(0, CENTRAL_SIG, true);
    cv.setUint16(4, 20, true); // version made by
    cv.setUint16(6, 20, true); // version needed
    cv.setUint16(8, flags, true);
    cv.setUint16(10, method, true);
    cv.setUint16(12, DOS_TIME, true);
    cv.setUint16(14, DOS_DATE, true);
    cv.setUint32(16, crc, true);
    cv.setUint32(20, payload.length, true);
    cv.setUint32(24, e.bytes.length, true);
    cv.setUint16(28, name.length, true);
    cv.setUint32(42, offset, true);
    central.set(name, 46);
    centrals.push(central);

    offset += local.length;
  }

  const centralSize = centrals.reduce((n, c) => n + c.length, 0);
  const eocd = new Uint8Array(22);
  const ev = new DataView(eocd.buffer);
  ev.setUint32(0, EOCD_SIG, true);
  ev.setUint16(8, entries.length, true);
  ev.setUint16(10, entries.length, true);
  ev.setUint32(12, centralSize, true);
  ev.setUint32(16, offset, true);

  const total = offset + centralSize + eocd.length;
  const out = new Uint8Array(total);
  let at = 0;
  for (const part of [...locals, ...centrals, eocd]) {
    out.set(part, at);
    at += part.length;
  }
  return out;
}
