import { deflateSync } from 'fflate';
import { crc32 } from '../digest/crc32.js';
import { guard, OpcError } from '../errors.js';
import {
  COMPRESSION_DEFLATE,
  COMPRESSION_STORE,
  type ZipArchive,
  type ZipEntry,
} from './zip-reader.js';

/**
 * The ZIP32 writer - the exact inverse of `zip-reader.ts`.
 *
 * Every header field below is measured against a file PowerPoint 365 wrote on
 * this machine - `corpus/authored/b01-blank.pptx` - rather than chosen from the
 * specification's range of legal values. On four of the seven we write what
 * PowerPoint writes. On three we do not, and those are the interesting rows:
 *
 * | field                   | PowerPoint            | us            |
 * | ----------------------- | --------------------- | ------------- |
 * | version made by         | 45, host 0            | **20**        |
 * | version needed          | 10 stored, 20 deflate | same          |
 * | general purpose flags   | 0x0006 deflated, 0 stored | **0**     |
 * | last modified           | 1980-01-01 00:00      | same          |
 * | internal/external attrs | 0                     | same          |
 * | extra field             | 0xA220 growth hint    | **none**      |
 * | entry comment           | empty                 | same          |
 *
 * The three divergences are deliberate, and each is a claim we decline to make
 * rather than a field we forgot:
 *
 *   - **Version made by** is the ZIP version an archive was *created* with.
 *     PowerPoint's zip library stamps 4.5; we emit no ZIP64 record, no data
 *     descriptor and no encryption, so nothing here needs above 2.0 and saying
 *     4.5 would overstate what a reader has to understand.
 *   - **Flags 0x0006** is bits 1 and 2, which for method 8 are a compression
 *     *level hint* and mean nothing to a decompressor - 0b11 is "super fast".
 *     We deflate at level 6, whose encoding is 0b00. Copying 0x0006 would
 *     describe a compressor we are not.
 *   - **The 0xA220 growth hint** is 512 bytes of padding so an editor can
 *     rewrite a part slightly larger *in place* without moving every entry
 *     after it. We rewrite the whole archive on every save, so the padding
 *     would buy nothing and announce a capability we do not have. PowerPoint
 *     writes five of them, 1832 bytes in all, which is the entire size
 *     difference between `b01-blank.pptx` and `c01-opc-writer.pptx`.
 *
 * None of the three is guesswork about whether PowerPoint minds:
 * `corpus/written/c01-opc-writer.pptx` is that same deck rewritten by this
 * writer, committed, and opened in PowerPoint 16.0.20326 with no repair prompt.
 * `tools/corpus/tiers/c-written/decks.test.ts` asserts all three in both directions, so
 * neither our drifting nor PowerPoint's can pass unnoticed.
 *
 * The UTF-8 flag is the fourth bit worth a note, and the one place where our
 * zero and PowerPoint's agree for the same reason. Office never sets bit 11 and
 * does not need to: the OPC part-name grammar admits only `pchar`, so a part
 * name is always ASCII and anything else is percent-encoded before it becomes
 * an entry name. We tested setting it anyway on a real deck and PowerPoint
 * opened the file without complaint - so this too is a choice about honesty
 * rather than compatibility, and the honest flag for an all-ASCII name is zero.
 *
 * `fflate.deflateSync` produces a **raw** DEFLATE stream, verified rather than
 * assumed: its output begins `cb 48` where a zlib-wrapped stream would begin
 * `78 9c`, and Node's `inflateRawSync` round-trips it. A zlib header inside a
 * ZIP entry is a silent corruption that no reader reports as one, so this is
 * checked rather than trusted.
 */

const SIG_LOCAL = 0x04034b50;
const SIG_CENTRAL = 0x02014b50;
const SIG_EOCD = 0x06054b50;

const VERSION_MADE_BY = 20;
const VERSION_NEEDED_STORE = 10;
const VERSION_NEEDED_DEFLATE = 20;
const LOCAL_HEADER_SIZE = 30;
const CENTRAL_HEADER_SIZE = 46;
const EOCD_SIZE = 22;

/** 1980-01-01 00:00:00, the earliest representable DOS timestamp. */
const DOS_TIME = 0;
const DOS_DATE = 33;

const U16_MAX = 0xffff;
const U32_MAX = 0xffffffff;

/** One entry, with its data already in the form it will be stored in. */
export interface ZipEntryInput {
  /** ZIP entry name: a part name without its leading slash. */
  readonly name: string;
  readonly method: number;
  /** Deflated bytes when `method` is 8, the bytes themselves when it is 0. */
  readonly data: Uint8Array;
  readonly crc32: number;
  readonly uncompressedSize: number;
}

class ByteWriter {
  #chunks: Uint8Array[] = [];
  #length = 0;
  readonly #scratch = new DataView(new ArrayBuffer(4));

  get length(): number {
    return this.#length;
  }

  bytes(b: Uint8Array): void {
    this.#chunks.push(b);
    this.#length += b.length;
  }

  u16(v: number): void {
    this.#scratch.setUint16(0, v, true);
    this.bytes(new Uint8Array(this.#scratch.buffer.slice(0, 2)));
  }

  u32(v: number): void {
    this.#scratch.setUint32(0, v >>> 0, true);
    this.bytes(new Uint8Array(this.#scratch.buffer.slice(0, 4)));
  }

  finish(): Uint8Array {
    const out = new Uint8Array(this.#length);
    let at = 0;
    for (const chunk of this.#chunks) {
      out.set(chunk, at);
      at += chunk.length;
    }
    return out;
  }
}

function encodeName(name: string): Uint8Array {
  const bytes = new TextEncoder().encode(name);
  for (const b of bytes) {
    if (b >= 0x80) {
      throw new OpcError(
        'ERR_INVALID_PART_NAME',
        JSON.stringify(name) +
          ' is not an ASCII entry name. Part names are percent-encoded before they reach the ' +
          'archive, so a non-ASCII byte here means something skipped that step.',
        { entry: name },
      );
    }
  }
  if (bytes.length > U16_MAX) {
    throw new OpcError('ERR_INVALID_PART_NAME', 'entry name is longer than 65535 bytes', {
      entry: name.slice(0, 80),
    });
  }
  return bytes;
}

/** An entry stored uncompressed. */
export function storedEntry(name: string, bytes: Uint8Array): ZipEntryInput {
  return {
    name,
    method: COMPRESSION_STORE,
    data: bytes,
    crc32: crc32(bytes),
    uncompressedSize: bytes.length,
  };
}

/**
 * An entry compressed with DEFLATE, or stored when that would not be smaller.
 *
 * The fallback is not an optimisation so much as a correctness convenience: it
 * means a caller never has to decide, and it reproduces what Office does with
 * already-compressed media, which it stores rather than inflating by a few
 * bytes for nothing.
 */
export function deflatedEntry(name: string, bytes: Uint8Array, level = 6): ZipEntryInput {
  const checksum = crc32(bytes);
  const deflated = guard('ERR_INFLATE_FAILED', 'could not deflate ' + name, { entry: name }, () =>
    deflateSync(bytes, { level: level as 0 | 1 | 2 | 3 | 4 | 5 | 6 | 7 | 8 | 9 }),
  );
  if (deflated.length >= bytes.length) {
    return {
      name,
      method: COMPRESSION_STORE,
      data: bytes,
      crc32: checksum,
      uncompressedSize: bytes.length,
    };
  }
  return {
    name,
    method: COMPRESSION_DEFLATE,
    data: deflated,
    crc32: checksum,
    uncompressedSize: bytes.length,
  };
}

/**
 * An entry copied out of a source archive without being decompressed.
 *
 * This is what makes an untouched package come back byte-for-byte: the stored
 * DEFLATE stream, its CRC and both sizes move across together, and nothing
 * re-compresses. The local header offset is deliberately *not* copied, because
 * it is the one field that depends on where the entry lands in the new file.
 */
export function passthroughEntry(archive: ZipArchive, entry: ZipEntry): ZipEntryInput {
  return {
    name: entry.name,
    method: entry.method,
    data: archive.raw(entry),
    crc32: entry.crc32,
    uncompressedSize: entry.uncompressedSize,
  };
}

/**
 * Emit a ZIP32 archive.
 *
 * No directory entries, no ZIP64, no data descriptors, no archive comment. The
 * central directory is written in the same order as the local headers, which
 * is what every real package we have measured does and what makes the two
 * orders interchangeable for a reader.
 */
export function writeZip(entries: readonly ZipEntryInput[]): Uint8Array {
  if (entries.length > U16_MAX) {
    throw new OpcError(
      'ERR_ZIP32_OVERFLOW',
      'a ZIP32 central directory holds at most ' +
        String(U16_MAX) +
        ' entries and this package has ' +
        String(entries.length),
      { limit: U16_MAX, actual: entries.length },
    );
  }

  const w = new ByteWriter();
  const directory: { input: ZipEntryInput; name: Uint8Array; offset: number }[] = [];
  const seen = new Set<string>();

  for (const input of entries) {
    if (seen.has(input.name)) {
      throw new OpcError(
        'ERR_DUPLICATE_ENTRY',
        'entry ' + JSON.stringify(input.name) + ' would be written twice',
        { entry: input.name },
      );
    }
    seen.add(input.name);

    const name = encodeName(input.name);
    const offset = w.length;
    if (offset > U32_MAX) {
      throw new OpcError(
        'ERR_ZIP32_OVERFLOW',
        'the package passes 4 GiB, which needs ZIP64; this writer emits ZIP32 only',
        { entry: input.name, limit: U32_MAX, actual: offset },
      );
    }
    const versionNeeded =
      input.method === COMPRESSION_STORE ? VERSION_NEEDED_STORE : VERSION_NEEDED_DEFLATE;

    w.u32(SIG_LOCAL);
    w.u16(versionNeeded);
    w.u16(0); // general purpose flags
    w.u16(input.method);
    w.u16(DOS_TIME);
    w.u16(DOS_DATE);
    w.u32(input.crc32);
    w.u32(input.data.length);
    w.u32(input.uncompressedSize);
    w.u16(name.length);
    w.u16(0); // extra field length
    w.bytes(name);
    w.bytes(input.data);

    directory.push({ input, name, offset });
  }

  const directoryOffset = w.length;
  if (directoryOffset > U32_MAX) {
    throw new OpcError(
      'ERR_ZIP32_OVERFLOW',
      'the central directory would start past 4 GiB, which needs ZIP64',
      { limit: U32_MAX, actual: directoryOffset },
    );
  }

  for (const { input, name, offset } of directory) {
    const versionNeeded =
      input.method === COMPRESSION_STORE ? VERSION_NEEDED_STORE : VERSION_NEEDED_DEFLATE;
    w.u32(SIG_CENTRAL);
    w.u16(VERSION_MADE_BY);
    w.u16(versionNeeded);
    w.u16(0); // general purpose flags
    w.u16(input.method);
    w.u16(DOS_TIME);
    w.u16(DOS_DATE);
    w.u32(input.crc32);
    w.u32(input.data.length);
    w.u32(input.uncompressedSize);
    w.u16(name.length);
    w.u16(0); // extra field length
    w.u16(0); // file comment length
    w.u16(0); // disk number start
    w.u16(0); // internal file attributes
    w.u32(0); // external file attributes
    w.u32(offset);
    w.bytes(name);
  }

  const directorySize = w.length - directoryOffset;
  w.u32(SIG_EOCD);
  w.u16(0); // this disk
  w.u16(0); // disk with the central directory
  w.u16(directory.length);
  w.u16(directory.length);
  w.u32(directorySize);
  w.u32(directoryOffset);
  w.u16(0); // archive comment length

  return w.finish();
}

/** The byte cost of an entry, for a size estimate that does not build one. */
export function entryOverhead(name: string): number {
  const nameLength = new TextEncoder().encode(name).length;
  return LOCAL_HEADER_SIZE + CENTRAL_HEADER_SIZE + nameLength * 2;
}

/** The fixed cost of an archive with no entries. */
export const EMPTY_ARCHIVE_SIZE = EOCD_SIZE;
