/**
 * A streaming ZIP32 writer.
 *
 * `tools/ground-truth/lib/zip.ts` already writes archives, and this is deliberately
 * not that one, for two reasons that only appear at this size:
 *
 * - **It never holds the archive.** That writer concatenates every local header
 *   and every payload into one buffer and then copies the lot into a second. At
 *   200 MB that is 400 MB of peak allocation to produce a file we are about to
 *   read back anyway. This one writes each entry to the file descriptor as it
 *   is built and keeps only the central directory, which is a few hundred
 *   kilobytes.
 * - **It lets the caller refuse DEFLATE.** That writer compresses every entry
 *   at level 9 to find out whether compression helped. Most of a realistic deck
 *   is media that is already compressed, and running DEFLATE over 195 MB of
 *   incompressible bytes to learn that it did not help costs minutes.
 *
 * ZIP32 only, as the plan requires of our own writer: no ZIP64, no directory
 * entries, no data descriptors. A benchmark deck at 200 MB with a few thousand
 * entries is nowhere near the 4 GB / 65535 ceilings.
 *
 * Timestamps are fixed, so the same recipe produces the same bytes and the
 * SHA-256 in `corpus/bench/manifest.json` is reproducible.
 */

import { closeSync, openSync, writeSync } from 'node:fs';
import { crc32, deflateRawSync } from 'node:zlib';

const LOCAL_SIG = 0x04034b50;
const CENTRAL_SIG = 0x02014b50;
const EOCD_SIG = 0x06054b50;

/** 1980-01-01T00:00:00, the earliest a DOS timestamp can express. */
const DOS_TIME = 0;
const DOS_DATE = 0x0021;

const U16_MAX = 0xffff;
const U32_MAX = 0xffffffff;

const encoder = new TextEncoder();

export interface AddOptions {
  /**
   * Store rather than deflate.
   *
   * Set for anything already compressed. It is not an optimisation of the
   * archive - it is an optimisation of *writing* the archive, and the
   * difference on 195 MB of PNG noise is minutes.
   */
  readonly store?: boolean;
  readonly level?: number;
}

interface CentralEntry {
  readonly name: Uint8Array;
  readonly method: number;
  readonly crc: number;
  readonly compressedSize: number;
  readonly uncompressedSize: number;
  readonly offset: number;
}

export class ZipStream {
  readonly #fd: number;
  readonly #central: CentralEntry[] = [];
  #offset = 0;
  #closed = false;

  constructor(path: string) {
    this.#fd = openSync(path, 'w');
  }

  /** Bytes written so far, not counting the central directory still to come. */
  get bytesWritten(): number {
    return this.#offset;
  }

  get entryCount(): number {
    return this.#central.length;
  }

  add(name: string, content: string | Uint8Array, options: AddOptions = {}): void {
    if (this.#closed) throw new Error('the archive is closed');
    const bytes = typeof content === 'string' ? encoder.encode(content) : content;
    const nameBytes = encoder.encode(name);
    if (nameBytes.length > U16_MAX) throw new Error('entry name is too long: ' + name);

    const checksum = crc32(bytes);
    let payload = bytes;
    let method = 0;
    if (options.store !== true) {
      const deflated = deflateRawSync(bytes, { level: options.level ?? 6 });
      if (deflated.length < bytes.length) {
        payload = new Uint8Array(deflated.buffer, deflated.byteOffset, deflated.byteLength);
        method = 8;
      }
    }

    const header = new Uint8Array(30 + nameBytes.length);
    const view = new DataView(header.buffer);
    view.setUint32(0, LOCAL_SIG, true);
    view.setUint16(4, 20, true); // version needed to extract: 2.0
    view.setUint16(6, 0, true); // no flags: no encryption, no data descriptor
    view.setUint16(8, method, true);
    view.setUint16(10, DOS_TIME, true);
    view.setUint16(12, DOS_DATE, true);
    view.setUint32(14, checksum, true);
    view.setUint32(18, payload.length, true);
    view.setUint32(22, bytes.length, true);
    view.setUint16(26, nameBytes.length, true);
    view.setUint16(28, 0, true); // no extra field
    header.set(nameBytes, 30);

    this.#central.push({
      name: nameBytes,
      method,
      crc: checksum,
      compressedSize: payload.length,
      uncompressedSize: bytes.length,
      offset: this.#offset,
    });

    writeSync(this.#fd, header);
    writeSync(this.#fd, payload);
    this.#offset += header.length + payload.length;
    if (this.#offset > U32_MAX) {
      throw new Error('the archive passed 4 GB, which ZIP32 cannot address');
    }
  }

  close(): { bytes: number; entries: number } {
    if (this.#closed) throw new Error('the archive is already closed');
    if (this.#central.length > U16_MAX) {
      throw new Error(
        'a ZIP32 central directory holds at most ' +
          String(U16_MAX) +
          ' entries and this archive has ' +
          String(this.#central.length),
      );
    }

    const directoryOffset = this.#offset;
    for (const entry of this.#central) {
      const record = new Uint8Array(46 + entry.name.length);
      const view = new DataView(record.buffer);
      view.setUint32(0, CENTRAL_SIG, true);
      view.setUint16(4, 20, true); // version made by
      view.setUint16(6, 20, true); // version needed
      view.setUint16(8, 0, true);
      view.setUint16(10, entry.method, true);
      view.setUint16(12, DOS_TIME, true);
      view.setUint16(14, DOS_DATE, true);
      view.setUint32(16, entry.crc, true);
      view.setUint32(20, entry.compressedSize, true);
      view.setUint32(24, entry.uncompressedSize, true);
      view.setUint16(28, entry.name.length, true);
      view.setUint32(42, entry.offset, true);
      record.set(entry.name, 46);
      writeSync(this.#fd, record);
      this.#offset += record.length;
    }

    const eocd = new Uint8Array(22);
    const view = new DataView(eocd.buffer);
    view.setUint32(0, EOCD_SIG, true);
    view.setUint16(8, this.#central.length, true);
    view.setUint16(10, this.#central.length, true);
    view.setUint32(12, this.#offset - directoryOffset, true);
    view.setUint32(16, directoryOffset, true);
    writeSync(this.#fd, eocd);
    this.#offset += eocd.length;

    closeSync(this.#fd);
    this.#closed = true;
    return { bytes: this.#offset, entries: this.#central.length };
  }
}
