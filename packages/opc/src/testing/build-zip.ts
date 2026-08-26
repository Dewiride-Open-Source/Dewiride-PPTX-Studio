import { deflateSync } from 'fflate';
import { crc32 } from '../crc32.js';

/**
 * A ZIP writer that can lie.
 *
 * Test-only, and not exported from the package entry point, so it never ships.
 * It exists because the reader's job is to survive archives no honest producer
 * would ever emit, and there is no way to test that against archives an honest
 * producer emitted. Every knob here corresponds to one attack the reader claims
 * to stop.
 *
 * The real writer lands in 0.3 and shares none of this code; it has the
 * opposite requirements.
 */

const SIG_LOCAL = 0x04034b50;
const SIG_CENTRAL = 0x02014b50;
const SIG_EOCD = 0x06054b50;
const SIG_ZIP64_EOCD = 0x06064b50;
const SIG_ZIP64_LOCATOR = 0x07064b50;

const U32_MAX = 0xffffffff;

export interface FixtureEntry {
  readonly name: string;
  /** Uncompressed content. Defaults to empty. */
  readonly data?: Uint8Array;
  /** Compress with DEFLATE. Defaults to true when there is data. */
  readonly deflate?: boolean;
  /**
   * DEFLATE level. Level 0 emits stored blocks, which is a different code path
   * inside the inflater and the one that throws a raw `RangeError` when the
   * output buffer is too small.
   */
  readonly deflateLevel?: 0 | 9;
  /** General-purpose bit flags to OR into both headers. */
  readonly flags?: number;
  /** Compression method written to the central directory. Defaults to 0 or 8. */
  readonly method?: number;
  /** Compression method written to the local header, when it should disagree. */
  readonly localMethod?: number;
  /** Entry name written to the local header, when it should disagree. */
  readonly localName?: string;
  /** Uncompressed size claimed by the central directory, overriding the truth. */
  readonly declaredUncompressedSize?: number;
  /** Compressed size claimed by the central directory, overriding the truth. */
  readonly declaredCompressedSize?: number;
  /** CRC-32 claimed by the central directory, overriding the truth. */
  readonly declaredCrc32?: number;
  /**
   * Extend this entry's declared compressed size to swallow the next entry's
   * local header and data - the quoted-overlap shape.
   */
  readonly overlapNext?: boolean;
  /** Write ZIP64 sentinels for the sizes and put the real values in extra field 0x0001. */
  readonly zip64?: boolean;
}

export interface BuildZipOptions {
  /** Archive comment. */
  readonly comment?: Uint8Array;
  /** Entry count written into the EOCD, overriding the truth. */
  readonly declaredEntryCount?: number;
  /** Emit a ZIP64 end-of-central-directory record and locator. */
  readonly zip64Eocd?: boolean;
}

class ByteWriter {
  #buffer = new Uint8Array(1024);
  #length = 0;

  get length(): number {
    return this.#length;
  }

  #need(extra: number): void {
    if (this.#length + extra <= this.#buffer.length) return;
    let size = this.#buffer.length * 2;
    while (size < this.#length + extra) size *= 2;
    const next = new Uint8Array(size);
    next.set(this.#buffer.subarray(0, this.#length));
    this.#buffer = next;
  }

  u8(value: number): void {
    this.#need(1);
    this.#buffer[this.#length++] = value & 0xff;
  }

  u16(value: number): void {
    this.u8(value);
    this.u8(value >>> 8);
  }

  u32(value: number): void {
    this.u16(value);
    this.u16(Math.floor(value / 0x10000));
  }

  u64(value: number): void {
    this.u32(value % 0x100000000);
    this.u32(Math.floor(value / 0x100000000));
  }

  bytes(value: Uint8Array): void {
    this.#need(value.length);
    this.#buffer.set(value, this.#length);
    this.#length += value.length;
  }

  finish(): Uint8Array {
    return this.#buffer.slice(0, this.#length);
  }
}

const encoder = new TextEncoder();

interface Laid {
  readonly fixture: FixtureEntry;
  readonly nameBytes: Uint8Array;
  readonly flags: number;
  readonly method: number;
  readonly payload: Uint8Array;
  readonly crc: number;
  readonly uncompressedSize: number;
  localHeaderOffset: number;
  dataStart: number;
  dataEnd: number;
}

function needsUtf8(bytes: Uint8Array): boolean {
  return bytes.some((b) => b >= 0x80);
}

export function buildZip(
  entries: readonly FixtureEntry[],
  options: BuildZipOptions = {},
): Uint8Array {
  const out = new ByteWriter();

  const laid: Laid[] = entries.map((fixture) => {
    const data = fixture.data ?? new Uint8Array(0);
    const deflate = fixture.deflate ?? data.length > 0;
    const payload = deflate ? deflateSync(data, { level: fixture.deflateLevel ?? 9 }) : data;
    const nameBytes = encoder.encode(fixture.name);
    return {
      fixture,
      nameBytes,
      flags: (fixture.flags ?? 0) | (needsUtf8(nameBytes) ? 0x0800 : 0),
      method: fixture.method ?? (deflate ? 8 : 0),
      payload,
      crc: crc32(data),
      uncompressedSize: data.length,
      localHeaderOffset: 0,
      dataStart: 0,
      dataEnd: 0,
    };
  });

  // Local file headers and data, in order.
  for (const entry of laid) {
    entry.localHeaderOffset = out.length;
    const localNameBytes = entry.fixture.localName
      ? encoder.encode(entry.fixture.localName)
      : entry.nameBytes;
    const usesDescriptor = (entry.flags & 0x0008) !== 0;

    out.u32(SIG_LOCAL);
    out.u16(20);
    out.u16(entry.flags);
    out.u16(entry.fixture.localMethod ?? entry.method);
    out.u16(0);
    out.u16(0);
    out.u32(usesDescriptor ? 0 : entry.crc);
    out.u32(usesDescriptor ? 0 : entry.payload.length);
    out.u32(usesDescriptor ? 0 : entry.uncompressedSize);
    out.u16(localNameBytes.length);
    out.u16(0);
    out.bytes(localNameBytes);
    entry.dataStart = out.length;
    out.bytes(entry.payload);
    entry.dataEnd = out.length;

    if (usesDescriptor) {
      out.u32(0x08074b50);
      out.u32(entry.crc);
      out.u32(entry.payload.length);
      out.u32(entry.uncompressedSize);
      entry.dataEnd = out.length;
    }
  }

  // Central directory.
  const centralOffset = out.length;
  for (let i = 0; i < laid.length; i++) {
    const entry = laid[i]!;
    const fixture = entry.fixture;

    let compressedSize = fixture.declaredCompressedSize ?? entry.payload.length;
    if (fixture.overlapNext) {
      const next = laid[i + 1];
      if (next) compressedSize = next.dataEnd - entry.dataStart;
    }
    const uncompressedSize = fixture.declaredUncompressedSize ?? entry.uncompressedSize;
    const zip64 = fixture.zip64 === true;

    out.u32(SIG_CENTRAL);
    out.u16(20);
    out.u16(zip64 ? 45 : 20);
    out.u16(entry.flags);
    out.u16(entry.method);
    out.u16(0);
    out.u16(0);
    out.u32(fixture.declaredCrc32 ?? entry.crc);
    out.u32(zip64 ? U32_MAX : compressedSize);
    out.u32(zip64 ? U32_MAX : uncompressedSize);
    out.u16(entry.nameBytes.length);
    out.u16(zip64 ? 20 : 0);
    out.u16(0);
    out.u16(0);
    out.u16(0);
    out.u32(0);
    out.u32(entry.localHeaderOffset);
    out.bytes(entry.nameBytes);
    if (zip64) {
      out.u16(0x0001);
      out.u16(16);
      out.u64(uncompressedSize);
      out.u64(compressedSize);
    }
  }
  const centralSize = out.length - centralOffset;

  if (options.zip64Eocd) {
    const zip64Start = out.length;
    out.u32(SIG_ZIP64_EOCD);
    out.u64(44);
    out.u16(45);
    out.u16(45);
    out.u32(0);
    out.u32(0);
    out.u64(laid.length);
    out.u64(laid.length);
    out.u64(centralSize);
    out.u64(centralOffset);

    out.u32(SIG_ZIP64_LOCATOR);
    out.u32(0);
    out.u64(zip64Start);
    out.u32(1);
  }

  const comment = options.comment ?? new Uint8Array(0);
  const count = options.declaredEntryCount ?? laid.length;
  out.u32(SIG_EOCD);
  out.u16(0);
  out.u16(0);
  out.u16(Math.min(count, 0xffff));
  out.u16(Math.min(count, 0xffff));
  out.u32(centralSize);
  out.u32(centralOffset);
  out.u16(comment.length);
  out.bytes(comment);

  return out.finish();
}

/** `n` bytes of a byte pattern that DEFLATE cannot compress. */
export function incompressible(n: number): Uint8Array {
  const out = new Uint8Array(n);
  let state = 0x9e3779b9;
  for (let i = 0; i < n; i++) {
    state = (Math.imul(state, 1664525) + 1013904223) >>> 0;
    out[i] = (state >>> 24) & 0xff;
  }
  return out;
}
