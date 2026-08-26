import { inflateSync } from 'fflate';
import { crc32 } from './crc32.js';
import { guard, OpcError } from './errors.js';
import { DEFAULT_ZIP_LIMITS, InflationBudget, type ZipLimits } from './limits.js';

/**
 * A ZIP reader for untrusted archives.
 *
 * This is deliberately not `fflate.unzipSync`, and the reasons are the design:
 *
 * 1. **`unzipSync` decompresses every entry eagerly.** The whole architecture
 *    rests on parsing a part only when something reads it, so eager inflation
 *    is both the wrong shape and the over-allocation vector we are defending
 *    against.
 * 2. **It returns a plain object keyed by name.** JavaScript object key order
 *    puts integer-like keys first, so an archive containing an entry named
 *    `1` silently reorders. Entry order is what the writer replays to keep an
 *    export byte-identical, so it is stored in an array here.
 * 3. **It discards everything except the bytes** - general-purpose flags,
 *    CRC-32, local header offsets, external attributes. Every one of those is
 *    needed either for safety or for round-tripping.
 * 4. **It verifies nothing.** No CRC, no cross-check of the local header
 *    against the central directory, no bound on total inflation.
 *
 * What we do keep from `fflate` is `inflateSync`, and one empirically verified
 * property of it: when handed a pre-allocated `out` buffer it never grows it.
 * Ten megabytes of zeros compressed to ten kilobytes, inflated into a
 * 1024-byte `out`, yields exactly 1024 bytes and allocates nothing further.
 * That makes the declared uncompressed size a hard allocation ceiling. The
 * surplus is dropped silently rather than reported, which is why every read
 * ends in a CRC-32 check - see `crc32.ts`.
 */

// --- ZIP signatures -------------------------------------------------------

const SIG_LOCAL = 0x04034b50;
const SIG_CENTRAL = 0x02014b50;
const SIG_EOCD = 0x06054b50;
const SIG_ZIP64_EOCD = 0x06064b50;
const SIG_ZIP64_LOCATOR = 0x07064b50;

/** OLE2 / Compound File Binary signature. */
const CFB_MAGIC = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1] as const;

const EOCD_SIZE = 22;
const ZIP64_LOCATOR_SIZE = 20;
const CENTRAL_HEADER_SIZE = 46;
const LOCAL_HEADER_SIZE = 30;
const MAX_COMMENT = 0xffff;

/** Sentinels that mean "the real value lives in the ZIP64 extra field". */
const U32_MAX = 0xffffffff;

export const COMPRESSION_STORE = 0;
export const COMPRESSION_DEFLATE = 8;

/** General-purpose bit flags this reader acts on. */
const FLAG_ENCRYPTED = 0x0001;
const FLAG_DATA_DESCRIPTOR = 0x0008;
const FLAG_STRONG_ENCRYPTION = 0x0040;
const FLAG_UTF8 = 0x0800;

/**
 * CP437 code points 0x80-0xFF, for entry names written without the UTF-8 flag.
 *
 * Such a name will not survive the OPC part-name grammar anyway, so this exists
 * so the *error* names the entry correctly rather than as mojibake.
 */
const CP437_HIGH =
  'ÇüéâäàåçêëèïîìÄÅ' +
  'ÉæÆôöòûùÿÖÜ¢£¥₧ƒ' +
  'áíóúñÑªº¿⌐¬½¼¡«»' +
  '░▒▓│┤╡╢╖╕╣║╗╝╜╛┐' +
  '└┴┬├─┼╞╟╚╔╩╦╠═╬╧' +
  '╨╤╥╙╘╒╓╫╪┘┌█▄▌▐▀' +
  'αßΓπΣσµτΦΘΩδ∞φε∩' +
  '≡±≥≤⌠⌡÷≈°∙·√ⁿ²■ ';

let utf8Decoder: TextDecoder | undefined;

function decodeEntryName(bytes: Uint8Array, flags: number): string {
  if ((flags & FLAG_UTF8) !== 0) {
    utf8Decoder ??= new TextDecoder('utf-8', { fatal: true });
    const decoder = utf8Decoder;
    return guard(
      'ERR_INVALID_PART_NAME',
      'entry name is flagged UTF-8 but is not valid UTF-8',
      {},
      () => decoder.decode(bytes),
    );
  }
  let out = '';
  for (let i = 0; i < bytes.length; i++) {
    const b = bytes[i]!;
    out += b < 0x80 ? String.fromCharCode(b) : CP437_HIGH[b - 0x80]!;
  }
  return out;
}

// --- Public shapes --------------------------------------------------------

/** One central-directory entry, resolved against its local file header. */
export interface ZipEntry {
  /** The raw ZIP entry name, e.g. `ppt/slides/slide1.xml`. No leading slash. */
  readonly name: string;
  /**
   * Position in the central directory.
   *
   * This is the export order. The writer replays it so an untouched package
   * comes back out with the same entry sequence, which is what makes a
   * round-trip diff meaningful.
   */
  readonly index: number;
  readonly method: number;
  readonly flags: number;
  readonly compressedSize: number;
  readonly uncompressedSize: number;
  /** The stored CRC-32, unsigned. Taken from the central directory, which is authoritative. */
  readonly crc32: number;
  /** Offset of the local file header. */
  readonly localHeaderOffset: number;
  /** Offset of the entry's compressed bytes, after the local header and its name and extra fields. */
  readonly dataOffset: number;
  readonly dosTime: number;
  readonly dosDate: number;
  readonly externalAttributes: number;
  /** A name ending in `/`. OPC packages should have none; real archives sometimes do. */
  readonly isDirectory: boolean;
  /** Bit 3 was set: the local header carried zeroed sizes and a trailing data descriptor. */
  readonly hasDataDescriptor: boolean;
}

export interface ReadZipOptions {
  /** Override any subset of the default budgets. */
  readonly limits?: Partial<ZipLimits>;
}

// --- Reading --------------------------------------------------------------

function readU64(view: DataView, offset: number): number {
  const value = view.getBigUint64(offset, true);
  if (value > BigInt(Number.MAX_SAFE_INTEGER)) {
    throw new OpcError(
      'ERR_ZIP64_UNSUPPORTED',
      'a ZIP64 field holds ' +
        value.toString() +
        ', which is past the exactly-representable ' +
        'integer range. Nothing this reader accepts is that large.',
      { offset },
    );
  }
  return Number(value);
}

function looksEncrypted(bytes: Uint8Array): boolean {
  if (bytes.length < CFB_MAGIC.length) return false;
  return CFB_MAGIC.every((b, i) => bytes[i] === b);
}

/**
 * Locate the end-of-central-directory record.
 *
 * The comment-length cross-check matters: the four signature bytes can occur
 * inside compressed data, and a scan that stops at the first match can be
 * steered onto an attacker-chosen central directory. Requiring the record to
 * end exactly at the end of the buffer also rejects archives with data appended
 * after the comment, which is the shape of a ZIP polyglot.
 */
function findEndOfCentralDirectory(view: DataView, length: number): number {
  const lowest = Math.max(0, length - EOCD_SIZE - MAX_COMMENT);
  for (let i = length - EOCD_SIZE; i >= lowest; i--) {
    if (view.getUint32(i, true) !== SIG_EOCD) continue;
    if (i + EOCD_SIZE + view.getUint16(i + 20, true) === length) return i;
  }
  throw new OpcError(
    'ERR_NOT_A_ZIP',
    'no end-of-central-directory record. The file is not a ZIP archive, or it has been ' +
      'truncated, or something has been appended after the archive comment.',
    { actual: length },
  );
}

interface Zip64Extra {
  readonly uncompressedSize?: number;
  readonly compressedSize?: number;
  readonly localHeaderOffset?: number;
}

/**
 * Parse the ZIP64 extended information extra field (header id 0x0001).
 *
 * Its fields are in a fixed order but present only when the corresponding
 * 32-bit field held the 0xFFFFFFFF sentinel - so the layout cannot be read
 * without knowing which sentinels fired, which is why those flags are
 * arguments.
 */
function parseZip64Extra(
  view: DataView,
  start: number,
  length: number,
  need: { size: boolean; compressed: boolean; offset: boolean },
): Zip64Extra {
  let cursor = start;
  const end = start + length;
  while (cursor + 4 <= end) {
    const id = view.getUint16(cursor, true);
    const size = view.getUint16(cursor + 2, true);
    const body = cursor + 4;
    if (body + size > end) break;
    if (id === 0x0001) {
      const out: {
        uncompressedSize?: number;
        compressedSize?: number;
        localHeaderOffset?: number;
      } = {};
      let field = body;
      if (need.size && field + 8 <= body + size) {
        out.uncompressedSize = readU64(view, field);
        field += 8;
      }
      if (need.compressed && field + 8 <= body + size) {
        out.compressedSize = readU64(view, field);
        field += 8;
      }
      if (need.offset && field + 8 <= body + size) {
        out.localHeaderOffset = readU64(view, field);
      }
      return out;
    }
    cursor = body + size;
  }
  return {};
}

/**
 * A parsed ZIP archive. Entries are enumerated eagerly; bytes are inflated only
 * when asked for.
 */
export class ZipArchive {
  readonly entries: readonly ZipEntry[];
  readonly limits: ZipLimits;
  readonly budget: InflationBudget;

  readonly #bytes: Uint8Array;
  readonly #byName: ReadonlyMap<string, ZipEntry>;
  readonly #charged = new Set<number>();

  constructor(bytes: Uint8Array, entries: readonly ZipEntry[], limits: ZipLimits) {
    this.#bytes = bytes;
    this.entries = entries;
    this.limits = limits;
    this.budget = new InflationBudget(limits.maxTotalInflatedBytes);
    const byName = new Map<string, ZipEntry>();
    for (const entry of entries) byName.set(entry.name, entry);
    this.#byName = byName;
  }

  get(name: string): ZipEntry | undefined {
    return this.#byName.get(name);
  }

  has(name: string): boolean {
    return this.#byName.has(name);
  }

  /**
   * The entry's bytes exactly as stored - still compressed if it was
   * compressed. Charges no budget and inflates nothing.
   *
   * This is what the writer streams through for parts we never edited, so that
   * preservation costs neither a decompress nor a recompress.
   */
  raw(entry: ZipEntry): Uint8Array {
    return this.#bytes.subarray(entry.dataOffset, entry.dataOffset + entry.compressedSize);
  }

  /**
   * Inflate one entry, verify its CRC-32, and charge the running budget.
   *
   * The budget is charged once per entry rather than once per call: re-reading
   * a part must not consume the archive's allowance a second time, or opening
   * the same deck twice in one session would fail the second time for no
   * reason a user could understand.
   */
  read(entry: ZipEntry): Uint8Array {
    if (entry.isDirectory) return new Uint8Array(0);

    if ((entry.flags & (FLAG_ENCRYPTED | FLAG_STRONG_ENCRYPTION)) !== 0) {
      throw new OpcError(
        'ERR_ENTRY_ENCRYPTED',
        entry.name +
          ' is encrypted. Password-protected Office files are normally a whole encrypted ' +
          'container rather than an encrypted ZIP entry, so this archive was produced by ' +
          'something other than Office.',
        { entry: entry.name },
      );
    }

    if (entry.method !== COMPRESSION_STORE && entry.method !== COMPRESSION_DEFLATE) {
      throw new OpcError(
        'ERR_UNSUPPORTED_COMPRESSION',
        entry.name +
          ' uses compression method ' +
          String(entry.method) +
          '. OPC packages use stored (0) or deflate (8); PowerPoint will not open anything else ' +
          'either.',
        { entry: entry.name, actual: entry.method },
      );
    }

    if (entry.uncompressedSize > this.limits.maxEntryInflatedBytes) {
      throw new OpcError(
        'ERR_ENTRY_TOO_LARGE',
        entry.name +
          ' declares ' +
          String(entry.uncompressedSize) +
          ' inflated bytes, past the ' +
          String(this.limits.maxEntryInflatedBytes) +
          '-byte ceiling for a single part.',
        {
          entry: entry.name,
          limit: this.limits.maxEntryInflatedBytes,
          actual: entry.uncompressedSize,
        },
      );
    }

    if (
      entry.uncompressedSize > this.limits.ratioFloorBytes &&
      entry.compressedSize > 0 &&
      entry.uncompressedSize / entry.compressedSize > this.limits.maxCompressionRatio
    ) {
      throw new OpcError(
        'ERR_RATIO_EXCEEDED',
        entry.name +
          ' expands ' +
          (entry.uncompressedSize / entry.compressedSize).toFixed(0) +
          ':1, past the ' +
          String(this.limits.maxCompressionRatio) +
          ':1 ceiling. Presentation XML compresses roughly 10:1; this does not.',
        {
          entry: entry.name,
          limit: this.limits.maxCompressionRatio,
          actual: Math.round(entry.uncompressedSize / entry.compressedSize),
        },
      );
    }

    if (!this.#charged.has(entry.index)) {
      this.budget.charge(entry.uncompressedSize, entry.name);
      this.#charged.add(entry.index);
    }

    const compressed = this.raw(entry);
    let out: Uint8Array;

    if (entry.method === COMPRESSION_STORE) {
      if (entry.compressedSize !== entry.uncompressedSize) {
        throw new OpcError(
          'ERR_HEADER_MISMATCH',
          entry.name +
            ' is stored uncompressed but declares different compressed and uncompressed sizes.',
          { entry: entry.name },
        );
      }
      out = compressed.slice();
    } else {
      out = guard(
        'ERR_INFLATE_FAILED',
        entry.name + ' is not a valid DEFLATE stream',
        { entry: entry.name },
        () => inflateSync(compressed, { out: new Uint8Array(entry.uncompressedSize) }),
      );
    }

    if (out.length !== entry.uncompressedSize) {
      throw new OpcError(
        'ERR_TRUNCATED',
        entry.name +
          ' inflated to ' +
          String(out.length) +
          ' bytes but declares ' +
          String(entry.uncompressedSize) +
          '.',
        { entry: entry.name, actual: out.length, limit: entry.uncompressedSize },
      );
    }

    const actual = crc32(out);
    if (actual !== entry.crc32) {
      throw new OpcError(
        'ERR_CRC_MISMATCH',
        entry.name +
          ' failed its CRC-32 check (stored 0x' +
          entry.crc32.toString(16).padStart(8, '0') +
          ', computed 0x' +
          actual.toString(16).padStart(8, '0') +
          '). Either the archive is damaged, or it declares a size smaller than the stream ' +
          'actually produces - which is how a decompression bomb hides behind an honest-looking ' +
          'header.',
        { entry: entry.name },
      );
    }

    return out;
  }

  /** `read`, addressed by ZIP entry name. */
  readByName(name: string): Uint8Array {
    const entry = this.#byName.get(name);
    if (!entry) {
      throw new OpcError('ERR_NOT_A_ZIP', 'no entry named ' + JSON.stringify(name), {
        entry: name,
      });
    }
    return this.read(entry);
  }
}

/**
 * Parse the central directory of `bytes` and validate the archive's structure.
 *
 * Nothing is inflated here. Every check below is either a bound (so the reader
 * cannot be steered out of the buffer) or a consistency check between the two
 * places ZIP stores the same fact.
 */
export function readZip(bytes: Uint8Array, options: ReadZipOptions = {}): ZipArchive {
  const limits: ZipLimits = { ...DEFAULT_ZIP_LIMITS, ...options.limits };
  const length = bytes.byteLength;

  if (looksEncrypted(bytes)) {
    throw new OpcError(
      'ERR_ENCRYPTED_PACKAGE',
      'this file is an OLE2 compound document, not a ZIP package. A password-protected ' +
        '.pptx is stored that way - the presentation lives in an EncryptedPackage stream ' +
        'inside it. A pre-2007 binary .ppt looks identical from here.',
    );
  }

  if (length > limits.maxArchiveBytes) {
    throw new OpcError(
      'ERR_ARCHIVE_TOO_LARGE',
      'archive is ' +
        String(length) +
        ' bytes, past the ' +
        String(limits.maxArchiveBytes) +
        '-byte ceiling.',
      { limit: limits.maxArchiveBytes, actual: length },
    );
  }

  if (length < EOCD_SIZE) {
    throw new OpcError(
      'ERR_TRUNCATED',
      'archive is ' +
        String(length) +
        ' bytes; the smallest possible ZIP is ' +
        String(EOCD_SIZE) +
        '.',
      { actual: length },
    );
  }

  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const eocd = findEndOfCentralDirectory(view, length);

  let entryCount = view.getUint16(eocd + 10, true);
  let centralSize = view.getUint32(eocd + 12, true);
  let centralOffset = view.getUint32(eocd + 16, true);

  // ZIP64. PowerPoint never writes it and our writer never will, but other
  // producers emit a ZIP64 record even for small archives, and refusing to
  // read one would fail files that are otherwise perfectly ordinary.
  if (
    eocd >= ZIP64_LOCATOR_SIZE &&
    view.getUint32(eocd - ZIP64_LOCATOR_SIZE, true) === SIG_ZIP64_LOCATOR
  ) {
    const zip64Offset = readU64(view, eocd - ZIP64_LOCATOR_SIZE + 8);
    if (zip64Offset + 56 <= length && view.getUint32(zip64Offset, true) === SIG_ZIP64_EOCD) {
      entryCount = readU64(view, zip64Offset + 32);
      centralSize = readU64(view, zip64Offset + 40);
      centralOffset = readU64(view, zip64Offset + 48);
    }
  }

  if (entryCount > limits.maxEntries) {
    throw new OpcError(
      'ERR_TOO_MANY_ENTRIES',
      'archive declares ' +
        String(entryCount) +
        ' entries, past the ' +
        String(limits.maxEntries) +
        '-entry ceiling. A 300-slide deck has a few thousand.',
      { limit: limits.maxEntries, actual: entryCount },
    );
  }

  if (centralOffset + centralSize > length) {
    throw new OpcError(
      'ERR_TRUNCATED',
      'the central directory is declared to run past the end of the archive.',
      { offset: centralOffset, actual: length },
    );
  }

  const entries: ZipEntry[] = [];
  const seen = new Set<string>();
  let cursor = centralOffset;

  for (let index = 0; index < entryCount; index++) {
    if (cursor + CENTRAL_HEADER_SIZE > length) {
      throw new OpcError('ERR_TRUNCATED', 'central directory ends mid-header.', { offset: cursor });
    }
    if (view.getUint32(cursor, true) !== SIG_CENTRAL) {
      throw new OpcError(
        'ERR_NOT_A_ZIP',
        'expected a central-directory header at offset ' + String(cursor) + '.',
        { offset: cursor },
      );
    }

    const flags = view.getUint16(cursor + 8, true);
    const method = view.getUint16(cursor + 10, true);
    const dosTime = view.getUint16(cursor + 12, true);
    const dosDate = view.getUint16(cursor + 14, true);
    const storedCrc = view.getUint32(cursor + 16, true) >>> 0;
    const rawCompressed = view.getUint32(cursor + 20, true);
    const rawUncompressed = view.getUint32(cursor + 24, true);
    const nameLength = view.getUint16(cursor + 28, true);
    const extraLength = view.getUint16(cursor + 30, true);
    const commentLength = view.getUint16(cursor + 32, true);
    const externalAttributes = view.getUint32(cursor + 38, true);
    const rawLocalOffset = view.getUint32(cursor + 42, true);

    const nameStart = cursor + CENTRAL_HEADER_SIZE;
    const extraStart = nameStart + nameLength;
    const next = extraStart + extraLength + commentLength;
    if (next > length) {
      throw new OpcError(
        'ERR_TRUNCATED',
        'central-directory entry runs past the end of the archive.',
        {
          offset: cursor,
        },
      );
    }

    if (nameLength > limits.maxNameLength) {
      throw new OpcError(
        'ERR_INVALID_PART_NAME',
        'entry name is ' +
          String(nameLength) +
          ' bytes, past the ' +
          String(limits.maxNameLength) +
          '-byte ceiling.',
        { offset: cursor, limit: limits.maxNameLength, actual: nameLength },
      );
    }

    const name = decodeEntryName(bytes.subarray(nameStart, extraStart), flags);

    const zip64 = parseZip64Extra(view, extraStart, extraLength, {
      size: rawUncompressed === U32_MAX,
      compressed: rawCompressed === U32_MAX,
      offset: rawLocalOffset === U32_MAX,
    });
    const uncompressedSize = zip64.uncompressedSize ?? rawUncompressed;
    const compressedSize = zip64.compressedSize ?? rawCompressed;
    const localHeaderOffset = zip64.localHeaderOffset ?? rawLocalOffset;

    if (
      uncompressedSize === U32_MAX ||
      compressedSize === U32_MAX ||
      localHeaderOffset === U32_MAX
    ) {
      throw new OpcError(
        'ERR_ZIP64_UNSUPPORTED',
        name + ' uses a ZIP64 sentinel with no matching ZIP64 extra field.',
        { entry: name },
      );
    }

    if (seen.has(name)) {
      throw new OpcError(
        'ERR_DUPLICATE_ENTRY',
        'two entries are both named ' +
          JSON.stringify(name) +
          '. Which one a consumer sees then depends on whether it reads the central directory ' +
          'forwards or backwards, and a file that means different things to different readers ' +
          'is not one we open.',
        { entry: name },
      );
    }
    seen.add(name);

    // Resolve the local header. Its name and extra-field lengths differ from
    // the central directory's - the extra fields legitimately do, the name
    // never does.
    if (localHeaderOffset + LOCAL_HEADER_SIZE > length) {
      throw new OpcError(
        'ERR_TRUNCATED',
        name + ': local header runs past the end of the archive.',
        {
          entry: name,
          offset: localHeaderOffset,
        },
      );
    }
    if (view.getUint32(localHeaderOffset, true) !== SIG_LOCAL) {
      throw new OpcError(
        'ERR_HEADER_MISMATCH',
        name +
          ': the central directory points at offset ' +
          String(localHeaderOffset) +
          ', which is not a local file header.',
        { entry: name, offset: localHeaderOffset },
      );
    }

    const localFlags = view.getUint16(localHeaderOffset + 6, true);
    const localMethod = view.getUint16(localHeaderOffset + 8, true);
    const localNameLength = view.getUint16(localHeaderOffset + 26, true);
    const localExtraLength = view.getUint16(localHeaderOffset + 28, true);
    const localNameStart = localHeaderOffset + LOCAL_HEADER_SIZE;
    const dataOffset = localNameStart + localNameLength + localExtraLength;

    if (dataOffset + compressedSize > length) {
      throw new OpcError(
        'ERR_TRUNCATED',
        name + ': its data runs past the end of the archive. The file is incomplete.',
        { entry: name, offset: dataOffset },
      );
    }

    const localName = decodeEntryName(
      bytes.subarray(localNameStart, localNameStart + localNameLength),
      localFlags,
    );
    if (localName !== name) {
      throw new OpcError(
        'ERR_HEADER_MISMATCH',
        'the central directory calls this entry ' +
          JSON.stringify(name) +
          ' and its local header calls it ' +
          JSON.stringify(localName) +
          '. Readers disagree about which one wins, so an archive that disagrees with itself ' +
          'is one that means different things to different programs.',
        { entry: name, offset: localHeaderOffset },
      );
    }
    if (localMethod !== method) {
      throw new OpcError(
        'ERR_HEADER_MISMATCH',
        name +
          ': central directory says compression method ' +
          String(method) +
          ', local header says ' +
          String(localMethod) +
          '.',
        { entry: name, offset: localHeaderOffset },
      );
    }

    entries.push({
      name,
      index,
      method,
      flags,
      compressedSize,
      uncompressedSize,
      crc32: storedCrc,
      localHeaderOffset,
      dataOffset,
      dosTime,
      dosDate,
      externalAttributes,
      isDirectory: name.endsWith('/'),
      hasDataDescriptor: (flags & FLAG_DATA_DESCRIPTOR) !== 0,
    });

    cursor = next;
  }

  assertNoOverlap(entries);

  return new ZipArchive(bytes, entries, limits);
}

/**
 * Reject archives whose entries claim overlapping byte ranges.
 *
 * This is the structural answer to Fifield's quoted-overlap bomb, where each
 * entry's compressed data contains the local file headers of the entries that
 * follow it, so a small archive declares an enormous total. The budget alone
 * survives that construction; this refuses it, which is a better answer,
 * because a legitimate archive never overlaps and so nothing is lost by saying
 * so.
 */
function assertNoOverlap(entries: readonly ZipEntry[]): void {
  const ranges = entries
    .map((entry) => ({
      entry,
      start: entry.localHeaderOffset,
      end: entry.dataOffset + entry.compressedSize,
    }))
    .sort((a, b) => a.start - b.start);

  for (let i = 1; i < ranges.length; i++) {
    const previous = ranges[i - 1]!;
    const current = ranges[i]!;
    if (current.start < previous.end) {
      throw new OpcError(
        'ERR_OVERLAPPING_ENTRY',
        'entries ' +
          JSON.stringify(previous.entry.name) +
          ' and ' +
          JSON.stringify(current.entry.name) +
          ' claim overlapping byte ranges. Legitimate archives never share bytes between ' +
          'entries; a bomb that expands quadratically does exactly that.',
        { entry: current.entry.name, offset: current.start },
      );
    }
  }
}
