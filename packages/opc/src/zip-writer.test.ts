import { describe, expect, it } from 'vitest';
import { crc32 } from './crc32.js';
import { isOpcError } from './errors.js';
import { readZip } from './zip-reader.js';
import { deflatedEntry, passthroughEntry, storedEntry, writeZip } from './zip-writer.js';

const encoder = new TextEncoder();
const decoder = new TextDecoder();

function codeOf(fn: () => unknown): string {
  try {
    fn();
  } catch (error) {
    if (isOpcError(error)) return error.code;
    return 'NOT_AN_OPC_ERROR: ' + (error instanceof Error ? error.constructor.name : typeof error);
  }
  return 'DID_NOT_THROW';
}

/** Read the raw central-directory fields the public `ZipEntry` does not expose. */
function centralDirectory(bytes: Uint8Array) {
  const v = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  let eocd = -1;
  for (let i = bytes.length - 22; i >= 0; i--) {
    if (v.getUint32(i, true) === 0x06054b50) {
      eocd = i;
      break;
    }
  }
  const offset = v.getUint32(eocd + 16, true);
  const count = v.getUint16(eocd + 10, true);
  const out = [];
  let p = offset;
  for (let i = 0; i < count; i++) {
    const nameLength = v.getUint16(p + 28, true);
    out.push({
      versionMadeBy: v.getUint16(p + 4, true),
      versionNeeded: v.getUint16(p + 6, true),
      flags: v.getUint16(p + 8, true),
      method: v.getUint16(p + 10, true),
      time: v.getUint16(p + 12, true),
      date: v.getUint16(p + 14, true),
      extraLength: v.getUint16(p + 30, true),
      commentLength: v.getUint16(p + 32, true),
      internalAttributes: v.getUint16(p + 36, true),
      externalAttributes: v.getUint32(p + 38, true),
      localOffset: v.getUint32(p + 42, true),
      name: decoder.decode(bytes.subarray(p + 46, p + 46 + nameLength)),
    });
    p += 46 + nameLength + out[i]!.extraLength + out[i]!.commentLength;
  }
  return { entries: out, commentLength: v.getUint16(eocd + 20, true) };
}

const HELLO = encoder.encode('hello world '.repeat(50));

describe('what we emit', () => {
  it('round-trips through our own reader', () => {
    const bytes = writeZip([
      deflatedEntry('[Content_Types].xml', encoder.encode('<Types/>')),
      deflatedEntry('ppt/presentation.xml', HELLO),
      storedEntry('ppt/media/image1.png', new Uint8Array([1, 2, 3, 4])),
    ]);
    const archive = readZip(bytes);
    expect(archive.entries.map((e) => e.name)).toEqual([
      '[Content_Types].xml',
      'ppt/presentation.xml',
      'ppt/media/image1.png',
    ]);
    expect(decoder.decode(archive.readByName('ppt/presentation.xml'))).toBe(decoder.decode(HELLO));
    expect([...archive.readByName('ppt/media/image1.png')]).toEqual([1, 2, 3, 4]);
  });

  it('writes the header fields PowerPoint writes', () => {
    const { entries, commentLength } = centralDirectory(
      writeZip([deflatedEntry('a.xml', HELLO), storedEntry('b.png', new Uint8Array([9]))]),
    );
    expect(commentLength).toBe(0);
    for (const e of entries) {
      expect(e.versionMadeBy).toBe(20); // ZIP 2.0, host 0 = MS-DOS/FAT
      expect(e.flags).toBe(0); // no UTF-8 bit, no data descriptor, no level hint
      expect(e.time).toBe(0);
      expect(e.date).toBe(33); // 1980-01-01, so output is reproducible
      expect(e.extraLength).toBe(0);
      expect(e.commentLength).toBe(0);
      expect(e.internalAttributes).toBe(0);
      expect(e.externalAttributes).toBe(0);
    }
    expect(entries[0]!.versionNeeded).toBe(20); // deflated
    expect(entries[1]!.versionNeeded).toBe(10); // stored
  });

  it('writes the central directory in local-header order', () => {
    const { entries } = centralDirectory(
      writeZip(['c', 'a', 'b', '1'].map((n) => deflatedEntry(n + '.xml', HELLO))),
    );
    expect(entries.map((e) => e.name)).toEqual(['c.xml', 'a.xml', 'b.xml', '1.xml']);
    const offsets = entries.map((e) => e.localOffset);
    expect([...offsets].sort((x, y) => x - y)).toEqual(offsets);
  });

  it('keeps an entry named like an integer in place', () => {
    // The reason entries are an array. A plain object would put `1` first.
    const archive = readZip(writeZip([deflatedEntry('z.xml', HELLO), deflatedEntry('1', HELLO)]));
    expect(archive.entries.map((e) => e.name)).toEqual(['z.xml', '1']);
  });

  it('writes an empty archive that reads back as empty', () => {
    const archive = readZip(writeZip([]));
    expect(archive.entries).toHaveLength(0);
  });

  it('is deterministic: the same input twice gives the same bytes', () => {
    const build = () => writeZip([deflatedEntry('a.xml', HELLO), storedEntry('b', HELLO)]);
    expect([...build()]).toEqual([...build()]);
  });
});

describe('choosing a compression method', () => {
  it('deflates text and records the right CRC and sizes', () => {
    const entry = deflatedEntry('a.xml', HELLO);
    expect(entry.method).toBe(8);
    expect(entry.data.length).toBeLessThan(HELLO.length);
    expect(entry.crc32).toBe(crc32(HELLO));
    expect(entry.uncompressedSize).toBe(HELLO.length);
  });

  it('stores rather than growing something that will not compress', () => {
    // What Office does with PNG and JPEG: a real deck we measured stores every
    // one of them at method 0.
    const random = new Uint8Array(4096);
    let seed = 12345;
    for (let i = 0; i < random.length; i++) {
      seed = (seed * 1103515245 + 12345) & 0x7fffffff;
      random[i] = (seed >> 16) & 0xff;
    }
    const entry = deflatedEntry('a.bin', random);
    expect(entry.method).toBe(0);
    expect(entry.data).toBe(random);
    expect(entry.uncompressedSize).toBe(random.length);
  });

  it('emits a raw DEFLATE stream, not a zlib-wrapped one', () => {
    // A zlib header inside a ZIP entry is a corruption no reader reports, so
    // this is checked rather than trusted. zlib would start 0x78.
    const entry = deflatedEntry('a.xml', HELLO);
    expect(entry.data[0]).not.toBe(0x78);
    expect(decoder.decode(readZip(writeZip([entry])).readByName('a.xml'))).toBe(
      decoder.decode(HELLO),
    );
  });

  it('handles an empty part', () => {
    const archive = readZip(writeZip([deflatedEntry('empty.xml', new Uint8Array(0))]));
    expect(archive.readByName('empty.xml')).toHaveLength(0);
  });
});

describe('copying an entry across without decompressing it', () => {
  it('moves the stored stream, its CRC and both sizes together', () => {
    const source = readZip(
      writeZip([deflatedEntry('a.xml', HELLO), storedEntry('b.png', new Uint8Array([7, 7, 7]))]),
    );
    const rebuilt = writeZip(source.entries.map((e) => passthroughEntry(source, e)).reverse());
    const back = readZip(rebuilt);
    expect(back.entries.map((e) => e.name)).toEqual(['b.png', 'a.xml']);
    for (const entry of back.entries) {
      const original = source.get(entry.name)!;
      expect(entry.method).toBe(original.method);
      expect(entry.crc32).toBe(original.crc32);
      expect(entry.uncompressedSize).toBe(original.uncompressedSize);
      expect(entry.compressedSize).toBe(original.compressedSize);
      expect([...back.raw(entry)]).toEqual([...source.raw(original)]);
    }
  });

  it('recomputes the local header offset rather than copying it', () => {
    const source = readZip(
      writeZip([deflatedEntry('a.xml', HELLO), deflatedEntry('b.xml', HELLO)]),
    );
    const back = readZip(
      writeZip([
        passthroughEntry(source, source.entries[1]!),
        passthroughEntry(source, source.entries[0]!),
      ]),
    );
    expect(back.entries[0]!.localHeaderOffset).toBe(0);
    expect(decoder.decode(back.readByName('a.xml'))).toBe(decoder.decode(HELLO));
  });
});

describe('what we refuse to emit', () => {
  it('refuses two entries with the same name', () => {
    expect(
      codeOf(() => writeZip([deflatedEntry('a.xml', HELLO), deflatedEntry('a.xml', HELLO)])),
    ).toBe('ERR_DUPLICATE_ENTRY');
  });

  it('refuses a non-ASCII entry name rather than quietly setting the UTF-8 flag', () => {
    // Part names are percent-encoded before they reach the archive, so a
    // non-ASCII byte here means something skipped that step.
    expect(codeOf(() => writeZip([deflatedEntry('ppt/média.xml', HELLO)]))).toBe(
      'ERR_INVALID_PART_NAME',
    );
  });

  it('refuses more entries than a ZIP32 central directory can count', () => {
    const many = {
      length: 65536,
      [Symbol.iterator]: function* () {
        for (let i = 0; i < 65536; i++) yield storedEntry('a' + String(i), new Uint8Array(0));
      },
    };
    expect(codeOf(() => writeZip([...many] as never))).toBe('ERR_ZIP32_OVERFLOW');
  });
});
