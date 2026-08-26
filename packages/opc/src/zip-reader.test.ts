import { describe, expect, it } from 'vitest';
import { crc32 } from './crc32.js';
import { isOpcError } from './errors.js';
import { partNameFromZipEntry } from './pack-uri.js';
import { buildZip, incompressible, type FixtureEntry } from './testing/build-zip.js';
import { readZip, type ReadZipOptions } from './zip-reader.js';

const encoder = new TextEncoder();
const decoder = new TextDecoder();

const KIB = 1024;
const MIB = 1024 * KIB;

function bytes(text: string): Uint8Array {
  return encoder.encode(text);
}

function text(data: Uint8Array): string {
  return decoder.decode(data);
}

/** The code of the `OpcError` thrown by `fn`, or a description of what it did instead. */
function codeOf(fn: () => unknown): string {
  try {
    fn();
  } catch (error) {
    if (isOpcError(error)) return error.code;
    return 'NOT_AN_OPC_ERROR: ' + (error instanceof Error ? error.constructor.name : typeof error);
  }
  return 'DID_NOT_THROW';
}

/** A three-part package that stands in for a real one. */
const MINIMAL: FixtureEntry[] = [
  { name: '[Content_Types].xml', data: bytes('<Types xmlns="..."/>') },
  { name: '_rels/.rels', data: bytes('<Relationships xmlns="..."/>') },
  { name: 'ppt/presentation.xml', data: bytes('<p:presentation xmlns:p="..."/>') },
];

function openMinimal(options?: ReadZipOptions) {
  return readZip(buildZip(MINIMAL), options);
}

// ---------------------------------------------------------------------------
// The archives we are supposed to open.
// ---------------------------------------------------------------------------

describe('reading an honest archive', () => {
  it('round-trips content through deflate', () => {
    const archive = openMinimal();
    expect(archive.entries.map((e) => e.name)).toEqual([
      '[Content_Types].xml',
      '_rels/.rels',
      'ppt/presentation.xml',
    ]);
    expect(text(archive.readByName('ppt/presentation.xml'))).toBe(
      '<p:presentation xmlns:p="..."/>',
    );
  });

  it('preserves central-directory order even for integer-like names', () => {
    // This is why entries are an array. `fflate.unzipSync` returns an object,
    // and JavaScript object key order puts integer-like keys first, so the
    // order a package was written in silently changes on the way in - which
    // then changes what the writer emits.
    const archive = readZip(
      buildZip([{ name: '10' }, { name: '2' }, { name: 'ppt/x.xml' }, { name: '1' }]),
    );
    expect(archive.entries.map((e) => e.name)).toEqual(['10', '2', 'ppt/x.xml', '1']);
    expect(Object.keys({ '10': 1, '2': 1, 'ppt/x.xml': 1, '1': 1 })).toEqual([
      '1',
      '2',
      '10',
      'ppt/x.xml',
    ]);
  });

  it('reads stored (uncompressed) entries', () => {
    const payload = incompressible(4 * KIB);
    const archive = readZip(
      buildZip([{ name: 'ppt/media/image1.png', data: payload, deflate: false }]),
    );
    expect(archive.entries[0]?.method).toBe(0);
    expect(archive.readByName('ppt/media/image1.png')).toEqual(payload);
  });

  it('reads entries written with a trailing data descriptor', () => {
    // Bit 3 zeroes the sizes in the local header and puts them in a trailer.
    // A reader that trusts the local header sees a zero-length part; the
    // central directory is authoritative and is what we read.
    const archive = readZip(
      buildZip([{ name: 'ppt/presentation.xml', data: bytes('<p:presentation/>'), flags: 0x0008 }]),
    );
    expect(archive.entries[0]?.hasDataDescriptor).toBe(true);
    expect(text(archive.readByName('ppt/presentation.xml'))).toBe('<p:presentation/>');
  });

  it('reads UTF-8 entry names', () => {
    const archive = readZip(buildZip([{ name: 'ppt/media/图片.png', data: bytes('x') }]));
    expect(archive.entries[0]?.name).toBe('ppt/media/图片.png');
    expect(archive.entries[0]?.flags).toBe(0x0800);
  });

  it('reads an archive with a comment', () => {
    const archive = readZip(buildZip(MINIMAL, { comment: bytes('written by something') }));
    expect(archive.entries).toHaveLength(3);
  });

  it('reads ZIP64 sizes and a ZIP64 end-of-central-directory record', () => {
    // Our writer will never emit ZIP64, but other producers emit the record
    // even for small archives, and refusing it would fail ordinary files.
    const archive = readZip(
      buildZip([{ name: 'ppt/presentation.xml', data: bytes('<p:presentation/>'), zip64: true }], {
        zip64Eocd: true,
      }),
    );
    expect(text(archive.readByName('ppt/presentation.xml'))).toBe('<p:presentation/>');
  });

  it('hands back stored bytes without inflating or charging budget', () => {
    // This is what the writer streams through for parts nobody edited.
    const archive = openMinimal();
    const entry = archive.entries[2]!;
    expect(archive.raw(entry)).toHaveLength(entry.compressedSize);
    expect(archive.budget.used).toBe(0);
  });

  it('tolerates directory entries by reading them as empty', () => {
    const archive = readZip(buildZip([{ name: 'ppt/' }, { name: 'ppt/x.xml', data: bytes('x') }]));
    expect(archive.entries[0]?.isDirectory).toBe(true);
    expect(archive.readByName('ppt/')).toHaveLength(0);
  });
});

// ---------------------------------------------------------------------------
// Decompression bombs.
// ---------------------------------------------------------------------------

describe('decompression bombs', () => {
  it('rejects an honest high-ratio bomb on the ratio ceiling, without inflating it', () => {
    // 8 MiB of zeros deflates to a few kilobytes: a genuine 1000:1 expansion
    // with completely truthful headers, so only the ratio catches it.
    const archive = readZip(buildZip([{ name: 'bomb', data: new Uint8Array(8 * MIB) }]));
    const start = performance.now();
    expect(codeOf(() => archive.readByName('bomb'))).toBe('ERR_RATIO_EXCEEDED');
    expect(performance.now() - start).toBeLessThan(50);
    expect(archive.budget.used).toBe(0);
  });

  it('rejects an entry that declares more than one part may hold, before allocating', () => {
    const archive = readZip(
      buildZip([{ name: 'bomb', data: bytes('x'), declaredUncompressedSize: 900 * MIB }]),
    );
    const start = performance.now();
    expect(codeOf(() => archive.readByName('bomb'))).toBe('ERR_ENTRY_TOO_LARGE');
    expect(performance.now() - start).toBeLessThan(50);
  });

  it('enforces the total budget as a running counter across entries', () => {
    // The counter is the defence that survives an archive of individually
    // reasonable entries. Incompressible data so the ratio test stays out of
    // the way and the budget is the only thing being measured.
    const size = 64 * KIB;
    const archive = readZip(
      buildZip([1, 2, 3, 4].map((n) => ({ name: 'part' + String(n), data: incompressible(size) }))),
      { limits: { maxTotalInflatedBytes: 3 * size } },
    );

    expect(archive.readByName('part1')).toHaveLength(size);
    expect(archive.readByName('part2')).toHaveLength(size);
    expect(archive.readByName('part3')).toHaveLength(size);
    expect(archive.budget.used).toBe(3 * size);
    expect(archive.budget.remaining).toBe(0);
    expect(codeOf(() => archive.readByName('part4'))).toBe('ERR_BUDGET_EXCEEDED');
  });

  it('charges the budget once per entry, not once per read', () => {
    // Otherwise re-reading a part would exhaust an archive's allowance for no
    // reason a user could act on.
    const size = 64 * KIB;
    const archive = readZip(buildZip([{ name: 'part1', data: incompressible(size) }]), {
      limits: { maxTotalInflatedBytes: size },
    });
    expect(archive.readByName('part1')).toHaveLength(size);
    expect(archive.readByName('part1')).toHaveLength(size);
    expect(archive.readByName('part1')).toHaveLength(size);
    expect(archive.budget.used).toBe(size);
  });

  it('refuses the quoted-overlap construction structurally', () => {
    // Fifield, WOOT '19: each entry's compressed data contains the local
    // headers of the entries that follow, so the declared total grows
    // quadratically against the file size. No legitimate archive overlaps.
    const archive = () =>
      readZip(
        buildZip([
          { name: 'a', data: bytes('aaaa'), overlapNext: true },
          { name: 'b', data: bytes('bbbb') },
        ]),
      );
    expect(codeOf(archive)).toBe('ERR_OVERLAPPING_ENTRY');
  });

  it('refuses an archive that declares more entries than a deck could have', () => {
    const archive = () => readZip(buildZip(MINIMAL, { declaredEntryCount: 25_000 }));
    expect(codeOf(archive)).toBe('ERR_TOO_MANY_ENTRIES');
  });

  it('refuses an archive larger than the ceiling before parsing anything', () => {
    expect(codeOf(() => openMinimal({ limits: { maxArchiveBytes: 32 } }))).toBe(
      'ERR_ARCHIVE_TOO_LARGE',
    );
  });

  it('does not apply the ratio ceiling to small entries', () => {
    // A 40-byte DEFLATE stream expanding to 12 KB of XML is a 300:1 ratio and
    // completely ordinary. A ratio only means something above an absolute
    // size, which is what ratioFloorBytes is for.
    const repetitive = bytes('<a:t>x</a:t>'.repeat(1000));
    const archive = readZip(buildZip([{ name: 'ppt/slides/slide1.xml', data: repetitive }]));
    expect(
      archive.entries[0]!.uncompressedSize / archive.entries[0]!.compressedSize,
    ).toBeGreaterThan(200);
    expect(archive.readByName('ppt/slides/slide1.xml')).toHaveLength(repetitive.length);
  });
});

// ---------------------------------------------------------------------------
// Lying headers.
// ---------------------------------------------------------------------------

describe('archives that disagree with themselves', () => {
  it('catches an entry that declares less than its stream produces', () => {
    // This is the case the output-buffer cap alone cannot see: `fflate` fills
    // the buffer, silently drops the surplus and reports success. Without the
    // CRC-32 check the caller gets 16 bytes of a 4096-byte part and no signal.
    const archive = readZip(
      buildZip([
        { name: 'ppt/presentation.xml', data: new Uint8Array(4096), declaredUncompressedSize: 16 },
      ]),
    );
    expect(codeOf(() => archive.readByName('ppt/presentation.xml'))).toBe('ERR_CRC_MISMATCH');
  });

  it('catches a corrupted CRC', () => {
    const archive = readZip(
      buildZip([{ name: 'ppt/x.xml', data: bytes('<x/>'), declaredCrc32: 0xdeadbeef }]),
    );
    expect(codeOf(() => archive.readByName('ppt/x.xml'))).toBe('ERR_CRC_MISMATCH');
  });

  it('catches a local header that names a different entry', () => {
    // Readers disagree about whether the central directory or the local header
    // wins, so an archive that disagrees with itself means different things to
    // different programs.
    const archive = () =>
      readZip(buildZip([{ name: 'ppt/x.xml', data: bytes('<x/>'), localName: 'ppt/evil.xml' }]));
    expect(codeOf(archive)).toBe('ERR_HEADER_MISMATCH');
  });

  it('catches a local header that disagrees about compression', () => {
    const archive = () =>
      readZip(buildZip([{ name: 'ppt/x.xml', data: bytes('<x/>'), localMethod: 0 }]));
    expect(codeOf(archive)).toBe('ERR_HEADER_MISMATCH');
  });

  it('catches two entries with the same name', () => {
    const archive = () =>
      readZip(
        buildZip([
          { name: 'ppt/x.xml', data: bytes('<a/>') },
          { name: 'ppt/x.xml', data: bytes('<b/>') },
        ]),
      );
    expect(codeOf(archive)).toBe('ERR_DUPLICATE_ENTRY');
  });

  it('catches a stored entry whose two sizes differ', () => {
    const archive = readZip(
      buildZip([
        { name: 'ppt/x.xml', data: bytes('<x/>'), deflate: false, declaredUncompressedSize: 99 },
      ]),
    );
    expect(codeOf(() => archive.readByName('ppt/x.xml'))).toBe('ERR_HEADER_MISMATCH');
  });
});

// ---------------------------------------------------------------------------
// Damaged and mistaken input.
// ---------------------------------------------------------------------------

describe('input that is not the archive it claims to be', () => {
  it('names a password-protected file for what it is', () => {
    // A protected .pptx is an OLE2 compound document with the presentation in
    // an EncryptedPackage stream - not a ZIP at all. "Not a ZIP archive" is a
    // true but useless thing to tell someone who typed a password once.
    const cfb = new Uint8Array(512);
    cfb.set([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]);
    expect(codeOf(() => readZip(cfb))).toBe('ERR_ENCRYPTED_PACKAGE');
  });

  it('rejects a truncated archive without hanging', () => {
    const whole = buildZip(MINIMAL);
    const start = performance.now();
    expect(codeOf(() => readZip(whole.subarray(0, whole.length >> 1)))).toBe('ERR_NOT_A_ZIP');
    expect(codeOf(() => readZip(whole.subarray(0, whole.length - 6)))).toBe('ERR_NOT_A_ZIP');
    expect(codeOf(() => readZip(whole.subarray(0, 4)))).toBe('ERR_TRUNCATED');
    expect(codeOf(() => readZip(new Uint8Array(0)))).toBe('ERR_TRUNCATED');
    expect(performance.now() - start).toBeLessThan(500);
  });

  it('rejects an entry whose data runs past the end of the archive', () => {
    const archive = () =>
      readZip(
        buildZip([{ name: 'ppt/x.xml', data: bytes('<x/>'), declaredCompressedSize: 1 << 20 }]),
      );
    expect(codeOf(archive)).toBe('ERR_TRUNCATED');
  });

  it('rejects a compression method PowerPoint would not open either', () => {
    const archive = readZip(
      buildZip([{ name: 'ppt/x.xml', data: bytes('<x/>'), method: 12, localMethod: 12 }]),
    );
    expect(codeOf(() => archive.readByName('ppt/x.xml'))).toBe('ERR_UNSUPPORTED_COMPRESSION');
  });

  it('rejects an encrypted entry', () => {
    const archive = readZip(buildZip([{ name: 'ppt/x.xml', data: bytes('<x/>'), flags: 0x0001 }]));
    expect(codeOf(() => archive.readByName('ppt/x.xml'))).toBe('ERR_ENTRY_ENCRYPTED');
  });

  it('rejects an entry name past the ceiling', () => {
    expect(codeOf(() => readZip(buildZip([{ name: 'a'.repeat(2000) }])))).toBe(
      'ERR_INVALID_PART_NAME',
    );
  });

  it('rejects a malformed DEFLATE stream as an OpcError, not a raw throw', () => {
    const good = buildZip([{ name: 'ppt/x.xml', data: bytes('<x/>'.repeat(200)) }]);
    const archive = readZip(good);
    const entry = archive.entries[0]!;
    // Corrupt the middle of the compressed stream, leaving the headers intact.
    good[entry.dataOffset + 4] = good[entry.dataOffset + 4]! ^ 0xff;
    good[entry.dataOffset + 5] = good[entry.dataOffset + 5]! ^ 0xff;
    const code = codeOf(() => readZip(good).readByName('ppt/x.xml'));
    expect(['ERR_INFLATE_FAILED', 'ERR_CRC_MISMATCH', 'ERR_TRUNCATED']).toContain(code);
  });

  it('wraps the raw RangeError fflate throws on a stored-block overrun', () => {
    // Empirically: with a caller-supplied output buffer, fflate silently drops
    // surplus bytes on the Huffman path but throws a bare
    // `RangeError: offset is out of bounds` on the stored-block path. A raw
    // RangeError escaping this package would break the parser invariant the
    // fuzzing harness in 12.5 asserts, so it is wrapped.
    const archive = readZip(
      buildZip([
        {
          name: 'ppt/x.xml',
          data: incompressible(64 * KIB),
          deflateLevel: 0,
          declaredUncompressedSize: 1024,
        },
      ]),
    );
    expect(codeOf(() => archive.readByName('ppt/x.xml'))).toBe('ERR_INFLATE_FAILED');
  });
});

// ---------------------------------------------------------------------------
// Zip slip.
// ---------------------------------------------------------------------------

describe('zip slip', () => {
  it.each([
    '../../../etc/passwd',
    'ppt/../../../evil.xml',
    '/etc/passwd',
    'C:/Windows/System32/evil.dll',
    'ppt\\slides\\slide1.xml',
    'ppt/slide1.xml\u0000.png',
    './hidden',
  ])('accepts %s as a ZIP entry name and refuses it as a part name', (name) => {
    // The layering is deliberate. ZIP has no opinion about entry names, so the
    // reader stays honest about what the archive contains; the OPC grammar is
    // where a name has to be usable, and it is the grammar - not a blocklist -
    // that refuses these.
    const archive = readZip(buildZip([{ name, data: bytes('x') }]));
    expect(archive.entries[0]?.name).toBe(name);
    expect(codeOf(() => partNameFromZipEntry(name))).toBe('ERR_INVALID_PART_NAME');
  });
});

// ---------------------------------------------------------------------------
// The invariant that makes all of the above worth having.
// ---------------------------------------------------------------------------

describe('the error invariant', () => {
  it('throws only OpcError, over a systematic sweep of single-byte corruptions', () => {
    // Phase 12.5 fuzzes this properly. This is the cheap version that runs on
    // every commit: the parser either returns data or throws an OpcError. It
    // never throws a raw RangeError, never hangs, never over-allocates.
    const original = buildZip([
      { name: '[Content_Types].xml', data: bytes('<Types/>') },
      { name: '_rels/.rels', data: bytes('<Relationships/>') },
      { name: 'ppt/presentation.xml', data: bytes('<p:presentation/>'.repeat(20)) },
      { name: 'ppt/media/image1.png', data: incompressible(512), deflate: false },
    ]);

    let seenThrows = 0;
    const start = performance.now();

    // A fixed stride rather than random offsets, so a failure is reproducible.
    for (let offset = 0; offset < original.length; offset += 3) {
      for (const mask of [0xff, 0x01, 0x80]) {
        const mutated = original.slice();
        mutated[offset] = (mutated[offset]! ^ mask) & 0xff;
        try {
          const archive = readZip(mutated);
          for (const entry of archive.entries) archive.read(entry);
        } catch (error) {
          seenThrows++;
          if (!isOpcError(error)) {
            throw new Error(
              'offset ' +
                String(offset) +
                ' mask 0x' +
                mask.toString(16) +
                ' threw ' +
                (error instanceof Error
                  ? error.constructor.name + ': ' + error.message
                  : String(error)),
              { cause: error },
            );
          }
        }
      }
    }

    expect(seenThrows).toBeGreaterThan(0);
    expect(performance.now() - start).toBeLessThan(20_000);
  });

  it('every entry read back equals what went in', () => {
    const parts = [
      { name: 'ppt/slides/slide1.xml', data: bytes('<p:sld/>') },
      { name: 'ppt/media/image1.png', data: incompressible(8 * KIB) },
      { name: 'ppt/media/image2.png', data: incompressible(3 * KIB), deflate: false },
      { name: 'ppt/fonts/font1.fntdata', data: incompressible(KIB) },
    ];
    const archive = readZip(buildZip(parts));
    for (const part of parts) {
      const read = archive.readByName(part.name);
      expect(read).toEqual(part.data);
      expect(crc32(read)).toBe(archive.get(part.name)?.crc32);
    }
  });
});
