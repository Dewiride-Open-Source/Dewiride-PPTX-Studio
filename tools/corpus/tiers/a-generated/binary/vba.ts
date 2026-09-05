/**
 * A VBA project, for `a32-macros`.
 *
 * ## What is measured here, and what is not
 *
 * This is the least-verified module in the generator and says so up front.
 *
 * **Measured.** That a macro-enabled package is nothing but a content type:
 * PowerPoint 16.0.20326, asked on 2026-08-28 to save a one-slide deck as
 * `.pptm`, wrote a package identical to the `.pptx` one except that
 * `/ppt/presentation.xml` is typed
 * `application/vnd.ms-powerpoint.presentation.macroEnabled.main+xml`. There was
 * **no `vbaProject.bin` at all**, because there were no macros. So the content
 * type and the part are independent, and a deck that probes the census's
 * `macros` feature - which keys on a part named `vbaProject.bin` and on nothing
 * else - has to carry a real one.
 *
 * **Not measured.** Everything inside the part. Authoring a VBA module through
 * COM needs "Trust access to the VBA project object model", which is off by
 * default and is a Trust Center setting on the user's machine rather than
 * something this project may change. So the container below is built from the
 * specification and checked structurally - `cfb.test.ts` reads it back, and the
 * compressor round-trips - and **whether the VBA engine would compile the
 * project is not verified by anything here.** PowerPoint does not load a VBA
 * project when it opens a file; macros are disabled by default and the project
 * is read when the editor opens or a macro runs. A deck that opens is therefore
 * evidence about the package and not about the project.
 *
 * `ROSTER.md` carries that as a declared gap. It closes with a Tier B deck on a
 * machine where the setting is on.
 *
 * ## The pieces
 *
 * ```
 * /PROJECT            plain text, the project's own manifest
 * /PROJECTwm          module names, MBCS then UTF-16, terminated
 * /VBA                a storage
 * /VBA/_VBA_PROJECT   a version stamp and a performance cache we omit
 * /VBA/dir            a compressed record stream describing the project
 * /VBA/Module1        a compressed copy of the source text
 * ```
 *
 * `dir` and the module stream are compressed with the scheme in MS-OVBA §2.4.1,
 * which is a small LZ77: a `0x01` signature byte, then chunks of at most 4096
 * decompressed bytes, each a 16-bit header followed by flag-byte-driven token
 * sequences. The part that is easy to get wrong, and the reason `compressOvba`
 * and `decompressOvba` are both here and tested against each other, is that the
 * **copy token's bit split changes as the chunk fills**: the number of bits
 * given to the offset is `max(4, ceil(log2(position)))`, recomputed before
 * every token, so the same two bytes mean different things at different points
 * in the same stream.
 *
 * ## The constants that are asserted rather than measured
 *
 * Six `Reserved` fields inside `dir` carry fixed values - `0x0040`, `0x003D`,
 * `0x003C`, `0x003E`, `0x0032`, `0x0048` - and two records use their 4-byte
 * field as a reserved constant rather than as a size. They are named in the
 * code below at the point they are written. A wrong one fails the project load
 * and nothing else, which is exactly the failure this module cannot see.
 */

const encoder = new TextEncoder();

/* -------------------------------------------------------------------------- */
/* MS-OVBA §2.4.1 compression                                                 */
/* -------------------------------------------------------------------------- */

/** The bit split for a copy token, recomputed before every token. */
function copyTokenShape(position: number): {
  bits: number;
  lengthBits: number;
  lengthMask: number;
} {
  let bits = 4;
  while (1 << bits < position) bits += 1;
  return { bits, lengthBits: 16 - bits, lengthMask: 0xffff >> bits };
}

/**
 * Compress a byte string into an MS-OVBA container.
 *
 * Greedy matching, which is what the format's tiny window rewards: the longest
 * back-reference at each position, or a literal.
 */
export function compressOvba(input: Uint8Array): Uint8Array {
  // The signature byte alone is a valid, empty container. Emitting a chunk for
  // no data would declare a length of 4096 - the size field holds `length - 1`,
  // so zero bytes and 4096 bytes are the same twelve bits.
  if (input.length === 0) return Uint8Array.from([0x01]);

  const out: number[] = [0x01];

  for (let chunkStart = 0; chunkStart < input.length; chunkStart += 4096) {
    const chunkEnd = Math.min(chunkStart + 4096, input.length);
    const data: number[] = [];

    let at = chunkStart;
    while (at < chunkEnd) {
      // One flag byte, then up to eight tokens. Bit 0 is the first token.
      const flagIndex = data.length;
      data.push(0);
      let flags = 0;
      for (let bit = 0; bit < 8 && at < chunkEnd; bit++) {
        const position = at - chunkStart;
        const { bits, lengthBits, lengthMask } = copyTokenShape(position);
        const maxOffset = 1 << bits;
        const maxLength = lengthMask + 3;

        let bestLength = 0;
        let bestOffset = 0;
        const earliest = Math.max(chunkStart, at - maxOffset);
        for (let candidate = at - 1; candidate >= earliest; candidate--) {
          let length = 0;
          while (
            length < maxLength &&
            at + length < chunkEnd &&
            input[candidate + length] === input[at + length]
          ) {
            length += 1;
          }
          if (length > bestLength) {
            bestLength = length;
            bestOffset = at - candidate;
          }
          if (bestLength === maxLength) break;
        }

        if (bestLength >= 3) {
          const token = ((bestOffset - 1) << lengthBits) | (bestLength - 3);
          data.push(token & 0xff, (token >> 8) & 0xff);
          flags |= 1 << bit;
          at += bestLength;
        } else {
          data.push(input[at] ?? 0);
          at += 1;
        }
      }
      data[flagIndex] = flags;
    }

    if (data.length > 4096) {
      // The chunk size field is twelve bits, so a compressed chunk's data
      // cannot exceed 4096 bytes. The format's answer is a raw chunk, whose
      // padding rules are ambiguous enough that this refuses instead - no
      // input this generator has needs one.
      throw new Error('a compressed chunk exceeded 4096 bytes; a raw chunk would be needed');
    }
    const header = 0xb000 | ((data.length - 1) & 0x0fff);
    out.push(header & 0xff, (header >> 8) & 0xff, ...data);
  }

  return Uint8Array.from(out);
}

/** The inverse, written independently so the pair can be tested against itself. */
export function decompressOvba(input: Uint8Array): Uint8Array {
  if (input[0] !== 0x01) throw new Error('not an MS-OVBA container: no 0x01 signature');
  const out: number[] = [];
  let at = 1;
  while (at + 1 < input.length) {
    const header = (input[at] ?? 0) | ((input[at + 1] ?? 0) << 8);
    at += 2;
    if (((header >> 12) & 0b111) !== 0b011) throw new Error('bad chunk signature');
    const dataLength = (header & 0x0fff) + 1;
    const compressed = (header & 0x8000) !== 0;
    const end = at + dataLength;
    if (!compressed) {
      for (; at < end; at++) out.push(input[at] ?? 0);
      continue;
    }
    const chunkStart = out.length;
    while (at < end) {
      const flags = input[at] ?? 0;
      at += 1;
      for (let bit = 0; bit < 8 && at < end; bit++) {
        if ((flags & (1 << bit)) === 0) {
          out.push(input[at] ?? 0);
          at += 1;
          continue;
        }
        const token = (input[at] ?? 0) | ((input[at + 1] ?? 0) << 8);
        at += 2;
        const { lengthBits, lengthMask } = copyTokenShape(out.length - chunkStart);
        const length = (token & lengthMask) + 3;
        const offset = ((token & ~lengthMask & 0xffff) >> lengthBits) + 1;
        // Read after write, so an overlapping copy repeats, LZ77 style.
        for (let i = 0; i < length; i++) out.push(out[out.length - offset] ?? 0);
      }
    }
  }
  return Uint8Array.from(out);
}

/* -------------------------------------------------------------------------- */
/* the dir stream                                                             */
/* -------------------------------------------------------------------------- */

const u16 = (value: number): number[] => [value & 0xff, (value >> 8) & 0xff];
const u32 = (value: number): number[] => [
  value & 0xff,
  (value >> 8) & 0xff,
  (value >> 16) & 0xff,
  (value >> 24) & 0xff,
];

/** MBCS, which for an ASCII project name is the same bytes as UTF-8. */
const mbcs = (text: string): number[] => [...encoder.encode(text)];

const utf16 = (text: string): number[] => {
  const out: number[] = [];
  for (let i = 0; i < text.length; i++) out.push(...u16(text.charCodeAt(i)));
  return out;
};

/** `Id`, a 4-byte size, then the payload. The common record shape. */
const record = (id: number, payload: readonly number[]): number[] => [
  ...u16(id),
  ...u32(payload.length),
  ...payload,
];

/**
 * `Id` then a 4-byte **Reserved** constant, which is not a size.
 *
 * `PROJECTVERSION`, `MODULETYPE` and both terminators are this shape, and
 * reading their second field as a length is the classic way to walk off the
 * end of the stream.
 */
const reservedRecord = (id: number, reserved: number, tail: readonly number[] = []): number[] => [
  ...u16(id),
  ...u32(reserved),
  ...tail,
];

/**
 * A record carrying an MBCS string and its UTF-16 twin, with a fixed `Reserved`
 * between them.
 *
 * Five records have this shape and each has its **own** reserved value. They
 * are the constants this module cannot check.
 */
const pairedRecord = (id: number, reserved: number, text: string): number[] => {
  const narrow = mbcs(text);
  const wide = utf16(text);
  return [
    ...u16(id),
    ...u32(narrow.length),
    ...narrow,
    ...u16(reserved),
    ...u32(wide.length),
    ...wide,
  ];
};

export interface VbaProjectSpec {
  /** `PROJECTNAME`, and the `Name=` line in `/PROJECT`. */
  readonly projectName: string;
  /** The module's name, which has to agree in six places. */
  readonly moduleName: string;
  /** The module source, without its `Attribute VB_Name` line. */
  readonly source: string;
}

/** `/VBA/dir`, decompressed. */
function dirStream(spec: VbaProjectSpec): Uint8Array {
  const out: number[] = [
    // PROJECTINFORMATION
    ...record(0x0001, u32(1)), // PROJECTSYSKIND: 1 = Win32
    ...record(0x0002, u32(0x0409)), // PROJECTLCID
    ...record(0x0014, u32(0x0409)), // PROJECTLCIDINVOKE
    ...record(0x0003, u16(0x04e4)), // PROJECTCODEPAGE: 1252
    ...record(0x0004, mbcs(spec.projectName)), // PROJECTNAME
    ...pairedRecord(0x0005, 0x0040, ''), // PROJECTDOCSTRING
    ...pairedRecord(0x0006, 0x003d, ''), // PROJECTHELPFILEPATH
    ...record(0x0007, u32(0)), // PROJECTHELPCONTEXT
    ...record(0x0008, u32(0)), // PROJECTLIBFLAGS
    // PROJECTVERSION: the 4-byte field is Reserved = 4, then major and minor.
    ...reservedRecord(0x0009, 4, [...u32(0), ...u16(3)]),
    ...pairedRecord(0x000c, 0x003c, ''), // PROJECTCONSTANTS

    // PROJECTREFERENCES: none. Whether a project with no references loads is
    // one of the things this module cannot check; `ROSTER.md` says so.

    // PROJECTMODULES
    ...record(0x000f, u16(1)), // one module
    ...record(0x0013, u16(0xffff)), // PROJECTCOOKIE, ignored on read
    ...record(0x0019, mbcs(spec.moduleName)), // MODULENAME
    ...record(0x0047, utf16(spec.moduleName)), // MODULENAMEUNICODE
    ...pairedRecord(0x001a, 0x0032, spec.moduleName), // MODULESTREAMNAME
    ...pairedRecord(0x001c, 0x0048, ''), // MODULEDOCSTRING
    ...record(0x0031, u32(0)), // MODULEOFFSET: no performance cache
    ...record(0x001e, u32(0)), // MODULEHELPCONTEXT
    ...record(0x002c, u16(0xffff)), // MODULECOOKIE
    ...reservedRecord(0x0021, 0), // MODULETYPE: procedural
    ...reservedRecord(0x002b, 0), // MODULE terminator
    ...reservedRecord(0x0010, 0), // dir terminator
  ];
  return Uint8Array.from(out);
}

/** `/PROJECT`, plain text with CRLF line endings. */
function projectStream(spec: VbaProjectSpec): Uint8Array {
  const lines = [
    'ID="{00000000-0000-0000-0000-000000000000}"',
    'Module=' + spec.moduleName,
    'Name="' + spec.projectName + '"',
    'HelpContextID="0"',
    'VersionCompatible32="393222000"',
    // CMG, DPB and GC encode the project's protection state with the "data
    // encryption" scheme in MS-OVBA §2.4.3, whose output is seeded and
    // therefore different on every save. They are omitted rather than
    // invented: a wrong value here means something specific about
    // protection, and a missing one means the project is unprotected.
    '',
    '[Host Extender Info]',
    '&H00000001={3832D640-CF90-11CF-8E43-00A0C911005A};VBE;&H00000000',
    '',
    '[Workspace]',
    spec.moduleName + '=0, 0, 0, 0, C',
    '',
  ];
  return encoder.encode(lines.join('\r\n'));
}

/** `/PROJECTwm`: each module's name in MBCS then UTF-16, then a terminator. */
function projectWmStream(spec: VbaProjectSpec): Uint8Array {
  return Uint8Array.from([
    ...mbcs(spec.moduleName),
    0x00,
    ...utf16(spec.moduleName),
    0x00,
    0x00,
    0x00,
    0x00,
  ]);
}

/**
 * `/VBA/_VBA_PROJECT`.
 *
 * `0x61CC`, a version word, and three reserved bytes. Everything after that is
 * a performance cache of compiled p-code that a reader is required to ignore,
 * so it is omitted and VBA recompiles from `dir` and the module source - which
 * is what a hand-authored project wants anyway.
 */
const VBA_PROJECT_STREAM = Uint8Array.from([0xcc, 0x61, 0xff, 0xff, 0x00, 0x00, 0x00]);

/** The module stream: no performance cache, so `MODULEOFFSET` is zero. */
function moduleStream(spec: VbaProjectSpec): Uint8Array {
  // `Attribute VB_Name` is stored in the file and hidden by the editor. Without
  // it the module has no name of its own, and the name has to match `dir`.
  const text = 'Attribute VB_Name = "' + spec.moduleName + '"\r\n' + spec.source;
  return compressOvba(encoder.encode(text));
}

/** The six streams and one storage a minimal VBA project is made of. */
export function vbaProjectNodes(spec: VbaProjectSpec): {
  readonly project: Uint8Array;
  readonly projectWm: Uint8Array;
  readonly vbaProject: Uint8Array;
  readonly dir: Uint8Array;
  readonly module: Uint8Array;
} {
  return {
    project: projectStream(spec),
    projectWm: projectWmStream(spec),
    vbaProject: VBA_PROJECT_STREAM,
    dir: compressOvba(dirStream(spec)),
    module: moduleStream(spec),
  };
}
