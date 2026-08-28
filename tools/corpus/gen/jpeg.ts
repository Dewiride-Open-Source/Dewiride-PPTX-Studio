/**
 * A baseline JPEG, written by hand, for `a33-thumbnail`.
 *
 * ## Why this exists rather than a committed `.jpeg`
 *
 * `docProps/thumbnail.jpeg` is the one part of a PowerPoint package that is a
 * photograph of the document rather than part of it, and `a33-thumbnail` needs
 * one that is ours. Committing an image somebody else made would put a second
 * licence in a manifest whose whole point is that every entry is CC0 and
 * self-authored; `png.ts` made the same call for the same reason.
 *
 * ## What PowerPoint writes, measured
 *
 * Build 16.0.20326 saving a one-slide deck on 2026-08-28 wrote a 1203-byte
 * `docProps/thumbnail.jpeg`: **baseline** (SOF0), JFIF 1.1 at 96 x 96 dpi,
 * 256 x 144 pixels, three components with 2x2 luma subsampling - 4:2:0 - two
 * quantization tables of its own, and the four **standard** Huffman tables
 * from ITU-T T.81 Annex K.
 *
 * Those four tables are what `LUMINANCE_DC_BITS` and friends below are, and
 * they are not recalled: they were read out of that file's `DHT` segments and
 * compared byte for byte. Annex K is where they come from - every encoder ever
 * written ships the same arrays - but having verified ours against a real
 * decoder's input is the difference between a fixture and a guess.
 *
 * ## What this writes, and the two deliberate differences
 *
 * **Greyscale, one component.** A thumbnail exists to be looked at, not to
 * probe chroma subsampling, and one component removes the `SOF0` sampling
 * factors, two of the four Huffman tables, one quantization table and the
 * component interleave from the entropy-coded segment. `a25-svg-blips` and
 * `a03-fills` are where image content is the probe; here the probe is the part.
 *
 * **An all-ones quantization table.** Legal - `ST_QuantizationValue` is 1..255
 * for 8-bit precision - and it makes the encoder exact rather than approximate.
 * With every divisor 1, the DC coefficient of a block of constant level `v` is
 * exactly `8 * (v - 128)`, so a picture built from 8 x 8 blocks of flat grey
 * round-trips through the file with no error at all and needs no DCT: every
 * AC coefficient is zero and every block is one DC coefficient followed by an
 * end-of-block. That is the whole encoder, and it is why this file is a few
 * hundred bytes and has no floating-point arithmetic in it.
 *
 * The cost is stated rather than hidden: this cannot encode a photograph. It
 * encodes a **blocky** image at 8 x 8 granularity, which is all a 256 x 144
 * thumbnail of a slide needs to be recognisable, and nothing in the corpus
 * asks it for more.
 */

/** ITU-T T.81 Annex K, table K.3 - the DC luminance code lengths. */
const DC_BITS = [0, 1, 5, 1, 1, 1, 1, 1, 1, 0, 0, 0, 0, 0, 0, 0] as const;
/** The twelve DC difference categories, in code order. */
const DC_VALUES = [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11] as const;

/** Annex K, table K.5 - the AC luminance code lengths. 162 symbols. */
const AC_BITS = [0, 2, 1, 3, 3, 2, 4, 3, 5, 5, 4, 4, 0, 0, 1, 0x7d] as const;

/**
 * Annex K, table K.5's symbol list: `(run << 4) | size`, in code order.
 *
 * `0x00` is end-of-block and `0xf0` is a run of sixteen zeroes. Every other
 * value is a run of 0..15 zero coefficients followed by a coefficient needing
 * 1..10 bits. Read out of the `DHT` segment of a PowerPoint-written thumbnail
 * rather than typed from memory.
 */
const AC_VALUES = [
  0x01, 0x02, 0x03, 0x00, 0x04, 0x11, 0x05, 0x12, 0x21, 0x31, 0x41, 0x06, 0x13, 0x51, 0x61, 0x07,
  0x22, 0x71, 0x14, 0x32, 0x81, 0x91, 0xa1, 0x08, 0x23, 0x42, 0xb1, 0xc1, 0x15, 0x52, 0xd1, 0xf0,
  0x24, 0x33, 0x62, 0x72, 0x82, 0x09, 0x0a, 0x16, 0x17, 0x18, 0x19, 0x1a, 0x25, 0x26, 0x27, 0x28,
  0x29, 0x2a, 0x34, 0x35, 0x36, 0x37, 0x38, 0x39, 0x3a, 0x43, 0x44, 0x45, 0x46, 0x47, 0x48, 0x49,
  0x4a, 0x53, 0x54, 0x55, 0x56, 0x57, 0x58, 0x59, 0x5a, 0x63, 0x64, 0x65, 0x66, 0x67, 0x68, 0x69,
  0x6a, 0x73, 0x74, 0x75, 0x76, 0x77, 0x78, 0x79, 0x7a, 0x83, 0x84, 0x85, 0x86, 0x87, 0x88, 0x89,
  0x8a, 0x92, 0x93, 0x94, 0x95, 0x96, 0x97, 0x98, 0x99, 0x9a, 0xa2, 0xa3, 0xa4, 0xa5, 0xa6, 0xa7,
  0xa8, 0xa9, 0xaa, 0xb2, 0xb3, 0xb4, 0xb5, 0xb6, 0xb7, 0xb8, 0xb9, 0xba, 0xc2, 0xc3, 0xc4, 0xc5,
  0xc6, 0xc7, 0xc8, 0xc9, 0xca, 0xd2, 0xd3, 0xd4, 0xd5, 0xd6, 0xd7, 0xd8, 0xd9, 0xda, 0xe1, 0xe2,
  0xe3, 0xe4, 0xe5, 0xe6, 0xe7, 0xe8, 0xe9, 0xea, 0xf1, 0xf2, 0xf3, 0xf4, 0xf5, 0xf6, 0xf7, 0xf8,
  0xf9, 0xfa,
] as const;

interface Code {
  readonly bits: number;
  readonly length: number;
}

/**
 * The canonical code assignment of T.81 Annex C: walk the lengths in order,
 * emitting consecutive codes and shifting left at every length boundary.
 */
function buildCodes(bits: readonly number[], values: readonly number[]): Map<number, Code> {
  const codes = new Map<number, Code>();
  let code = 0;
  let index = 0;
  for (let length = 1; length <= 16; length++) {
    for (let n = 0; n < (bits[length - 1] ?? 0); n++) {
      const value = values[index++];
      if (value === undefined) throw new Error('huffman table is shorter than its BITS say');
      codes.set(value, { bits: code, length });
      code += 1;
    }
    code <<= 1;
  }
  return codes;
}

const DC_CODES = buildCodes([...DC_BITS], [...DC_VALUES]);
const AC_CODES = buildCodes([...AC_BITS], [...AC_VALUES]);

/** A bit sink that does JPEG's `0xFF` -> `0xFF 0x00` stuffing as it goes. */
class BitWriter {
  private readonly out: number[] = [];
  private accumulator = 0;
  private width = 0;

  write(bits: number, length: number): void {
    for (let i = length - 1; i >= 0; i--) {
      this.accumulator = (this.accumulator << 1) | ((bits >> i) & 1);
      this.width += 1;
      if (this.width !== 8) continue;
      this.out.push(this.accumulator & 0xff);
      // Inside entropy-coded data an 0xFF byte would look like the start of a
      // marker, so a zero byte is stuffed after it. A decoder that skips this
      // step reads the next segment as image data.
      if ((this.accumulator & 0xff) === 0xff) this.out.push(0x00);
      this.accumulator = 0;
      this.width = 0;
    }
  }

  /** Pad to a byte boundary with 1 bits, which T.81 F.1.2.3 requires. */
  finish(): number[] {
    while (this.width !== 0) this.write(1, 1);
    return this.out;
  }
}

/** The number of bits a nonzero DC difference needs: its T.81 "category". */
function category(value: number): number {
  let magnitude = Math.abs(value);
  let bits = 0;
  while (magnitude > 0) {
    magnitude >>= 1;
    bits += 1;
  }
  return bits;
}

export interface JpegSpec {
  readonly width: number;
  readonly height: number;
  /** The grey level, 0..255, of the 8 x 8 block at this block coordinate. */
  readonly level: (blockX: number, blockY: number) => number;
  /** Pixels per inch, into the JFIF segment. PowerPoint writes 96. */
  readonly density?: number;
}

/**
 * A greyscale baseline JPEG of flat 8 x 8 blocks.
 *
 * Marker order is the one every decoder expects and the one PowerPoint writes:
 * `SOI`, `APP0`, `DQT`, `SOF0`, `DHT`, `DHT`, `SOS`, entropy-coded data, `EOI`.
 */
export function probeJpeg(spec: JpegSpec): Uint8Array {
  const { width, height } = spec;
  if (width < 1 || height < 1 || width > 65535 || height > 65535) {
    throw new Error('a JPEG frame is 1..65535 in each axis');
  }
  const density = spec.density ?? 96;
  const out: number[] = [];
  const u8 = (...bytes: number[]): void => {
    out.push(...bytes);
  };
  const u16 = (value: number): void => {
    out.push((value >> 8) & 0xff, value & 0xff);
  };

  u8(0xff, 0xd8); // SOI

  // APP0/JFIF. Not required by T.81 - a bare SOI/DQT/SOF0 file decodes - but
  // it is what every writer emits, it is where the density lives, and a file
  // without it is the kind of technically-legal that shell thumbnailers
  // occasionally decline.
  u8(0xff, 0xe0);
  u16(16);
  u8(0x4a, 0x46, 0x49, 0x46, 0x00); // "JFIF\0"
  u8(1, 1); // version 1.1
  u8(1); // units: pixels per inch
  u16(density);
  u16(density);
  u8(0, 0); // no embedded thumbnail of the thumbnail

  // DQT, one 8-bit table of all ones. See the header for why this is exact.
  u8(0xff, 0xdb);
  u16(2 + 1 + 64);
  u8(0x00); // precision 0 (8-bit), table id 0
  for (let i = 0; i < 64; i++) u8(1);

  // SOF0: baseline, 8-bit, one component at 1x1 using quantization table 0.
  u8(0xff, 0xc0);
  u16(8 + 3);
  u8(8);
  u16(height);
  u16(width);
  u8(1);
  u8(1, 0x11, 0);

  const huffman = (className: 0 | 1, bits: readonly number[], values: readonly number[]): void => {
    u8(0xff, 0xc4);
    u16(2 + 1 + 16 + values.length);
    u8((className << 4) | 0);
    for (const b of bits) u8(b);
    for (const v of values) u8(v);
  };
  huffman(0, DC_BITS, DC_VALUES);
  huffman(1, AC_BITS, AC_VALUES);

  // SOS: one component, DC table 0, AC table 0, full spectral selection.
  u8(0xff, 0xda);
  u16(6 + 2);
  u8(1);
  u8(1, 0x00);
  u8(0, 63, 0x00);

  const writer = new BitWriter();
  const columns = Math.ceil(width / 8);
  const rows = Math.ceil(height / 8);
  let previous = 0;
  for (let by = 0; by < rows; by++) {
    for (let bx = 0; bx < columns; bx++) {
      const level = Math.max(0, Math.min(255, Math.round(spec.level(bx, by))));
      // The DCT of a constant block, with every quantization divisor 1.
      const dc = 8 * (level - 128);
      const diff = dc - previous;
      previous = dc;

      const size = diff === 0 ? 0 : category(diff);
      const dcCode = DC_CODES.get(size);
      if (dcCode === undefined) throw new Error('DC category ' + String(size) + ' has no code');
      writer.write(dcCode.bits, dcCode.length);
      if (size > 0) {
        // A negative difference is written as its value minus one, in `size`
        // bits - which is the two's-complement low bits of `diff - 1`.
        writer.write((diff < 0 ? diff - 1 : diff) & ((1 << size) - 1), size);
      }

      // Every AC coefficient is zero, so the block is one end-of-block symbol.
      const eob = AC_CODES.get(0x00);
      if (eob === undefined) throw new Error('the AC table has no end-of-block symbol');
      writer.write(eob.bits, eob.length);
    }
  }
  out.push(...writer.finish());

  u8(0xff, 0xd9); // EOI
  return Uint8Array.from(out);
}
