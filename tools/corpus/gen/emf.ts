/**
 * A tiny Enhanced Metafile, for the probes that need a vector preview.
 *
 * An OLE object's preview in a modern PowerPoint package is an **EMF**, not a
 * PNG - experiment E6 measured that - and there is no EMF anywhere else in this
 * corpus. Committing PowerPoint's own would mean committing a rendering of a
 * Microsoft Excel worksheet, which is not ours to license, so this writes one.
 *
 * ## What it is
 *
 * Five records, 168 bytes, and it genuinely draws: a header, a brush, a select,
 * a filled rectangle, and the end-of-file record. Opened by anything that reads
 * EMF it produces a blue rectangle with a black outline on a 400 x 300 logical
 * canvas. Nothing is elided or approximated.
 *
 * ## What it is not
 *
 * Not a stress test for sub-phase 10.6. The hard part of EMF is the ~70 record
 * types `rtf.js` routes into a `default:` no-op - text output, bitmap blits,
 * clipping, EMF+ inside `EMR_COMMENT` - and a real Visio or Excel metafile is
 * tens of kilobytes of exactly those. This has none of them, and a renderer
 * that draws this correctly has proved almost nothing. `ROSTER.md` carries that
 * as a declared gap against a Tier B deck.
 *
 * ## Layout notes
 *
 * Every record is `iType`, `nSize`, then its own fields, and `nSize` counts the
 * two header words as well - so the smallest possible record is 8 bytes and
 * every one of these is a multiple of 4. All little-endian.
 *
 * `nHandles` in the header is **one more than the highest handle index used**,
 * because index 0 is reserved. One brush at index 1 means `nHandles` is 2, and
 * getting that wrong is the classic way to write an EMF that GDI rejects.
 */

/** 400 x 300 logical units, and the same in device pixels. */
const WIDTH = 400;
const HEIGHT = 300;

/** A 96 dpi device, described in the units EMF asks for. */
const DEVICE_PIXELS = { cx: 1920, cy: 1080 } as const;
const DEVICE_MM = { cx: 508, cy: 286 } as const;

class Writer {
  private readonly bytes: number[] = [];

  u16(value: number): void {
    this.bytes.push(value & 0xff, (value >>> 8) & 0xff);
  }

  u32(value: number): void {
    this.bytes.push(
      value & 0xff,
      (value >>> 8) & 0xff,
      (value >>> 16) & 0xff,
      (value >>> 24) & 0xff,
    );
  }

  /** `RECTL`: left, top, right, bottom, all signed 32-bit. */
  rect(left: number, top: number, right: number, bottom: number): void {
    this.u32(left);
    this.u32(top);
    this.u32(right);
    this.u32(bottom);
  }

  /** `SIZEL`. */
  size(cx: number, cy: number): void {
    this.u32(cx);
    this.u32(cy);
  }

  get length(): number {
    return this.bytes.length;
  }

  finish(): Uint8Array {
    return new Uint8Array(this.bytes);
  }
}

/** `COLORREF`: blue, green, red, then a reserved byte. Not RGB order. */
function colorRef(red: number, green: number, blue: number): number {
  return (red & 0xff) | ((green & 0xff) << 8) | ((blue & 0xff) << 16);
}

const RECORD_COUNT = 5;
const HANDLE_COUNT = 2;
const TOTAL_BYTES = 168;

/** A five-record EMF that draws one filled rectangle. See the file comment. */
export function probeEmf(): Uint8Array {
  const w = new Writer();

  // --- EMR_HEADER, the 88-byte original form ------------------------------
  w.u32(1); // iType: EMR_HEADER
  w.u32(88); // nSize
  w.rect(0, 0, WIDTH, HEIGHT); // rclBounds, device units
  w.rect(0, 0, WIDTH * 100, HEIGHT * 100); // rclFrame, 0.01 mm units
  w.u32(0x464d4520); // dSignature: ' EMF'
  w.u32(0x00010000); // nVersion
  w.u32(TOTAL_BYTES); // nBytes: the whole file
  w.u32(RECORD_COUNT); // nRecords
  w.u16(HANDLE_COUNT); // nHandles: highest index used, plus one
  w.u16(0); // sReserved
  w.u32(0); // nDescription
  w.u32(0); // offDescription
  w.u32(0); // nPalEntries
  w.size(DEVICE_PIXELS.cx, DEVICE_PIXELS.cy); // szlDevice
  w.size(DEVICE_MM.cx, DEVICE_MM.cy); // szlMillimeters

  // --- EMR_CREATEBRUSHINDIRECT -------------------------------------------
  w.u32(39);
  w.u32(24);
  w.u32(1); // ihBrush: handle 1
  w.u32(0); // lbStyle: BS_SOLID
  w.u32(colorRef(0x00, 0x70, 0xc0));
  w.u32(0); // lbHatch, unused for a solid brush

  // --- EMR_SELECTOBJECT ---------------------------------------------------
  w.u32(37);
  w.u32(12);
  w.u32(1); // the brush just created

  // --- EMR_RECTANGLE ------------------------------------------------------
  w.u32(43);
  w.u32(24);
  w.rect(20, 20, WIDTH - 20, HEIGHT - 20);

  // --- EMR_EOF ------------------------------------------------------------
  w.u32(14);
  w.u32(20);
  w.u32(0); // nPalEntries
  w.u32(16); // offPalEntries, relative to this record
  w.u32(20); // nSizeLast: this record's size again, so a reader can walk back

  const bytes = w.finish();
  if (bytes.length !== TOTAL_BYTES) {
    // `nBytes` in the header is written before the body exists, so it is a
    // constant that has to be checked rather than computed.
    throw new Error(
      `probeEmf wrote ${String(bytes.length)} bytes, header claims ${String(TOTAL_BYTES)}`,
    );
  }
  return bytes;
}
