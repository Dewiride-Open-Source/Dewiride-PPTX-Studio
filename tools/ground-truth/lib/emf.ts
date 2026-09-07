/**
 * Enough of the EMF container to read the text PowerPoint drew, with its face,
 * size and colour.
 *
 * The records that carry a face, an origin, a transform or an advance;
 * everything else is skipped by `nSize`. ADRs 0031 and 0034.
 */

/** The record types this reader understands. Everything else is skipped. */
const EMR_HEADER = 1;
const EMR_SETTEXTALIGN = 22;
const EMR_SAVEDC = 33;
const EMR_RESTOREDC = 34;
const EMR_SETTEXTCOLOR = 24;
const EMR_SETWORLDTRANSFORM = 35;
const EMR_MODIFYWORLDTRANSFORM = 36;
const EMR_SELECTOBJECT = 37;
const EMR_DELETEOBJECT = 40;
const EMR_EXTCREATEFONTINDIRECTW = 82;
const EMR_EXTTEXTOUTA = 83;
const EMR_EXTTEXTOUTW = 84;

/** `LOGFONTW`, as much of it as says which face at which size. */
export interface EmfFont {
  /** `lfFaceName`, up to 32 UTF-16 units, NUL-terminated. */
  readonly face: string;
  /** `lfHeight` in logical units; negative is GDI for the em size, so the sign is kept. */
  readonly height: number;
  readonly weight: number;
  readonly italic: boolean;
  /** `lfUnderline`: GDI drew the underline itself, from the face's own metrics. */
  readonly underline: boolean;
  /** `lfStrikeOut`, likewise. */
  readonly strikeOut: boolean;
  readonly charSet: number;
  /** `lfEscapement`, tenths of a degree counter-clockwise: the baseline angle. */
  readonly escapement: number;
  /** `lfOrientation`, tenths of a degree: the angle of each character. */
  readonly orientation: number;
  /** `lfFaceName` began with `@`: GDI names the vertical variant of a face that way. */
  readonly verticalFace: boolean;
}

/**
 * An `XFORM`: `[m11, m12, m21, m22, dx, dy]`, logical units to page units.
 *
 * A negative determinant is a mirror, which is what question `flip` turns on.
 */
export type EmfTransform = readonly [number, number, number, number, number, number];

/** The identity, which is what a page with no `EMR_SETWORLDTRANSFORM` is in. */
export const IDENTITY_TRANSFORM: EmfTransform = [1, 0, 0, 1, 0, 0];

/** `b` applied after `a`, in GDI's row-vector convention. */
export function composeTransform(a: EmfTransform, b: EmfTransform): EmfTransform {
  return [
    a[0] * b[0] + a[1] * b[2],
    a[0] * b[1] + a[1] * b[3],
    a[2] * b[0] + a[3] * b[2],
    a[2] * b[1] + a[3] * b[3],
    a[4] * b[0] + a[5] * b[2] + b[4],
    a[4] * b[1] + a[5] * b[3] + b[5],
  ];
}

/** A point through an `XFORM`. */
export function applyTransform(t: EmfTransform, x: number, y: number): readonly [number, number] {
  return [t[0] * x + t[2] * y + t[4], t[1] * x + t[3] * y + t[5]];
}

/** One `ExtTextOutW` call: the characters, the face, the colour. */
export interface EmfText {
  /** The string PowerPoint drew, exactly. */
  readonly text: string;
  /** The face selected when it was drawn, or `null` if none had been. */
  readonly font: EmfFont | null;
  /** `EMR_SETTEXTCOLOR`, as `0x00bbggrr`, or `null` if never set. */
  readonly colorBgr: number | null;
  /** `rclBounds`, the record's own device-space box: left, top, right, bottom. */
  readonly bounds: readonly [number, number, number, number];
  /** `ptlReference`, in logical units, before the world transform. */
  readonly reference: readonly [number, number];
  /** The world transform in force, which is where a flip or a rotation shows. */
  readonly transform: EmfTransform;
  /** `ptlReference` through `transform`: the origin in page units. */
  readonly origin: readonly [number, number];
  /** The per-character advances at `offDx`, empty when the record carries none. */
  readonly dx: readonly number[];
  /** `EMR_SETTEXTALIGN`'s `iMode`, which says what `reference` is a corner of. */
  readonly textAlign: number;
  /** Position in the record stream, so two identical strings stay distinct. */
  readonly index: number;
  /** `ETO_GLYPH_INDEX`: `text` holds glyph ids, not characters - a shaped run. */
  readonly glyphIndices: boolean;
  /** `fOptions`, whole, so a flag nothing here reads yet is still in the fixture. */
  readonly options: number;
}

export interface EmfReading {
  /** `rclFrame` from the header, in 0.01mm units. */
  readonly frame: readonly [number, number, number, number];
  readonly records: number;
  readonly texts: readonly EmfText[];
}

class EmfError extends Error {
  constructor(message: string) {
    super(message);
    this.name = 'EmfError';
  }
}

function utf16(bytes: Uint8Array, view: DataView, at: number, max: number): string {
  let out = '';
  for (let i = 0; i < max; i++) {
    const offset = at + i * 2;
    if (offset + 2 > bytes.length) break;
    const unit = view.getUint16(offset, true);
    if (unit === 0) break;
    out += String.fromCharCode(unit);
  }
  return out;
}

/**
 * Read every text record in an EMF, in draw order.
 *
 * The object table is tracked because `EMR_EXTTEXTOUTW` names no font: the face is
 * the last one a `EMR_SELECTOBJECT` chose. Stock handles select nothing.
 */
export function readEmf(bytes: Uint8Array): EmfReading {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  if (bytes.length < 88)
    throw new EmfError(`EMF is ${String(bytes.length)} bytes, too short for a header`);
  if (view.getUint32(0, true) !== EMR_HEADER)
    throw new EmfError('EMF does not begin with EMR_HEADER');
  const frame: [number, number, number, number] = [
    view.getInt32(24, true),
    view.getInt32(28, true),
    view.getInt32(32, true),
    view.getInt32(36, true),
  ];

  const fonts = new Map<number, EmfFont>();
  let current: EmfFont | null = null;
  let colorBgr: number | null = null;
  let world: EmfTransform = IDENTITY_TRANSFORM;
  let textAlign = 0;
  const texts: EmfText[] = [];
  // Without the DC stack the world transform accumulates across drawing calls,
  // and every text record after the first reports an origin twice too far.
  const saved: { world: EmfTransform; font: EmfFont | null; colorBgr: number | null }[] = [];

  const xform = (at: number): EmfTransform => [
    view.getFloat32(at, true),
    view.getFloat32(at + 4, true),
    view.getFloat32(at + 8, true),
    view.getFloat32(at + 12, true),
    view.getFloat32(at + 16, true),
    view.getFloat32(at + 20, true),
  ];

  let offset = 0;
  let records = 0;
  while (offset + 8 <= bytes.length) {
    const type = view.getUint32(offset, true);
    const size = view.getUint32(offset + 4, true);
    if (size < 8 || size % 4 !== 0 || offset + size > bytes.length) {
      throw new EmfError(
        `record ${String(records)} at ${String(offset)} has nSize ${String(size)}, which does not fit`,
      );
    }

    if (type === EMR_EXTCREATEFONTINDIRECTW) {
      // ihFont, then LOGFONTW: 28 bytes of scalars, then lfFaceName[32].
      const handle = view.getUint32(offset + 8, true);
      // LOGFONTW: lfHeight(0) lfWidth(4) lfEscapement(8) lfOrientation(12)
      // lfWeight(16) lfItalic(20) lfUnderline(21) lfStrikeOut(22) lfCharSet(23)
      // ... lfFaceName(28), 32 UTF-16 units.
      const logfont = offset + 12;
      const face = utf16(bytes, view, logfont + 28, 32);
      fonts.set(handle, {
        face,
        height: view.getInt32(logfont, true),
        weight: view.getInt32(logfont + 16, true),
        italic: bytes[logfont + 20] !== 0,
        underline: bytes[logfont + 21] !== 0,
        strikeOut: bytes[logfont + 22] !== 0,
        charSet: bytes[logfont + 23] ?? 0,
        escapement: view.getInt32(logfont + 8, true),
        orientation: view.getInt32(logfont + 12, true),
        verticalFace: face.startsWith('@'),
      });
    } else if (type === EMR_SELECTOBJECT) {
      const handle = view.getUint32(offset + 8, true);
      const font = fonts.get(handle);
      if (font !== undefined) current = font;
    } else if (type === EMR_DELETEOBJECT) {
      fonts.delete(view.getUint32(offset + 8, true));
    } else if (type === EMR_SETTEXTCOLOR) {
      colorBgr = view.getUint32(offset + 8, true) & 0xffffff;
    } else if (type === EMR_SAVEDC) {
      saved.push({ world, font: current, colorBgr });
    } else if (type === EMR_RESTOREDC) {
      // `iRelative` is negative and counts back from the top of the stack.
      const relative = view.getInt32(offset + 8, true);
      const depth = relative < 0 ? -relative : saved.length - relative;
      for (let i = 1; i < depth && saved.length > 1; i++) saved.pop();
      const state = saved.pop();
      if (state !== undefined) {
        world = state.world;
        current = state.font;
        colorBgr = state.colorBgr;
      }
    } else if (type === EMR_SETTEXTALIGN) {
      textAlign = view.getUint32(offset + 8, true);
    } else if (type === EMR_SETWORLDTRANSFORM) {
      world = xform(offset + 8);
    } else if (type === EMR_MODIFYWORLDTRANSFORM) {
      const next = xform(offset + 8);
      const mode = view.getUint32(offset + 32, true);
      if (mode === 1) world = IDENTITY_TRANSFORM;
      else if (mode === 2) world = composeTransform(next, world);
      else if (mode === 3) world = composeTransform(world, next);
      else world = next;
    } else if (type === EMR_EXTTEXTOUTW || type === EMR_EXTTEXTOUTA) {
      // rclBounds(16) iGraphicsMode(4) exScale(4) eyScale(4), then EMRTEXT:
      // ptlReference(8) nChars(4) offString(4) fOptions(4) rcl(16) offDx(4).
      const chars = view.getUint32(offset + 44, true);
      const stringAt = offset + view.getUint32(offset + 48, true);
      let text = '';
      if (type === EMR_EXTTEXTOUTW) {
        text = utf16(bytes, view, stringAt, chars);
      } else {
        for (let i = 0; i < chars && stringAt + i < bytes.length; i++) {
          text += String.fromCharCode(bytes[stringAt + i] ?? 0);
        }
      }
      const options = view.getUint32(offset + 52, true);
      // ETO_PDY writes an (x, y) pair per character; only the advance is read.
      const stride = (options & 0x2000) === 0 ? 1 : 2;
      const dxAt = view.getUint32(offset + 72, true);
      const dx: number[] = [];
      if (dxAt !== 0) {
        for (let i = 0; i < chars; i++) {
          const at = offset + dxAt + i * stride * 4;
          if (at + 4 > offset + size) break;
          dx.push(view.getInt32(at, true));
        }
      }
      const reference: [number, number] = [
        view.getInt32(offset + 36, true),
        view.getInt32(offset + 40, true),
      ];
      texts.push({
        text,
        glyphIndices: (options & 0x0010) !== 0,
        options,
        font: current,
        colorBgr,
        bounds: [
          view.getInt32(offset + 8, true),
          view.getInt32(offset + 12, true),
          view.getInt32(offset + 16, true),
          view.getInt32(offset + 20, true),
        ],
        reference,
        transform: world,
        origin: applyTransform(world, reference[0], reference[1]),
        dx,
        textAlign,
        index: records,
      });
    }

    offset += size;
    records++;
    if (type === 14) break; // EMR_EOF
  }

  return { frame, records, texts };
}
