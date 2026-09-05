/**
 * Enough of the EMF container to read the text PowerPoint drew, with its face,
 * size and colour.
 *
 * Six record types; everything else is skipped by `nSize`. Reasoning in
 * docs/adr/phase-3-text/0031-bullets-fields-and-script-runs.md.
 */

/** The record types this reader understands. Everything else is skipped. */
const EMR_HEADER = 1;
const EMR_SETTEXTCOLOR = 24;
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
  readonly charSet: number;
  /** `lfEscapement`, tenths of a degree counter-clockwise: the baseline angle. */
  readonly escapement: number;
  /** `lfOrientation`, tenths of a degree: the angle of each character. */
  readonly orientation: number;
  /** `lfFaceName` began with `@`: GDI names the vertical variant of a face that way. */
  readonly verticalFace: boolean;
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
  const texts: EmfText[] = [];

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
        reference: [view.getInt32(offset + 36, true), view.getInt32(offset + 40, true)],
        index: records,
      });
    }

    offset += size;
    records++;
    if (type === 14) break; // EMR_EOF
  }

  return { frame, records, texts };
}
