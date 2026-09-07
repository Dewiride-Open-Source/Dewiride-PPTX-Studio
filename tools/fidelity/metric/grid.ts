/**
 * The committed oracle, as bytes.
 *
 * The reference is a file and files do not have versions that surprise you.
 * What is committed is the reduced grid rather than PowerPoint's PNG, for three
 * reasons: it is a quarter of the size, it needs no image decoder in CI, and a
 * PNG round-trips through a compressor whose output is not part of any
 * specification we can hold anyone to. ADR 0035.
 */

import { FidelityError } from '../errors.ts';

import type { Grid } from './reduce.ts';

/** `PXFG`, so a truncated or unrelated file fails at byte 0 rather than later. */
const MAGIC = 0x50584647;

const VERSION = 1;

/** Magic, version, cell, width, height. */
const HEADER_BYTES = 10;

export function chromaDims(width: number, height: number): { width: number; height: number } {
  return { width: Math.ceil(width / 2), height: Math.ceil(height / 2) };
}

export function encodeGrid(grid: Grid): Uint8Array {
  const chroma = chromaDims(grid.width, grid.height);
  const planes = grid.width * grid.height + 2 * chroma.width * chroma.height;
  const out = new Uint8Array(HEADER_BYTES + planes);
  const view = new DataView(out.buffer);
  view.setUint32(0, MAGIC, false);
  out[4] = VERSION;
  out[5] = grid.cell;
  view.setUint16(6, grid.width, true);
  view.setUint16(8, grid.height, true);

  let at = HEADER_BYTES;
  for (const value of grid.luma) out[at++] = value;
  for (const value of grid.cb) out[at++] = value;
  for (const value of grid.cr) out[at++] = value;
  return out;
}

/** `PXFS`: one deck's slides in one file, so the manifest has one entry per deck. */
const SET_MAGIC = 0x50584653;

/**
 * Every grid of one deck, keyed by slide.
 *
 * A file per slide would be 149 files and 149 manifest entries for what is one
 * measurement per deck; a single file for the whole corpus would be 1.8 MiB and
 * over the corpus's own per-file cap. Per deck is the grain the corpus is
 * already organised in.
 */
export function encodeGridSet(grids: readonly { key: string; grid: Grid }[]): Uint8Array {
  const parts: Uint8Array[] = [];
  const header = new Uint8Array(6);
  new DataView(header.buffer).setUint32(0, SET_MAGIC, false);
  header[4] = VERSION;
  header[5] = grids.length;
  parts.push(header);

  for (const { key, grid } of grids) {
    const name = new TextEncoder().encode(key);
    if (name.length > 255) throw new FidelityError('FID_GRID_MALFORMED', `${key} is too long`, key);
    const body = encodeGrid(grid);
    const frame = new Uint8Array(5 + name.length);
    frame[0] = name.length;
    frame.set(name, 1);
    new DataView(frame.buffer).setUint32(1 + name.length, body.length, true);
    parts.push(frame, body);
  }

  const total = parts.reduce((n, part) => n + part.length, 0);
  const out = new Uint8Array(total);
  let at = 0;
  for (const part of parts) {
    out.set(part, at);
    at += part.length;
  }
  return out;
}

export function decodeGridSet(bytes: Uint8Array, subject: string): ReadonlyMap<string, Grid> {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  if (bytes.length < 6 || view.getUint32(0, false) !== SET_MAGIC) {
    throw new FidelityError('FID_GRID_MALFORMED', `${subject} does not start with PXFS`, subject);
  }
  if ((bytes[4] ?? 0) !== VERSION) {
    throw new FidelityError(
      'FID_GRID_MALFORMED',
      `${subject} is not version ${String(VERSION)}`,
      subject,
    );
  }
  const count = bytes[5] ?? 0;
  const out = new Map<string, Grid>();
  let at = 6;
  for (let i = 0; i < count; i++) {
    const nameLength = bytes[at] ?? 0;
    const key = new TextDecoder().decode(bytes.subarray(at + 1, at + 1 + nameLength));
    const length = view.getUint32(at + 1 + nameLength, true);
    const from = at + 5 + nameLength;
    out.set(key, decodeGrid(bytes.subarray(from, from + length), `${subject}:${key}`));
    at = from + length;
  }
  if (at !== bytes.length) {
    throw new FidelityError(
      'FID_GRID_MALFORMED',
      `${subject} has ${String(bytes.length - at)} trailing bytes`,
      subject,
    );
  }
  return out;
}

export function decodeGrid(bytes: Uint8Array, subject: string): Grid {
  if (bytes.length < HEADER_BYTES) {
    throw new FidelityError(
      'FID_GRID_MALFORMED',
      `${subject} is ${String(bytes.length)} bytes`,
      subject,
    );
  }
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  if (view.getUint32(0, false) !== MAGIC) {
    throw new FidelityError('FID_GRID_MALFORMED', `${subject} does not start with PXFG`, subject);
  }
  const version = bytes[4] ?? 0;
  if (version !== VERSION) {
    throw new FidelityError(
      'FID_GRID_MALFORMED',
      `${subject} is version ${String(version)}, not ${String(VERSION)}`,
      subject,
    );
  }
  const cell = bytes[5] ?? 0;
  const width = view.getUint16(6, true);
  const height = view.getUint16(8, true);
  const chroma = chromaDims(width, height);
  const lumaCount = width * height;
  const chromaCount = chroma.width * chroma.height;
  const expected = HEADER_BYTES + lumaCount + 2 * chromaCount;
  if (bytes.length !== expected) {
    throw new FidelityError(
      'FID_GRID_MALFORMED',
      `${subject} is ${String(bytes.length)} bytes, not the ${String(expected)} ` +
        `a ${String(width)}x${String(height)} grid needs`,
      subject,
    );
  }

  const at = HEADER_BYTES;
  return {
    width,
    height,
    cell,
    luma: [...bytes.subarray(at, at + lumaCount)],
    chromaWidth: chroma.width,
    chromaHeight: chroma.height,
    cb: [...bytes.subarray(at + lumaCount, at + lumaCount + chromaCount)],
    cr: [...bytes.subarray(at + lumaCount + chromaCount, expected)],
  };
}
