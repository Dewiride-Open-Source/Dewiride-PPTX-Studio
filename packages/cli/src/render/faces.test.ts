import { mkdtempSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { describe, expect, it } from 'vitest';

import { buildFont } from '../../../../tools/ground-truth/lib/truetype.ts';
import { indexFonts, type FontLibrary } from './faces.js';

/**
 * Which face draws a typeface the machine does not have.
 *
 * Every library here is built from fonts written to a temporary directory and
 * indexed with `system: false`, so the answers do not depend on what happens to
 * be installed on the machine running the suite.
 */
function libraryOf(files: Readonly<Record<string, string>>): FontLibrary {
  const dir = mkdtempSync(join(tmpdir(), 'pptx-studio-faces-'));
  for (const [file, family] of Object.entries(files)) {
    writeFileSync(join(dir, file), buildFont({ familyName: family }).bytes);
  }
  return indexFonts({ extra: [dir], system: false });
}

/**
 * The same font with its `name` ID 3 rewritten as an ID 16 of zero length.
 *
 * A record header is 12 bytes of fixed width, so only the two fields change and
 * nothing in the table moves - which is what makes a blank ID 16 buildable at
 * all, `truetype.ts` writing no ID 16 of its own.
 */
function withBlankTypographicFamily(bytes: Uint8Array): Uint8Array {
  const out = new Uint8Array(bytes);
  const view = new DataView(out.buffer);
  let table = -1;
  for (let i = 0; i < view.getUint16(4); i++) {
    const at = 12 + i * 16;
    const tag = String.fromCharCode(
      out[at] ?? 0,
      out[at + 1] ?? 0,
      out[at + 2] ?? 0,
      out[at + 3] ?? 0,
    );
    if (tag === 'name') table = view.getUint32(at + 8);
  }
  if (table < 0) throw new Error('the probe font has no name table');
  for (let i = 0; i < view.getUint16(table + 2); i++) {
    const at = table + 6 + i * 12;
    if (view.getUint16(at + 6) !== 3) continue;
    view.setUint16(at + 6, 16);
    view.setUint16(at + 8, 0);
    return out;
  }
  throw new Error('the probe font has no name ID 3 to rewrite');
}

describe('what the library resolves a typeface to', () => {
  it('draws a family it holds as itself', () => {
    const found = libraryOf({ 'a.ttf': 'Roboto' }).resolve('Roboto', false, false);
    expect(found?.drawn).toBe('Roboto');
    expect(found?.substituted).toBe(false);
  });

  it('takes the substitution table before anything else', () => {
    // `Aptos` -> `Carlito` is the table's; `Aaa Filler` sorts first and must not win.
    const found = libraryOf({ 'a.ttf': 'Aaa Filler', 'b.ttf': 'Carlito' }).resolve(
      'Aptos',
      false,
      false,
    );
    expect(found?.drawn).toBe('Carlito');
    expect(found?.asked).toBe('Aptos');
    expect(found?.substituted).toBe(true);
  });

  it('takes what PowerPoint falls back to before the asked-for name shortened', () => {
    // ADR 0033 measured the last resort as Calibri, 22 of 22. `Roboto` is only
    // nearer by name, and a guess must never displace a measurement.
    const found = libraryOf({ 'a.ttf': 'Roboto', 'b.ttf': 'Carlito' }).resolve(
      'Roboto Light',
      false,
      false,
    );
    expect(found?.drawn).toBe('Carlito');
  });

  it('falls to the family the asked-for name is a variation of', () => {
    // `Aaa Filler` sorts first, so only the shortened name can produce `Roboto`.
    const found = libraryOf({ 'a.ttf': 'Aaa Filler', 'b.ttf': 'Roboto' }).resolve(
      'Roboto Light',
      false,
      false,
    );
    expect(found?.drawn).toBe('Roboto');
    expect(found?.asked).toBe('Roboto Light');
    expect(found?.substituted).toBe(true);
  });

  it('shortens a name one word at a time, longest first', () => {
    // `Aaa Filler` is what the last resort would answer, so a rule that gives up
    // after dropping one word answers it here rather than `Segoe`.
    const library = libraryOf({ 'a.ttf': 'Segoe', 'b.ttf': 'Segoe UI', 'c.ttf': 'Aaa Filler' });
    expect(library.resolve('Segoe UI Semilight', false, false)?.drawn).toBe('Segoe UI');
    expect(library.resolve('Segoe Nothing Like It', false, false)?.drawn).toBe('Segoe');
  });
});

describe('the last resort, where no named face is on the machine', () => {
  const library = libraryOf({ 'a.ttf': 'Zzz Probe', 'b.ttf': 'Aaa Probe' });

  it('draws in a face the library holds rather than refusing', () => {
    const found = library.resolve('Nothing Like It At All', false, false);
    expect(found).toBeDefined();
    expect(found?.face.family).toBe('Aaa Probe');
  });

  it('reports the typeface the deck asked for, and that it was substituted', () => {
    const found = library.resolve('Nothing Like It At All', false, false);
    expect(found?.asked).toBe('Nothing Like It At All');
    expect(found?.drawn).toBe('Aaa Probe');
    expect(found?.substituted).toBe(true);
  });

  it('names the file it drew from, so the report can be checked', () => {
    expect(library.resolve('Nothing Like It At All', false, false)?.file).toMatch(/b\.ttf$/);
  });

  it('chooses by family name, so the order the files were read in cannot matter', () => {
    // The two directories hold the same faces under swapped file names. Under
    // "the first face indexed wins" they answer differently; they must not.
    const forward = libraryOf({ 'a.ttf': 'Zzz Probe', 'b.ttf': 'Aaa Probe' });
    const backward = libraryOf({ 'a.ttf': 'Aaa Probe', 'b.ttf': 'Zzz Probe' });
    const asked = 'Nothing Like It At All';
    expect(forward.resolve(asked, false, false)?.drawn).toBe('Aaa Probe');
    expect(backward.resolve(asked, false, false)?.drawn).toBe(
      forward.resolve(asked, false, false)?.drawn,
    );
  });

  it('is the same answer for every typeface nothing stands in for', () => {
    const one = library.resolve('Zzz Absent One', false, false);
    const two = library.resolve('Zzz Absent Two', false, false);
    expect(one?.drawn).toBe(two?.drawn);
    expect(one?.file).toBe(two?.file);
  });

  it('reads a directory in name order rather than the order it is listed in', () => {
    // Two files claim one family and only one can fill the slot. NTFS lists
    // `a.ttf` first, case-insensitively, and ext4 lists by a hash of the name;
    // neither may decide which face a deck is drawn in.
    const library = libraryOf({ 'B.ttf': 'One Family', 'a.ttf': 'One Family' });
    expect(library.resolve('One Family', false, false)?.file).toMatch(/B\.ttf$/);
  });

  it('is never a face whose typographic family is blank, which names nothing', () => {
    const dir = mkdtempSync(join(tmpdir(), 'pptx-studio-blank-'));
    const font = buildFont({ familyName: 'Zzz Only Face' }).bytes;
    writeFileSync(join(dir, 'a.ttf'), withBlankTypographicFamily(font));
    const library = indexFonts({ extra: [dir], system: false });
    expect(library.indexed).toHaveLength(1);
    // A blank name sorts before every real one, so it would win the moment it
    // was indexed, and the report would name the face the deck was drawn in ''.
    expect(library.resolve('Nothing Like It At All', false, false)?.drawn).toBe('Zzz Only Face');
  });

  it('reports nothing only when the library is empty', () => {
    const empty = indexFonts({
      extra: [mkdtempSync(join(tmpdir(), 'pptx-studio-none-'))],
      system: false,
    });
    expect(empty.indexed).toHaveLength(0);
    expect(empty.resolve('Roboto', false, false)).toBeUndefined();
  });
});
