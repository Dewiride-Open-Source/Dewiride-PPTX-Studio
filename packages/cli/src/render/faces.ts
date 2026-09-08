/**
 * Which file on this machine is the typeface the deck asked for.
 *
 * The browser answers this with a font stack and never tells us what it chose;
 * here the choice is ours, so it is also reportable - `renderDeck` can say
 * "Aptos was drawn in Carlito" because this module knows it was. The
 * substitution table itself is `@pptx-studio/text`'s, so the two renderers
 * cannot fall back differently. ADR 0042.
 */

import { readdirSync, readFileSync, statSync } from 'node:fs';
import { homedir } from 'node:os';
import { join } from 'node:path';

import { LAST_RESORT_FAMILIES, substituteFor } from '@pptx-studio/text';

import { RenderError } from './errors.js';
import { facesIn, type Face } from './sfnt.js';

const FONT_FILE = /\.(ttf|ttc|otf|otc)$/i;

/** Deep enough for `/usr/share/fonts/truetype/dejavu`, and no deeper. */
const MAX_DEPTH = 4;

/** Case and whitespace are not part of a typeface's identity for a lookup. */
function key(family: string): string {
  return family.trim().replace(/\s+/g, ' ').toLowerCase();
}

/** Where fonts live on this platform, whether or not the directories exist. */
export function systemFontDirectories(platform: string = process.platform): readonly string[] {
  const home = homedir();
  if (platform === 'win32') {
    const root = process.env['SystemRoot'] ?? 'C:\\Windows';
    const local = process.env['LOCALAPPDATA'];
    return [
      join(root, 'Fonts'),
      ...(local === undefined ? [] : [join(local, 'Microsoft', 'Windows', 'Fonts')]),
    ];
  }
  if (platform === 'darwin') {
    return ['/System/Library/Fonts', '/Library/Fonts', join(home, 'Library', 'Fonts')];
  }
  return [
    '/usr/share/fonts',
    '/usr/local/share/fonts',
    join(home, '.local', 'share', 'fonts'),
    join(home, '.fonts'),
  ];
}

function filesUnder(directory: string, depth: number, out: string[]): void {
  if (depth > MAX_DEPTH) return;
  let entries;
  try {
    entries = readdirSync(directory, { withFileTypes: true });
  } catch {
    // A directory that is not there, or not readable, contributes no fonts.
    // Only a directory the caller named explicitly is an error, and `indexFonts`
    // checks those before it gets here.
    return;
  }
  for (const entry of entries) {
    const path = join(directory, entry.name);
    if (entry.isDirectory()) filesUnder(path, depth + 1, out);
    else if (FONT_FILE.test(entry.name)) out.push(path);
  }
}

export interface IndexedFace {
  readonly face: Face;
  readonly file: string;
}

/** What the library found for a typeface the deck named. */
export interface Resolved {
  readonly face: Face;
  /** The typeface the deck asked for. */
  readonly asked: string;
  /** The family actually drawn in, which differs when a substitute was used. */
  readonly drawn: string;
  readonly file: string;
  readonly substituted: boolean;
}

export interface FontLibrary {
  /** Every face found, for the diagnostics and for a test to count. */
  readonly indexed: readonly IndexedFace[];
  /** The directories that were scanned, in order. */
  readonly directories: readonly string[];
  resolve(family: string, bold: boolean, italic: boolean): Resolved | undefined;
}

export interface IndexOptions {
  /** Directories to scan before the system ones. Each must exist. */
  readonly extra?: readonly string[];
  /** Whether to scan this platform's own font directories. Default true. */
  readonly system?: boolean;
  readonly platform?: string;
}

/**
 * The style slot a face fills.
 *
 * Two bits rather than a weight axis, because DrawingML has no weight axis:
 * `a:rPr/@b` is a boolean and "Roboto Light" is a typeface name, not a weight.
 */
function slot(bold: boolean, italic: boolean): number {
  return (bold ? 1 : 0) | (italic ? 2 : 0);
}

/**
 * Read every font file under the given directories, once.
 *
 * A file that will not parse is skipped rather than fatal: a font directory on
 * a real machine holds `.ttf` files that are bitmap-only, damaged, or not fonts
 * at all, and one of them must not stop a deck from rendering.
 */
export function indexFonts(options: IndexOptions = {}): FontLibrary {
  for (const directory of options.extra ?? []) {
    try {
      if (!statSync(directory).isDirectory()) {
        throw new RenderError('CLI_FONT_DIR', `${directory} is not a directory`, directory);
      }
    } catch (error) {
      if (error instanceof RenderError) throw error;
      throw new RenderError('CLI_FONT_DIR', `no such directory: ${directory}`, directory);
    }
  }

  const directories = [
    ...(options.extra ?? []),
    ...(options.system === false ? [] : systemFontDirectories(options.platform)),
  ];

  const files: string[] = [];
  for (const directory of directories) filesUnder(directory, 0, files);

  const indexed: IndexedFace[] = [];
  const byFamily = new Map<string, (IndexedFace | undefined)[]>();
  const claim = (family: string, entry: IndexedFace): void => {
    const slots = byFamily.get(key(family)) ?? [undefined, undefined, undefined, undefined];
    const at = slot(entry.face.bold, entry.face.italic);
    // First file wins, so an earlier `--font-dir` beats a system face of the
    // same name and the caller can override without uninstalling anything.
    slots[at] ??= entry;
    byFamily.set(key(family), slots);
  };

  for (const file of files) {
    let faces: readonly Face[];
    try {
      faces = facesIn(new Uint8Array(readFileSync(file)), file);
    } catch {
      continue;
    }
    for (const face of faces) {
      const entry: IndexedFace = { face, file };
      indexed.push(entry);
      claim(face.family, entry);
      // Name ID 16 is what a large family is known by once it outgrows the four
      // style slots, and it is the name a deck spells - `Segoe UI` for the face
      // whose ID 1 is `Segoe UI Semibold`.
      if (face.typographicFamily !== undefined) claim(face.typographicFamily, entry);
    }
  }

  /** The nearest slot to the one asked for: exact, then drop italic, then bold. */
  const pick = (family: string, bold: boolean, italic: boolean): IndexedFace | undefined => {
    const slots = byFamily.get(key(family));
    if (slots === undefined) return undefined;
    const wanted = [
      slot(bold, italic),
      slot(bold, false),
      slot(false, italic),
      slot(false, false),
      0,
      1,
      2,
      3,
    ];
    for (const at of wanted) {
      const found = slots[at];
      if (found !== undefined) return found;
    }
    return undefined;
  };

  return {
    indexed,
    directories,
    resolve(family: string, bold: boolean, italic: boolean): Resolved | undefined {
      const direct = pick(family, bold, italic);
      if (direct !== undefined) {
        return {
          face: direct.face,
          asked: family,
          drawn: family,
          file: direct.file,
          substituted: false,
        };
      }
      const chain = [substituteFor(family)?.use, ...LAST_RESORT_FAMILIES].filter(
        (name): name is string => name !== undefined,
      );
      for (const name of chain) {
        const found = pick(name, bold, italic);
        if (found !== undefined) {
          return {
            face: found.face,
            asked: family,
            drawn: name,
            file: found.file,
            substituted: true,
          };
        }
      }
      return undefined;
    },
  };
}
