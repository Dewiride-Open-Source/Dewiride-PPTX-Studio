/**
 * A minimal Compound File Binary writer and reader, for `a32-macros`.
 *
 * ## Why this is here at all
 *
 * `ppt/vbaProject.bin` is the one part of a PowerPoint package that is not XML
 * and not an image: it is an OLE2 compound file, a whole filesystem-in-a-file
 * with its own allocation tables and directory tree. `a32-macros` needs one,
 * and the corpus authors its own bytes rather than committing somebody else's -
 * the same call `png.ts`, `emf.ts`, `xlsx.ts` and `jpeg.ts` all made.
 *
 * The must-not-break rules say **never touch `ppt/embeddings/*.bin`**, and that
 * still holds: this writes a compound file, it does not edit one. Writing one
 * from nothing is safe; rewriting one somebody else wrote is a guaranteed
 * repair, because a compound file's sector chains are position-dependent and
 * nothing in the format tolerates a re-layout.
 *
 * ## The shape of the format
 *
 * A 512-byte header, then sectors of 512 bytes numbered from zero, where sector
 * *N* begins at byte `(N + 1) * 512`. Two allocation tables and a directory:
 *
 * - The **FAT** is one flat array of 4-byte "next sector" links. Following a
 *   chain from a start sector until `ENDOFCHAIN` gives a stream.
 * - The **directory** is a tree of 128-byte entries, entry 0 being the root.
 *   Children of a storage are held in a **red-black tree**, not a list.
 * - Streams smaller than 4096 bytes do not get sectors of their own. They live
 *   in the **mini stream** - one byte array owned by the root entry, itself a
 *   normal sector chain - cut into 64-byte mini sectors and allocated by a
 *   second table, the **MiniFAT**. Every stream in a small VBA project is under
 *   the cutoff, so in practice the whole payload is in the mini stream and the
 *   regular sectors hold nothing but bookkeeping.
 *
 * ## The sibling ordering is the part that bites
 *
 * Directory children are ordered by **name length first**, then by a
 * case-insensitive comparison of the uppercased UTF-16 code units - not by
 * ordinary lexicographic order, and not by anything a `sort()` does by default.
 * `VBA` sorts before `PROJECT` because it is shorter. Readers enforce that
 * ordering while being lenient about the red-black colouring itself, so this
 * writer builds a balanced tree from the correctly sorted list and colours
 * every node black.
 *
 * ## What is verified
 *
 * `cfb.test.ts` writes a tree and reads it back with `readCfb`, which is a
 * separate implementation of the walk rather than the writer run backwards:
 * it follows the FAT and MiniFAT chains and the directory tree the way any
 * reader would. That catches a chain that does not terminate, a mini sector
 * miscount, and a sibling tree in the wrong order - which is most of what can
 * go wrong here and none of which a hash comparison would show.
 */

const SECTOR = 512;
const MINI_SECTOR = 64;
const MINI_CUTOFF = 4096;
const DIRS_PER_SECTOR = SECTOR / 128;
const FAT_PER_SECTOR = SECTOR / 4;

const FREESECT = 0xffffffff;
const ENDOFCHAIN = 0xfffffffe;
const FATSECT = 0xfffffffd;
const NOSTREAM = 0xffffffff;

const SIGNATURE = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1] as const;

export interface CfbStream {
  readonly kind: 'stream';
  readonly name: string;
  readonly bytes: Uint8Array;
}

export interface CfbStorage {
  readonly kind: 'storage';
  readonly name: string;
  readonly children: readonly CfbNode[];
}

export type CfbNode = CfbStream | CfbStorage;

export const stream = (name: string, bytes: Uint8Array): CfbStream => ({
  kind: 'stream',
  name,
  bytes,
});

export const storage = (name: string, children: readonly CfbNode[]): CfbStorage => ({
  kind: 'storage',
  name,
  children,
});

/**
 * The sibling comparison readers actually enforce: shorter names first, then a
 * case-insensitive comparison of UTF-16 code units.
 *
 * `localeCompare` and a plain `<` both get this wrong, in different ways and
 * only for some names, which is the worst kind of wrong.
 */
function compareNames(a: string, b: string): number {
  if (a.length !== b.length) return a.length - b.length;
  const upperA = a.toUpperCase();
  const upperB = b.toUpperCase();
  for (let i = 0; i < upperA.length; i++) {
    const ca = upperA.charCodeAt(i);
    const cb = upperB.charCodeAt(i);
    if (ca !== cb) return ca - cb;
  }
  return 0;
}

interface Entry {
  name: string;
  type: 1 | 2 | 5;
  left: number;
  right: number;
  child: number;
  start: number;
  size: number;
}

/** A balanced binary tree over a sorted list, returning the root's index. */
function buildTree(indices: readonly number[], entries: readonly Entry[]): number {
  if (indices.length === 0) return NOSTREAM;
  const middle = indices.length >> 1;
  const root = indices[middle];
  if (root === undefined) return NOSTREAM;
  const node = entries[root];
  if (node === undefined) return NOSTREAM;
  node.left = buildTree(indices.slice(0, middle), entries);
  node.right = buildTree(indices.slice(middle + 1), entries);
  return root;
}

/** Write a compound file whose root storage holds `children`. */
export function writeCfb(children: readonly CfbNode[]): Uint8Array {
  // --- the directory, depth first, root at index 0 -------------------------
  const entries: Entry[] = [
    {
      name: 'Root Entry',
      type: 5,
      left: NOSTREAM,
      right: NOSTREAM,
      child: NOSTREAM,
      start: 0,
      size: 0,
    },
  ];
  const payloads = new Map<number, Uint8Array>();

  const addAll = (nodes: readonly CfbNode[], parent: number): void => {
    const indices: number[] = [];
    for (const node of [...nodes].sort((a, b) => compareNames(a.name, b.name))) {
      if (node.name.length > 31) {
        throw new Error('a compound file name is at most 31 characters: ' + node.name);
      }
      const index = entries.length;
      entries.push({
        name: node.name,
        type: node.kind === 'storage' ? 1 : 2,
        left: NOSTREAM,
        right: NOSTREAM,
        child: NOSTREAM,
        start: 0,
        size: node.kind === 'stream' ? node.bytes.length : 0,
      });
      indices.push(index);
      if (node.kind === 'stream') payloads.set(index, node.bytes);
    }
    // The children are added before recursing so that a storage's own
    // subtree does not interleave with its siblings' entries. Nothing
    // requires that; it just makes a hex dump readable.
    const parentEntry = entries[parent];
    if (parentEntry !== undefined) parentEntry.child = buildTree(indices, entries);
    let cursor = 0;
    for (const node of [...nodes].sort((a, b) => compareNames(a.name, b.name))) {
      const index = indices[cursor++];
      if (node.kind === 'storage' && index !== undefined) addAll(node.children, index);
    }
  };
  addAll(children, 0);

  // --- the mini stream ------------------------------------------------------
  // Every stream under the 4096-byte cutoff goes here, padded to a 64-byte
  // boundary. A VBA project this size puts everything in it.
  const miniPieces: Uint8Array[] = [];
  let miniSectors = 0;
  const large: { index: number; bytes: Uint8Array }[] = [];
  for (const [index, bytes] of payloads) {
    const entry = entries[index];
    if (entry === undefined) continue;
    if (bytes.length >= MINI_CUTOFF) {
      large.push({ index, bytes });
      continue;
    }
    entry.start = miniSectors;
    const padded = new Uint8Array(Math.ceil(bytes.length / MINI_SECTOR) * MINI_SECTOR);
    padded.set(bytes, 0);
    miniPieces.push(padded);
    miniSectors += padded.length / MINI_SECTOR;
  }
  const miniStream = new Uint8Array(miniSectors * MINI_SECTOR);
  {
    let at = 0;
    for (const piece of miniPieces) {
      miniStream.set(piece, at);
      at += piece.length;
    }
  }

  // --- sector counts --------------------------------------------------------
  const directorySectors = Math.max(1, Math.ceil(entries.length / DIRS_PER_SECTOR));
  const miniFatSectors = miniSectors === 0 ? 0 : Math.ceil(miniSectors / FAT_PER_SECTOR);
  const miniStreamSectors = Math.ceil(miniStream.length / SECTOR);
  const largeSectors = large.reduce((n, item) => n + Math.ceil(item.bytes.length / SECTOR), 0);

  let fatSectors = 1;
  for (;;) {
    const total = fatSectors + directorySectors + miniFatSectors + miniStreamSectors + largeSectors;
    if (total <= fatSectors * FAT_PER_SECTOR) break;
    fatSectors += 1;
  }
  if (fatSectors > 109) throw new Error('this writer does not emit DIFAT sectors');

  // --- sector assignment ----------------------------------------------------
  let next = 0;
  const fatStart = next;
  next += fatSectors;
  const directoryStart = next;
  next += directorySectors;
  const miniFatStart = next;
  next += miniFatSectors;
  const miniStreamStart = next;
  next += miniStreamSectors;
  for (const item of large) {
    const entry = entries[item.index];
    if (entry !== undefined) entry.start = next;
    next += Math.ceil(item.bytes.length / SECTOR);
  }
  const totalSectors = next;

  const root = entries[0];
  if (root === undefined) throw new Error('unreachable: no root entry');
  root.start = miniStreamSectors === 0 ? ENDOFCHAIN : miniStreamStart;
  root.size = miniStream.length;

  // --- the FAT --------------------------------------------------------------
  const fat = new Uint32Array(fatSectors * FAT_PER_SECTOR).fill(FREESECT);
  const chain = (start: number, count: number): void => {
    for (let i = 0; i < count; i++) fat[start + i] = i === count - 1 ? ENDOFCHAIN : start + i + 1;
  };
  for (let i = 0; i < fatSectors; i++) fat[fatStart + i] = FATSECT;
  chain(directoryStart, directorySectors);
  if (miniFatSectors > 0) chain(miniFatStart, miniFatSectors);
  if (miniStreamSectors > 0) chain(miniStreamStart, miniStreamSectors);
  {
    let at = miniStreamStart + miniStreamSectors;
    for (const item of large) {
      const count = Math.ceil(item.bytes.length / SECTOR);
      chain(at, count);
      at += count;
    }
  }

  // --- the MiniFAT ----------------------------------------------------------
  // One chain per small stream, in the order they were laid into the mini
  // stream, so each is a run of consecutive mini sectors.
  const miniFat = new Uint32Array(Math.max(miniFatSectors, 0) * FAT_PER_SECTOR).fill(FREESECT);
  {
    let at = 0;
    for (const piece of miniPieces) {
      const count = piece.length / MINI_SECTOR;
      for (let i = 0; i < count; i++) {
        miniFat[at + i] = i === count - 1 ? ENDOFCHAIN : at + i + 1;
      }
      at += count;
    }
  }

  // --- bytes ----------------------------------------------------------------
  const out = new Uint8Array(SECTOR + totalSectors * SECTOR);
  const view = new DataView(out.buffer);
  const sectorAt = (n: number): number => SECTOR + n * SECTOR;

  out.set(SIGNATURE, 0);
  // 8..23 header CLSID, all zero.
  view.setUint16(24, 0x003e, true); // minor version
  view.setUint16(26, 0x0003, true); // major version: 512-byte sectors
  view.setUint16(28, 0xfffe, true); // byte order mark, little-endian
  view.setUint16(30, 9, true); // sector shift: 2^9
  view.setUint16(32, 6, true); // mini sector shift: 2^6
  // 34..39 reserved, zero.
  view.setUint32(40, 0, true); // directory sector count: MUST be 0 in v3
  view.setUint32(44, fatSectors, true);
  view.setUint32(48, directoryStart, true);
  view.setUint32(52, 0, true); // transaction signature
  view.setUint32(56, MINI_CUTOFF, true);
  view.setUint32(60, miniFatSectors === 0 ? ENDOFCHAIN : miniFatStart, true);
  view.setUint32(64, miniFatSectors, true);
  view.setUint32(68, ENDOFCHAIN, true); // first DIFAT sector: none
  view.setUint32(72, 0, true); // DIFAT sector count
  for (let i = 0; i < 109; i++) {
    view.setUint32(76 + i * 4, i < fatSectors ? fatStart + i : FREESECT, true);
  }

  for (let i = 0; i < fat.length; i++) {
    view.setUint32(sectorAt(fatStart) + i * 4, fat[i] ?? FREESECT, true);
  }
  for (let i = 0; i < miniFat.length; i++) {
    view.setUint32(sectorAt(miniFatStart) + i * 4, miniFat[i] ?? FREESECT, true);
  }
  out.set(miniStream, sectorAt(miniStreamStart));
  {
    let at = miniStreamStart + miniStreamSectors;
    for (const item of large) {
      out.set(item.bytes, sectorAt(at));
      at += Math.ceil(item.bytes.length / SECTOR);
    }
  }

  const directoryBase = sectorAt(directoryStart);
  for (let i = 0; i < directorySectors * DIRS_PER_SECTOR; i++) {
    const at = directoryBase + i * 128;
    const entry = entries[i];
    if (entry === undefined) {
      // An unallocated slot. Its ids still have to be NOSTREAM rather than 0,
      // which is a real directory index.
      view.setUint32(at + 68, NOSTREAM, true);
      view.setUint32(at + 72, NOSTREAM, true);
      view.setUint32(at + 76, NOSTREAM, true);
      continue;
    }
    for (let c = 0; c < entry.name.length; c++) {
      view.setUint16(at + c * 2, entry.name.charCodeAt(c), true);
    }
    // The length is in **bytes** and includes the two-byte terminator, which
    // is the field most often written as a character count.
    view.setUint16(at + 64, (entry.name.length + 1) * 2, true);
    out[at + 66] = entry.type;
    out[at + 67] = 1; // black. Readers enforce the ordering, not the colouring.
    view.setUint32(at + 68, entry.left, true);
    view.setUint32(at + 72, entry.right, true);
    view.setUint32(at + 76, entry.child, true);
    // 80..95 CLSID, 96..99 state bits, 100..115 timestamps: all zero.
    view.setUint32(at + 116, entry.start, true);
    view.setUint32(at + 120, entry.size, true);
    view.setUint32(at + 124, 0, true); // v3: the high half MUST be zero
  }

  return out;
}

/**
 * Read a compound file back into a tree.
 *
 * Deliberately written as a reader rather than as `writeCfb` run backwards: it
 * follows the header's pointers, the FAT and MiniFAT chains and the directory
 * tree, so a writer that put the right bytes in the wrong place fails here.
 */
export function readCfb(bytes: Uint8Array): CfbNode[] {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  for (let i = 0; i < 8; i++) {
    if (bytes[i] !== SIGNATURE[i]) throw new Error('not a compound file');
  }
  if (view.getUint16(26, true) !== 3) throw new Error('only version 3 is supported');

  const sectorAt = (n: number): number => SECTOR + n * SECTOR;
  const fatSectors = view.getUint32(44, true);
  const directoryStart = view.getUint32(48, true);
  const miniFatStart = view.getUint32(60, true);
  const miniFatSectors = view.getUint32(64, true);

  const readTable = (start: number, count: number): Uint32Array => {
    const table = new Uint32Array(count * FAT_PER_SECTOR);
    for (let s = 0; s < count; s++) {
      for (let i = 0; i < FAT_PER_SECTOR; i++) {
        table[s * FAT_PER_SECTOR + i] = view.getUint32(sectorAt(start + s) + i * 4, true);
      }
    }
    return table;
  };
  // The FAT sectors are named by the DIFAT, which for a file this size is
  // entirely inside the header.
  const fat = new Uint32Array(fatSectors * FAT_PER_SECTOR);
  for (let s = 0; s < fatSectors; s++) {
    const sector = view.getUint32(76 + s * 4, true);
    for (let i = 0; i < FAT_PER_SECTOR; i++) {
      fat[s * FAT_PER_SECTOR + i] = view.getUint32(sectorAt(sector) + i * 4, true);
    }
  }
  const miniFat =
    miniFatSectors === 0 ? new Uint32Array(0) : readTable(miniFatStart, miniFatSectors);

  const follow = (table: Uint32Array, start: number, limit: number): number[] => {
    const chain: number[] = [];
    let cursor = start;
    while (cursor !== ENDOFCHAIN && cursor !== FREESECT && chain.length <= limit) {
      chain.push(cursor);
      cursor = table[cursor] ?? ENDOFCHAIN;
    }
    if (chain.length > limit) throw new Error('a sector chain does not terminate');
    return chain;
  };

  const totalSectors = Math.floor((bytes.length - SECTOR) / SECTOR);
  const directoryChain = follow(fat, directoryStart, totalSectors);
  const directory: Uint8Array[] = [];
  for (const sector of directoryChain) {
    for (let i = 0; i < DIRS_PER_SECTOR; i++) {
      directory.push(bytes.subarray(sectorAt(sector) + i * 128, sectorAt(sector) + (i + 1) * 128));
    }
  }

  const entryAt = (index: number): DataView => {
    const raw = directory[index];
    if (raw === undefined) throw new Error('directory entry ' + String(index) + ' is missing');
    return new DataView(raw.buffer, raw.byteOffset, raw.byteLength);
  };

  const rootView = entryAt(0);
  const miniStreamChain = follow(fat, rootView.getUint32(116, true), totalSectors);
  const miniStream = new Uint8Array(miniStreamChain.length * SECTOR);
  miniStreamChain.forEach((sector, i) => {
    miniStream.set(bytes.subarray(sectorAt(sector), sectorAt(sector) + SECTOR), i * SECTOR);
  });

  const readStream = (start: number, size: number): Uint8Array => {
    if (size === 0) return new Uint8Array(0);
    if (size < MINI_CUTOFF) {
      const chain = follow(miniFat, start, miniStream.length / MINI_SECTOR);
      const out = new Uint8Array(chain.length * MINI_SECTOR);
      chain.forEach((mini, i) => {
        out.set(miniStream.subarray(mini * MINI_SECTOR, (mini + 1) * MINI_SECTOR), i * MINI_SECTOR);
      });
      return out.subarray(0, size);
    }
    const chain = follow(fat, start, totalSectors);
    const out = new Uint8Array(chain.length * SECTOR);
    chain.forEach((sector, i) => {
      out.set(bytes.subarray(sectorAt(sector), sectorAt(sector) + SECTOR), i * SECTOR);
    });
    return out.subarray(0, size);
  };

  const walk = (index: number): CfbNode[] => {
    if (index === NOSTREAM) return [];
    const entry = entryAt(index);
    const nameLength = entry.getUint16(64, true);
    let name = '';
    for (let i = 0; i < nameLength / 2 - 1; i++)
      name += String.fromCharCode(entry.getUint16(i * 2, true));
    const type = entry.getUint8(66);
    const self: CfbNode =
      type === 1
        ? storage(name, walk(entry.getUint32(76, true)))
        : stream(name, readStream(entry.getUint32(116, true), entry.getUint32(120, true)));
    // In-order, so the result comes back in the sibling order the file states
    // rather than in the order the writer happened to add them.
    return [...walk(entry.getUint32(68, true)), self, ...walk(entry.getUint32(72, true))];
  };

  return walk(rootView.getUint32(76, true));
}
