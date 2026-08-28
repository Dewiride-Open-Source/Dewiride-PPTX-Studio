import { describe, expect, it } from 'vitest';
import { readCfb, storage, stream, writeCfb, type CfbNode } from './cfb.ts';
import { compressOvba, decompressOvba, vbaProjectNodes } from './vba.ts';

/**
 * The two things `a32-macros` needs that no other test can see.
 *
 * A `.pptm`'s `vbaProject.bin` is a compound file holding compressed streams,
 * and neither layer shows up anywhere else in this repository's tests: the
 * census classifies the part as binary and does not read it, `C-REGEN` compares
 * one hash, and PowerPoint does not load a VBA project when it opens a file.
 * So a compound file with a broken sector chain, or a compressor that cannot
 * decompress its own output, would sail through every other gate.
 *
 * `readCfb` is deliberately a reader rather than the writer run backwards, and
 * `decompressOvba` was written from the format rather than from
 * `compressOvba`. Testing a pair against each other is weaker than testing
 * either against a measurement - which is why `ROSTER.md` carries the VBA
 * project itself as a declared gap - but it is what is available here, and it
 * catches everything structural.
 */

const bytes = (text: string): Uint8Array => new TextEncoder().encode(text);
const text = (raw: Uint8Array): string => new TextDecoder().decode(raw);

function find(nodes: readonly CfbNode[], path: readonly string[]): CfbNode {
  const [head, ...rest] = path;
  const node = nodes.find((child) => child.name === head);
  if (node === undefined) throw new Error('no such node: ' + path.join('/'));
  if (rest.length === 0) return node;
  if (node.kind !== 'storage') throw new Error(String(head) + ' is not a storage');
  return find(node.children, rest);
}

describe('compound file', () => {
  it('round-trips a tree of storages and streams', () => {
    const tree: CfbNode[] = [
      stream('PROJECT', bytes('ID="{0}"\r\nName="Probe"\r\n')),
      stream('PROJECTwm', Uint8Array.from([0x41, 0x00, 0x41, 0x00, 0x00, 0x00, 0x00, 0x00])),
      storage('VBA', [
        stream('_VBA_PROJECT', Uint8Array.from([0xcc, 0x61, 0xff, 0xff, 0, 0, 0])),
        stream('dir', bytes('a compressed record stream, notionally')),
        stream('Module1', bytes('Attribute VB_Name = "Module1"')),
      ]),
    ];

    const back = readCfb(writeCfb(tree));

    expect(text((find(back, ['PROJECT']) as { bytes: Uint8Array }).bytes)).toBe(
      'ID="{0}"\r\nName="Probe"\r\n',
    );
    expect(text((find(back, ['VBA', 'Module1']) as { bytes: Uint8Array }).bytes)).toBe(
      'Attribute VB_Name = "Module1"',
    );
    expect(find(back, ['VBA']).kind).toBe('storage');
  });

  it('orders siblings by name length first, then case-insensitively', () => {
    // The rule readers enforce, and the one a default `sort()` gets wrong:
    // `VBA` precedes `PROJECT` because it is shorter, not because of `V` < `P`.
    const back = readCfb(
      writeCfb([
        stream('PROJECTwm', bytes('c')),
        stream('PROJECT', bytes('b')),
        storage('VBA', [stream('zzz', bytes('x')), stream('dir', bytes('y'))]),
      ]),
    );
    expect(back.map((node) => node.name)).toEqual(['VBA', 'PROJECT', 'PROJECTwm']);
    const vba = find(back, ['VBA']);
    if (vba.kind !== 'storage') throw new Error('VBA should be a storage');
    expect(vba.children.map((node) => node.name)).toEqual(['dir', 'zzz']);
  });

  it('keeps a stream that is larger than the mini-stream cutoff', () => {
    // 4096 bytes is the boundary: below it a stream lives in the mini stream
    // and its start is a mini-sector index; at or above it the same field is a
    // regular sector number. Getting that backwards reads garbage.
    const large = new Uint8Array(9000);
    for (let i = 0; i < large.length; i++) large[i] = (i * 37) & 0xff;
    const back = readCfb(writeCfb([stream('big', large), stream('sm', bytes('small'))]));
    const found = find(back, ['big']);
    if (found.kind !== 'stream') throw new Error('big should be a stream');
    expect(found.bytes).toHaveLength(9000);
    expect([...found.bytes.subarray(0, 8)]).toEqual([...large.subarray(0, 8)]);
    expect([...found.bytes.subarray(8992)]).toEqual([...large.subarray(8992)]);
  });

  it('refuses a name longer than the directory entry can hold', () => {
    expect(() => writeCfb([stream('x'.repeat(32), bytes('y'))])).toThrow(/at most 31/);
  });
});

describe('MS-OVBA compression', () => {
  it('round-trips an empty input', () => {
    expect([...decompressOvba(compressOvba(new Uint8Array(0)))]).toEqual([]);
  });

  it('round-trips text with long repeats, which is what makes copy tokens fire', () => {
    const source = bytes(
      'Attribute VB_Name = "Module1"\r\nSub Noop()\r\nEnd Sub\r\n'.repeat(20) +
        'x'.repeat(300) +
        'Attribute VB_Name = "Module1"\r\n',
    );
    const packed = compressOvba(source);
    expect(packed[0]).toBe(0x01);
    expect(packed.length).toBeLessThan(source.length);
    expect([...decompressOvba(packed)]).toEqual([...source]);
  });

  it('round-trips input that spans several chunks', () => {
    // Chunks are at most 4096 decompressed bytes and the back-reference window
    // resets at every boundary, so a stream long enough to need three chunks is
    // the case where a compressor that forgot to reset produces offsets
    // pointing before the chunk start.
    const source = new Uint8Array(10_000);
    for (let i = 0; i < source.length; i++) source[i] = (i % 251) & 0xff;
    expect([...decompressOvba(compressOvba(source))]).toEqual([...source]);
  });

  it('round-trips incompressible input', () => {
    const source = new Uint8Array(2000);
    let state = 0x12345678;
    for (let i = 0; i < source.length; i++) {
      state = (Math.imul(state, 1664525) + 1013904223) >>> 0;
      source[i] = (state >>> 16) & 0xff;
    }
    expect([...decompressOvba(compressOvba(source))]).toEqual([...source]);
  });

  it('rejects a container with no signature byte', () => {
    expect(() => decompressOvba(Uint8Array.from([0x02, 0x00, 0xb0]))).toThrow(/signature/);
  });
});

describe('the VBA project a32 carries', () => {
  const spec = {
    projectName: 'VBAProject',
    moduleName: 'Module1',
    source: 'Sub Noop()\r\nEnd Sub\r\n',
  };

  it('decompresses back to the source it was given', () => {
    const nodes = vbaProjectNodes(spec);
    expect(text(decompressOvba(nodes.module))).toBe(
      'Attribute VB_Name = "Module1"\r\nSub Noop()\r\nEnd Sub\r\n',
    );
  });

  it('writes a dir stream whose records walk to the terminator', () => {
    // Reading the record sequence back is what catches the trap the module
    // comment names: PROJECTVERSION, MODULETYPE and both terminators use their
    // 4-byte field as a fixed Reserved value rather than as a size, so a walker
    // that treats every record the same runs off the end of the stream.
    // The value is the tail after the 2-byte id and the 4-byte Reserved.
    const RESERVED_NOT_SIZE = new Map([
      [0x0009, 6], // PROJECTVERSION: VersionMajor (4) then VersionMinor (2)
      [0x0021, 0], // MODULETYPE
      [0x002b, 0], // MODULE terminator
      [0x0010, 0], // dir terminator
    ]);
    // The five records that carry an MBCS string, a Reserved, then a UTF-16
    // copy - each with its own Reserved value.
    const PAIRED = new Map([
      [0x0005, 0x0040],
      [0x0006, 0x003d],
      [0x000c, 0x003c],
      [0x001a, 0x0032],
      [0x001c, 0x0048],
    ]);

    const dir = decompressOvba(vbaProjectNodes(spec).dir);
    const view = new DataView(dir.buffer, dir.byteOffset, dir.byteLength);
    const seen: number[] = [];
    let at = 0;
    while (at < dir.length) {
      const id = view.getUint16(at, true);
      seen.push(id);
      const fixed = RESERVED_NOT_SIZE.get(id);
      if (fixed !== undefined) {
        at += 6 + fixed;
        continue;
      }
      const size = view.getUint32(at + 2, true);
      at += 6 + size;
      const reserved = PAIRED.get(id);
      if (reserved === undefined) continue;
      expect(view.getUint16(at, true), 'reserved of record 0x' + id.toString(16)).toBe(reserved);
      at += 2;
      at += 4 + view.getUint32(at, true);
    }

    expect(at, 'the walk ends exactly at the end of the stream').toBe(dir.length);
    expect(seen[0], 'PROJECTSYSKIND first').toBe(0x0001);
    expect(seen.at(-1), 'the dir terminator last').toBe(0x0010);
    expect(seen).toContain(0x000f); // PROJECTMODULES
    expect(seen).toContain(0x0047); // MODULENAMEUNICODE
  });

  it('assembles into a compound file that reads back', () => {
    const nodes = vbaProjectNodes(spec);
    const back = readCfb(
      writeCfb([
        stream('PROJECT', nodes.project),
        stream('PROJECTwm', nodes.projectWm),
        storage('VBA', [
          stream('_VBA_PROJECT', nodes.vbaProject),
          stream('dir', nodes.dir),
          stream(spec.moduleName, nodes.module),
        ]),
      ]),
    );
    const module = find(back, ['VBA', 'Module1']);
    if (module.kind !== 'stream') throw new Error('Module1 should be a stream');
    expect(text(decompressOvba(module.bytes))).toContain('Sub Noop()');
    expect(text((find(back, ['PROJECT']) as { bytes: Uint8Array }).bytes)).toContain(
      'Name="VBAProject"',
    );
  });
});
