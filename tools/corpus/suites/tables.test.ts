import { readFileSync } from 'node:fs';
import { join } from 'node:path';

import { REPO_ROOT as ROOT } from '../../repo/root.ts';
import { PartStore } from '@pptx-studio/opc';
import {
  BUILTIN_TABLE_STYLES,
  loadDocument,
  tableGrid,
  tableStyleOf,
  type Shape,
  type TableGrid,
  type TableStyle,
  type TableStyleRef,
} from '@pptx-studio/model';
import { describe, expect, it } from 'vitest';

/**
 * Every table in every committed deck through the parser, the grid (4.1) and the style (4.2), and
 * the two decks that carry tables held to what PowerPoint drew and wrote. The census counts are the
 * floor, so a table the parser silently skipped would show here.
 */

const COLLECTIONS = ['decks', 'authored', 'written'];

interface Deck {
  readonly id: string;
  readonly tables: number;
  readonly bytes: Uint8Array;
}

function committedDecks(): Deck[] {
  const decks: Deck[] = [];
  for (const collection of COLLECTIONS) {
    const dir = join(ROOT, 'corpus', collection);
    const envelope = JSON.parse(readFileSync(join(dir, 'manifest.json'), 'utf8')) as {
      entries: {
        id: string;
        kind: string;
        storage: string;
        path?: string;
        features?: Record<string, number>;
      }[];
    };
    for (const entry of envelope.entries) {
      if (entry.kind !== 'deck' || entry.storage !== 'committed') continue;
      decks.push({
        id: entry.id,
        tables: entry.features?.['table'] ?? 0,
        bytes: new Uint8Array(readFileSync(join(dir, entry.path ?? ''))),
      });
    }
  }
  return decks;
}

function allShapes(shapes: readonly Shape[]): Shape[] {
  const out: Shape[] = [];
  for (const shape of shapes) {
    out.push(shape);
    if (shape.children.length > 0) out.push(...allShapes(shape.children));
  }
  return out;
}

interface Found {
  readonly deck: string;
  readonly part: string;
  readonly name: string;
  readonly grid: TableGrid;
  readonly ref: TableStyleRef | undefined;
  readonly style: TableStyle | null;
}

const found: Found[] = [];
const failures: string[] = [];
const DECKS = committedDecks();
for (const deck of DECKS) {
  try {
    const document = loadDocument(PartStore.open(deck.bytes));
    for (const sheet of [...document.slides, ...document.layouts, ...document.masters]) {
      for (const shape of allShapes(sheet.shapes)) {
        if (shape.table === undefined) continue;
        found.push({
          deck: deck.id,
          part: sheet.partName,
          name: shape.name,
          grid: tableGrid(shape.table),
          ref: shape.table.props?.style,
          style: tableStyleOf(shape.table),
        });
      }
    }
  } catch (error) {
    failures.push(`${deck.id}: ${error instanceof Error ? error.message : String(error)}`);
  }
}

const spans = (grid: TableGrid): string[] =>
  grid.anchors
    .filter((a) => a.rows > 1 || a.cols > 1)
    .map((a) => `${String(a.row + 1)},${String(a.col + 1)}:${String(a.rows)}x${String(a.cols)}`);

describe('every table in every committed deck', () => {
  it('parses without throwing, as many as the census counted', () => {
    expect(failures).toEqual([]);
    expect(DECKS).toHaveLength(55);
    const counted = DECKS.reduce((n, d) => n + d.tables, 0);
    expect(found).toHaveLength(counted);
    expect(counted).toBeGreaterThanOrEqual(5);
  });

  it('resolves every grid to positions that all have an owner inside the grid', () => {
    for (const { deck, name, grid } of found) {
      const where = `${deck} ${name}`;
      expect(grid.rows, where).toBeGreaterThan(0);
      expect(grid.cols, where).toBeGreaterThan(0);
      expect(grid.positions, where).toHaveLength(grid.rows);
      for (const line of grid.positions) {
        expect(line, where).toHaveLength(grid.cols);
        for (const owner of line) {
          expect(owner.row + owner.rows, where).toBeLessThanOrEqual(grid.rows);
          expect(owner.col + owner.cols, where).toBeLessThanOrEqual(grid.cols);
        }
      }
      expect(
        grid.widths.every((w) => Number.isInteger(w) && w >= 0),
        where,
      ).toBe(true);
      expect(
        grid.heights.every((h) => Number.isInteger(h) && h >= 0),
        where,
      ).toBe(true);
    }
  });

  it('reads b04-table, which PowerPoint wrote, as the grid its object model reported', () => {
    // ADR 0056 opens with this deck's object-model reading: Cell(2,3) answers with Cell(2,2)'s
    // rectangle and Cell(4,1) with Cell(3,1)'s.
    const b04 = found.find((f) => f.deck === 'b04-table');
    expect(b04).toBeDefined();
    expect(b04!.grid.rows).toBe(4);
    expect(b04!.grid.cols).toBe(4);
    expect(b04!.grid.widths).toEqual([2667000, 2667000, 2667000, 2667000]);
    expect(b04!.grid.heights).toEqual([952500, 952500, 952500, 952500]);
    expect(spans(b04!.grid)).toEqual(['2,2:1x2', '3,1:2x1']);
    expect(b04!.grid.positions[1]![2]!).toBe(b04!.grid.positions[1]![1]!);
    expect(b04!.grid.positions[3]![0]!).toBe(b04!.grid.positions[2]![0]!);
    expect(b04!.grid.anchors.every((a) => a.cell !== null)).toBe(true);
  });

  it('reads a20-tables: a 2x2 block beside a row span, every a:tcPr child, and both style sources', () => {
    const a20 = found.filter((f) => f.deck === 'a20-tables');
    expect(a20.map((f) => f.name)).toEqual([
      'Merged table',
      'Cell properties table',
      'Styled by id',
      'Styled inline',
    ]);
    const [merged, props, byId] = a20;
    expect(spans(merged!.grid)).toEqual(['1,1:1x2', '2,1:2x1', '2,3:2x2']);
    const borders = props!.grid.positions[0]!.map((p) => p.cell?.props?.borders);
    expect(borders[0]).toMatchObject({ left: { w: 57150 }, right: { w: 57150 }, top: undefined });
    expect(borders[3]).toMatchObject({ tlToBr: { w: 12700 }, blToTr: { w: 12700 } });
    expect(props!.grid.positions[1]!.map((p) => p.cell?.props?.anchor)).toEqual([
      undefined,
      'b',
      'ctr',
      undefined,
    ]);
    expect(props!.grid.positions[2]!.map((p) => p.cell?.props?.vert)).toEqual([
      'vert270',
      'vert',
      'eaVert',
      'wordArtVert',
    ]);
    expect(props!.grid.positions[3]!.map((p) => p.cell?.props?.fill?.type)).toEqual([
      'solid',
      'gradient',
      'pattern',
      'solid',
    ]);
    expect(byId!.grid.anchors[0]!.cell?.text?.paragraphs[0]?.content[0]).toMatchObject({
      text: 'tableStyleId',
    });
    expect(merged!.grid.anchors[0]!.cell?.props?.marL).toBeUndefined();
  });

  it('resolves every table to the built-in its GUID names, or to none', () => {
    const builtins = new Set(BUILTIN_TABLE_STYLES.map((s) => s.id));
    for (const { deck, name, ref, style } of found) {
      const named = ref !== undefined && builtins.has(ref.id.toUpperCase());
      expect(style?.id.toUpperCase(), `${deck} ${name}`).toBe(
        named ? ref.id.toUpperCase() : undefined,
      );
    }
    expect(found.filter((f) => f.style !== null).length).toBeGreaterThanOrEqual(3);
  });

  it("finds b04-table's Medium Style 2 - Accent 1 written byte for byte as C8 has it", () => {
    // Two PowerPoint sessions weeks apart, one through the gallery and one through AddTable.
    const b04 = DECKS.find((d) => d.id === 'b04-table')!;
    const part = new TextDecoder().decode(PartStore.open(b04.bytes).read('/ppt/tableStyles.xml'));
    const medium = BUILTIN_TABLE_STYLES.find((s) => s.name === 'Medium Style 2 - Accent 1')!;
    expect(part).toContain(medium.xml);
    expect(found.find((f) => f.deck === 'b04-table')?.style?.name).toBe(
      'Medium Style 2 - Accent 1',
    );
  });

  it("reads a20-tables' two style sources as PowerPoint draws them", () => {
    const byId = found.find((f) => f.deck === 'a20-tables' && f.name === 'Styled by id')!;
    const inline = found.find((f) => f.deck === 'a20-tables' && f.name === 'Styled inline')!;
    expect(byId.ref?.kind).toBe('id');
    expect(byId.style?.name).toBe('Medium Style 2 - Accent 1');
    // Its own a:tableStyle is kept in the file and never drawn: the GUID is no built-in (C8).
    expect(inline.ref).toMatchObject({
      kind: 'inline',
      id: '{1B1D64F0-0000-4000-8000-0000000A200A}',
    });
    expect(inline.style).toBeNull();
  });
});
