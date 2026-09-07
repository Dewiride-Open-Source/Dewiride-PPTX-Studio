import { readFileSync } from 'node:fs';
import { join } from 'node:path';

import { REPO_ROOT as ROOT } from '../../repo/root.ts';
import { PartStore } from '@pptx-studio/opc';
import {
  loadDocument,
  resolveIndent,
  resolveMarginLeft,
  resolveRun,
  resolveSize,
  textLevels,
  type Document,
  type Origin,
  type Paragraph,
  type Shape,
  type Sheet,
  type TextContext,
} from '@pptx-studio/model';
import { describe, expect, it } from 'vitest';

/**
 * Sub-phase 3.1's honest check: the cascade against fifty-four real decks.
 *
 * The unit suite is written from `corpus/ground-truth/text-cascade.json` and
 * proves the resolver reproduces what PowerPoint did on 39 packages built to
 * ask one question each. It cannot prove the *parser* survives markup nobody
 * designed - a `p:txBody` with no paragraphs, an `a:rPr` carrying a gradient, a
 * master whose `p:txStyles` declares seven of nine levels, an `a:fld` with a
 * type nobody has heard of. Real decks have all of that and the probe packages
 * have none of it, so this is the check that fails the way the feature will.
 *
 * It asserts three things, and the third is the one that matters:
 *
 * 1. Every paragraph on every sheet of every deck resolves without throwing.
 * 2. Every resolved size, margin and indent is a finite number in a range a
 *    slide could hold.
 * 3. The corpus actually exercises the cascade rather than trivially passing:
 *    the counts are asserted non-zero, and every origin a real deck can produce
 *    is asserted to have been produced by one.
 */

const COLLECTIONS = ['decks', 'authored', 'written'];

interface Deck {
  readonly id: string;
  readonly bytes: Uint8Array;
}

function committedDecks(): Deck[] {
  const decks: Deck[] = [];
  for (const collection of COLLECTIONS) {
    const dir = join(ROOT, 'corpus', collection);
    const envelope = JSON.parse(readFileSync(join(dir, 'manifest.json'), 'utf8')) as {
      entries: { id: string; kind: string; storage: string; path?: string }[];
    };
    for (const entry of envelope.entries) {
      if (entry.kind !== 'deck' || entry.storage !== 'committed') continue;
      decks.push({
        id: entry.id,
        bytes: new Uint8Array(readFileSync(join(dir, entry.path ?? ''))),
      });
    }
  }
  return decks;
}

/** Every shape on a sheet, groups flattened, because a group holds text too. */
function allShapes(shapes: readonly Shape[]): Shape[] {
  const out: Shape[] = [];
  for (const shape of shapes) {
    out.push(shape);
    if (shape.children.length > 0) out.push(...allShapes(shape.children));
  }
  return out;
}

interface Tally {
  paragraphs: number;
  runs: number;
  origins: Map<Origin, number>;
  sizes: number[];
  problems: string[];
}

/** A size a slide could hold: DrawingML bounds `a:defRPr/@sz` at 1pt to 4000pt. */
const MIN_SZ = 100;
const MAX_SZ = 400000;
/** A margin a slide could hold. Ten metres of EMU either way is generous. */
const MAX_EMU = 360000000;

function sweep(document: Document, sheet: Sheet, deck: string, tally: Tally): void {
  for (const shape of allShapes(sheet.shapes)) {
    const body = shape.text;
    if (body === undefined) continue;
    const context: TextContext = {
      sheet,
      shape,
      defaultTextStyle: document.defaultTextStyle,
    };
    for (const paragraph of body.paragraphs) {
      tally.paragraphs += 1;
      check(context, paragraph, `${deck} ${sheet.partName} ${shape.name}`, tally);
    }
  }
}

function check(context: TextContext, paragraph: Paragraph, where: string, tally: Tally): void {
  const size = resolveSize(context, paragraph, undefined);
  const marL = resolveMarginLeft(context, paragraph);
  const indent = resolveIndent(context, paragraph);
  tally.origins.set(size.origin, (tally.origins.get(size.origin) ?? 0) + 1);
  tally.sizes.push(size.value);

  if (!Number.isFinite(size.value) || size.value < MIN_SZ || size.value > MAX_SZ) {
    tally.problems.push(`${where}: size ${String(size.value)}`);
  }
  if (!Number.isFinite(marL.value) || Math.abs(marL.value) > MAX_EMU) {
    tally.problems.push(`${where}: marL ${String(marL.value)}`);
  }
  if (!Number.isFinite(indent.value) || Math.abs(indent.value) > MAX_EMU) {
    tally.problems.push(`${where}: indent ${String(indent.value)}`);
  }
  // The walk must be finite and must end somewhere: a chain that returned to a
  // sheet it had visited would loop here rather than in a Web Worker.
  const levels = textLevels(context);
  if (levels.length === 0 || levels.length > 8) {
    tally.problems.push(`${where}: ${String(levels.length)} levels`);
  }

  for (const run of paragraph.content) {
    tally.runs += 1;
    const runSize = resolveSize(context, paragraph, run);
    tally.origins.set(runSize.origin, (tally.origins.get(runSize.origin) ?? 0) + 1);
    if (!Number.isFinite(runSize.value) || runSize.value < MIN_SZ || runSize.value > MAX_SZ) {
      tally.problems.push(`${where}: run size ${String(runSize.value)}`);
    }
    // A typeface may be genuinely unresolved - measured in 3.1, and it is a 3.7
    // question - but if one resolves it must be a non-empty name.
    const latin = resolveRun(context, paragraph, run, (p) => p.latin);
    if (latin !== undefined && latin.value.typeface === '') {
      tally.problems.push(`${where}: empty typeface`);
    }
  }
}

const DECKS = committedDecks();

describe('the text cascade over every committed deck', () => {
  const tally: Tally = {
    paragraphs: 0,
    runs: 0,
    origins: new Map(),
    sizes: [],
    problems: [],
  };
  const failures: string[] = [];

  for (const deck of DECKS) {
    try {
      const document = loadDocument(PartStore.open(deck.bytes));
      for (const sheet of [...document.slides, ...document.layouts, ...document.masters]) {
        sweep(document, sheet, deck.id, tally);
      }
    } catch (error) {
      failures.push(`${deck.id}: ${error instanceof Error ? error.message : String(error)}`);
    }
  }

  it('parses and resolves all fifty-four without throwing', () => {
    expect(failures).toEqual([]);
    expect(DECKS).toHaveLength(54);
  });

  it('resolves every paragraph to a size, a margin and an indent a slide could hold', () => {
    expect(tally.problems.slice(0, 12)).toEqual([]);
  });

  it('is not passing vacuously', () => {
    // A sweep over a corpus that happened to contain no text would report no
    // problems and prove nothing, so the counts are assertions too.
    expect(tally.paragraphs).toBeGreaterThan(500);
    expect(tally.runs).toBeGreaterThan(500);
    expect(new Set(tally.sizes).size).toBeGreaterThan(5);
  });

  it('exercises every origin a real deck can produce', () => {
    // `builtin` is deliberately absent: every PowerPoint-authored master writes
    // `p:txStyles`, which is exactly why the built-in styles needed a probe
    // package of their own rather than a corpus deck.
    const histogram = [...tally.origins]
      .sort((a, b) => b[1] - a[1])
      .map(([origin, n]) => origin + '=' + String(n))
      .join(' ');
    const missing = ['run', 'paragraph', 'shape', 'layoutPh', 'masterPh', 'txStyles'].filter(
      (origin) => (tally.origins.get(origin as Origin) ?? 0) === 0,
    );
    expect(missing, histogram).toEqual([]);
    expect(tally.origins.get('builtin') ?? 0).toBe(0);

    // What the corpus currently exercises, as a fact and not only as a gate:
    //
    //   txStyles 1453  schemaDefault 899  run 887  masterPh 556
    //   shape 362  defaultTextStyle 64  layoutPh 50  paragraph 2
    //
    // Pinned as a set with floors rather than as exact counts, because a deck
    // added to the corpus should not turn this red - but an origin that stops
    // being exercised at all should.
    expect([...tally.origins.keys()].sort(), histogram).toEqual([
      'defaultTextStyle',
      'layoutPh',
      'masterPh',
      'paragraph',
      'run',
      'schemaDefault',
      'shape',
      'txStyles',
    ]);
    const floors: Readonly<Record<string, number>> = {
      txStyles: 1000,
      schemaDefault: 500,
      run: 500,
      masterPh: 300,
      shape: 200,
      defaultTextStyle: 30,
      layoutPh: 20,
      paragraph: 1,
    };
    for (const [origin, floor] of Object.entries(floors)) {
      expect(tally.origins.get(origin as Origin) ?? 0, origin).toBeGreaterThanOrEqual(floor);
    }
  });
});
