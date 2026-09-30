/**
 * Experiment C9, step 1 - one package per deck of probes, and `cascade-inputs.json` for `read.ps1`.
 *
 * ```
 * node tools/ground-truth/model/tables/cascade/build-deck.ts <out-dir>
 * ```
 *
 * Before PowerPoint is asked, every pair of readings is either told apart by some probe or searched
 * for a table that would: one found throws, naming it; none found is recorded as indistinguishable.
 */

import { mkdirSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';

import { buildSheetPackage } from '../../../lib/sheet-pptx.ts';
import {
  allDecks,
  APPLICABILITY,
  BASELINE,
  BUILTIN_TABLE_STYLE_XML,
  BUILTINS,
  EDGE_MODELS,
  EXPORT,
  fillStack,
  flagsOf,
  maskOf,
  MASTER_BG,
  MERGES,
  ORDERS,
  partsGrid,
  predictions,
  SCALE,
  styledTables,
  tableFrame,
  THEMES,
  type DeckSpec,
  type Scenario,
} from './probes.ts';

const dir = process.argv[2];
if (dir === undefined) {
  throw new Error('usage: tools/ground-truth/model/tables/cascade/build-deck.ts <out-dir>');
}

function build(deck: DeckSpec): Uint8Array {
  const { theme, clrMap } = THEMES[deck.theme];
  return buildSheetPackage({
    themes: [theme],
    masters: [
      {
        theme: 0,
        clrMap,
        bg: MASTER_BG,
        ...(deck.txStyles === undefined ? {} : { txStyles: deck.txStyles }),
        ...(deck.masterShapes === undefined ? {} : { shapes: deck.masterShapes }),
      },
    ],
    layouts: [
      {
        master: 0,
        type: deck.layoutShapes === undefined ? 'blank' : 'obj',
        name: deck.layoutShapes === undefined ? 'Blank' : 'Title and Content',
        ...(deck.layoutShapes === undefined ? {} : { shapes: deck.layoutShapes }),
      },
    ],
    slides: deck.slides.map((table, k) => ({
      layout: 0,
      shapes: [tableFrame(`${deck.id}#${String(k)}`, table)],
    })),
    ...(deck.defaultTextStyle === undefined ? {} : { defaultTextStyle: deck.defaultTextStyle }),
  });
}

/* -------------------------------------------------------------------------- */
/* separability                                                               */
/* -------------------------------------------------------------------------- */

type UnitClass = 'fill' | 'text' | 'edge';

const classOf = (key: string): UnitClass =>
  key.endsWith(':fill') ? 'fill' : key.endsWith(':text') ? 'text' : 'edge';

interface Dimension {
  readonly name: string;
  readonly values: readonly string[];
  readonly classes: readonly UnitClass[];
  readonly apply: (value: string) => Scenario;
}

const DIMENSIONS: readonly Dimension[] = [
  {
    name: 'applicability',
    values: APPLICABILITY,
    classes: ['fill', 'text', 'edge'],
    apply: (v) => ({ ...BASELINE, reading: { ...BASELINE.reading, applicability: v as never } }),
  },
  {
    name: 'order',
    values: Object.keys(ORDERS),
    classes: ['fill', 'text', 'edge'],
    apply: (v) => ({ ...BASELINE, reading: { ...BASELINE.reading, order: v as never } }),
  },
  {
    name: 'merge',
    values: MERGES,
    classes: ['fill', 'text', 'edge'],
    apply: (v) => ({ ...BASELINE, reading: { ...BASELINE.reading, merge: v as never } }),
  },
  {
    name: 'edges',
    values: EDGE_MODELS,
    classes: ['edge'],
    apply: (v) => ({ ...BASELINE, edges: v as never }),
  },
  {
    name: 'source',
    values: ['explicit', 'fontRef'],
    classes: ['text'],
    apply: (v) => ({ ...BASELINE, source: v as never }),
  },
];

/** The first key of class `unit` where two scenarios predict differently over `decks`, or `null`. */
function firstDifference(
  decks: readonly DeckSpec[],
  a: Scenario,
  b: Scenario,
  unit: UnitClass,
): string | null {
  const other = predictions(decks, b);
  for (const [key, value] of predictions(decks, a)) {
    const next = other.next();
    if (next.done === true || next.value[0] !== key)
      throw new Error(`unit order diverged at ${key}`);
    if (classOf(key) === unit && next.value[1] !== value) return key;
  }
  return null;
}

/** One built-in per structure: accent variants are one template (ADR 0063), so one stands for all. */
function templates(): string[] {
  const seen = new Map<string, string>();
  for (const style of BUILTINS) {
    const xml = BUILTIN_TABLE_STYLE_XML.get(style.id) ?? '';
    const key = xml.replace(/ style(Id|Name)="[^"]*"/g, '').replace(/accent\d/g, 'accentN');
    if (!seen.has(key)) seen.set(key, style.id);
  }
  return [...seen.values()];
}

const SHAPES = Array.from(
  { length: 36 },
  (_, k) => [Math.floor(k / 6) + 1, (k % 6) + 1] as const,
).sort((p, q) => p[0] * p[1] - q[0] * q[1] || p[0] - q[0]);

/** The smallest table of the search space on which two scenarios differ for `unit`, or `null`. */
function searchSpace(
  a: Scenario,
  b: Scenario,
  unit: UnitClass,
  styles: readonly string[],
): string | null {
  for (const [rows, cols] of SHAPES) {
    for (const style of styles) {
      const deck: DeckSpec = {
        id: 'space',
        group: 'sweep',
        theme: 'A',
        read: 'sweep',
        slides: Array.from({ length: 64 }, (_, mask) => ({
          rows,
          cols,
          style,
          flags: flagsOf(mask),
        })),
      };
      const key = firstDifference([deck], a, b, unit);
      if (key !== null) {
        const slide = Number(/#(\d+)@/.exec(key)?.[1]);
        const name = BUILTINS.find((s) => s.id === style)?.name ?? style;
        return `${name}, ${String(rows)}x${String(cols)}, mask ${String(slide)}, ${key.slice(key.indexOf('@'))}`;
      }
    }
  }
  return null;
}

interface Verdict {
  readonly dimension: string;
  readonly pair: readonly [string, string];
  readonly unit: UnitClass;
  readonly probe: string | null;
}

function separability(decks: readonly DeckSpec[]): Verdict[] {
  const probes = decks.filter(
    (d) => d.group === 'sweep' || d.group === 'ends' || d.group === 'sides',
  );
  const styles = templates();
  const verdicts: Verdict[] = [];
  for (const dimension of DIMENSIONS) {
    for (const [i, u] of dimension.values.entries()) {
      for (const v of dimension.values.slice(i + 1)) {
        for (const unit of dimension.classes) {
          const a = dimension.apply(u);
          const b = dimension.apply(v);
          const probe = firstDifference(probes, a, b, unit);
          if (probe === null) {
            const found = searchSpace(a, b, unit, styles);
            if (found !== null) {
              throw new Error(
                `${dimension.name} ${u} / ${v} (${unit}): no probe separates them, but ${found} does`,
              );
            }
          }
          verdicts.push({ dimension: dimension.name, pair: [u, v], unit, probe });
        }
      }
    }
  }
  return verdicts;
}

/** Stacked and winning translucent fills must differ on some probe cell, or the paint question asks nothing. */
function assertStackable(decks: readonly DeckSpec[]): void {
  const translucent = (fill: ReturnType<typeof fillStack>[number] | undefined): boolean =>
    fill?.kind === 'value' &&
    fill.value.type === 'solid' &&
    fill.value.color.transforms.some((t) => t.op === 'alpha');
  for (const { table, style } of styledTables(decks)) {
    for (const row of partsGrid('spec', table)) {
      for (const parts of row) {
        const stack = fillStack(style, parts, 'schema');
        if (stack.length >= 2 && translucent(stack.at(-1))) return;
      }
    }
  }
  throw new Error('fill painting: no probe cell has a translucent fill over another');
}

const decks = allDecks();
const verdicts = separability(decks);
assertStackable(decks);
for (const v of verdicts.filter((x) => x.probe === null)) {
  console.log(
    `indistinguishable over every built-in template, 64 flag sets, 1x1-6x6: ${v.dimension} ${v.pair.join(' / ')} (${v.unit})`,
  );
}
mkdirSync(dir, { recursive: true });
for (const deck of decks) writeFileSync(join(dir, `${deck.id}.pptx`), build(deck));

const inputs = {
  emuPerPoint: 12700,
  scale: SCALE,
  export: EXPORT,
  /** One BMP twin per group proves the PNG decoder against the same slide. */
  twins: [...new Set(decks.map((d) => d.group))].map((g) => decks.find((d) => d.group === g)?.id),
  separability: verdicts,
  decks: decks.map((deck) => ({
    id: deck.id,
    file: `${deck.id}.pptx`,
    group: deck.group,
    theme: deck.theme,
    read: deck.read,
    resave: deck.resave === true,
    slides: deck.slides.map((s) => ({
      rows: s.rows,
      cols: s.cols,
      style: s.style,
      mask: maskOf(s.flags),
      rtl: s.rtl === true,
    })),
  })),
};
writeFileSync(join(dir, 'cascade-inputs.json'), JSON.stringify(inputs, null, 2));
console.log(
  `wrote ${String(decks.length)} packages, ${String(decks.reduce((n, d) => n + d.slides.length, 0))} slides`,
);
