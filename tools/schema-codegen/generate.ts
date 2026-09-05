/**
 * Generate `packages/xml/src/edit/schema-order.gen.ts` from the ECMA-376
 * Transitional XSDs.
 *
 * ## Getting the input
 *
 * The schemas are not in this repository. They are a free download from Ecma
 * International, and the codegen is run by hand when they change - which, for a
 * standard whose last edition is December 2016, is not often:
 *
 * ```
 * curl -LO https://ecma-international.org/wp-content/uploads/ECMA-376-4_5th_edition_december_2016.zip
 * # extract OfficeOpenXML-XMLSchema-Transitional.zip from it, then extract that
 * node tools/schema-codegen/generate.ts --schemas <dir> --write
 * ```
 *
 * The generated file records the SHA-256 of every input, so "which schemas is
 * this table from" has an answer that does not depend on anyone's memory. Only
 * the generated file is committed; no Ecma material enters the tree.
 *
 * Run without `--write` for the report alone, which is what the ADR quotes.
 *
 * ## Which schemas, and why not all of them
 *
 * PresentationML, DrawingML main, and the two one-element schemas that sit on
 * top of them. Not `dml-chart`, not `dml-diagram`, not `sml`, not `wml`, and
 * that is a design rule rather than an omission: **we never rewrite a chart, a
 * diagram or an embedding.** A chart part is copied byte-for-byte, so an order
 * table for it would be code that ships to every browser to serve an operation
 * the plan forbids. If a later phase needs one, it is a line in `SCHEMA_FILES`.
 */

import { createHash } from 'node:crypto';
import { readFileSync, writeFileSync } from 'node:fs';
import { join, resolve } from 'node:path';
import { readXsd, type XsdSchema } from './read-xsd.ts';
import { ambiguities, contextTableOf, flatten, tableOf, universeOf } from './order.ts';

const SCHEMA_FILES = ['pml.xsd', 'dml-main.xsd', 'dml-picture.xsd', 'dml-lockedCanvas.xsd'];

const OUT = 'packages/xml/src/edit/schema-order.gen.ts';

function parseArgs(argv: readonly string[]): { schemas: string; out: string; write: boolean } {
  let schemas = '';
  let out = OUT;
  let write = false;
  for (let i = 0; i < argv.length; i++) {
    const arg = argv[i];
    if (arg === '--schemas') schemas = argv[++i] ?? '';
    else if (arg === '--out') out = argv[++i] ?? OUT;
    else if (arg === '--write') write = true;
    else throw new Error(`unknown argument ${String(arg)}`);
  }
  if (schemas === '') {
    throw new Error('--schemas <dir> is required: the directory holding pml.xsd and dml-main.xsd');
  }
  return { schemas, out, write };
}

const { schemas, out, write } = parseArgs(process.argv.slice(2));

const sources: { file: string; sha256: string }[] = [];
const parsed: XsdSchema[] = [];
for (const file of SCHEMA_FILES) {
  const bytes = readFileSync(join(schemas, file));
  sources.push({ file, sha256: createHash('sha256').update(bytes).digest('hex') });
  parsed.push(readXsd(bytes.toString('utf8'), file));
}

const universe = universeOf(parsed);
const flattened = flatten(universe);
const collisions = [...flattened.orders]
  .filter(([, order]) => order.collisions.length > 0)
  .map(([typeKey, order]) => `${typeKey}: ${order.collisions.join(', ')}`);
const ambiguous = ambiguities(universe, flattened);
const ambiguousKeys = new Set(ambiguous.map((a) => a.elementKey));
const table = tableOf(universe, flattened, ambiguousKeys);
const context = contextTableOf(universe, flattened, ambiguousKeys);

// ------------------------------------------------------------------ the report

console.log(`schemas      ${String(parsed.length)}`);
for (const { file, sha256 } of sources) console.log(`  ${file.padEnd(22)} ${sha256}`);
console.log(`complex types ${String(universe.complexTypes.size)}`);
console.log(`groups        ${String(universe.groups.size)}`);
console.log(`global elems  ${String(universe.globalElements.size)}`);
console.log(
  `open types    ${String([...flattened.orders.values()].filter((o) => o.open).length)} ` +
    `(xsd:any - never ordered)`,
);
console.log(`rank collisions ${String(collisions.length)}`);
for (const line of collisions) console.log(`  ${line}`);

console.log(`\nambiguous element names ${String(ambiguous.length)}`);
for (const a of ambiguous) {
  console.log(`  ${a.elementKey}`);
  for (const [typeKey, signature] of a.variants) {
    const sites = a.declaredIn.get(typeKey) ?? [];
    console.log(
      `    ${typeKey}  under ${sites.length > 3 ? `${String(sites.length)} types` : sites.join(', ')}`,
    );
    console.log(`      ${signature === '' ? '(no element children)' : signature}`);
  }
}

console.log(`\ncontext entries   ${String(context.entries.size)} (grandparent + parent)`);
for (const pair of [...context.entries.keys()].sort()) console.log(`  ${pair}`);
console.log(`unresolved pairs  ${String(context.unresolved.length)}`);
for (const line of context.unresolved) console.log(`  ${line}`);
if (context.unresolved.length > 0) {
  throw new Error('one more level of context is not enough; the lookup design needs revisiting');
}

const ranked = [...table.values()].reduce((n, ranks) => n + ranks.size, 0);
console.log(`\norderable parents ${String(table.size)}`);
console.log(`ranked children   ${String(ranked)}`);

// ------------------------------------------------------------------- encoding

/**
 * The table as one string.
 *
 * Parents are separated by `;`, and each parent reads `<parent>=<children>`. In
 * the child list `,` starts the next rank and `|` continues the current one, so
 * ranks are positional and never written down. Every name is prefixed with a
 * single digit indexing `NAMESPACES` - unambiguous because an XML name cannot
 * begin with a digit.
 *
 * A plain object literal would be several times the size and would cost a parse
 * of every entry at module load whether or not anything is ever inserted. This
 * is one string, decoded on first use.
 */
const namespaces: string[] = [];
const indexOf = (ns: string): number => {
  let i = namespaces.indexOf(ns);
  if (i < 0) i = namespaces.push(ns) - 1;
  if (i > 9) throw new Error('more than ten namespaces: the single-digit index no longer works');
  return i;
};
const shorten = (k: string): string => {
  const close = k.indexOf('}');
  return String(indexOf(k.slice(1, close))) + k.slice(close + 1);
};

function encodeRanks(ranks: ReadonlyMap<string, number>): string {
  const byRank = new Map<number, string[]>();
  for (const [child, rank] of ranks) {
    const bucket = byRank.get(rank) ?? [];
    bucket.push(shorten(child));
    byRank.set(rank, bucket);
  }
  return [...byRank.keys()]
    .sort((a, b) => a - b)
    .map((rank) => byRank.get(rank)!.sort().join('|'))
    .join(',');
}

function encode(entries: ReadonlyMap<string, ReadonlyMap<string, number>>, pair: boolean): string {
  return [...entries.keys()]
    .sort()
    .map((parent) => {
      const head = pair
        ? parent
            .split('|')
            .map((k) => shorten(k))
            .join('>')
        : shorten(parent);
      return `${head}=${encodeRanks(entries.get(parent)!)}`;
    })
    .join(';');
}

const data = encode(table, false);
const contextData = encode(context.entries, true);
console.log(`namespaces        ${String(namespaces.length)}`);
console.log(`encoded table     ${String(data.length)} characters`);
console.log(`encoded context   ${String(contextData.length)} characters`);

const banner = `/* eslint-disable */
// Generated by tools/schema-codegen/generate.ts. Do not edit.
//
// Source: the ECMA-376 Transitional schemas, Part 4, 5th edition (December
// 2016), as published by Ecma International. The schemas themselves are not
// redistributed here; these are the child-ordering facts extracted from them.
//
${sources.map((s) => `//   ${s.file.padEnd(22)} sha256 ${s.sha256}`).join('\n')}
//
// Encoding: parents separated by ";", each "<parent>=<children>". In a child
// list "," advances the rank and "|" holds it, so a rank is a position and is
// never written. Every name carries a one-digit index into NAMESPACES.
//
// CONTEXT_DATA is the same, keyed "<grandparent>><parent>", for the handful of
// names whose content model the parent alone does not determine.
`;

const body = `${banner}
export const NAMESPACES: readonly string[] = ${JSON.stringify(namespaces, null, 2)};

export const SCHEMA_SOURCES: readonly { readonly file: string; readonly sha256: string }[] = ${JSON.stringify(
  sources,
  null,
  2,
)};

/** ${String(table.size)} parents, ${String(ranked)} ranked children. */
export const ORDER_DATA =
  ${JSON.stringify(data)};

/** ${String(context.entries.size)} grandparent-qualified parents. */
export const CONTEXT_DATA =
  ${JSON.stringify(contextData)};
`;

if (write) {
  const path = resolve(out);
  writeFileSync(path, body, 'utf8');
  console.log(`\nwrote ${path} (${String(body.length)} characters)`);
} else {
  console.log('\n(dry run - pass --write to emit the file)');
}
