#!/usr/bin/env node
/**
 * Measure what PowerPoint's serializer writes, from the two decks
 * `probe-conventions.ps1` produced.
 *
 * ```
 * powershell -File tools/corpus/authored/probe-conventions.ps1 -Out <dir>
 * node tools/corpus/authored/analyse-conventions.ts <dir> \
 *   --json corpus/ground-truth/powerpoint-conventions.json
 * ```
 *
 * What lands in the repository is the measurement, never the decks: facts
 * about Microsoft-authored files rather than the files themselves, the same
 * posture as `eot-headers.json`.
 *
 * Why it exists at all. Sub-phase 0.5's gate is that a part nobody edited
 * re-emits byte for byte, and that is a statement about lexical conventions -
 * quote characters, escaping, self-closing form, line endings, ZIP flags. Not
 * one of them is in the schema. They can only be measured from a file the
 * application itself wrote, and until now nothing in the repository recorded
 * what those conventions actually are.
 */

import { createHash } from 'node:crypto';
import { readFileSync, writeFileSync } from 'node:fs';
import { join, resolve } from 'node:path';
// The repository's own Node-side reader rather than fflate: `tools/` has no
// node_modules of its own, and this has to run from a bare clone.
import { readZip } from '../../ground-truth/zip.ts';

const args = process.argv.slice(2);
const dir = args[0];
if (dir === undefined) {
  console.error('usage: analyse-conventions.ts <dir> [--json <out>]');
  process.exit(2);
}
const jsonAt = args.indexOf('--json');
const jsonOut = jsonAt === -1 ? undefined : args[jsonAt + 1];

const decoder = new TextDecoder();

type Parts = Map<string, Uint8Array>;

interface Deck {
  readonly bytes: Uint8Array;
  readonly parts: Parts;
}

function open(name: string): Deck {
  const bytes = new Uint8Array(readFileSync(join(dir as string, name)));
  const parts: Parts = new Map();
  for (const item of readZip(bytes)) parts.set(item.name, item.bytes);
  return { bytes, parts };
}

function text(parts: Parts, name: string): string {
  const raw = parts.get(name);
  return raw === undefined ? '' : decoder.decode(raw);
}

const isXml = (name: string): boolean => name.endsWith('.xml') || name.endsWith('.rels');

/** Lexical facts, measured across every XML part rather than one. */
function lexical(parts: Parts): Record<string, unknown> {
  const names = [...parts.keys()].filter(isXml).sort();
  let bom = 0;
  let tight = 0;
  let spaced = 0;
  let singleQuoted = 0;
  let crlfTotal = 0;
  let brokenAcrossLines = 0;
  const declarations = new Set<string>();
  const afterDeclaration = new Set<string>();
  const entities = new Set<string>();

  for (const name of names) {
    const raw = parts.get(name);
    if (raw === undefined) continue;
    if (raw[0] === 0xef && raw[1] === 0xbb && raw[2] === 0xbf) bom += 1;
    const body = decoder.decode(raw);

    const end = body.indexOf('?>');
    if (end !== -1) {
      declarations.add(body.slice(0, end + 2));
      afterDeclaration.add(JSON.stringify(body.slice(end + 2, end + 4)));
    }
    tight += (body.match(/[^\s]\/>/g) ?? []).length;
    spaced += (body.match(/\s\/>/g) ?? []).length;
    singleQuoted += (body.match(/\s[a-zA-Z:]+='/g) ?? []).length;

    const crlf = (body.match(/\r\n/g) ?? []).length;
    crlfTotal += crlf;
    // One CRLF is the break after the declaration. More than one means the
    // document is broken across lines, which is what "pretty-printed" means
    // for a file whose whitespace has to be reproduced exactly.
    if (crlf > 1) brokenAcrossLines += 1;

    for (const match of body.matchAll(/&[a-zA-Z#][a-zA-Z0-9]*;/g)) entities.add(match[0]);
  }

  return {
    xmlParts: names.length,
    partsWithBom: bom,
    declarations: [...declarations],
    bytesAfterDeclaration: [...afterDeclaration],
    selfClosingTight: tight,
    selfClosingSpaced: spaced,
    singleQuotedAttributes: singleQuoted,
    partsBrokenAcrossLines: brokenAcrossLines,
    crlfTotal,
    entitiesUsed: [...entities].sort(),
  };
}

/** ZIP facts, from the first local file header and the entry order. */
function archive(deck: Deck): Record<string, unknown> {
  const { bytes, parts } = deck;
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const nameLength = view.getUint16(26, true);
  const extraLength = view.getUint16(28, true);
  const extra = bytes.slice(30 + nameLength, 30 + nameLength + extraLength);
  const flag = view.getUint16(6, true);
  const hex = (value: number): string => '0x' + value.toString(16).padStart(4, '0');

  return {
    entries: parts.size,
    firstEntry: decoder.decode(bytes.slice(30, 30 + nameLength)),
    entryOrder: [...parts.keys()].slice(0, 10),
    generalPurposeFlag: hex(flag),
    utf8FlagSet: (flag & 0x0800) !== 0,
    compressionMethod: view.getUint16(8, true),
    dosTime: hex(view.getUint16(10, true)),
    dosDate: hex(view.getUint16(12, true)),
    firstExtraFieldLength: extraLength,
    firstExtraFieldId:
      extra.byteLength >= 4
        ? hex(new DataView(extra.buffer, extra.byteOffset).getUint16(0, true))
        : null,
    hasDirectoryEntries: [...parts.keys()].some((name) => name.endsWith('/')),
  };
}

const plain = open('plain.pptx');
const ole = open('ole.pptx');

const slide = text(ole.parts, 'ppt/slides/slide1.xml');
const oleRels = text(ole.parts, 'ppt/slides/_rels/slide1.xml.rels');
const plainSlide = text(plain.parts, 'ppt/slides/slide1.xml');

const runs = [...plainSlide.matchAll(/<a:t([^>]*)>([\s\S]*?)<\/a:t>/g)].map((match) => ({
  attributes: (match[1] ?? '').trim(),
  text: match[2] ?? '',
}));

const oleAttributes = (/<p:oleObj([^>]*)>/.exec(slide)?.[1] ?? '')
  .trim()
  .split(/\s+/)
  .map((pair) => pair.split('=')[0])
  .filter((name): name is string => name !== undefined && name !== '');

const report = {
  measuredOn: {
    application: 'Microsoft PowerPoint',
    version: '16.0',
    build: '20326.20100',
    platform: 'Windows 11 Enterprise 10.0.26200',
  },
  lexical: lexical(plain.parts),
  archive: archive(plain),
  text: {
    $comment:
      'The second run was authored with two leading and two trailing spaces. ' +
      'PowerPoint preserves them and writes NO xml:space attribute, unlike ' +
      'WordprocessingML, where a w:t with edge whitespace carries one.',
    runs,
    writesXmlSpacePreserve: /<a:t[^>]*xml:space/.test(plainSlide),
  },
  ole: {
    $comment:
      'Experiment E6. Build 2205 stopped writing the VML fallback for OLE. ' +
      'This build still writes the mc:AlternateContent wrapper, and its ' +
      'mc:Choice still requires the VML prefix, but there is no vmlDrawing ' +
      'part and no @spid for it to point at.',
    writesAlternateContent: slide.includes('mc:AlternateContent'),
    choiceRequires: /<mc:Choice[^>]*Requires="([^"]+)"/.exec(slide)?.[1] ?? null,
    choiceDeclaresNamespaceOnItself: /<mc:Choice[^>]*xmlns:v=/.test(slide),
    writesVmlDrawingPart: [...ole.parts.keys()].some((name) => /vmlDrawing/i.test(name)),
    writesSpid: /\bspid=/.test(slide),
    oleObjAttributes: oleAttributes,
    embeddingPart: [...ole.parts.keys()].find((name) => name.includes('embeddings')) ?? null,
    embeddingRelationshipType: /Type="([^"]*relationships\/package)"/.exec(oleRels)?.[1] ?? null,
    previewPart: [...ole.parts.keys()].find((name) => name.startsWith('ppt/media/')) ?? null,
    previewInFallbackOnly:
      slide.includes('<mc:Fallback>') && slide.indexOf('<p:pic>') > slide.indexOf('<mc:Fallback>'),
    writesPackageThumbnail: [...ole.parts.keys()].some((name) =>
      name.startsWith('docProps/thumbnail'),
    ),
  },
  sourceDecks: {
    $comment: 'Not committed. Hashes so a re-run can be checked against this measurement.',
    'plain.pptx': createHash('sha256').update(plain.bytes).digest('hex'),
    'ole.pptx': createHash('sha256').update(ole.bytes).digest('hex'),
  },
};

const serialized = JSON.stringify(report, null, 2) + '\n';
if (jsonOut === undefined) {
  console.log(serialized);
} else {
  writeFileSync(resolve(jsonOut), serialized);
  console.log('wrote ' + resolve(jsonOut));
}
