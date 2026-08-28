import { mkdtempSync, readFileSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { afterAll, describe, expect, it } from 'vitest';
import { readZip, type ZipEntry } from '../ground-truth/zip.ts';
import { buildPptx } from '../ground-truth/pptx.ts';
import { RECIPES, writeDeck, type DeckRecipe } from '../bench/deck.ts';
import { ZipStream } from '../bench/zip-stream.ts';
import { buildProbePackage } from './gen/package.ts';
import { PROBE_DECKS } from './gen/decks/index.ts';

/**
 * Every producer in this repository writes the bytes PowerPoint writes.
 *
 * There are three hand-written PresentationML builders here - the benchmark
 * generator, the sub-phase 0.7 ground-truth builder, and the corpus chassis -
 * and the obvious worry about three is that they drift into three different
 * notions of correct markup. Sharing code is the weak answer: it makes them
 * agree without making any of them right, and it does nothing about the fourth
 * one somebody writes next year.
 *
 * This is the strong answer. `corpus/ground-truth/powerpoint-conventions.json`
 * is a measurement of what PowerPoint 16.0.20326 actually wrote across the 38
 * XML parts of a deck it saved, taken in sub-phase 1.1's experiment E8. Not one
 * of the facts in it is in any schema. This file reads that measurement and
 * demands the same of us.
 *
 * Why it matters more than it looks: sub-phase 0.5's gate is "parse, serialize,
 * byte-identical for 100% of parts", and every fixture that gate runs against
 * comes out of one of these three builders. A convention we never emit is a
 * convention the gate never tests, and `C-LEX` exists precisely because two
 * producers that agree with each other prove nothing.
 *
 * The one convention deliberately *not* asserted is the self-closing spacing.
 * ADR 0005 measured 70,822 of 98,777 self-closing tags in the real-world corpus
 * written spaced, `<a:off ... />`; this PowerPoint build writes zero spaced out
 * of 1,305. Both forms exist in the wild and no producer available to this
 * project emits the dominant one, which is the roster's open problem rather
 * than a rule anything here can satisfy.
 */

const ROOT = resolve(fileURLToPath(import.meta.url), '../../..');
const MEASURED = JSON.parse(
  readFileSync(join(ROOT, 'corpus/ground-truth/powerpoint-conventions.json'), 'utf8'),
) as {
  lexical: {
    partsWithBom: number;
    declarations: string[];
    bytesAfterDeclaration: string[];
    selfClosingSpaced: number;
    singleQuotedAttributes: number;
    partsBrokenAcrossLines: number;
    entitiesUsed: string[];
  };
  archive: { firstEntry: string; hasDirectoryEntries: boolean };
};

const decoder = new TextDecoder();
const isXml = (name: string): boolean => name.endsWith('.xml') || name.endsWith('.rels');

// At module scope, not in `beforeAll`: `describe.each` below evaluates its
// table while the file is being collected, which is before any hook has run.
const directory = mkdtempSync(join(tmpdir(), 'pptx-studio-conventions-'));
afterAll(() => {
  rmSync(directory, { recursive: true, force: true });
});

/** Every producer, each reduced to the entries of one small package. */
function producers(): { name: string; entries: ZipEntry[] }[] {
  const benchRecipe: DeckRecipe = {
    ...(RECIPES['small'] as DeckRecipe),
    id: 'conventions',
    title: 'Conventions probe',
    slides: 2,
    imageBytes: 0,
    shapesPerSlide: 2,
    embedFont: false,
  };
  const benchPath = join(directory, 'bench.pptx');
  const stream = new ZipStream(benchPath);
  writeDeck(stream, benchRecipe, null);

  return [
    {
      name: 'tools/corpus/gen (the corpus chassis)',
      entries: readZip(buildProbePackage(PROBE_DECKS[1]!.build()).bytes),
    },
    {
      name: 'tools/ground-truth/pptx.ts (sub-phase 0.7)',
      entries: readZip(
        buildPptx({
          slides: [
            '<p:sp><p:nvSpPr><p:cNvPr id="2" name="Probe"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>' +
              '<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm>' +
              '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>' +
              '<a:solidFill><a:srgbClr val="4472C4"/></a:solidFill></p:spPr></p:sp>',
          ],
        }),
      ),
    },
    {
      name: 'tools/bench/deck.ts (the benchmark generator)',
      entries: readZip(new Uint8Array(readFileSync(benchPath))),
    },
  ];
}

describe.each(producers().map((p) => [p.name, p] as const))('%s', (_name, producer) => {
  const xmlParts = producer.entries.filter((entry) => isXml(entry.name));

  it('writes no byte-order mark', () => {
    const withBom = xmlParts.filter(
      (part) => part.bytes[0] === 0xef && part.bytes[1] === 0xbb && part.bytes[2] === 0xbf,
    );
    expect(withBom.map((part) => part.name)).toEqual([]);
    expect(MEASURED.lexical.partsWithBom).toBe(0);
  });

  it('writes one declaration form, and CRLF after it', () => {
    const forms = new Set<string>();
    const after = new Set<string>();
    for (const part of xmlParts) {
      const body = decoder.decode(part.bytes);
      const end = body.indexOf('?>');
      expect(end, part.name).toBeGreaterThan(0);
      forms.add(body.slice(0, end + 2));
      after.add(JSON.stringify(body.slice(end + 2, end + 4)));
    }
    expect([...forms]).toEqual(MEASURED.lexical.declarations);
    expect([...after]).toEqual(MEASURED.lexical.bytesAfterDeclaration);
  });

  it('never single-quotes an attribute', () => {
    let singleQuoted = 0;
    for (const part of xmlParts) {
      singleQuoted += (decoder.decode(part.bytes).match(/\s[a-zA-Z:]+='/g) ?? []).length;
    }
    expect(singleQuoted).toBe(MEASURED.lexical.singleQuotedAttributes);
  });

  it('never breaks a part across lines', () => {
    // Exactly one CRLF per part - the one after the declaration. More than that
    // is a pretty-printed document, and every byte of that whitespace is
    // significant to a gate that re-emits parts unchanged.
    const broken: string[] = [];
    for (const part of xmlParts) {
      const crlf = (decoder.decode(part.bytes).match(/\r\n/g) ?? []).length;
      if (crlf !== 1) broken.push(part.name + ' (' + String(crlf) + ')');
    }
    expect(broken).toEqual([]);
    expect(MEASURED.lexical.partsBrokenAcrossLines).toBe(0);
  });

  it('uses only the three entity references PowerPoint uses', () => {
    const used = new Set<string>();
    for (const part of xmlParts) {
      for (const match of decoder.decode(part.bytes).matchAll(/&[a-zA-Z#][a-zA-Z0-9]*;/g)) {
        used.add(match[0]);
      }
    }
    for (const entity of used) expect(MEASURED.lexical.entitiesUsed).toContain(entity);
  });

  it('never writes xml:space on an a:t', () => {
    // Measured, and against the plan's appendix, which carried Word's rule.
    // PowerPoint stores edge whitespace in an `a:t` with no attribute at all,
    // so writing one would be synthesizing markup we did not read.
    for (const part of xmlParts) {
      expect(decoder.decode(part.bytes), part.name).not.toMatch(/<a:t[^>]*xml:space/);
    }
  });

  it('opens with [Content_Types].xml and has no directory entries', () => {
    expect(producer.entries[0]?.name).toBe(MEASURED.archive.firstEntry);
    expect(producer.entries.some((entry) => entry.name.endsWith('/'))).toBe(
      MEASURED.archive.hasDirectoryEntries,
    );
  });
});
