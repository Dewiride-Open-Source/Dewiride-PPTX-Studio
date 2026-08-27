/**
 * Experiment A - read back everything a `.pptx` says about its embedded fonts.
 *
 * ```
 * node tools/ground-truth/dump-fntdata.ts <deck.pptx> [<deck.pptx> ...] [--json <file>]
 * ```
 *
 * Reports, per deck: the `fntdata` content-type Default entry, the
 * `embedTrueTypeFonts` attribute, every `p:embeddedFont` with its attributes and
 * relationship targets, and for every `ppt/fonts/*.fntdata` the full EOT header
 * plus what the wrapped SFNT says about itself.
 *
 * The question the whole experiment exists to answer is one bit: is
 * `TTEMBED_TTCOMPRESSED` (0x4) set in the flags word at 0x0C? If it is, the
 * payload is MicroType Express and there is no decoder for it in JavaScript or
 * WebAssembly, which would put roughly four weeks of Phase 8 on a different
 * footing.
 */

import { readFileSync, writeFileSync } from 'node:fs';
import { basename } from 'node:path';
import { entry, readZip, type ZipEntry } from './zip.ts';
import { eotFontData, flagNames, readEot, EOT_FLAG, type EotHeader } from './eot.ts';
import {
  NAME_ID,
  numGlyphs,
  readCoverage,
  readHead,
  readNames,
  readOs2,
  readSfnt,
  SFNT_TAG_OTTO,
  SFNT_VERSION_TRUETYPE,
} from './sfnt.ts';

const text = (bytes: Uint8Array | undefined): string =>
  bytes === undefined ? '' : new TextDecoder().decode(bytes);

function attr(tag: string, name: string): string | undefined {
  const m = new RegExp(`\\s${name}="([^"]*)"`).exec(tag);
  return m?.[1];
}

function hexdump(bytes: Uint8Array, length: number): string[] {
  const lines: string[] = [];
  for (let at = 0; at < Math.min(length, bytes.length); at += 16) {
    const slice = bytes.subarray(at, Math.min(at + 16, bytes.length));
    const hex = [...slice].map((b) => b.toString(16).padStart(2, '0')).join(' ');
    const ascii = [...slice]
      .map((b) => (b >= 0x20 && b < 0x7f ? String.fromCharCode(b) : '.'))
      .join('');
    lines.push(`  ${at.toString(16).padStart(4, '0')}  ${hex.padEnd(47)}  ${ascii}`);
  }
  return lines;
}

interface FontReport {
  readonly part: string;
  readonly bytes: number;
  readonly eot: Omit<EotHeader, 'panose'> & { panose: string };
  readonly flagNames: string[];
  readonly compressed: boolean;
  readonly subset: boolean;
  readonly payload: {
    readonly bytes: number;
    readonly sfntVersion: string;
    readonly outlines: 'glyf' | 'CFF' | 'unknown';
    readonly tables: string[];
    readonly numGlyphs: number | undefined;
    readonly codePointsCovered: number;
    readonly nameFamily: string | undefined;
    readonly namePostScript: string | undefined;
    readonly os2Version: number | undefined;
    readonly os2FsType: number | undefined;
    readonly os2WeightClass: number | undefined;
    readonly headCheckSumAdjustment: number | undefined;
  } | null;
  readonly agreement: string[];
  readonly payloadHead: string;
  readonly payloadLooksLikeSfnt: boolean;
}

function reportFont(part: string, bytes: Uint8Array): FontReport {
  const header = readEot(bytes);
  const names = flagNames(header.flags);
  const compressed = (header.flags & EOT_FLAG.TTCOMPRESSED) !== 0;
  const subset = (header.flags & EOT_FLAG.SUBSET) !== 0;

  const agreement: string[] = [];
  let payload: FontReport['payload'] = null;

  // Peek at the payload whatever the flags claim. A flag is a claim about the
  // bytes, not the bytes themselves, and "the writer sets TTCOMPRESSED but ships
  // raw SFNT" is exactly the sort of thing that would be cheap to believe and
  // expensive to be wrong about.
  const rawPayload = eotFontData(bytes, header);
  const payloadHead = [...rawPayload.subarray(0, 16)]
    .map((b) => b.toString(16).padStart(2, '0'))
    .join(' ');
  const looksLikeSfnt =
    rawPayload.length >= 4 &&
    ((rawPayload[0] === 0x00 &&
      rawPayload[1] === 0x01 &&
      rawPayload[2] === 0x00 &&
      rawPayload[3] === 0x00) ||
      (rawPayload[0] === 0x4f &&
        rawPayload[1] === 0x54 &&
        rawPayload[2] === 0x54 &&
        rawPayload[3] === 0x4f) ||
      (rawPayload[0] === 0x74 &&
        rawPayload[1] === 0x72 &&
        rawPayload[2] === 0x75 &&
        rawPayload[3] === 0x65));

  if (!compressed) {
    const raw = rawPayload;
    const font = readSfnt(raw);
    const os2 = readOs2(font);
    const head = readHead(font);
    const nameTable = readNames(font);
    payload = {
      bytes: raw.length,
      sfntVersion: '0x' + font.version.toString(16).padStart(8, '0'),
      outlines:
        font.version === SFNT_VERSION_TRUETYPE
          ? 'glyf'
          : font.version === SFNT_TAG_OTTO
            ? 'CFF'
            : 'unknown',
      tables: font.tables.map((t) => t.tag).sort(),
      numGlyphs: numGlyphs(font),
      codePointsCovered: readCoverage(font).size,
      nameFamily: nameTable.get(NAME_ID.family),
      namePostScript: nameTable.get(NAME_ID.postScript),
      os2Version: os2?.version,
      os2FsType: os2?.fsType,
      os2WeightClass: os2?.usWeightClass,
      headCheckSumAdjustment: head?.checkSumAdjustment,
    };

    // The EOT header duplicates four things the SFNT already knows. A reader
    // that trusts the header is much simpler than one that parses the font, so
    // it matters a great deal whether the copy is faithful.
    const check = (what: string, inHeader: unknown, inFont: unknown): void => {
      agreement.push(
        `${inHeader === inFont ? 'agree  ' : 'DIFFER '} ${what}: eot=${String(inHeader)} sfnt=${String(inFont)}`,
      );
    };
    check('familyName', header.familyName, payload.nameFamily);
    check('fsType', header.fsType, payload.os2FsType);
    check('weight', header.weight, payload.os2WeightClass);
    check('checkSumAdjustment', header.checkSumAdjustment, payload.headCheckSumAdjustment);
  }

  const { panose, ...rest } = header;
  return {
    part,
    bytes: bytes.length,
    eot: { ...rest, panose: [...panose].map((b) => b.toString(16).padStart(2, '0')).join(' ') },
    flagNames: names,
    compressed,
    subset,
    payload,
    agreement,
    payloadHead,
    payloadLooksLikeSfnt: looksLikeSfnt,
  };
}

function reportDeck(path: string): unknown {
  const entries: ZipEntry[] = readZip(new Uint8Array(readFileSync(path)));
  const name = basename(path);
  console.log('='.repeat(78));
  console.log(name);
  console.log('='.repeat(78));

  const contentTypes = text(entry(entries, '[Content_Types].xml'));
  const fntDefault = /<Default[^>]*Extension="fntdata"[^>]*\/>/.exec(contentTypes)?.[0];
  console.log(`\n[Content_Types].xml fntdata Default : ${fntDefault ?? 'ABSENT'}`);

  const presentation = text(entry(entries, 'ppt/presentation.xml'));
  const embedAttr = attr(
    /<p:presentation[^>]*>/.exec(presentation)?.[0] ?? '',
    'embedTrueTypeFonts',
  );
  console.log(`p:presentation/@embedTrueTypeFonts  : ${embedAttr ?? 'ABSENT'}`);

  const rels = text(entry(entries, 'ppt/_rels/presentation.xml.rels'));
  const fontRels = new Map<string, string>();
  for (const m of rels.matchAll(/<Relationship\b[^>]*\/>/g)) {
    const tag = m[0];
    const type = attr(tag, 'Type') ?? '';
    if (!type.endsWith('/font')) continue;
    fontRels.set(attr(tag, 'Id') ?? '', attr(tag, 'Target') ?? '');
  }
  console.log(`font relationships                  : ${String(fontRels.size)}`);
  if (fontRels.size > 0) {
    const [, sample] = [...fontRels][0]!;
    const type = /Type="([^"]*\/font)"/.exec(rels)?.[1];
    console.log(`  relationship type                 : ${type ?? '?'}`);
    console.log(`  example target                    : ${sample}`);
  }

  const embedded = [...presentation.matchAll(/<p:embeddedFont>[\s\S]*?<\/p:embeddedFont>/g)];
  console.log(`p:embeddedFont entries              : ${String(embedded.length)}`);
  const embeddedFonts = embedded.map((m) => {
    const block = m[0];
    const font = /<p:font\b[^>]*\/>/.exec(block)?.[0] ?? '';
    const slots: Record<string, string> = {};
    for (const s of block.matchAll(/<p:(regular|bold|italic|boldItalic)\b[^>]*\/>/g)) {
      slots[s[1]!] = attr(s[0], 'r:id') ?? '';
    }
    return {
      typeface: attr(font, 'typeface'),
      panose: attr(font, 'panose'),
      pitchFamily: attr(font, 'pitchFamily'),
      charset: attr(font, 'charset'),
      slots,
    };
  });
  for (const f of embeddedFonts) {
    console.log(
      `  ${String(f.typeface).padEnd(18)} panose=${String(f.panose)} pitchFamily=${String(f.pitchFamily)} ` +
        `charset=${String(f.charset)} slots=${Object.keys(f.slots).join(',')}`,
    );
  }

  const fonts = entries
    .filter((e) => e.name.startsWith('ppt/fonts/'))
    .sort((a, b) => a.name.localeCompare(b.name));
  console.log(`\nppt/fonts/* parts                   : ${String(fonts.length)}\n`);

  const reports: FontReport[] = [];
  for (const font of fonts) {
    const report = reportFont(font.name, font.bytes);
    reports.push(report);
    console.log(`--- ${font.name}  (${String(font.bytes.length)} bytes) ---`);
    console.log(hexdump(font.bytes, 48).join('\n'));
    console.log(
      `  version=0x${report.eot.version.toString(16).padStart(8, '0')}` +
        `  flags=0x${report.eot.flags.toString(16).padStart(8, '0')} [${report.flagNames.join(' ') || 'none'}]`,
    );
    console.log(
      `  eotSize=${String(report.eot.eotSize)} fontDataSize=${String(report.eot.fontDataSize)}` +
        `  headerEnd=${String(report.eot.headerEnd)}` +
        `  (eotSize-fontDataSize=${String(report.eot.eotSize - report.eot.fontDataSize)})`,
    );
    console.log(
      `  charset=${String(report.eot.charset)} italic=${String(report.eot.italic)}` +
        ` weight=${String(report.eot.weight)} fsType=0x${report.eot.fsType.toString(16)}` +
        ` panose=[${report.eot.panose}]`,
    );
    console.log(
      `  family=${JSON.stringify(report.eot.familyName)} style=${JSON.stringify(report.eot.styleName)}` +
        ` version=${JSON.stringify(report.eot.versionName)} full=${JSON.stringify(report.eot.fullName)}`,
    );
    if (report.payload !== null) {
      const p = report.payload;
      console.log(
        `  payload: ${p.outlines} ${p.sfntVersion} ${String(p.numGlyphs)} glyphs,` +
          ` ${String(p.codePointsCovered)} code points, ${String(p.tables.length)} tables`,
      );
      console.log(
        `  tables: ${p.tables.slice(0, 40).join(' ')}` +
          (p.tables.length > 40 ? ` ...(${String(p.tables.length)} total)` : ''),
      );
      for (const line of report.agreement) console.log(`  ${line}`);
    } else {
      console.log('  payload: NOT PARSED - TTEMBED_TTCOMPRESSED is set (MicroType Express)');
    }
    console.log(`  payload first 16 bytes: ${report.payloadHead}`);
    console.log(
      `  payload begins with an SFNT signature: ${report.payloadLooksLikeSfnt ? 'YES' : 'no'}`,
    );
    console.log('');
  }

  return {
    deck: name,
    contentTypeDefault: fntDefault ?? null,
    embedTrueTypeFonts: embedAttr ?? null,
    fontRelationshipCount: fontRels.size,
    embeddedFonts,
    fonts: reports,
  };
}

const argv = process.argv.slice(2);
const jsonAt = argv.indexOf('--json');
const jsonOut = jsonAt >= 0 ? argv[jsonAt + 1] : undefined;
const decks = argv.filter((_, i) => jsonAt < 0 || (i !== jsonAt && i !== jsonAt + 1));
if (decks.length === 0) throw new Error('usage: dump-fntdata.ts <deck.pptx> ... [--json <file>]');

const all = decks.map(reportDeck);

console.log('='.repeat(78));
console.log('THE ONE BIT THAT MATTERS');
console.log('='.repeat(78));
let anyCompressed = false;
for (const deck of all as { deck: string; fonts: FontReport[] }[]) {
  for (const font of deck.fonts) {
    if (font.compressed) anyCompressed = true;
    console.log(
      `${deck.deck.padEnd(20)} ${font.part.padEnd(24)} ` +
        `TTCOMPRESSED=${font.compressed ? 'SET' : 'clear'}  SUBSET=${font.subset ? 'SET' : 'clear'}`,
    );
  }
}
console.log(
  `\nMicroType Express present anywhere: ${anyCompressed ? 'YES - Phase 8 needs a decoder' : 'NO'}`,
);

if (jsonOut !== undefined) {
  // Header fields only. The `.fntdata` payloads are Microsoft's fonts and do
  // not go into this repository under any circumstances - what is committed is
  // a record of what PowerPoint wrote *about* them, which is a measurement.
  const fixture = {
    $comment:
      'What PowerPoint writes into ppt/fonts/*.fntdata. EOT header fields only - no font ' +
      'data. See docs/adr/0007-ground-truth.md. Generated by tools/ground-truth/dump-fntdata.ts.',
    decks: (
      all as {
        deck: string;
        contentTypeDefault: string | null;
        embedTrueTypeFonts: string | null;
        embeddedFonts: unknown[];
        fonts: FontReport[];
      }[]
    ).map((deck) => ({
      deck: deck.deck,
      contentTypeDefault: deck.contentTypeDefault,
      embedTrueTypeFonts: deck.embedTrueTypeFonts,
      embeddedFonts: deck.embeddedFonts,
      fonts: deck.fonts.map((f) => ({
        part: f.part,
        bytes: f.bytes,
        flags: '0x' + f.eot.flags.toString(16).padStart(8, '0'),
        flagNames: f.flagNames,
        version: '0x' + f.eot.version.toString(16).padStart(8, '0'),
        eotSize: f.eot.eotSize,
        fontDataSize: f.eot.fontDataSize,
        headerBytes: f.eot.eotSize - f.eot.fontDataSize,
        charset: f.eot.charset,
        italic: f.eot.italic,
        weight: f.eot.weight,
        fsType: f.eot.fsType,
        panose: f.eot.panose,
        // Kept exactly as read, NUL terminators included: that a family name
        // arrives as "Cambria " is the finding, not a defect to tidy.
        familyName: f.eot.familyName,
        styleName: f.eot.styleName,
        versionName: f.eot.versionName,
        fullName: f.eot.fullName,
        unicodeRange: f.eot.unicodeRange.map((r) => '0x' + r.toString(16).padStart(8, '0')),
        codePageRange: f.eot.codePageRange.map((r) => '0x' + r.toString(16).padStart(8, '0')),
        checkSumAdjustment: '0x' + f.eot.checkSumAdjustment.toString(16).padStart(8, '0'),
        payloadFirst16Bytes: f.payloadHead,
        payloadBeginsWithSfntSignature: f.payloadLooksLikeSfnt,
      })),
    })),
  };
  writeFileSync(jsonOut, JSON.stringify(fixture, null, 2) + '\n');
  console.log(`\nwrote ${jsonOut}`);
}
