/**
 * `C-LEX`: an inventory of every lexical form the committed corpus contains,
 * and which serializer wrote it.
 *
 * The rule this file exists to make checkable is the one sub-phase 1.1 stated
 * and could not enforce until there was more than one producer: **every lexical
 * convention is exercised by at least two decks from at least two distinct
 * serializers.** Feature coverage was never the binding constraint on whether
 * the round-trip gate means anything; this is. A gate that re-emits bytes it
 * only ever saw its own generator write is a proof of idempotence.
 *
 * ## Serializers, not files, and not tiers
 *
 * The unit is the code that decided the bytes. Counting *files* would score
 * `corpus/written/c01-opc-writer.pptx` as independent evidence about XML
 * lexical form when every byte inside every one of its parts is Microsoft's -
 * it is `b01-blank` passed through `PartStore` unchanged. Counting *tiers*
 * makes the same mistake with a directory name attached.
 *
 * So each deck has two serializers rather than one, and they are declared per
 * collection in the manifest envelope's `serializers` (`C018`):
 *
 * | collection | `xml`                          | `container`                 |
 * | ---------- | ------------------------------ | --------------------------- |
 * | `decks`    | `tools/corpus/gen`             | `tools/ground-truth/zip.ts` |
 * | `authored` | `Microsoft PowerPoint 16.0…`   | `Microsoft PowerPoint 16.0…` |
 * | `written`  | `Microsoft PowerPoint 16.0…`   | `packages/opc`              |
 *
 * Tier C's XML serializer is spelled identically to Tier B's on purpose. The
 * collision is the mechanism: two collections that share a serializer count
 * once, so the caveat is enforced rather than written in a paragraph and
 * trusted.
 *
 * `tools/corpus/gen` delegates its container to `tools/ground-truth/zip.ts` and
 * the envelope says so, because a ZIP header written by `writeZip` is the same
 * evidence whichever generator called it.
 *
 * ## Why this scans bytes rather than using `@pptx-studio/xml`
 *
 * Deliberately, and it is the whole point. `packages/xml` records the quote
 * character, the self-closing form and the whitespace before `>` because
 * ADR 0005 decided it must; a rule that asked the tokenizer what forms a file
 * contains would agree with the tokenizer by construction and could never catch
 * it being wrong. This scanner is an independent reading of the same bytes.
 *
 * It is a lexer and not a regex sweep for the same reason a regex would be
 * wrong: `/>` occurs inside attribute values, `&gt;` occurs in text, and a
 * comment may contain anything at all. Forms are recorded per token, with the
 * scanner knowing which construct it is inside.
 */

import { inflateRawSync } from 'node:zlib';

// ------------------------------------------------------------------- headers

/**
 * One ZIP entry, with every central-directory field `C-LEX` reads.
 *
 * `tools/ground-truth/zip.ts` surfaces `flags` and the local extra field and
 * stops there, which is right for a reader: `versionMadeBy` and the DOS
 * timestamp are fields nothing acts on. They are exactly the fields a
 * divergence hides in, so this reads the central directory itself.
 */
export interface ZipHeader {
  readonly name: string;
  readonly versionMadeBy: number;
  readonly versionNeeded: number;
  readonly flags: number;
  readonly method: number;
  readonly dosTime: number;
  readonly dosDate: number;
  readonly localExtraLength: number;
  readonly localExtraId: number | null;
  readonly centralExtraLength: number;
  readonly commentLength: number;
  /** Inflated, whatever the method. */
  readonly bytes: Uint8Array;
}

const EOCD_SIG = 0x06054b50;
const CENTRAL_SIG = 0x02014b50;
const EOCD64_LOCATOR_SIG = 0x07064b50;

export interface Archive {
  readonly entries: readonly ZipHeader[];
  readonly hasZip64Locator: boolean;
}

/** Every entry of an archive, in central-directory order, headers included. */
export function readHeaders(archive: Uint8Array): Archive {
  const view = new DataView(archive.buffer, archive.byteOffset, archive.byteLength);

  let eocd = -1;
  for (let at = archive.length - 22; at >= 0; at--) {
    if (view.getUint32(at, true) === EOCD_SIG) {
      eocd = at;
      break;
    }
  }
  if (eocd < 0) throw new Error('not a zip archive: no end-of-central-directory record');

  const count = view.getUint16(eocd + 10, true);
  let offset = view.getUint32(eocd + 16, true);

  const entries: ZipHeader[] = [];
  for (let i = 0; i < count; i++) {
    if (view.getUint32(offset, true) !== CENTRAL_SIG) {
      throw new Error('central directory entry ' + String(i) + ' has a bad signature');
    }
    const nameLength = view.getUint16(offset + 28, true);
    const extraLength = view.getUint16(offset + 30, true);
    const commentLength = view.getUint16(offset + 32, true);
    const local = view.getUint32(offset + 42, true);
    const method = view.getUint16(offset + 10, true);
    const compressedSize = view.getUint32(offset + 20, true);

    const localNameLength = view.getUint16(local + 26, true);
    const localExtraLength = view.getUint16(local + 28, true);
    const start = local + 30 + localNameLength + localExtraLength;
    const raw = archive.subarray(start, start + compressedSize);

    entries.push({
      name: new TextDecoder().decode(archive.subarray(offset + 46, offset + 46 + nameLength)),
      versionMadeBy: view.getUint16(offset + 4, true),
      versionNeeded: view.getUint16(offset + 6, true),
      flags: view.getUint16(offset + 8, true),
      method,
      dosTime: view.getUint16(offset + 12, true),
      dosDate: view.getUint16(offset + 14, true),
      localExtraLength,
      localExtraId:
        localExtraLength >= 2 ? view.getUint16(local + 30 + localNameLength, true) : null,
      centralExtraLength: extraLength,
      commentLength,
      bytes: method === 0 ? raw : new Uint8Array(inflateRawSync(raw)),
    });
    offset += 46 + nameLength + extraLength + commentLength;
  }

  let hasZip64Locator = false;
  for (let at = eocd - 20; at >= 0 && at > eocd - 128; at--) {
    if (view.getUint32(at, true) === EOCD64_LOCATOR_SIG) {
      hasZip64Locator = true;
      break;
    }
  }

  return { entries, hasZip64Locator };
}

// ---------------------------------------------------------------- the scanner

const isXmlPart = (name: string): boolean => name.endsWith('.xml') || name.endsWith('.rels');

const WS = new Set([' ', '\t', '\r', '\n']);

/** What separates the declaration from the root element, named. */
const GAPS = new Map([
  ['\r\n', 'crlf'],
  ['', 'none'],
  ['\n', 'lf'],
]);

/**
 * One pseudo-attribute of the XML declaration, reduced to a named form.
 *
 * `known` lists the spellings that get their own form; anything else is
 * `other`, and a declaration that omits the pseudo-attribute is `omitted`.
 * `encoding` and `standalone` are both optional in production [23], and both
 * are omitted somewhere in the 2834 parts ADR 0004 measured.
 */
function pseudoAttribute(declaration: string, name: string, known: readonly string[]): string {
  const found = new RegExp('\\b' + name + '\\s*=\\s*(?:"([^"]*)"|\'([^\']*)\')').exec(declaration);
  const value = found?.[1] ?? found?.[2];
  if (value === undefined) return 'omitted';
  return known.includes(value) ? value : 'other';
}

/**
 * Every lexical form one XML part exercises.
 *
 * A hand-written lexer over the text, not a parse: it never builds a tree, has
 * no opinion about well-formedness beyond what it must track to stay oriented,
 * and refuses nothing. A part this cannot make sense of contributes the forms
 * it did recognise before it lost the thread, which is the right failure for an
 * inventory - the alternative is a rule that goes quiet on exactly the file
 * that is strange enough to be worth cataloguing.
 */
export function scanXmlPart(bytes: Uint8Array): Set<string> {
  const forms = new Set<string>();
  const put = (dimension: string, form: string): void => {
    forms.add(dimension + '=' + form);
  };

  const hasBom = bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf;
  put('xml.bom', hasBom ? 'present' : 'absent');

  const text = new TextDecoder().decode(hasBom ? bytes.subarray(3) : bytes);

  let at = 0;
  if (text.startsWith('<?xml')) {
    // Decomposed rather than compared whole against the one declaration Office
    // writes, because ADR 0004's real-world table counts *three* declarations
    // across 2834 parts and the difference between them is one pseudo-attribute
    // at a time: `utf-8` for `UTF-8`, and `standalone` present or omitted. An
    // office/other split would report "we have one of three" as "we have one of
    // two" and hide which two are missing.
    const end = text.indexOf('?>');
    const declaration = end < 0 ? text : text.slice(0, end + 2);
    put('xml.declaration', 'present');

    put('xml.declarationEncoding', pseudoAttribute(declaration, 'encoding', ['UTF-8', 'utf-8']));
    put('xml.declarationStandalone', pseudoAttribute(declaration, 'standalone', ['yes', 'no']));

    at = end < 0 ? text.length : end + 2;

    let after = 0;
    while (at + after < text.length && WS.has(text[at + after]!)) after += 1;
    put('xml.afterDeclaration', GAPS.get(text.slice(at, at + after)) ?? 'other');
  } else {
    put('xml.declaration', 'absent');
  }

  const crlf = (text.match(/\r\n/g) ?? []).length;
  const lf = (text.match(/\n/g) ?? []).length;
  if (lf === 0) put('xml.lineBreaks', 'one-line');
  else if (crlf === 1 && lf === 1) put('xml.lineBreaks', 'declaration-break');
  else put('xml.lineBreaks', 'multi-line');

  /** Named and numeric references, wherever they are allowed to appear. */
  const references = (slice: string): void => {
    for (const match of slice.matchAll(/&(#x[0-9a-fA-F]+|#\d+|[A-Za-z][A-Za-z0-9]*);/g)) {
      const body = match[1]!;
      if (body.startsWith('#x')) put('xml.entity', 'hex-lower');
      else if (body.startsWith('#')) put('xml.entity', 'decimal');
      else put('xml.entity', body);
    }
    // `&#X41;` is **not** a character reference: production [66] spells `'&#x'`
    // as a terminal in a case-sensitive grammar, so an upper-case `X` is
    // ill-formed and `packages/xml` refuses it (ADR 0004). It is matched
    // separately, and never by the pattern above, because a deck that grew one
    // would be a well-formedness failure smuggled into the corpus rather than a
    // convention anything should cover. `C-LEX` declares it with no producer so
    // that finding one fails loudly.
    if (/&#X[0-9a-fA-F]+;/.test(slice)) put('xml.entity', 'hex-upper-ill-formed');
  };

  /** Whether the tag that just closed had no content before its end tag. */
  let openedAt: string | null = null;

  while (at < text.length) {
    const next = text.indexOf('<', at);
    if (next < 0) {
      references(text.slice(at));
      break;
    }
    if (next > at) {
      const between = text.slice(at, next);
      references(between);
      if (between.trim() !== '') openedAt = null;
    }

    if (text.startsWith('<!--', next)) {
      put('xml.markup', 'comment');
      const end = text.indexOf('-->', next);
      at = end < 0 ? text.length : end + 3;
      openedAt = null;
      continue;
    }
    if (text.startsWith('<![CDATA[', next)) {
      put('xml.markup', 'cdata');
      const end = text.indexOf(']]>', next);
      at = end < 0 ? text.length : end + 3;
      openedAt = null;
      continue;
    }
    if (text.startsWith('<!DOCTYPE', next)) {
      put('xml.markup', 'doctype');
      const end = text.indexOf('>', next);
      at = end < 0 ? text.length : end + 1;
      openedAt = null;
      continue;
    }
    if (text.startsWith('<?', next)) {
      put('xml.markup', 'processing-instruction');
      const end = text.indexOf('?>', next);
      at = end < 0 ? text.length : end + 2;
      openedAt = null;
      continue;
    }
    if (text.startsWith('</', next)) {
      let cursor = next + 2;
      const from = cursor;
      while (cursor < text.length && !WS.has(text[cursor]!) && text[cursor] !== '>') cursor += 1;
      const name = text.slice(from, cursor);
      const spaceFrom = cursor;
      while (cursor < text.length && WS.has(text[cursor]!)) cursor += 1;
      put('xml.tagClose', cursor > spaceFrom ? 'spaced-end' : 'tight');
      // Only the two ways of writing an element with no content are forms. An
      // element that *has* content is not evidence about either.
      if (openedAt === name) put('xml.emptyElement', 'paired');
      openedAt = null;
      at = cursor < text.length ? cursor + 1 : text.length;
      continue;
    }

    // A start tag. Walk its attributes so `/>`, quotes and whitespace are read
    // where they mean something rather than wherever they occur.
    let cursor = next + 1;
    const from = cursor;
    while (cursor < text.length && !WS.has(text[cursor]!) && !'>/'.includes(text[cursor]!)) {
      cursor += 1;
    }
    const name = text.slice(from, cursor);
    // Whitespace run immediately before whatever ends the tag. Reset each time
    // round, because only the last one - the run before `>` or `/>` - is the
    // form; the runs between attributes are just attribute separators.
    let trailing: number;

    for (;;) {
      trailing = 0;
      while (cursor < text.length && WS.has(text[cursor]!)) {
        trailing += 1;
        cursor += 1;
      }
      if (cursor >= text.length) break;
      const here = text[cursor]!;
      if (here === '/' || here === '>') break;

      const nameFrom = cursor;
      while (cursor < text.length && !WS.has(text[cursor]!) && !'=>/'.includes(text[cursor]!)) {
        cursor += 1;
      }
      if (cursor === nameFrom) {
        cursor += 1; // Not a name: step over it rather than spin.
        continue;
      }
      let spaced = false;
      while (cursor < text.length && WS.has(text[cursor]!)) {
        spaced = true;
        cursor += 1;
      }
      if (text[cursor] !== '=') continue;
      cursor += 1;
      while (cursor < text.length && WS.has(text[cursor]!)) {
        spaced = true;
        cursor += 1;
      }
      put('xml.attrEquals', spaced ? 'spaced' : 'tight');

      const quote = text[cursor];
      if (quote !== '"' && quote !== "'") continue;
      put('xml.attrQuote', quote === '"' ? 'double' : 'single');
      const valueFrom = cursor + 1;
      const valueEnd = text.indexOf(quote, valueFrom);
      references(text.slice(valueFrom, valueEnd < 0 ? text.length : valueEnd));
      cursor = valueEnd < 0 ? text.length : valueEnd + 1;
    }

    if (text[cursor] === '/') {
      put('xml.selfClosing', trailing > 0 ? 'spaced' : 'tight');
      put('xml.emptyElement', 'self-closed');
      openedAt = null;
      cursor += 1;
      at = cursor < text.length ? cursor + 1 : text.length;
      continue;
    }
    if (trailing > 0) put('xml.tagClose', 'spaced-start');
    openedAt = name;
    at = cursor < text.length ? cursor + 1 : text.length;
  }

  return forms;
}

/** Every lexical form one package's ZIP headers exercise. */
export function scanArchive(archive: Archive): Set<string> {
  const forms = new Set<string>();
  const put = (dimension: string, form: string): void => {
    forms.add(dimension + '=' + form);
  };

  put(
    'zip.firstEntry',
    archive.entries[0]?.name === '[Content_Types].xml' ? 'content-types' : 'other',
  );
  put('zip.zip64', archive.hasZip64Locator ? 'present' : 'absent');
  put(
    'zip.directoryEntries',
    archive.entries.some((entry) => entry.name.endsWith('/')) ? 'present' : 'absent',
  );

  for (const entry of archive.entries) {
    put('zip.versionMadeBy', String(entry.versionMadeBy));
    put('zip.versionNeeded', String(entry.versionNeeded));
    put('zip.method', entry.method === 0 ? 'stored' : entry.method === 8 ? 'deflate' : 'other');
    put(
      'zip.dosDateTime',
      entry.dosDate === 0x0021 && entry.dosTime === 0x0000 ? '1980-01-01T00:00' : 'other',
    );
    put(
      'zip.localExtra',
      entry.localExtraLength === 0
        ? 'none'
        : entry.localExtraId === 0xa220
          ? 'growth-hint'
          : 'other',
    );
    put('zip.centralExtra', entry.centralExtraLength === 0 ? 'none' : 'present');
    put('zip.entryComment', entry.commentLength === 0 ? 'none' : 'present');
    put('zip.dataDescriptor', (entry.flags & 0x0008) === 0 ? 'absent' : 'present');
    put('zip.utf8NameFlag', (entry.flags & 0x0800) === 0 ? 'absent' : 'present');

    // Bits 1 and 2 are the deflate level hint and mean nothing to a
    // decompressor. On a stored entry they mean nothing at all, so reading them
    // there would invent a convention: PowerPoint writes 0x0006 on every
    // deflated entry and 0x0000 on every stored one, and that is one habit and
    // not two.
    if (entry.method === 8) {
      const level = (entry.flags >> 1) & 0b11;
      put(
        'zip.deflateLevelHint',
        (['normal', 'maximum', 'fast', 'super-fast'] as const)[level] ?? 'normal',
      );
    }
  }
  return forms;
}

// ------------------------------------------------------------------ inventory

export type Axis = 'xml' | 'container';

export interface LexicalDeck {
  readonly id: string;
  readonly collection: string;
  readonly serializers: { readonly xml: string; readonly container: string };
  readonly bytes: Uint8Array;
}

export interface FormEvidence {
  readonly dimension: string;
  readonly form: string;
  readonly axis: Axis;
  /** Decks exercising it, sorted. */
  readonly decks: readonly string[];
  /** Distinct serializers on this form's axis, sorted. */
  readonly producers: readonly string[];
}

export const axisOf = (dimension: string): Axis =>
  dimension.startsWith('zip.') ? 'container' : 'xml';

/** What the committed corpus actually contains, one row per observed form. */
export function inventory(decks: readonly LexicalDeck[]): FormEvidence[] {
  const rows = new Map<string, { decks: Set<string>; producers: Set<string> }>();

  for (const deck of decks) {
    const archive = readHeaders(deck.bytes);
    const forms = scanArchive(archive);
    for (const entry of archive.entries) {
      if (!isXmlPart(entry.name)) continue;
      for (const form of scanXmlPart(entry.bytes)) forms.add(form);
    }

    for (const key of forms) {
      let row = rows.get(key);
      if (row === undefined) {
        row = { decks: new Set(), producers: new Set() };
        rows.set(key, row);
      }
      row.decks.add(deck.id);
      const axis = axisOf(key.slice(0, key.indexOf('=')));
      row.producers.add(axis === 'xml' ? deck.serializers.xml : deck.serializers.container);
    }
  }

  return [...rows.entries()]
    .map(([key, row]) => {
      const split = key.indexOf('=');
      const dimension = key.slice(0, split);
      return {
        dimension,
        form: key.slice(split + 1),
        axis: axisOf(dimension),
        decks: [...row.decks].sort((a, b) => a.localeCompare(b)),
        producers: [...row.producers].sort((a, b) => a.localeCompare(b)),
      };
    })
    .sort((a, b) =>
      a.dimension === b.dimension
        ? a.form.localeCompare(b.form)
        : a.dimension.localeCompare(b.dimension),
    );
}
