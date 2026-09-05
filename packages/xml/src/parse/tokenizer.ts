/**
 * A pull tokenizer over a decoded XML document.
 *
 * Two properties define it, and everything else follows from them:
 *
 * 1. **Every token carries the span it came from**, and the spans tile the
 *    input exactly - no gaps, no overlaps, first token starts at the BOM
 *    boundary, last token ends at `source.length`. That invariant is checkable
 *    on any document (see {@link checkSpanCoverage}) and it is what makes
 *    sub-phase 0.5's byte-identical serializer a slice rather than a rewrite.
 *
 * 2. **A token's source and a token's value are different things.** XML 1.0
 *    mandates transformations of the input - end-of-line normalization (§2.11),
 *    attribute-value normalization (§3.3.3), reference expansion (§4.6) - and a
 *    tokenizer that applies them in place can never reproduce the bytes it was
 *    given. So the span is the truth and the value is derived from it.
 *
 * The lexical detail this preserves is not academic. Handed a slide part,
 * PowerPoint's own writer discards comments, discards processing instructions,
 * rewrites `<![CDATA[x]]>` as plain text, resolves `&#72;` to `H`, converts
 * single-quoted attributes to double, collapses `<p:spPr></p:spPr>` to
 * `<p:spPr/>`, strips a BOM and drops an unused namespace declaration - all
 * measured, by resaving files through PowerPoint 365 and diffing. Every one of
 * those is something we are able to hand back untouched, and being lossier
 * than that would make us PowerPoint.
 */

import {
  isNameChar,
  isNameStartChar,
  isUnpairedSurrogate,
  isXmlChar,
  isXmlWhitespace,
} from './chars.js';
import { XmlError } from '../errors.js';
import { decodeCharacterData, isLiteralRun, normalizeAttributeValue } from '../mce/references.js';

const LT = 0x3c;
const GT = 0x3e;
const SLASH = 0x2f;
const BANG = 0x21;
const QUESTION = 0x3f;
const EQUALS = 0x3d;
const QUOTE = 0x22;
const APOSTROPHE = 0x27;
const BRACKET_CLOSE = 0x5d;

export type XmlTokenType =
  | 'declaration'
  | 'processingInstruction'
  | 'comment'
  | 'startTag'
  | 'emptyElementTag'
  | 'endTag'
  | 'text'
  | 'cdata';

/**
 * One attribute, with everything needed to reproduce it character for
 * character.
 *
 * This is the same object the tree holds: `XElement.attributes` adopts the
 * array the tokenizer produced rather than copying it. Attributes are the most
 * numerous thing in an OOXML part - 170 019 of them across our corpus, against
 * 194 148 elements - so allocating each one twice is a cost worth not paying.
 *
 * Offsets are flat numbers rather than a nested `{ start, end }` object for the
 * same reason.
 */
export interface XAttribute {
  qname: string;
  /** Empty string when the name is unprefixed. */
  prefix: string;
  local: string;
  /**
   * Start of this attribute's **leading whitespace**, not of its name.
   *
   * Including the whitespace is what lets sub-phase 0.6 rewrite one attribute
   * of a tag and leave the spacing of the others alone. Office's own writer is
   * consistent, but 70 822 of the 98 777 self-closing tags in our corpus come
   * from a producer that writes `<a:off x="0" y="0" />` with a space before the
   * slash, so "the spacing is always the same" is false in practice.
   */
  readonly start: number;
  readonly nameStart: number;
  readonly nameEnd: number;
  /** First character inside the opening quote. */
  readonly valueStart: number;
  /** The closing quote. */
  readonly valueEnd: number;
  /** One past the closing quote. */
  readonly end: number;
  quote: '"' | "'";
  /**
   * Normalized per XML 1.0 §3.3.3.
   *
   * Deliberately not the same string as `source.slice(valueStart, valueEnd)`.
   * `name="a&#9;b"` and `name="a<TAB>b"` have the same normalized value only if
   * you get §3.3.3 wrong: a literal tab becomes a space, a character reference
   * to a tab stays a tab.
   */
  value: string;
  /**
   * Set by `markAttributeDirty` when the value no longer matches the source.
   *
   * While it is false the attribute re-emits by slicing `[start, end)`, which
   * keeps its original spacing and quote character. Setting it by hand without
   * marking the owning element is a silent no-op - the element would still be
   * clean and would re-emit the whole start tag as a slice.
   */
  dirty: boolean;
}

interface TokenBase {
  readonly type: XmlTokenType;
  readonly start: number;
  readonly end: number;
}

export interface XmlDeclarationToken extends TokenBase {
  readonly type: 'declaration';
  readonly version: string;
  readonly encoding: string | undefined;
  readonly standalone: string | undefined;
}

export interface XmlProcessingInstructionToken extends TokenBase {
  readonly type: 'processingInstruction';
  readonly target: string;
  /** Everything between the target and `?>`, verbatim. */
  readonly data: string;
}

export interface XmlCommentToken extends TokenBase {
  readonly type: 'comment';
  /** The text between `<!--` and `-->`, with line endings normalized. */
  readonly value: string;
}

export interface XmlElementToken extends TokenBase {
  readonly type: 'startTag' | 'emptyElementTag';
  readonly qname: string;
  readonly prefix: string;
  readonly local: string;
  readonly nameEnd: number;
  /**
   * Where the whitespace before `>` or `/>` begins.
   *
   * Equal to the last attribute's `end` (or to `nameEnd`) when there is none.
   */
  readonly trailingSpaceStart: number;
  readonly attributes: readonly XAttribute[];
  /**
   * True if any attribute is an `xmlns` or `xmlns:*` declaration.
   *
   * Computed here because the attribute loop has already looked at every name.
   * It turns `resolvePrefix`'s walk from "scan every ancestor's whole attribute
   * list" into "step over ancestors that declare nothing", which is the
   * difference between 37 seconds and milliseconds on a deep tree whose
   * elements carry many non-namespace attributes.
   */
  readonly hasNamespaceDeclarations: boolean;
}

export interface XmlEndTagToken extends TokenBase {
  readonly type: 'endTag';
  readonly qname: string;
  readonly prefix: string;
  readonly local: string;
}

export interface XmlTextToken extends TokenBase {
  readonly type: 'text';
  /** References expanded and line endings normalized. */
  readonly value: string;
  /** True if the run is `S` characters only - inter-element formatting. */
  readonly whitespaceOnly: boolean;
}

export interface XmlCdataToken extends TokenBase {
  readonly type: 'cdata';
  /** The text between `<![CDATA[` and `]]>`, with line endings normalized. */
  readonly value: string;
}

export type XmlToken =
  | XmlDeclarationToken
  | XmlProcessingInstructionToken
  | XmlCommentToken
  | XmlElementToken
  | XmlEndTagToken
  | XmlTextToken
  | XmlCdataToken;

/**
 * Structural ceilings.
 *
 * None of these defends against a *malformed* document - malformed input fails
 * on its own merits. They bound what a *well-formed* but hostile one can cost,
 * which is the case a hand-rolled parser actually has to survive.
 */
export interface XmlTokenizerLimits {
  /** Attributes on one element. The most any element in our corpus carries is 15. */
  readonly maxAttributes: number;
  /** Tokens in one part. Our largest part yields 2303 elements. */
  readonly maxTokens: number;
  /**
   * Length of the source, in UTF-16 code units.
   *
   * The only ceiling here denominated in the quantity that actually governs
   * allocation. Counting tokens, nodes or attributes-per-element each bounds
   * one dimension and misses the product: 20 000 elements carrying 256
   * attributes apiece is 41.9 MiB of source, 20 001 nodes - half a percent of
   * `maxNodes` - and 817 MB of heap. Under a constrained heap that is a
   * `FATAL ERROR: Ineffective mark-compacts` process abort, which no
   * `try`/`catch` can see.
   *
   * The largest XML part in our 37-package corpus is 69 KB; the largest seen
   * anywhere in the wild is under a megabyte. 16 MiB is far past both.
   */
  readonly maxSourceLength: number;
}

export const DEFAULT_TOKENIZER_LIMITS: XmlTokenizerLimits = {
  maxAttributes: 256,
  maxTokens: 4_000_000,
  maxSourceLength: 16 * 1024 * 1024,
};

function fail(message: string, offset: number, name?: string): never {
  throw new XmlError('ERR_MALFORMED_XML', message, { offset, name });
}

function unpaired(offset: number): never {
  return fail(
    'an unpaired surrogate code unit is not a character; it cannot be encoded back to UTF-8',
    offset,
  );
}

/**
 * The tokenizer.
 *
 * Pull-shaped: {@link next} returns the token at the cursor and advances, or
 * `undefined` at the end. It is a plain loop with no recursion anywhere, so
 * nesting depth costs nothing here and is bounded by the tree builder instead.
 */
export class XmlTokenizer {
  readonly source: string;
  readonly #limits: XmlTokenizerLimits;
  #pos: number;
  #emitted = 0;
  /** Set once the declaration has been consumed, or once anything else has. */
  #atStart = true;

  constructor(source: string, limits: XmlTokenizerLimits = DEFAULT_TOKENIZER_LIMITS) {
    if (source.length > limits.maxSourceLength) {
      throw new XmlError('ERR_LIMIT_EXCEEDED', 'this part is larger than the limit allows', {
        limit: limits.maxSourceLength,
        actual: source.length,
      });
    }
    this.source = source;
    this.#limits = limits;
    // A BOM survived decoding as U+FEFF so the source could be re-encoded
    // byte-identically. It is not markup, so the token stream starts after it.
    this.#pos = source.charCodeAt(0) === 0xfeff ? 1 : 0;
  }

  /** Index of the first token. Non-zero exactly when the source carries a BOM. */
  get contentStart(): number {
    return this.source.charCodeAt(0) === 0xfeff ? 1 : 0;
  }

  get position(): number {
    return this.#pos;
  }

  next(): XmlToken | undefined {
    const source = this.source;
    if (this.#pos >= source.length) return undefined;
    if (++this.#emitted > this.#limits.maxTokens) {
      throw new XmlError(
        'ERR_LIMIT_EXCEEDED',
        'this part has more XML tokens than the limit allows',
        {
          offset: this.#pos,
          limit: this.#limits.maxTokens,
        },
      );
    }

    const start = this.#pos;
    if (source.charCodeAt(start) !== LT) return this.#readText(start);

    const second = source.charCodeAt(start + 1);
    if (second === SLASH) return this.#readEndTag(start);
    if (second === BANG) return this.#readBangConstruct(start);
    if (second === QUESTION) return this.#readQuestionConstruct(start);
    return this.#readStartTag(start);
  }

  /** Every remaining token, in order. */
  *[Symbol.iterator](): IterableIterator<XmlToken> {
    for (let token = this.next(); token !== undefined; token = this.next()) yield token;
  }

  // -------------------------------------------------------------- character data

  #readText(start: number): XmlTextToken {
    const source = this.source;
    const end = source.indexOf('<', start);
    const stop = end < 0 ? source.length : end;

    let whitespaceOnly = true;
    for (let i = start; i < stop; i++) {
      const code = source.charCodeAt(i);
      if (!isXmlChar(code)) {
        fail(
          'U+' +
            code.toString(16).toUpperCase().padStart(4, '0') +
            ' is not a character XML permits; it must be removed, not escaped',
          i,
        );
      }
      if (isUnpairedSurrogate(source, i)) unpaired(i);
      if (whitespaceOnly && !isXmlWhitespace(code)) whitespaceOnly = false;
      // XML 1.0 §2.4: the literal sequence "]]>" may not appear in content,
      // because a parser cannot tell it from the end of a CDATA section.
      if (
        code === BRACKET_CLOSE &&
        source.charCodeAt(i + 1) === BRACKET_CLOSE &&
        source.charCodeAt(i + 2) === GT
      ) {
        fail('the sequence "]]>" may not appear in text; write "]]&gt;"', i);
      }
    }

    this.#pos = stop;
    return {
      type: 'text',
      start,
      end: stop,
      whitespaceOnly,
      value: isLiteralRun(source, start, stop, false)
        ? source.slice(start, stop)
        : decodeCharacterData(source, start, stop),
    };
  }

  // ------------------------------------------------------------------ <! ... >

  #readBangConstruct(start: number): XmlCommentToken | XmlCdataToken {
    const source = this.source;
    if (source.startsWith('<!--', start)) return this.#readComment(start);
    if (source.startsWith('<![CDATA[', start)) return this.#readCdata(start);
    // Case-insensitively. `<!DocType r>` is the same construct and the same
    // attack surface, and reporting it as generic malformed XML loses the one
    // machine-readable signal a security reviewer greps for.
    if (source.slice(start, start + 9).toUpperCase() === '<!DOCTYPE') {
      throw new XmlError(
        'ERR_DOCTYPE_FORBIDDEN',
        'this document declares a DOCTYPE. @pptx-studio/xml refuses one outright: a DTD is where XXE and entity expansion live, and PowerPoint refuses one too',
        { offset: start },
      );
    }
    fail('"<!" must begin a comment, a CDATA section or a DOCTYPE', start);
  }

  #readComment(start: number): XmlCommentToken {
    const source = this.source;
    const from = start + 4;
    const close = source.indexOf('-->', from);
    if (close < 0) fail('unterminated comment', start);

    // XML 1.0 §2.5: "--" may not occur within a comment, which also makes
    // "--->" invalid - the comment would end with a hyphen.
    const inner = source.indexOf('--', from);
    if (inner >= 0 && inner < close) {
      fail('"--" may not appear inside a comment, and a comment may not end with "-"', inner);
    }
    this.#assertChars(from, close);

    this.#pos = close + 3;
    return {
      type: 'comment',
      start,
      end: this.#pos,
      value: this.#normalized(from, close),
    };
  }

  #readCdata(start: number): XmlCdataToken {
    const source = this.source;
    const from = start + 9;
    const close = source.indexOf(']]>', from);
    if (close < 0) fail('unterminated CDATA section', start);
    this.#assertChars(from, close);

    this.#pos = close + 3;
    return { type: 'cdata', start, end: this.#pos, value: this.#normalized(from, close) };
  }

  // ------------------------------------------------------------------ <? ... ?>

  #readQuestionConstruct(start: number): XmlDeclarationToken | XmlProcessingInstructionToken {
    const source = this.source;
    const close = source.indexOf('?>', start + 2);
    if (close < 0) fail('unterminated processing instruction', start);

    let i = start + 2;
    const nameStart = i;
    if (!isNameStartChar(source.charCodeAt(i))) fail('a processing instruction needs a target', i);
    i++;
    while (i < close && isNameChar(source.charCodeAt(i))) i++;
    const target = source.slice(nameStart, i);

    // Before anything is believed about the content. This used to sit on the
    // processing-instruction branch only, so a literal U+000B or NUL inside the
    // declaration - the one construct nothing else validates - went straight
    // through, and we would faithfully write the forbidden byte back out for
    // PowerPoint to reject at the far end of the pipeline.
    this.#assertChars(i, close);

    // XML 1.0 §2.8: the declaration is not a processing instruction, it only
    // looks like one, and it may appear only at the very start of the document.
    // `xml` there is a lowercase terminal; `<?XML` is simply a reserved target.
    if (target === 'xml' && this.#atStart && start === this.contentStart) {
      this.#atStart = false;
      return this.#readDeclaration(start, i, close);
    }
    if (target.toLowerCase() === 'xml') {
      fail('"' + target + '" is a reserved processing-instruction target', nameStart, target);
    }
    this.#atStart = false;

    if (i < close && !isXmlWhitespace(source.charCodeAt(i))) {
      fail('a processing-instruction target must be followed by whitespace or "?>"', i);
    }

    this.#pos = close + 2;
    return {
      type: 'processingInstruction',
      start,
      end: this.#pos,
      target,
      data: source.slice(i, close),
    };
  }

  /**
   * `XMLDecl ::= '<?xml' VersionInfo EncodingDecl? SDDecl? S? '?>'`.
   *
   * A fixed grammar with three pseudo-attributes in a fixed order, and nothing
   * else permitted. It used to be scraped with three independent regular
   * expressions over the body, which accepted reordered attributes, unknown
   * attributes, arbitrary junk, and - because the pattern was not anchored -
   * read `version="1.0"` out of the *inside* of an encoding value.
   */
  #readDeclaration(start: number, afterTarget: number, close: number): XmlDeclarationToken {
    const source = this.source;
    let i = afterTarget;

    const space = (): boolean => {
      const from = i;
      while (i < close && isXmlWhitespace(source.charCodeAt(i))) i++;
      return i > from;
    };

    const pseudoAttribute = (name: string): string | undefined => {
      const mark = i;
      if (!space()) return undefined;
      if (!source.startsWith(name, i)) {
        i = mark;
        return undefined;
      }
      i += name.length;
      while (i < close && isXmlWhitespace(source.charCodeAt(i))) i++;
      if (source.charCodeAt(i) !== EQUALS) fail('"' + name + '" needs a value', i, name);
      i++;
      while (i < close && isXmlWhitespace(source.charCodeAt(i))) i++;
      const quote = source.charCodeAt(i);
      if (quote !== QUOTE && quote !== APOSTROPHE) {
        fail('the value of "' + name + '" must be quoted', i, name);
      }
      const from = i + 1;
      const to = source.indexOf(quote === QUOTE ? '"' : "'", from);
      if (to < 0 || to > close) fail('unterminated value for "' + name + '"', i, name);
      i = to + 1;
      return source.slice(from, to);
    };

    const version = pseudoAttribute('version');
    if (version === undefined) fail('the XML declaration must state a version', start);
    if (!/^1\.[0-9]+$/.test(version)) {
      fail('"' + version + '" is not an XML version number', start, version);
    }
    const encoding = pseudoAttribute('encoding');
    if (encoding !== undefined && !/^[A-Za-z][A-Za-z0-9._-]*$/.test(encoding)) {
      fail('"' + encoding + '" is not an encoding name', start, encoding);
    }
    const standalone = pseudoAttribute('standalone');
    if (standalone !== undefined && standalone !== 'yes' && standalone !== 'no') {
      fail('standalone must be "yes" or "no", not "' + standalone + '"', start, standalone);
    }
    space();
    if (i !== close) {
      fail('the XML declaration carries something that is not version, encoding or standalone', i);
    }

    this.#pos = close + 2;
    return {
      type: 'declaration',
      start,
      end: this.#pos,
      version,
      encoding,
      standalone,
    };
  }

  // -------------------------------------------------------------------- tags

  #readEndTag(start: number): XmlEndTagToken {
    const source = this.source;
    this.#atStart = false;
    let i = start + 2;
    const nameStart = i;
    if (!isNameStartChar(source.charCodeAt(i))) fail('an end tag needs a name', i);
    i++;
    while (i < source.length && isNameChar(source.charCodeAt(i))) i++;
    const qname = source.slice(nameStart, i);

    while (i < source.length && isXmlWhitespace(source.charCodeAt(i))) i++;
    if (source.charCodeAt(i) !== GT) fail('an end tag may not carry attributes', i, qname);

    this.#pos = i + 1;
    const { prefix, local } = this.#splitName(qname, nameStart);
    return { type: 'endTag', start, end: this.#pos, qname, prefix, local };
  }

  #readStartTag(start: number): XmlElementToken {
    const source = this.source;
    const length = source.length;
    this.#atStart = false;

    let i = start + 1;
    if (!isNameStartChar(source.charCodeAt(i))) fail('a tag needs a name', i);
    i++;
    while (i < length && isNameChar(source.charCodeAt(i))) i++;
    const nameEnd = i;
    const qname = source.slice(start + 1, nameEnd);

    const attributes: XAttribute[] = [];

    for (;;) {
      const spaceStart = i;
      while (i < length && isXmlWhitespace(source.charCodeAt(i))) i++;
      const code = source.charCodeAt(i);

      if (code === GT) {
        this.#pos = i + 1;
        return this.#element('startTag', start, this.#pos, qname, nameEnd, spaceStart, attributes);
      }
      if (code === SLASH) {
        if (source.charCodeAt(i + 1) !== GT) fail('"/" in a tag must be followed by ">"', i, qname);
        this.#pos = i + 2;
        return this.#element(
          'emptyElementTag',
          start,
          this.#pos,
          qname,
          nameEnd,
          spaceStart,
          attributes,
        );
      }
      if (i >= length) fail('unterminated tag', start, qname);
      if (spaceStart === i) fail('attributes must be separated by whitespace', i, qname);

      attributes.push(this.#readAttribute(spaceStart, i, qname, attributes));
      if (attributes.length > this.#limits.maxAttributes) {
        throw new XmlError(
          'ERR_LIMIT_EXCEEDED',
          '<' + qname + '> carries more attributes than the limit allows',
          { offset: start, name: qname, limit: this.#limits.maxAttributes },
        );
      }
      i = attributes[attributes.length - 1]!.end;
    }
  }

  #readAttribute(
    start: number,
    nameStart: number,
    owner: string,
    siblings: readonly XAttribute[],
  ): XAttribute {
    const source = this.source;
    const length = source.length;

    let i = nameStart;
    if (!isNameStartChar(source.charCodeAt(i))) fail('an attribute needs a name', i, owner);
    i++;
    while (i < length && isNameChar(source.charCodeAt(i))) i++;
    const nameEnd = i;
    const qname = source.slice(nameStart, nameEnd);

    while (i < length && isXmlWhitespace(source.charCodeAt(i))) i++;
    if (source.charCodeAt(i) !== EQUALS) fail('attribute "' + qname + '" has no value', i, qname);
    i++;
    while (i < length && isXmlWhitespace(source.charCodeAt(i))) i++;

    const quoteCode = source.charCodeAt(i);
    if (quoteCode !== QUOTE && quoteCode !== APOSTROPHE) {
      fail('the value of "' + qname + '" must be quoted', i, qname);
    }
    const valueStart = i + 1;
    const valueEnd = source.indexOf(quoteCode === QUOTE ? '"' : "'", valueStart);
    if (valueEnd < 0) fail('unterminated value for attribute "' + qname + '"', i, qname);

    for (let k = valueStart; k < valueEnd; k++) {
      const code = source.charCodeAt(k);
      // XML 1.0 §3.1: "<" may never appear literally in an attribute value.
      // A "&" must begin a reference, which normalizeAttributeValue checks.
      if (code === LT) fail('"<" may not appear in an attribute value; write "&lt;"', k, qname);
      if (isUnpairedSurrogate(source, k)) unpaired(k);
      if (!isXmlChar(code)) {
        fail(
          'U+' +
            code.toString(16).toUpperCase().padStart(4, '0') +
            ' is not a character XML permits',
          k,
        );
      }
    }

    // XML 1.0 §3.1 well-formedness constraint "Unique Att Spec". PowerPoint
    // enforces it: `<a:off x="0" y="0" x="0"/>` is refused with 0x80070570.
    //
    // A linear scan rather than a Set, because a Set would have to be allocated
    // once per element - 262 429 of them in an 8 MB part - to dedupe a list
    // whose longest instance anywhere in our corpus is 15 entries.
    if (siblings.some((existing) => existing.qname === qname)) {
      throw new XmlError(
        'ERR_DUPLICATE_ATTRIBUTE',
        '<' + owner + '> carries the attribute "' + qname + '" twice',
        { offset: nameStart, name: qname },
      );
    }

    const { prefix, local } = this.#splitName(qname, nameStart);
    return {
      qname,
      prefix,
      local,
      start,
      nameStart,
      nameEnd,
      valueStart,
      valueEnd,
      end: valueEnd + 1,
      quote: quoteCode === QUOTE ? '"' : "'",
      value: isLiteralRun(source, valueStart, valueEnd, true)
        ? source.slice(valueStart, valueEnd)
        : normalizeAttributeValue(source, valueStart, valueEnd),
      dirty: false,
    };
  }

  /**
   * Split a qualified name, rejecting the shapes *Namespaces in XML* forbids.
   *
   * This used to be applied to element names only, so `<r a:b:c="1"/>` yielded
   * an attribute with prefix `a` and local name `b:c`, `<r a="1" :a="2"/>`
   * produced two attributes that both resolved to `('', 'a')` while sliding
   * past the duplicate-attribute check (which compares qnames), and
   * `<r xmlns:="u"/>` was reported by `declaredNamespaces` as a binding of the
   * *default* namespace.
   */
  #splitName(qname: string, offset: number): { prefix: string; local: string } {
    const colon = qname.indexOf(':');
    if (colon === 0 || colon === qname.length - 1 || qname.indexOf(':', colon + 1) > 0) {
      fail('"' + qname + '" is not a usable qualified name', offset, qname);
    }
    return colon < 0
      ? { prefix: '', local: qname }
      : { prefix: qname.slice(0, colon), local: qname.slice(colon + 1) };
  }

  #element(
    type: 'startTag' | 'emptyElementTag',
    start: number,
    end: number,
    qname: string,
    nameEnd: number,
    trailingSpaceStart: number,
    attributes: readonly XAttribute[],
  ): XmlElementToken {
    const { prefix, local } = this.#splitName(qname, start);
    return {
      type,
      start,
      end,
      qname,
      prefix,
      local,
      nameEnd,
      trailingSpaceStart,
      attributes,
      hasNamespaceDeclarations: attributes.some((a) => a.qname === 'xmlns' || a.prefix === 'xmlns'),
    };
  }

  // ------------------------------------------------------------------ helpers

  #assertChars(start: number, end: number): void {
    const source = this.source;
    for (let i = start; i < end; i++) {
      const code = source.charCodeAt(i);
      if (isUnpairedSurrogate(source, i)) unpaired(i);
      if (!isXmlChar(code)) {
        fail(
          'U+' +
            code.toString(16).toUpperCase().padStart(4, '0') +
            ' is not a character XML permits',
          i,
        );
      }
    }
  }

  #normalized(start: number, end: number): string {
    const source = this.source;
    for (let i = start; i < end; i++) {
      if (source.charCodeAt(i) === 0x0d) return source.slice(start, end).replace(/\r\n?/g, '\n');
    }
    return source.slice(start, end);
  }
}

/** Every token in `source`, in order. */
export function tokenize(source: string, limits?: XmlTokenizerLimits): XmlToken[] {
  const tokenizer = new XmlTokenizer(source, limits);
  const tokens: XmlToken[] = [];
  for (let token = tokenizer.next(); token !== undefined; token = tokenizer.next()) {
    tokens.push(token);
  }
  return tokens;
}

/** A place where the token spans fail to tile the source. */
export interface SpanGap {
  readonly kind: 'gap' | 'overlap' | 'prologue' | 'epilogue';
  readonly at: number;
  readonly text: string;
}

/**
 * Check that a token stream accounts for every character of its source.
 *
 * This is the real gate for sub-phase 0.4, and it is stronger than "tokenize
 * the corpus without error" by the margin that matters: a tokenizer with an
 * off-by-one in a span raises no error at all, tokenizes every part of every
 * deck happily, and then silently drops or duplicates a character the first
 * time 0.5 re-serializes from spans. Coverage is total and falsifiable, so it
 * catches that on the first part.
 */
export function checkSpanCoverage(source: string, tokens: readonly XmlToken[]): SpanGap[] {
  const gaps: SpanGap[] = [];
  const contentStart = source.charCodeAt(0) === 0xfeff ? 1 : 0;
  let cursor = contentStart;

  for (const token of tokens) {
    if (token.start > cursor) {
      gaps.push({ kind: 'gap', at: cursor, text: source.slice(cursor, token.start) });
    } else if (token.start < cursor) {
      gaps.push({ kind: 'overlap', at: token.start, text: source.slice(token.start, cursor) });
    }
    cursor = token.end;
  }

  if (cursor < source.length) {
    gaps.push({ kind: 'epilogue', at: cursor, text: source.slice(cursor) });
  }
  if (contentStart === 1 && tokens.length > 0 && tokens[0]!.start !== 1) {
    gaps.push({ kind: 'prologue', at: 0, text: source.slice(0, tokens[0]!.start) });
  }
  return gaps;
}

/** Split a qualified name. Exported because callers outside the tokenizer need it too. */
export function splitQName(qname: string): { prefix: string; local: string } {
  const colon = qname.indexOf(':');
  return colon < 0
    ? { prefix: '', local: qname }
    : { prefix: qname.slice(0, colon), local: qname.slice(colon + 1) };
}
