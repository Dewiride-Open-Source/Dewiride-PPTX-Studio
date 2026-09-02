import { CONTENT_TYPES_PART, OPC_NS } from './constants.js';
import { OpcError } from './errors.js';
import { escapeAttribute, readFlatXml, XML_DECLARATION } from './flat-xml.js';
import { normalizePartName, partExtension, toPartName } from './pack-uri.js';

/**
 * `[Content_Types].xml` - the one part that says what every other part is.
 *
 * OPC has no filename convention that carries meaning. `ppt/slides/slide1.xml`
 * is a slide because this stream says so, and for no other reason. Which makes
 * it the single most load-bearing part in the package, and the one whose
 * failure mode we measured directly rather than inferred: strip the
 * `<Default Extension="png"/>` from a deck that contains a PNG and PowerPoint
 * refuses the file outright with `0x80CB8002`. Delete the stream entirely and
 * you get the same code. That is a different error from a broken relationship
 * graph (`0x80070570`), and the difference tells you the content-type layer
 * rejects the package before the presentation layer ever sees it.
 *
 * The asymmetry is worth knowing, because it is not what you would guess:
 * removing an `Override` for a slide is *tolerated* - the slide falls back to
 * the `Default` for `xml` and PowerPoint opens the deck anyway - while removing
 * a `Default` that some part relies on is fatal. A part with no content type at
 * all is the thing that cannot be forgiven, and that is exactly the assertion
 * this module makes at write time.
 */

/** Resolution order, and where a part's type came from. */
export type ContentTypeOrigin = 'override' | 'default';

export interface DefaultEntry {
  /** The extension as written, without a dot. Office writes these lowercase. */
  readonly extension: string;
  readonly contentType: string;
}

export interface OverrideEntry {
  /** An absolute part name, leading slash included. */
  readonly partName: string;
  readonly contentType: string;
}

export interface ResolvedContentType {
  readonly contentType: string;
  readonly origin: ContentTypeOrigin;
}

/**
 * A media type is `type/subtype` with optional parameters.
 *
 * Deliberately loose about parameters and strict about the slash: a content
 * type with no slash is not a media type under any reading, while a parameter
 * we do not recognise is somebody else's extension, not our problem.
 */
const MEDIA_TYPE = /^[A-Za-z0-9!#$%&'*+\-.^_`|~]+\/[A-Za-z0-9!#$%&'*+\-.^_`|~]+(\s*;[\s\S]*)?$/;

function checkContentType(value: string, where: string): string {
  const trimmed = value.trim();
  if (!MEDIA_TYPE.test(trimmed)) {
    throw new OpcError(
      'ERR_INVALID_CONTENT_TYPE',
      JSON.stringify(value) + ' is not a media type (' + where + ')',
      { entry: CONTENT_TYPES_PART },
    );
  }
  return trimmed;
}

/**
 * Compare two content types.
 *
 * The media type is case-insensitive per RFC 2045, so `Application/XML` and
 * `application/xml` are the same type and comparing them with `===` would
 * report a conflict where there is none. Parameters are compared verbatim -
 * nothing in a `.pptx` uses them, and inventing a normalisation for something
 * we have never seen is how you get a rule nobody can justify later.
 */
function sameContentType(a: string, b: string): boolean {
  return normalizePartName(a.trim()) === normalizePartName(b.trim());
}

/**
 * Media types that are XML but do not say so.
 *
 * The `+xml` structured-syntax suffix of RFC 6839 answers the question for
 * almost every part in a package. These two are the exceptions in a `.pptx`,
 * and the first one is the one that bites: `vmlDrawing` carries no suffix at
 * all, and every OLE object in every deck has one. A reader that trusts the
 * suffix alone treats VML - which is where an OLE object's on-slide preview
 * lives, and where `@spid` resolves to - as an opaque blob.
 */
const XML_CONTENT_TYPES: ReadonlySet<string> = new Set([
  'application/xml',
  'text/xml',
  // Lower case, because the lookup lower-cases first. Office writes this one
  // `vmlDrawing`, and a set that copied that spelling would never match.
  'application/vnd.openxmlformats-officedocument.vmldrawing',
  'application/inkml+xml',
]);

/**
 * Whether a part with this content type holds XML.
 *
 * Parameters are stripped and the media type is lower-cased first, for the same
 * reason {@link sameContentType} does it: RFC 2045 makes the media type
 * case-insensitive, and a deck in the wild spells things how it likes.
 */
export function isXmlContentType(contentType: string | undefined): boolean {
  if (contentType === undefined) return false;
  const base = (contentType.split(';')[0] ?? '').trim().toLowerCase();
  return base.endsWith('+xml') || XML_CONTENT_TYPES.has(base);
}

export class ContentTypes {
  #defaults: DefaultEntry[] = [];
  #overrides: OverrideEntry[] = [];
  /** Normalised extension -> index into `#defaults`. */
  #defaultIndex = new Map<string, number>();
  /** Normalised part name -> index into `#overrides`. */
  #overrideIndex = new Map<string, number>();
  #dirty = false;

  private constructor() {}

  /** An empty map, for a package being built from nothing. */
  static empty(): ContentTypes {
    const ct = new ContentTypes();
    ct.#dirty = true;
    return ct;
  }

  /**
   * Read the stream.
   *
   * Order is preserved so that regenerating an unchanged map reproduces the
   * original layout, which is what keeps an export diff readable.
   */
  static parse(bytes: Uint8Array): ContentTypes {
    const ct = new ContentTypes();
    const elements = readFlatXml(bytes, CONTENT_TYPES_PART);
    const root = elements[0]!;
    if (root.name !== 'Types') {
      throw new OpcError(
        'ERR_MALFORMED_XML',
        CONTENT_TYPES_PART + ': root element is <' + root.qname + '>, expected <Types>',
        { entry: CONTENT_TYPES_PART },
      );
    }

    for (const el of elements) {
      if (el.depth !== 1) continue;
      if (el.name === 'Default') {
        const extension = el.attrs.get('Extension');
        const contentType = el.attrs.get('ContentType');
        if (extension === undefined || contentType === undefined) {
          throw new OpcError(
            'ERR_MALFORMED_XML',
            CONTENT_TYPES_PART + ': <Default> needs both Extension and ContentType',
            { entry: CONTENT_TYPES_PART },
          );
        }
        if (extension === '') {
          throw new OpcError(
            'ERR_MALFORMED_XML',
            CONTENT_TYPES_PART + ': <Default> has an empty Extension',
            { entry: CONTENT_TYPES_PART },
          );
        }
        ct.#addDefault(extension, checkContentType(contentType, 'Default ' + extension), false);
      } else if (el.name === 'Override') {
        const partName = el.attrs.get('PartName');
        const contentType = el.attrs.get('ContentType');
        if (partName === undefined || contentType === undefined) {
          throw new OpcError(
            'ERR_MALFORMED_XML',
            CONTENT_TYPES_PART + ': <Override> needs both PartName and ContentType',
            { entry: CONTENT_TYPES_PART },
          );
        }
        ct.#addOverride(
          toPartName(partName),
          checkContentType(contentType, 'Override ' + partName),
          false,
        );
      }
      // Anything else at this level is somebody's extension. Ignoring it cannot
      // give a part the wrong type, so it is ignored.
    }
    return ct;
  }

  get defaults(): readonly DefaultEntry[] {
    return this.#defaults;
  }

  get overrides(): readonly OverrideEntry[] {
    return this.#overrides;
  }

  /** True once anything has changed, which is what makes the writer regenerate. */
  get dirty(): boolean {
    return this.#dirty;
  }

  #addDefault(extension: string, contentType: string, dirty: boolean): void {
    const key = normalizePartName(extension);
    const existing = this.#defaultIndex.get(key);
    if (existing !== undefined) {
      const previous = this.#defaults[existing]!;
      if (!sameContentType(previous.contentType, contentType)) {
        throw new OpcError(
          'ERR_CONTENT_TYPE_CONFLICT',
          'extension ' +
            JSON.stringify(extension) +
            ' has two different Default content types, ' +
            JSON.stringify(previous.contentType) +
            ' and ' +
            JSON.stringify(contentType) +
            '. Which one applies is not decidable, and guessing would type every part with ' +
            'that extension wrongly.',
          { entry: CONTENT_TYPES_PART },
        );
      }
      this.#defaults[existing] = { extension, contentType };
      // A duplicate that agrees is still a duplicate. M2.5 is unconditional and
      // PowerPoint enforces it literally: two `<Default>`s for one extension
      // are refused with `0x80CB8000` even when the two elements are
      // byte-identical. We tolerate it on read - but marking dirty is what
      // stops us passing the original stream through untouched and re-emitting
      // a package that does not open. Anything we accept but would not write
      // has to regenerate, or lenient reading quietly launders corruption.
      this.#dirty = true;
      return;
    }
    this.#defaultIndex.set(key, this.#defaults.length);
    this.#defaults.push({ extension, contentType });
    if (dirty) this.#dirty = true;
  }

  #addOverride(partName: string, contentType: string, dirty: boolean): void {
    const key = normalizePartName(partName);
    const existing = this.#overrideIndex.get(key);
    if (existing !== undefined) {
      const previous = this.#overrides[existing]!;
      if (!sameContentType(previous.contentType, contentType)) {
        throw new OpcError(
          'ERR_CONTENT_TYPE_CONFLICT',
          'part ' +
            JSON.stringify(partName) +
            ' has two different Override content types, ' +
            JSON.stringify(previous.contentType) +
            ' and ' +
            JSON.stringify(contentType),
          { entry: partName },
        );
      }
      this.#overrides[existing] = { partName, contentType };
      // Same rule as `#addDefault`: PowerPoint refuses a duplicate `<Override>`
      // with `0x80CB8001` whether or not the two agree, so a tolerated
      // duplicate must regenerate rather than pass through.
      this.#dirty = true;
      return;
    }
    this.#overrideIndex.set(key, this.#overrides.length);
    this.#overrides.push({ partName, contentType });
    if (dirty) this.#dirty = true;
  }

  /**
   * The content type of `partName`, and where it came from.
   *
   * `Override` beats `Default`; both comparisons are ASCII case-insensitive,
   * matching the case-insensitive equivalence OPC already requires of part
   * names. Returns `undefined` when neither applies - that is a package error,
   * but reporting it is the caller's job, because a deck that PowerPoint opens
   * should open here even while we refuse to *write* one in that state.
   */
  resolve(partName: string): ResolvedContentType | undefined {
    const override = this.#overrideIndex.get(normalizePartName(partName));
    if (override !== undefined) {
      return { contentType: this.#overrides[override]!.contentType, origin: 'override' };
    }
    const extension = partExtension(partName);
    if (extension === '') return undefined;
    const fallback = this.#defaultIndex.get(normalizePartName(extension));
    if (fallback === undefined) return undefined;
    return { contentType: this.#defaults[fallback]!.contentType, origin: 'default' };
  }

  /** The content type, or `undefined`. */
  for(partName: string): string | undefined {
    return this.resolve(partName)?.contentType;
  }

  /** The content type, or throw. This is what the writer calls. */
  require(partName: string): string {
    const found = this.resolve(partName);
    if (found !== undefined) return found.contentType;
    const extension = partExtension(partName);
    throw new OpcError(
      'ERR_MISSING_CONTENT_TYPE',
      'part ' +
        partName +
        ' has no content type: no <Override> names it' +
        (extension === ''
          ? ' and it has no extension for a <Default> to match'
          : ' and no <Default Extension="' + extension + '"/> exists') +
        '. PowerPoint refuses a package in this state outright, so we do not write one.',
      { entry: partName },
    );
  }

  setDefault(extension: string, contentType: string): void {
    if (extension === '' || extension.includes('.') || extension.includes('/')) {
      throw new OpcError(
        'ERR_INVALID_CONTENT_TYPE',
        JSON.stringify(extension) + ' is not an extension',
        { entry: CONTENT_TYPES_PART },
      );
    }
    this.#addDefault(extension, checkContentType(contentType, 'Default ' + extension), true);
  }

  setOverride(partName: string, contentType: string): void {
    this.#addOverride(
      toPartName(partName),
      checkContentType(contentType, 'Override ' + partName),
      true,
    );
  }

  /** Remove the `Override` for a part, if there is one. Returns whether there was. */
  removeOverride(partName: string): boolean {
    const key = normalizePartName(partName);
    const at = this.#overrideIndex.get(key);
    if (at === undefined) return false;
    this.#overrides.splice(at, 1);
    this.#overrideIndex.delete(key);
    for (const [k, v] of this.#overrideIndex) {
      if (v > at) this.#overrideIndex.set(k, v - 1);
    }
    this.#dirty = true;
    return true;
  }

  /**
   * Make sure `partName` resolves to `contentType`, doing as little as possible.
   *
   * An `Override` rather than a `Default`, because an `Override` is always
   * correct and a `Default` is a claim about every part sharing that extension
   * - one that a caller adding a single part is in no position to make. Callers
   * that *do* know the extension is uniform, such as the font embedder writing
   * `fntdata`, call `setDefault` and say so.
   */
  ensureFor(partName: string, contentType: string): void {
    const current = this.resolve(partName);
    if (current !== undefined && sameContentType(current.contentType, contentType)) return;
    this.setOverride(partName, contentType);
  }

  /**
   * Serialise, in the layout Office uses: declaration, `Types`, every `Default`
   * in order, then every `Override` in order, no whitespace anywhere.
   */
  serialize(): Uint8Array {
    let xml = XML_DECLARATION + '<Types xmlns="' + OPC_NS.contentTypes + '">';
    for (const d of this.#defaults) {
      xml +=
        '<Default Extension="' +
        escapeAttribute(d.extension) +
        '" ContentType="' +
        escapeAttribute(d.contentType) +
        '"/>';
    }
    for (const o of this.#overrides) {
      xml +=
        '<Override PartName="' +
        escapeAttribute(o.partName) +
        '" ContentType="' +
        escapeAttribute(o.contentType) +
        '"/>';
    }
    return new TextEncoder().encode(xml + '</Types>');
  }
}
