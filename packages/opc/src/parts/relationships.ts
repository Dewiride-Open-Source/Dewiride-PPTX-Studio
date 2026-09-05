import { OPC_NS } from '../constants.js';
import { OpcError } from '../errors.js';
import { escapeAttribute, readFlatXml, XML_DECLARATION } from './flat-xml.js';
import {
  partDirectory,
  relsPartNameFor,
  resolveRelativeTarget,
  toPartName,
  type PartName,
} from './pack-uri.js';

/**
 * Relationship parts - the edges of the package graph.
 *
 * Nothing in a `.pptx` finds anything by filename. `ppt/presentation.xml` is
 * the presentation because `_rels/.rels` says so; `slide1.xml` uses a
 * particular layout because `ppt/slides/_rels/slide1.xml.rels` says so, and
 * **the slide part itself does not name its layout anywhere**. That single fact
 * is why "change layout" later turns out to be a relationship rewrite rather
 * than an edit to the slide.
 *
 * Three properties of this file are worth stating because getting any of them
 * wrong produces a package PowerPoint refuses to open - each verified against
 * the installed PowerPoint by breaking a real deck and watching it fail:
 *
 * - **`Id` is an `xsd:ID`, so it must be an XML `NCName`.** `Id="1rId"` is not,
 *   and PowerPoint rejects the file. Not with a message about identifiers -
 *   with `0x80070570`, "corrupted and unreadable".
 * - **`Id` is unique within one `.rels` part and means nothing outside it.**
 *   Every part's relationships start again at `rId1`. A package-global
 *   allocator is not an optimisation, it is a bug that shows up as a duplicate
 *   in the one file that mattered.
 * - **Targets resolve against the folder of the *source part*, not of the
 *   `.rels` part.** For `/ppt/slides/_rels/slide1.xml.rels` the base is
 *   `/ppt/slides/`, which is what makes `../slideLayouts/slideLayout1.xml`
 *   mean the right thing. For `/_rels/.rels` the source part is the package
 *   root, so the base is `/`.
 */

export type TargetMode = 'Internal' | 'External';

/**
 * Where a relationship came from.
 *
 * The distinction is what lets the writer be strict without being wrong. A
 * relationship pointing at a part the package does not contain is not
 * automatically fatal to PowerPoint - one that nothing dereferences opens
 * fine, and real packages ship them. Refusing to export such a deck would
 * break the one property this layer exists to provide.
 *
 * So the question is not "is this edge broken" but "did we break it". An
 * `archive` relationship arrived that way and is preserved; an `added` one is
 * ours, and if it points at nothing that is a bug in the code that added it,
 * caught while the cause is still nameable.
 */
export type RelationshipOrigin = 'archive' | 'added';

export interface Relationship {
  readonly id: string;
  /** The relationship type URI, verbatim. */
  readonly type: string;
  /** The target exactly as written - relative, absolute, or an external URI. */
  readonly target: string;
  readonly targetMode: TargetMode;
  readonly origin: RelationshipOrigin;
}

/**
 * XML `NCName`: an XML `Name` with no colon.
 *
 * Written out rather than approximated with `/\w+/` because the failure is
 * silent on the way out and fatal on the way in. Astral-plane name characters
 * are omitted: no producer emits one and matching surrogate pairs here would
 * cost more than it buys.
 */
const NC_NAME_START = 'A-Z_a-z\\u00C0-\\u00D6\\u00D8-\\u00F6\\u00F8-\\u02FF\\u0370-\\u037D';
const NC_NAME_START_2 = '\\u037F-\\u1FFF\\u200C-\\u200D\\u2070-\\u218F\\u2C00-\\u2FEF';
const NC_NAME_START_3 = '\\u3001-\\uD7FF\\uF900-\\uFDCF\\uFDF0-\\uFFFD';
const NC_NAME_EXTRA = '\\-.0-9\\u00B7\\u0300-\\u036F\\u203F-\\u2040';
const START = NC_NAME_START + NC_NAME_START_2 + NC_NAME_START_3;
// `no-misleading-character-class` fires on `̀-ͯ`: a combining mark
// directly after another class member reads as one grapheme, and the rule
// cannot tell a typo from a deliberate code-point range. Here it is deliberate
// - these are the combining marks XML's own `NameChar` production admits, so a
// name in Devanagari or Vietnamese is a legal xsd:ID and we say so. The `u`
// flag makes the range unambiguous to the engine even where it looks ambiguous
// to a reader.
// eslint-disable-next-line no-misleading-character-class -- XML NameChar, not a grapheme
const NC_NAME = new RegExp('^[' + START + '][' + START + NC_NAME_EXTRA + ']*$', 'u');

/** True when `id` is a legal `xsd:ID`. */
export function isValidRelationshipId(id: string): boolean {
  return NC_NAME.test(id);
}

const R_ID = /^rId(\d+)$/;

/**
 * The part whose relationships a `.rels` part describes.
 *
 * The inverse of `relsPartNameFor`. `/ppt/slides/_rels/slide1.xml.rels` ->
 * `/ppt/slides/slide1.xml`, and `/_rels/.rels` -> `/`, the package root.
 */
export function sourcePartNameForRels(relsPartName: string): string {
  const m = /^(\/(?:.*\/)?)_rels\/(.*)\.rels$/.exec(relsPartName);
  if (m === null) {
    throw new OpcError(
      'ERR_INVALID_PART_NAME',
      JSON.stringify(relsPartName) + ' is not a relationship part name',
      { entry: relsPartName },
    );
  }
  const [, dir, base] = m;
  return base === '' ? '/' : dir! + base!;
}

/**
 * A relative `Target` from `sourcePartName` to `targetPartName`.
 *
 * Office writes relative targets (`../slideLayouts/slideLayout1.xml`) and so do
 * we for anything new, but absolute ones are legal and common in the wild -
 * roughly three quarters of the relationships in one third-party corpus we
 * measured are written as `/ppt/...`. Both are read; only this form is written.
 */
export function relativeTargetFor(sourcePartName: string, targetPartName: string): string {
  const from = partDirectory(sourcePartName).split('/').filter(Boolean);
  const to = targetPartName.split('/').filter(Boolean);
  let shared = 0;
  while (shared < from.length && shared < to.length - 1 && from[shared] === to[shared]) shared++;
  const up = '../'.repeat(from.length - shared);
  return up + to.slice(shared).join('/');
}

export class Relationships {
  /** The part these relationships belong to. `/` for the package root. */
  readonly sourcePartName: string;
  /** The `.rels` part itself. */
  readonly partName: PartName;

  #items: Relationship[] = [];
  #byId = new Map<string, number>();
  #dirty = false;

  private constructor(sourcePartName: string) {
    this.sourcePartName = sourcePartName;
    this.partName = relsPartNameFor(sourcePartName);
  }

  /** An empty collection for a part that has no `.rels` yet. */
  static empty(sourcePartName: string): Relationships {
    return new Relationships(sourcePartName);
  }

  static parse(bytes: Uint8Array, sourcePartName: string): Relationships {
    const rels = new Relationships(sourcePartName);
    const context = rels.partName;
    const elements = readFlatXml(bytes, context);
    const root = elements[0]!;
    if (root.name !== 'Relationships') {
      throw new OpcError(
        'ERR_MALFORMED_XML',
        context + ': root element is <' + root.qname + '>, expected <Relationships>',
        { entry: context },
      );
    }

    for (const el of elements) {
      if (el.depth !== 1 || el.name !== 'Relationship') continue;
      const id = el.attrs.get('Id');
      const type = el.attrs.get('Type');
      const target = el.attrs.get('Target');
      if (id === undefined || type === undefined || target === undefined) {
        throw new OpcError(
          'ERR_MALFORMED_XML',
          context + ': <Relationship> needs Id, Type and Target',
          { entry: context },
        );
      }
      const rawMode = el.attrs.get('TargetMode');
      if (rawMode !== undefined && rawMode !== 'Internal' && rawMode !== 'External') {
        throw new OpcError(
          'ERR_MALFORMED_XML',
          context + ': TargetMode="' + rawMode + '" is neither Internal nor External',
          { entry: context },
        );
      }
      rels.#insert(
        { id, type, target, targetMode: rawMode ?? 'Internal', origin: 'archive' },
        context,
        /* dirty */ false,
      );
    }
    return rels;
  }

  get all(): readonly Relationship[] {
    return this.#items;
  }

  get dirty(): boolean {
    return this.#dirty;
  }

  get size(): number {
    return this.#items.length;
  }

  byId(id: string): Relationship | undefined {
    const at = this.#byId.get(id);
    return at === undefined ? undefined : this.#items[at];
  }

  byType(type: string): Relationship[] {
    return this.#items.filter((r) => r.type === type);
  }

  firstOfType(type: string): Relationship | undefined {
    return this.#items.find((r) => r.type === type);
  }

  /**
   * The part a relationship points at.
   *
   * Internal only - an external target is a URI naming something outside the
   * package and has no part name, so asking for one is a caller bug rather
   * than a data problem.
   */
  resolve(rel: Relationship): PartName {
    if (rel.targetMode === 'External') {
      throw new OpcError(
        'ERR_INVALID_PART_NAME',
        'relationship ' +
          rel.id +
          ' in ' +
          this.partName +
          ' is TargetMode="External"; ' +
          JSON.stringify(rel.target) +
          ' names something outside the package and has no part name',
        { entry: this.partName },
      );
    }
    return resolveRelativeTarget(this.sourcePartName, rel.target);
  }

  /** The part `id` points at, or `undefined` if there is no such internal relationship. */
  targetOf(id: string): PartName | undefined {
    const rel = this.byId(id);
    if (rel === undefined || rel.targetMode === 'External') return undefined;
    return this.resolve(rel);
  }

  #insert(rel: Relationship, context: string, dirty: boolean): Relationship {
    if (!isValidRelationshipId(rel.id)) {
      throw new OpcError(
        'ERR_INVALID_RELATIONSHIP_ID',
        context +
          ': Id=' +
          JSON.stringify(rel.id) +
          ' is not an xsd:ID. An xsd:ID is an XML NCName, so it cannot begin with a digit or ' +
          'contain a space or a colon. PowerPoint refuses the whole package over this.',
        { entry: context },
      );
    }
    if (this.#byId.has(rel.id)) {
      throw new OpcError(
        'ERR_DUPLICATE_RELATIONSHIP_ID',
        context +
          ': Id=' +
          JSON.stringify(rel.id) +
          ' appears twice. Ids are unique within one relationship part, and which target an ' +
          'r:id refers to would otherwise be undecidable.',
        { entry: context },
      );
    }
    if (rel.targetMode === 'Internal' && rel.target === '') {
      throw new OpcError('ERR_MALFORMED_XML', context + ': ' + rel.id + ' has an empty Target', {
        entry: context,
      });
    }
    this.#byId.set(rel.id, this.#items.length);
    this.#items.push(rel);
    if (dirty) this.#dirty = true;
    return rel;
  }

  /**
   * The next free `rIdN`.
   *
   * One past the highest number already used, **not** one past the count. Those
   * differ more often than you would expect: of 1296 relationship parts in the
   * corpus we measured, 99 have gaps in their numbering and 175 list their
   * relationships out of numeric order. Allocating from the count would have
   * collided in every one of them.
   */
  nextId(): string {
    let highest = 0;
    for (const rel of this.#items) {
      const m = R_ID.exec(rel.id);
      if (m !== null) highest = Math.max(highest, Number(m[1]));
    }
    let candidate = highest + 1;
    while (this.#byId.has('rId' + String(candidate))) candidate++;
    return 'rId' + String(candidate);
  }

  /** Add a relationship with a freshly allocated id. */
  add(type: string, target: string, targetMode: TargetMode = 'Internal'): Relationship {
    return this.#insert(
      { id: this.nextId(), type, target, targetMode, origin: 'added' },
      this.partName,
      true,
    );
  }

  /** Add a relationship to a part, computing the relative target for you. */
  addTo(type: string, targetPartName: string): Relationship {
    return this.add(type, relativeTargetFor(this.sourcePartName, toPartName(targetPartName)));
  }

  /** Add with a caller-chosen id - for reproducing an existing package exactly. */
  addWithId(
    id: string,
    type: string,
    target: string,
    targetMode: TargetMode = 'Internal',
  ): Relationship {
    return this.#insert({ id, type, target, targetMode, origin: 'added' }, this.partName, true);
  }

  remove(id: string): boolean {
    const at = this.#byId.get(id);
    if (at === undefined) return false;
    this.#items.splice(at, 1);
    this.#byId.delete(id);
    for (const [k, v] of this.#byId) {
      if (v > at) this.#byId.set(k, v - 1);
    }
    this.#dirty = true;
    return true;
  }

  /**
   * Serialise in Office's layout: declaration, `Relationships`, then each
   * `Relationship` in document order with attributes ordered Id, Type, Target,
   * TargetMode. `TargetMode` is written only when External, because `Internal`
   * is the schema default and Office omits it.
   */
  serialize(): Uint8Array {
    let xml = XML_DECLARATION + '<Relationships xmlns="' + OPC_NS.relationships + '">';
    for (const r of this.#items) {
      xml +=
        '<Relationship Id="' +
        escapeAttribute(r.id) +
        '" Type="' +
        escapeAttribute(r.type) +
        '" Target="' +
        escapeAttribute(r.target) +
        '"' +
        (r.targetMode === 'External' ? ' TargetMode="External"' : '') +
        '/>';
    }
    return new TextEncoder().encode(xml + '</Relationships>');
  }
}
