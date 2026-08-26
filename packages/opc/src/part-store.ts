import { CONTENT_TYPE, CONTENT_TYPES_PART, ROOT_RELS_PART } from './constants.js';
import { ContentTypes } from './content-types.js';
import { OpcError } from './errors.js';
import {
  checkPartNameCollisions,
  isContentTypesStreamName,
  isRelationshipPartName,
  normalizePartName,
  partNameFromZipEntry,
  relsPartNameFor,
  toPartName,
  validatePartName,
  zipEntryNameFor,
  type PartName,
} from './pack-uri.js';
import { Relationships, sourcePartNameForRels, type Relationship } from './relationships.js';
import { readZip, type ReadZipOptions, type ZipArchive, type ZipEntry } from './zip-reader.js';
import { deflatedEntry, passthroughEntry, writeZip, type ZipEntryInput } from './zip-writer.js';

/**
 * The package: every part, what each one is, and how they refer to each other.
 *
 * This is the first architectural bet made concrete. **The package is the
 * document.** A part is a name, a content type and some bytes; it is not
 * parsed until something reads it, and an untouched part keeps its original
 * DEFLATE stream all the way through an export without ever being inflated.
 * That is what makes preserving charts, SmartArt, animations and OLE the
 * default state rather than a feature - we do not have to understand a part to
 * carry it across, and so there is no category of content we silently drop.
 *
 * What the store deliberately does not do, because each belongs to a later
 * sub-phase and doing it early would be worse than not doing it:
 *
 * - **It does not parse part content.** That is `@pptx-studio/xml` in 0.4. The
 *   two package-level grammars are read by `flat-xml.ts`, which is not a
 *   general XML parser and is not trying to be.
 * - **It does not collect garbage.** Removing a part removes the part, its
 *   relationship part and its content-type entry, and nothing else. A
 *   relationship left pointing at it becomes a dangling one, and `write`
 *   refuses to emit that - loudly, naming the part - because *we* broke it.
 *   A relationship that was already dangling when the package was opened is
 *   preserved instead, and reported by `danglingRelationships`; PowerPoint
 *   tolerates those and refusing them would make such a deck unexportable.
 *   Deciding what is safe to delete is 1.3's job and it needs the whole graph.
 * - **It does not validate markup.** The 29 rules are 1.2. What `write`
 *   asserts here is only the OPC layer, and only the parts of it we have
 *   watched PowerPoint reject.
 */

/** What the store knows about a part without reading it. */
export interface PartInfo {
  readonly name: PartName;
  readonly contentType: string;
  /** Uncompressed size in bytes. */
  readonly size: number;
  /** False once the part's bytes have been replaced or it was added here. */
  readonly fromArchive: boolean;
}

export interface WritePackageOptions {
  /**
   * Move `[Content_Types].xml` to the front and `_rels/.rels` behind it,
   * the layout Office writes. Off by default - see `write`.
   */
  readonly normalizeEntryOrder?: boolean;
  /** DEFLATE level for parts we serialise ourselves. */
  readonly deflateLevel?: number;
}

type PartSource =
  | { readonly kind: 'archive'; readonly entry: ZipEntry }
  | { readonly kind: 'bytes'; readonly bytes: Uint8Array };

interface Part {
  readonly name: PartName;
  source: PartSource;
}

export class PartStore {
  /** The content-type map. Mutating it is how a part's type changes. */
  readonly contentTypes: ContentTypes;

  #archive: ZipArchive | undefined;
  /** ZIP entry names in the order they will be written. */
  #order: string[];
  /** Normalised part name -> part. */
  #parts = new Map<string, Part>();
  /** Normalised *source* part name -> its relationships, once anyone has asked. */
  #rels = new Map<string, Relationships>();
  #contentTypesEntry: ZipEntry | undefined;
  /**
   * Normalised names of parts this session removed.
   *
   * Kept so `write` can tell "this edge was already broken when we opened the
   * file" from "we broke this edge by deleting its target". The first is
   * preserved, the second is refused. A set of names costs nothing; the
   * alternative - walking every `.rels` on every removal to sever inbound
   * edges - would parse hundreds of parts per delete.
   */
  #removed = new Set<string>();

  private constructor(contentTypes: ContentTypes, order: string[]) {
    this.contentTypes = contentTypes;
    this.#order = order;
  }

  /** A package with nothing in it but an empty content-type map. */
  static create(): PartStore {
    return new PartStore(ContentTypes.empty(), [CONTENT_TYPES_PART]);
  }

  /** Open an archive. Nothing is inflated but `[Content_Types].xml`. */
  static open(bytes: Uint8Array, options: ReadZipOptions = {}): PartStore {
    const archive = readZip(bytes, options);

    // Located case-insensitively, not by exact name. ZIP entry names are
    // case-sensitive and the spec fixes the spelling, but PowerPoint opens a
    // package whose stream is named `[content_types].xml` - so refusing one
    // would fail a file that works everywhere else, for a reason no user could
    // act on. It is written back under the canonical spelling: read leniently,
    // write strictly, the same split as everywhere else here.
    const contentTypesEntry = archive.entries.find((e) => isContentTypesStreamName(e.name));
    if (contentTypesEntry === undefined) {
      throw new OpcError(
        'ERR_MISSING_PACKAGE_PART',
        'this archive has no ' +
          CONTENT_TYPES_PART +
          '. Without it no part has a content type, which is the one thing an OPC package ' +
          'cannot do without - PowerPoint refuses such a file outright.',
        { entry: CONTENT_TYPES_PART },
      );
    }
    const contentTypes = ContentTypes.parse(archive.read(contentTypesEntry));

    const order: string[] = [];
    const store = new PartStore(contentTypes, order);
    store.#archive = archive;
    store.#contentTypesEntry = contentTypesEntry;

    const names: string[] = [];
    for (const entry of archive.entries) {
      // A directory entry is not a part - no OPC part name can end in a slash.
      // Real Office packages contain none; PowerPoint tolerates them in files
      // that do, and we drop them rather than carry a thing with no meaning.
      if (entry.isDirectory) continue;
      // The one place a name is normalised on the way in, so that everything
      // downstream can compare against `CONTENT_TYPES_PART` and be right.
      if (isContentTypesStreamName(entry.name)) {
        order.push(CONTENT_TYPES_PART);
        continue;
      }
      order.push(entry.name);
      const name = partNameFromZipEntry(entry.name);
      const key = normalizePartName(name);
      if (store.#parts.has(key)) {
        throw new OpcError('ERR_DUPLICATE_ENTRY', 'two entries resolve to the part name ' + name, {
          entry: entry.name,
        });
      }
      store.#parts.set(key, { name, source: { kind: 'archive', entry } });
      names.push(name);
    }

    const collisions = checkPartNameCollisions(names);
    if (collisions.length > 0) {
      throw new OpcError(
        'ERR_INVALID_PART_NAME',
        'part names collide: ' + collisions.map((c) => c.message).join('; '),
      );
    }
    return store;
  }

  /** Part names in write order. Excludes `[Content_Types].xml`, which is not a part. */
  get partNames(): PartName[] {
    const out: PartName[] = [];
    for (const entryName of this.#entryNames()) {
      if (entryName === CONTENT_TYPES_PART) continue;
      const part = this.#parts.get(normalizePartName('/' + entryName));
      if (part !== undefined) out.push(part.name);
    }
    return out;
  }

  get size(): number {
    return this.#parts.size;
  }

  has(partName: string): boolean {
    return this.#parts.has(normalizePartName(partName));
  }

  #require(partName: string): Part {
    const part = this.#parts.get(normalizePartName(partName));
    if (part === undefined) {
      throw new OpcError('ERR_PART_NOT_FOUND', 'no part named ' + partName, { entry: partName });
    }
    return part;
  }

  contentTypeOf(partName: string): string | undefined {
    return this.contentTypes.for(partName);
  }

  info(partName: string): PartInfo | undefined {
    const part = this.#parts.get(normalizePartName(partName));
    if (part === undefined) return undefined;
    return {
      name: part.name,
      contentType: this.contentTypes.for(part.name) ?? '',
      size:
        part.source.kind === 'archive'
          ? part.source.entry.uncompressedSize
          : part.source.bytes.length,
      fromArchive: part.source.kind === 'archive',
    };
  }

  /**
   * The bytes of a part.
   *
   * An archive-backed part is inflated on every call and its CRC-32 checked on
   * every call; the decompression budget is charged once. Nothing is cached
   * here on purpose - what gets cached one layer up is the *parsed* form, and
   * holding both would double the memory of a 200 MB deck to keep a copy
   * nobody reads twice.
   */
  read(partName: string): Uint8Array {
    const part = this.#require(partName);
    if (part.source.kind === 'bytes') return part.source.bytes;
    return this.#archive!.read(part.source.entry);
  }

  /**
   * The relationships of a part, parsed on first ask and cached after.
   *
   * Pass `'/'` for the package root, whose relationships live in
   * `_rels/.rels`. A part with no relationship part gets an empty collection
   * rather than `undefined`, so callers add to it without a special case; an
   * empty, untouched collection is never written.
   */
  relationships(sourcePartName: string): Relationships {
    const source = sourcePartName === '/' ? '/' : toPartName(sourcePartName);
    const key = normalizePartName(source);
    const cached = this.#rels.get(key);
    if (cached !== undefined) return cached;

    const relsName = relsPartNameFor(source);
    const rels = this.has(relsName)
      ? Relationships.parse(this.read(relsName), source)
      : Relationships.empty(source);
    this.#rels.set(key, rels);
    return rels;
  }

  /** The package-level relationships: the only way in to a `.pptx`. */
  rootRelationships(): Relationships {
    return this.relationships('/');
  }

  /**
   * Every internal relationship that points at a part this package does not
   * contain.
   *
   * `write` preserves the ones that arrived broken rather than refusing to
   * export the deck - see `#assertRelationshipsResolve` for why. That makes
   * this the only way to know they are there, and they are worth knowing
   * about: a stale image relationship is exactly the defect that shows a user
   * a red X and a message about a missing part, with nothing in the file to
   * explain it.
   *
   * Reads and parses every relationship part, so it is a deliberate act rather
   * than something `open` does for you. Never throws.
   */
  danglingRelationships(): {
    source: PartName;
    relationship: Relationship;
    target: string;
  }[] {
    const found: { source: PartName; relationship: Relationship; target: string }[] = [];
    for (const partName of this.partNames) {
      if (!isRelationshipPartName(partName)) continue;
      const source = sourcePartNameForRels(partName);
      let rels: Relationships;
      try {
        rels = this.relationships(source);
      } catch {
        continue;
      }
      for (const rel of rels.all) {
        if (rel.targetMode === 'External') continue;
        try {
          const target = rels.resolve(rel);
          if (!this.has(target)) found.push({ source: partName, relationship: rel, target });
        } catch {
          found.push({ source: partName, relationship: rel, target: rel.target });
        }
      }
    }
    return found;
  }

  // --- mutation ------------------------------------------------------------

  /** Add a part that is not already there. */
  addPart(partName: string, contentType: string, bytes: Uint8Array): PartName {
    const name = toPartName(partName);
    const key = normalizePartName(name);
    if (this.#parts.has(key)) {
      throw new OpcError('ERR_PART_EXISTS', 'the package already has a part named ' + name, {
        entry: name,
      });
    }
    const collisions = checkPartNameCollisions([...this.partNames, name]);
    if (collisions.length > 0) {
      throw new OpcError('ERR_INVALID_PART_NAME', collisions[0]!.message, { entry: name });
    }
    this.#parts.set(key, { name, source: { kind: 'bytes', bytes } });
    this.#order.push(zipEntryNameFor(name));
    this.contentTypes.ensureFor(name, contentType);
    return name;
  }

  /** Replace a part's bytes, keeping its name, content type and position. */
  replacePart(partName: string, bytes: Uint8Array): void {
    const part = this.#require(partName);
    part.source = { kind: 'bytes', bytes };
    if (isRelationshipPartName(part.name)) {
      // Its parsed form is now stale; drop it so the next reader re-parses.
      this.#rels.delete(normalizePartName(sourcePartNameForRels(part.name)));
    }
  }

  /**
   * Remove a part, its relationship part and its content-type `Override`.
   *
   * Nothing scans for relationships that pointed at it. That is not an
   * oversight: finding them means walking every `.rels` in the package, and
   * deciding what to do about each is a question about the document rather
   * than the container. `write` refuses to emit a dangling relationship, so
   * the mistake surfaces at export with the part named, rather than as a file
   * PowerPoint declines to open.
   */
  removePart(partName: string): boolean {
    const key = normalizePartName(partName);
    const part = this.#parts.get(key);
    if (part === undefined) return false;

    this.#parts.delete(key);
    this.#removed.add(key);
    this.#order = this.#order.filter((n) => normalizePartName('/' + n) !== key);
    this.contentTypes.removeOverride(part.name);
    this.#rels.delete(key);

    const relsName = relsPartNameFor(part.name);
    if (this.has(relsName)) this.removePart(relsName);
    return true;
  }

  // --- writing -------------------------------------------------------------

  /** Entry names in write order, including parts added since the archive was read. */
  #entryNames(): string[] {
    const out: string[] = [];
    const seen = new Set<string>();
    for (const name of this.#order) {
      if (seen.has(name)) continue;
      seen.add(name);
      out.push(name);
    }
    // A relationship collection that grew from nothing has no entry yet.
    for (const rels of this.#rels.values()) {
      if (!rels.dirty || rels.size === 0) continue;
      const entryName = zipEntryNameFor(rels.partName);
      if (seen.has(entryName)) continue;
      seen.add(entryName);
      out.push(entryName);
    }
    return out;
  }

  /**
   * Serialise the package.
   *
   * **Entry order is preserved, not normalised.** The plan called for
   * `[Content_Types].xml` first and `_rels/.rels` second, which is what Office
   * writes - but it is a convention rather than a requirement, and we checked:
   * a real deck rebuilt with the content-type stream written *last* opens in
   * PowerPoint with every slide, shape and picture intact. Meanwhile roughly
   * four fifths of the third-party packages we measured do not put it first
   * either. Reordering them would buy nothing measurable and would break the
   * one property this sub-phase exists to establish, that a package we did not
   * edit comes back out exactly as it went in. Pass `normalizeEntryOrder` to
   * get Office's layout anyway; a package built with `create` gets it already,
   * since it has no original order to preserve.
   */
  write(options: WritePackageOptions = {}): Uint8Array {
    const relsBytes = this.#serializeDirtyRelationships();
    const entryNames = this.#orderedEntryNames(options.normalizeEntryOrder ?? false, relsBytes);
    this.#assertWritable(entryNames, relsBytes);

    const level = options.deflateLevel ?? 6;
    const entries: ZipEntryInput[] = [];
    for (const entryName of entryNames) {
      if (entryName === CONTENT_TYPES_PART) {
        entries.push(this.#contentTypesEntryInput(level));
        continue;
      }
      const key = normalizePartName('/' + entryName);
      const replacement = relsBytes.get(key);
      if (replacement !== undefined) {
        entries.push(deflatedEntry(entryName, replacement, level));
        continue;
      }
      const part = this.#parts.get(key);
      if (part === undefined) continue;
      entries.push(
        part.source.kind === 'archive'
          ? passthroughEntry(this.#archive!, part.source.entry)
          : deflatedEntry(entryName, part.source.bytes, level),
      );
    }
    return writeZip(entries);
  }

  #contentTypesEntryInput(level: number): ZipEntryInput {
    if (!this.contentTypes.dirty && this.#contentTypesEntry !== undefined) {
      // The canonical name, even when the archive spelled it differently. The
      // bytes are untouched; only the entry name is normalised.
      return {
        ...passthroughEntry(this.#archive!, this.#contentTypesEntry),
        name: CONTENT_TYPES_PART,
      };
    }
    return deflatedEntry(CONTENT_TYPES_PART, this.contentTypes.serialize(), level);
  }

  /**
   * Serialise every relationship collection that changed.
   *
   * A collection emptied of everything loses its part rather than being written
   * as an empty `<Relationships/>` - Office does not keep those around, and a
   * part that exists only to say nothing is one more thing to explain.
   */
  #serializeDirtyRelationships(): Map<string, Uint8Array> {
    const out = new Map<string, Uint8Array>();
    for (const rels of this.#rels.values()) {
      if (!rels.dirty) continue;
      const key = normalizePartName(rels.partName);
      if (rels.size === 0) {
        this.#parts.delete(key);
        this.contentTypes.removeOverride(rels.partName);
        continue;
      }
      this.contentTypes.ensureFor(rels.partName, CONTENT_TYPE.relationships);
      out.set(key, rels.serialize());
    }
    return out;
  }

  #orderedEntryNames(normalize: boolean, relsBytes: ReadonlyMap<string, Uint8Array>): string[] {
    const live = this.#entryNames().filter((entryName) => {
      if (entryName === CONTENT_TYPES_PART) return true;
      const key = normalizePartName('/' + entryName);
      return this.#parts.has(key) || relsBytes.has(key);
    });
    if (!normalize) return live;
    const first: string[] = [CONTENT_TYPES_PART, zipEntryNameFor(ROOT_RELS_PART as PartName)];
    return [...first.filter((n) => live.includes(n)), ...live.filter((n) => !first.includes(n))];
  }

  /**
   * The assertions that stand between us and a file PowerPoint refuses.
   *
   * Each one corresponds to a package we built deliberately broken and watched
   * the installed PowerPoint reject. They are not configurable, for the same
   * reason the CRC check in the reader is not: a flag that turns off a
   * correctness check is a flag somebody sets in a hot path.
   */
  #assertWritable(entryNames: readonly string[], relsBytes: ReadonlyMap<string, Uint8Array>): void {
    const present = new Set<string>();
    for (const entryName of entryNames) present.add(normalizePartName('/' + entryName));
    this.#assertPackagePartsPresent(entryNames, present);
    this.#assertEveryPartIsTypedAndNameable(entryNames);
    this.#assertRelationshipsResolve(entryNames, present, relsBytes);
  }

  /** `[Content_Types].xml` and `_rels/.rels`: neither is optional. */
  #assertPackagePartsPresent(entryNames: readonly string[], present: ReadonlySet<string>): void {
    if (!entryNames.includes(CONTENT_TYPES_PART)) {
      throw new OpcError('ERR_MISSING_PACKAGE_PART', 'the package has no ' + CONTENT_TYPES_PART, {
        entry: CONTENT_TYPES_PART,
      });
    }
    if (!present.has(normalizePartName(ROOT_RELS_PART))) {
      throw new OpcError(
        'ERR_MISSING_PACKAGE_PART',
        'the package has no ' +
          ROOT_RELS_PART +
          '. Nothing could find the presentation inside it; PowerPoint reports such a file as ' +
          'corrupted and unreadable.',
        { entry: ROOT_RELS_PART },
      );
    }
  }

  /** Every part carries a content type, and a name PowerPoint will accept. */
  #assertEveryPartIsTypedAndNameable(entryNames: readonly string[]): void {
    for (const entryName of entryNames) {
      if (entryName === CONTENT_TYPES_PART) continue;
      const partName = '/' + entryName;

      // Every part must have a content type. This is the one PowerPoint refuses
      // with its own error code rather than the generic corruption one.
      this.contentTypes.require(partName);

      // And every part name must be fully conformant, not merely non-fatal.
      // ADR 0002 rated a non-`pchar` character a warning, reasoning that a
      // media part named `my image.png` is out of spec and harmless. The first
      // half is right and the second is not: PowerPoint refuses a package
      // containing that name outright, as it does for `#` and for any non-ASCII
      // byte - while accepting `%20`, which is the same name spelled properly.
      // So the grammar stays lenient for reading and the writer does not.
      const nonPchar = validatePartName(partName).filter((v) => v.rule === 'M1.6');
      if (nonPchar.length > 0) {
        throw new OpcError(
          'ERR_INVALID_PART_NAME',
          partName +
            ' cannot be written: ' +
            nonPchar[0]!.message +
            '. Percent-encode it - PowerPoint accepts %20 and refuses the literal character.',
          { entry: partName },
        );
      }
    }
  }

  /** Every internal relationship lands on a part that is actually there. */
  #assertRelationshipsResolve(
    entryNames: readonly string[],
    present: ReadonlySet<string>,
    relsBytes: ReadonlyMap<string, Uint8Array>,
  ): void {
    for (const entryName of entryNames) {
      const partName = '/' + entryName;
      if (!isRelationshipPartName(partName)) continue;
      const key = normalizePartName(partName);
      const source = sourcePartNameForRels(partName);
      const rels =
        this.#rels.get(normalizePartName(source)) ??
        Relationships.parse(relsBytes.get(key) ?? this.read(partName), source);
      for (const rel of rels.all) {
        if (rel.targetMode === 'External') continue;
        const target = rels.resolve(rel);

        // A relationship may never point at a relationship part. The spec says
        // implementers "shall treat any such relationship as invalid", and
        // PowerPoint agrees by refusing the package.
        if (isRelationshipPartName(target)) {
          throw new OpcError(
            'ERR_INVALID_RELATIONSHIP_TARGET',
            partName +
              ': ' +
              rel.id +
              ' targets the relationship part ' +
              target +
              '. Relationship parts are reached by naming convention, never by relationship, and ' +
              'a package that claims otherwise is invalid.',
            { entry: partName },
          );
        }

        if (present.has(normalizePartName(target))) continue;

        // The edge is broken. Whether that is fatal depends on who broke it.
        //
        // PowerPoint tolerates a dangling relationship nothing dereferences -
        // we added one to a real deck and it opened with every slide and
        // picture intact - and refuses one that something follows. We cannot
        // tell those apart, because deciding it means reading `r:id` out of
        // part markup and this layer cannot see inside a part until 0.4.
        //
        // Refusing all of them was the first thing this code did, and it was
        // wrong: it makes a deck that arrived with a stale relationship
        // impossible to export, which is precisely the preservation this
        // package exists to provide. Refusing none of them is also wrong,
        // because then a delete silently produces the broken file.
        //
        // So the test is not "is this edge broken" but "did we break it".
        const weBrokeIt = rel.origin === 'added' || this.#removed.has(normalizePartName(target));
        if (!weBrokeIt) continue;

        throw new OpcError(
          'ERR_DANGLING_RELATIONSHIP',
          partName +
            ': ' +
            rel.id +
            ' targets ' +
            target +
            ', which this package does not contain' +
            (rel.origin === 'added'
              ? '. The relationship was added here, so the part it names should have been too.'
              : '. That part was removed from this package while the relationship pointing at it ' +
                'was left in place.'),
          { entry: partName },
        );
      }
    }
  }
}
