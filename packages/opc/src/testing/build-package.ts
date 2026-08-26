import { CONTENT_TYPE, CONTENT_TYPES_PART, REL_TYPE } from '../constants.js';
import { XML_DECLARATION } from '../flat-xml.js';
import { deflatedEntry, writeZip, type ZipEntryInput } from '../zip-writer.js';

/**
 * A minimal OPC package, and every way of breaking one.
 *
 * Test-only, and never reachable from `src/index.ts` so it cannot ship. The
 * companion to `build-zip.ts`: that one builds archives that lie about
 * themselves at the ZIP layer, this one builds packages that are structurally
 * wrong at the OPC layer. Both exist because the real corpus is clean - 37
 * packages written by two different producers contain not one dangling
 * relationship, duplicate id, missing content type or malformed part name
 * between them, so every defence in this package is untested by real files.
 */

const encoder = new TextEncoder();

export interface PackagePart {
  /** ZIP entry name, no leading slash. */
  readonly name: string;
  readonly content: string | Uint8Array;
}

export interface PackageFixture {
  /** Replaces the generated `[Content_Types].xml` wholesale. */
  readonly contentTypesXml?: string;
  /** Replaces the generated `_rels/.rels` wholesale. */
  readonly rootRelsXml?: string;
  /** Parts added on top of the minimal set. */
  readonly parts?: readonly PackagePart[];
  /** Entry names to leave out. */
  readonly drop?: readonly string[];
  /** Rearrange the final entry order. */
  readonly order?: (names: string[]) => string[];
}

export const MINIMAL_CONTENT_TYPES =
  XML_DECLARATION +
  '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
  '<Default Extension="rels" ContentType="' +
  CONTENT_TYPE.relationships +
  '"/>' +
  '<Default Extension="xml" ContentType="' +
  CONTENT_TYPE.xml +
  '"/>' +
  '<Override PartName="/ppt/presentation.xml" ContentType="' +
  CONTENT_TYPE.presentation +
  '"/>' +
  '</Types>';

export const MINIMAL_ROOT_RELS =
  XML_DECLARATION +
  '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
  '<Relationship Id="rId1" Type="' +
  REL_TYPE.officeDocument +
  '" Target="ppt/presentation.xml"/>' +
  '</Relationships>';

/** A stub presentation part. Nothing at the OPC layer reads inside it. */
export const MINIMAL_PRESENTATION =
  XML_DECLARATION +
  '<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>';

/** A `.rels` document with the given `<Relationship>` bodies. */
export function relsXml(...relationships: readonly string[]): string {
  return (
    XML_DECLARATION +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    relationships.join('') +
    '</Relationships>'
  );
}

/** One `<Relationship>` element, with no validation of any kind. */
export function relXml(
  id: string,
  type: string,
  target: string,
  targetMode?: 'Internal' | 'External',
): string {
  return (
    '<Relationship Id="' +
    id +
    '" Type="' +
    type +
    '" Target="' +
    target +
    '"' +
    (targetMode === undefined ? '' : ' TargetMode="' + targetMode + '"') +
    '/>'
  );
}

/** A `[Content_Types].xml` from explicit `<Default>` and `<Override>` bodies. */
export function contentTypesXml(...entries: readonly string[]): string {
  return (
    XML_DECLARATION +
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    entries.join('') +
    '</Types>'
  );
}

export function buildPackage(fixture: PackageFixture = {}): Uint8Array {
  const drop = new Set(fixture.drop ?? []);
  const base: PackagePart[] = [
    { name: CONTENT_TYPES_PART, content: fixture.contentTypesXml ?? MINIMAL_CONTENT_TYPES },
    { name: '_rels/.rels', content: fixture.rootRelsXml ?? MINIMAL_ROOT_RELS },
    { name: 'ppt/presentation.xml', content: MINIMAL_PRESENTATION },
  ];

  const byName = new Map<string, PackagePart>();
  for (const part of base) byName.set(part.name, part);
  for (const part of fixture.parts ?? []) byName.set(part.name, part);

  let names = [...byName.keys()].filter((n) => !drop.has(n));
  if (fixture.order !== undefined) names = fixture.order(names);

  const entries: ZipEntryInput[] = names.map((name) => {
    const content = byName.get(name)!.content;
    const bytes = typeof content === 'string' ? encoder.encode(content) : content;
    return deflatedEntry(name, bytes);
  });
  return writeZip(entries);
}
