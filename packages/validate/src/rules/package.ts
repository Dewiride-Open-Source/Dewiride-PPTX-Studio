import {
  checkPartNameCollisions,
  CONTENT_TYPES_PART,
  CONTENT_TYPE,
  FONT_DATA_CONTENT_TYPE,
  FONT_DATA_EXTENSION,
  partExtension,
  REL_TYPE,
  validatePartName,
} from '@pptx-studio/opc';
import { attributeValue, childElements } from '@pptx-studio/xml';
import type { Context } from './context.js';
import { elementLocation, PACKAGE_LOCATION, partLocation } from './location.js';
import { readRelsParts, resolveTarget } from './rels.js';

/**
 * `V001` … `V005`: the container.
 *
 * These five run before any markup is looked at, and they are the ones whose
 * failure mode is "PowerPoint declines to open the file at all" rather than
 * "the slide renders wrong". The package layer has no partial credit: a part
 * with no content type is not a part with an unknown type, it is a stream
 * nothing is allowed to interpret.
 */

/** Every part resolves to a content type. */
export function v001ContentTypeCoverage(ctx: Context): void {
  for (const part of ctx.parts()) {
    if (ctx.contentType(part) !== undefined) continue;
    const extension = partExtension(part);
    ctx.add(
      'V001',
      partLocation(part),
      'no content type: nothing in [Content_Types].xml covers it. Add an Override for the part, ' +
        (extension === ''
          ? 'or give it an extension a Default covers.'
          : 'or a Default for the "' + extension + '" extension.'),
    );
  }
}

/**
 * The content-type map holds together.
 *
 * Four separate things, one rule, because they are one question - "does this
 * map say exactly one thing about every part" - and a caller who has to
 * remember four rule ids to know whether their content types are sound is a
 * caller who will remember three.
 */
export function v002ContentTypeMap(ctx: Context): void {
  const document = ctx.contentTypesDocument();

  // Duplicates are only visible in the markup. See `Context.contentTypesDocument`.
  if (document === null) {
    if (ctx.archive === null) {
      // Said out loud rather than quietly skipped. The rest of this rule still
      // runs, so it is a narrower check than it looks like, and a report that
      // does not say so is a report that overstates itself.
      ctx.problem(
        CONTENT_TYPES_PART,
        'not available: only a PartStore was supplied, so V002 checked the parsed map and not ' +
          'the markup. A duplicate <Default> or <Override> is visible only in the markup.',
      );
    } else {
      ctx.add(
        'V002',
        partLocation(CONTENT_TYPES_PART),
        'the content-type stream is missing or would not parse.',
      );
    }
  } else {
    const seenDefault = new Map<string, number>();
    const seenOverride = new Map<string, number>();
    for (const child of childElements(document.root)) {
      if (child.local === 'Default') {
        const extension = (attributeValue(child, 'Extension') ?? '').toLowerCase();
        const count = (seenDefault.get(extension) ?? 0) + 1;
        seenDefault.set(extension, count);
        if (count === 2) {
          // M2.5 is unconditional and PowerPoint enforces it literally: two
          // `<Default>`s for one extension are refused with `0x80CB8000` even
          // when the two elements are byte-identical.
          ctx.add(
            'V002',
            elementLocation(CONTENT_TYPES_PART, child),
            'the "' +
              extension +
              '" extension has two <Default> entries. PowerPoint refuses a duplicate even when ' +
              'the two agree.',
          );
        }
      } else if (child.local === 'Override') {
        const partName = attributeValue(child, 'PartName') ?? '';
        const key = partName.toLowerCase();
        const count = (seenOverride.get(key) ?? 0) + 1;
        seenOverride.set(key, count);
        if (count === 2) {
          ctx.add(
            'V002',
            elementLocation(CONTENT_TYPES_PART, child),
            partName + ' has two <Override> entries.',
          );
        }
        if (partName !== '' && !ctx.store.has(partName)) {
          ctx.add(
            'V002',
            elementLocation(CONTENT_TYPES_PART, child),
            'an <Override> names ' + partName + ', which is not in the package.',
          );
        }
      }
    }
  }

  // The `fntdata` Default. Omitting it is the canonical "PowerPoint found a
  // problem with content" bug, and it is invisible until a user without the
  // font opens the file - the deck is otherwise complete and the relationship,
  // the `p:embeddedFont` entry and the part itself are all correct.
  const fontParts = ctx
    .parts()
    .filter((part) => partExtension(part).toLowerCase() === FONT_DATA_EXTENSION);
  if (fontParts.length > 0) {
    const declared = ctx.store.contentTypes.defaults.find(
      (entry) => entry.extension.toLowerCase() === FONT_DATA_EXTENSION,
    );
    if (declared === undefined) {
      ctx.add(
        'V002',
        partLocation(CONTENT_TYPES_PART),
        'the package has ' +
          String(fontParts.length) +
          ' .fntdata part(s) and no <Default Extension="fntdata" ContentType="' +
          FONT_DATA_CONTENT_TYPE +
          '"/>. PowerPoint reports the file as having a problem with its content.',
      );
    } else if (declared.contentType.toLowerCase() !== FONT_DATA_CONTENT_TYPE.toLowerCase()) {
      ctx.add(
        'V002',
        partLocation(CONTENT_TYPES_PART),
        'the fntdata Default is typed "' +
          declared.contentType +
          '"; embedded fonts are "' +
          FONT_DATA_CONTENT_TYPE +
          '".',
      );
    }
  }
}

/** No directory entries, no ZIP64 record. */
export function v003ArchiveShape(ctx: Context): void {
  const archive = ctx.archive;
  if (archive === null) return;

  for (const entry of archive.entries) {
    if (!entry.isDirectory) continue;
    ctx.add(
      'V003',
      PACKAGE_LOCATION,
      'the archive carries a directory entry, "' +
        entry.name +
        '". No OPC part name may end in a slash, so nothing in the package can name it.',
    );
  }

  if (ctx.bytes !== null && hasZip64Locator(ctx.bytes)) {
    ctx.add(
      'V003',
      PACKAGE_LOCATION,
      'the archive carries a ZIP64 end-of-central-directory locator. Our writer emits ZIP32 ' +
        'only, so one in a package we produced means a size or offset field overflowed.',
    );
  }
}

/**
 * The ZIP64 locator signature, searched backwards from the end.
 *
 * The locator sits immediately before the end-of-central-directory record,
 * which is within 64 KB + 22 bytes of the end once a comment is allowed for.
 * Same computation `@pptx-studio/census` makes; duplicated rather than shared
 * because `census` is a sibling at layer 1 and depending on it to answer one
 * boolean would put a whole feature scanner in the export path.
 */
function hasZip64Locator(bytes: Uint8Array): boolean {
  if (bytes.byteLength < 20) return false;
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const from = Math.max(0, bytes.byteLength - 0x10000 - 40);
  for (let at = bytes.byteLength - 20; at >= from; at--) {
    if (view.getUint32(at, true) === 0x07064b50) return true;
  }
  return false;
}

/**
 * Part names, with `M1.8` raised from warning to fatal.
 *
 * `@pptx-studio/opc` rates a percent-escape of an unreserved character a
 * warning, and that is right *there*: the package layer has to open what it is
 * given, and a file that arrives with `%5F` in a part name is a file the user
 * still wants to see. Here we are about to write one, and PowerPoint refuses
 * it with `0x808D1005` - measured nine for nine on one-name packages. Raising
 * the severity at the point of writing rather than at the point of reading is
 * the whole shape of the read-leniently/write-strictly split this project uses
 * everywhere.
 */
export function v004PartNames(ctx: Context): void {
  for (const part of ctx.parts()) {
    for (const violation of validatePartName(part)) {
      // Every violation, warning or not. `validatePartName` grades severity for
      // a *reader*, which has to open what it is given; here we are about to
      // write one, and a part name that is merely non-conformant is a part name
      // we should not be emitting. An odd name a deck arrived with is not
      // trapped by this - it comes back `inherited` and does not block the
      // export. See `report.ts`.
      ctx.add('V004', partLocation(part), violation.rule + ': ' + violation.message);
    }
  }

  for (const collision of checkPartNameCollisions([...ctx.parts()])) {
    ctx.add('V004', partLocation(collision.names[0]), collision.rule + ': ' + collision.message);
  }
}

/** The main presentation part: exactly one, of a type that matches the file. */
const MAIN_PART_TYPES = new Set<string>([
  CONTENT_TYPE.presentation,
  CONTENT_TYPE.template,
  CONTENT_TYPE.slideshow,
  'application/vnd.ms-powerpoint.presentation.macroEnabled.main+xml',
  'application/vnd.ms-powerpoint.template.macroEnabled.main+xml',
  'application/vnd.ms-powerpoint.slideshow.macroEnabled.main+xml',
]);

export function v005MainPart(ctx: Context): void {
  const rootRels = readRelsParts(ctx).find((rels) => rels.source === '/');
  if (rootRels === undefined) {
    ctx.add(
      'V005',
      partLocation('/_rels/.rels'),
      'the package has no root relationship part. It is the only way in to a .pptx; without it ' +
        'nothing can be reached and PowerPoint reports the file as corrupted and unreadable.',
    );
    return;
  }

  const officeDocuments = rootRels.relationships.filter(
    (rel) => rel.type === REL_TYPE.officeDocument,
  );
  if (officeDocuments.length === 0) {
    ctx.add(
      'V005',
      partLocation(rootRels.partName),
      'no officeDocument relationship. Nothing names the main presentation part.',
    );
    return;
  }
  if (officeDocuments.length > 1) {
    for (const rel of officeDocuments.slice(1)) {
      ctx.add(
        'V005',
        elementLocation(rootRels.partName, rel.element),
        'a second officeDocument relationship (' +
          rel.id +
          ' -> ' +
          rel.target +
          '). Exactly one part is the main part.',
      );
    }
  }

  const main = officeDocuments[0]!;
  const resolved = resolveTarget(rootRels.source, main.target);
  if (resolved.part === null) {
    ctx.add(
      'V005',
      elementLocation(rootRels.partName, main.element),
      'the officeDocument target does not resolve: ' + (resolved.failure ?? 'unknown reason'),
    );
    return;
  }
  if (!ctx.store.has(resolved.part)) {
    ctx.add(
      'V005',
      elementLocation(rootRels.partName, main.element),
      'the officeDocument target ' + resolved.part + ' is not in the package.',
    );
    return;
  }

  const contentType = ctx.contentType(resolved.part);
  if (contentType === undefined || !MAIN_PART_TYPES.has(contentType)) {
    ctx.add(
      'V005',
      partLocation(resolved.part),
      'the main part is typed ' +
        (contentType === undefined ? '(nothing)' : '"' + contentType + '"') +
        ', which is not a PresentationML main-part type. PowerPoint also checks this against ' +
        'the file extension and refuses the pair when they disagree, which is why a ' +
        'macro-enabled deck has to be written .pptm.',
    );
  }
}
