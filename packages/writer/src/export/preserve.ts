import {
  CONTENT_TYPES_PART,
  isContentTypesStreamName,
  normalizePartName,
  readZip,
  type ReadZipOptions,
  type ZipEntry,
} from '@pptx-studio/opc';
import { WriterError } from '../errors.js';

/**
 * The one assertion the writer makes about its own output.
 *
 * A part nobody edited must come out with the same bytes it went in with. That
 * is the first architectural bet stated as a testable sentence, and it is worth
 * a check of its own even though `V027` is nominally about the same thing,
 * because the two are looking at different objects.
 *
 * `V027` asks the **store**: it compares what the store would hand you for a
 * part against what the baseline store would hand you, and for an untouched
 * part both answers come from the same archive. So it can catch a part that was
 * replaced without being marked edited, and it cannot - even in principle -
 * catch a writer that re-serialised a clean part on the way out, because by the
 * time the bytes exist the rule has already run against a store that knows
 * nothing about them.
 *
 * This check reads the archive we actually emitted. It compares the **stored**
 * bytes, still compressed, against the stored bytes of the source archive:
 * same DEFLATE stream, same CRC, same declared sizes. Comparing compressed
 * bytes is both cheaper - nothing is inflated on either side - and stricter,
 * because two different DEFLATE encodings of identical content would pass an
 * inflate-and-compare and fail this. Stricter is what is wanted: any difference
 * at all means something re-compressed, and re-compressing is the failure.
 *
 * It needs the bytes the package was opened from. Without them there is nothing
 * to compare against and the check reports itself as not run, which is not the
 * same as reporting that it passed.
 */

export interface PreservationCheck {
  /** Entries compared and found identical. */
  readonly checked: number;
  /** Entries deliberately not compared, because this session wrote them. */
  readonly rewritten: number;
  /** Why the check did not run, or `null` when it did. */
  readonly skipped: string | null;
}

/** The key an entry is matched on across the two archives. */
function entryKey(entry: ZipEntry): string {
  // The content-type stream is located case-insensitively on the way in and
  // written under the canonical spelling on the way out, so an archive that
  // said `[content_types].xml` still matches the entry we emitted.
  return isContentTypesStreamName(entry.name)
    ? CONTENT_TYPES_PART
    : normalizePartName('/' + entry.name);
}

/**
 * Compare the emitted archive against the one it came from.
 *
 * Throws `ERR_PRESERVATION_BROKEN` on the first difference, naming the part.
 * There is no "report and continue" mode: one part quietly re-serialised is the
 * whole category of failure this package is built to make impossible, and every
 * later one is likely the same cause.
 */
export function assertPreserved(
  output: Uint8Array,
  original: Uint8Array | undefined,
  rewritten: ReadonlySet<string>,
  zip: ReadZipOptions = {},
): PreservationCheck {
  if (original === undefined) {
    return {
      checked: 0,
      rewritten: rewritten.size,
      skipped:
        'the bytes this package was opened from were not supplied, so there is nothing to compare ' +
        'the output against.',
    };
  }

  let before, after;
  try {
    before = readZip(original, zip);
    after = readZip(output, zip);
  } catch (error) {
    return {
      checked: 0,
      rewritten: rewritten.size,
      skipped:
        'one of the two archives could not be read: ' +
        (error instanceof Error ? error.message : String(error)),
    };
  }

  const source = new Map<string, ZipEntry>();
  for (const entry of before.entries) {
    if (entry.isDirectory) continue;
    source.set(entryKey(entry), entry);
  }

  let checked = 0;
  for (const entry of after.entries) {
    const key = entryKey(entry);
    // A part this session wrote is meant to differ, and a part that is not in
    // the source archive is one we added. Neither is evidence of anything.
    if (rewritten.has(key)) continue;
    const was = source.get(key);
    if (was === undefined) continue;

    const now = after.raw(entry);
    const then = before.raw(was);
    if (
      entry.method === was.method &&
      entry.crc32 === was.crc32 &&
      entry.uncompressedSize === was.uncompressedSize &&
      sameBytes(now, then)
    ) {
      checked++;
      continue;
    }

    throw new WriterError(
      'ERR_PRESERVATION_BROKEN',
      entry.name +
        ' was not edited and came out different anyway (' +
        describe(was) +
        ' in, ' +
        describe(entry) +
        ' out). An untouched part is streamed through still compressed and is never ' +
        're-serialised: the moment that stops being true, every feature we cannot parse is a ' +
        'feature we can lose.',
      { part: '/' + entry.name },
    );
  }

  return { checked, rewritten: rewritten.size, skipped: null };
}

function describe(entry: ZipEntry): string {
  return (
    String(entry.compressedSize) +
    'B method ' +
    String(entry.method) +
    ' crc ' +
    entry.crc32.toString(16)
  );
}

function sameBytes(a: Uint8Array, b: Uint8Array): boolean {
  if (a.byteLength !== b.byteLength) return false;
  for (let i = 0; i < a.byteLength; i++) if (a[i] !== b[i]) return false;
  return true;
}
