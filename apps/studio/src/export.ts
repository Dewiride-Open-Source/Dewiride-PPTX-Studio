import { PartStore } from '@pptx-studio/opc';
import {
  applyEdit,
  childElements,
  namespaceOf,
  newText,
  parseXml,
  serializeXml,
  type XElement,
  type XText,
} from '@pptx-studio/xml';
import { comparePackages, exportPackage, openPackage } from '@pptx-studio/writer';
import type { Report } from '@pptx-studio/validate';

/**
 * Gate 1, in a browser tab: read a package, hand it back, and prove nothing
 * moved.
 *
 * Every part of this has been exercised in a browser already - the writer's own
 * suite runs in real Chromium, not jsdom, and has since 0.1. What has not
 * existed until now is the *path*: a `File` the user chose, through a Worker,
 * out through a `Blob` and into their downloads folder, with the four checks
 * that make the bytes worth handing over run in between. That path is what the
 * gate asks about, and it is the part nothing in `packages/` can test, because
 * nothing in `packages/` is allowed to know a `File` exists.
 *
 * ## What runs, in order
 *
 * 1. **`openPackage`** - two independent `PartStore`s over the same bytes. The
 *    second is the baseline, and without it three of the validator's rules, the
 *    media sweep and the preservation check all quietly skip themselves rather
 *    than pass.
 * 2. **The edit**, if one was asked for. See below.
 * 3. **`exportPackage`** - prepare hooks, media collection, the write, then the
 *    preservation check and the twenty-nine rules. The bytes are returned only
 *    if both pass; a failure throws and nothing reaches the page.
 * 4. **`comparePackages`** - the 1.4 oracle, run over the archive that was just
 *    emitted rather than over the store's opinion of it. Redundant with the
 *    preservation check on a no-op and not redundant at all after an edit: the
 *    preservation check skips the parts this session wrote, and this is what
 *    then says what changed about them.
 *
 * ## Why there is an edit at all
 *
 * Gate 1's wording is "re-saves it ... with zero visible difference", which a
 * no-op satisfies by construction - and satisfying it by construction is the
 * point of the architecture. But a demo that can only ever emit the same bytes
 * it read never exercises the half of the writer that matters: the part where
 * one part is serialised afresh and every other part is streamed out of the
 * source archive still compressed. `rewritten: 1, streamed: 55` is the claim,
 * and a page that never rewrites anything cannot make it.
 *
 * So there is a second mode, and it is deliberately the smallest edit in the
 * package: the text of `cp:lastModifiedBy` in `docProps/core.xml`. Nothing on
 * any slide changes, which keeps "zero visible difference" true; the part is
 * one PowerPoint reads and shows in File - Info, so the change is checkable by
 * a human; and it is reached through `applyEdit`, which is the only sanctioned
 * way to change a tree in this project and hands back its own inverse.
 *
 * It is emphatically **not** a preview of Phase 5. There is no command, no
 * history and no model here - `packages/model` does not exist yet. This is one
 * XML edit, written where the demo needs it and nowhere else.
 */

/** `cp:`, matched by URI. The prefix in the file is not ours to rely on. */
const NS_CORE_PROPERTIES =
  'http://schemas.openxmlformats.org/package/2006/metadata/core-properties';

const CORE_PROPERTIES_PART = '/docProps/core.xml';

/** What `cp:lastModifiedBy` is set to. Also what a reviewer greps for. */
export const STAMP = 'PPTX Studio';

export type EditKind = 'none' | 'stamp';

/** One entry of `comparePackages`' answer, flattened for `postMessage`. */
export interface ExportDifference {
  readonly kind: string;
  readonly part: string;
  readonly detail: string;
}

/**
 * Everything the page is told about an export.
 *
 * Plain JSON throughout, and that is a constraint rather than a description:
 * this crosses a `postMessage` boundary, so a `PartStore` or an `XNode` in here
 * would be a structured-clone failure at runtime and not a type error at build
 * time. `RoundTripResult.written` is exactly such a field, which is why the
 * comparison is unpacked here rather than forwarded.
 */
export interface ExportOutcome {
  /** Parts serialised afresh. Empty on a no-op export - that is the point. */
  readonly rewritten: readonly string[];
  /** Parts copied out of the source archive without being decompressed. */
  readonly streamed: number;
  readonly preservation: {
    readonly checked: number;
    readonly rewritten: number;
    readonly skipped: string | null;
  };
  /** `null` only if validation was turned off, which this path never does. */
  readonly report: Report | null;
  readonly comparison: {
    readonly parts: number;
    readonly same: number;
    readonly xml: number;
    readonly relationships: number;
    readonly binary: number;
    readonly differences: readonly ExportDifference[];
  };
  /** The part the edit touched, or `null` when none was asked for or found. */
  readonly edited: string | null;
  readonly bytesIn: number;
  readonly bytesOut: number;
  readonly ms: number;
}

export interface ExportProduct {
  readonly bytes: Uint8Array;
  readonly outcome: ExportOutcome;
}

/** The first child element in `uri` named `local`, or `undefined`. */
function findChild(parent: XElement, uri: string, local: string): XElement | undefined {
  return childElements(parent).find((child) => child.local === local && namespaceOf(child) === uri);
}

/**
 * Set the text of `cp:lastModifiedBy`, and return the part it changed.
 *
 * Returns `null` when the package has no core-properties part or no such
 * element, rather than creating one. Adding an element would mean deciding
 * where it goes in `CT_CoreProperties`, and that is a question with a real
 * answer this function has no reason to be the place that answers.
 */
export function stampLastModifiedBy(store: PartStore, value: string): string | null {
  if (!store.has(CORE_PROPERTIES_PART)) return null;

  const document = parseXml(store.read(CORE_PROPERTIES_PART));
  const target = findChild(document.root, NS_CORE_PROPERTIES, 'lastModifiedBy');
  if (target === undefined) return null;

  const text = target.children.find((child): child is XText => child.type === 'text');
  applyEdit(
    text === undefined
      ? { kind: 'insertChild', parent: target, index: 0, node: newText(value) }
      : { kind: 'setValue', node: text, value },
  );

  store.replacePart(CORE_PROPERTIES_PART, serializeXml(document));
  return CORE_PROPERTIES_PART;
}

/**
 * Read a package and write it back out, with every check the writer offers.
 *
 * Throws rather than returning a failure. An export that produced bytes and a
 * complaint would leave the page deciding whether to offer them, and there is
 * one right answer to that: the firewall exists so that bytes PowerPoint would
 * refuse never leave the function.
 */
export function exportDeck(bytes: Uint8Array, edit: EditKind = 'none'): ExportProduct {
  const startedAt = performance.now();
  const { store, baseline, baselineBytes } = openPackage(bytes);

  const edited = edit === 'stamp' ? stampLastModifiedBy(store, STAMP) : null;

  const result = exportPackage({ store, baseline, baselineBytes });
  const comparison = comparePackages(PartStore.open(baselineBytes), PartStore.open(result.bytes));

  return {
    bytes: result.bytes,
    outcome: {
      rewritten: [...result.rewritten],
      streamed: result.streamed,
      preservation: {
        checked: result.preservation.checked,
        rewritten: result.preservation.rewritten,
        skipped: result.preservation.skipped,
      },
      report: result.report,
      comparison: {
        parts: comparison.parts.length,
        same: comparison.counts.same,
        xml: comparison.counts.xml,
        relationships: comparison.counts.relationships,
        binary: comparison.counts.binary,
        differences: comparison.differences.map((difference) => ({
          kind: difference.kind,
          part: difference.part,
          detail: difference.detail,
        })),
      },
      edited,
      bytesIn: bytes.byteLength,
      bytesOut: result.bytes.byteLength,
      ms: performance.now() - startedAt,
    },
  };
}
