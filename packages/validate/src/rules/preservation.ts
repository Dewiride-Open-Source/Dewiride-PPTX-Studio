import { isRelationshipPartName } from '@pptx-studio/opc';
import {
  attributeValue,
  descendantElements,
  sourceOf,
  textContent,
  type XDocument,
  type XElement,
} from '@pptx-studio/xml';
import type { Context } from '../context.js';
import { elementLocation, partLocation, xpathOf } from '../report/location.js';

/**
 * `V027` … `V029`: what has to come back out unchanged.
 *
 * These are the rules the architecture exists for. Preserving charts, SmartArt,
 * animations, OLE objects and macros is unachievable in any design where the
 * writer has to *understand* a feature in order to emit it - and the moment
 * preservation is a feature rather than the default state, it is a feature with
 * a coverage percentage, and the percentage is never a hundred.
 *
 * So the default is that nothing changes, and these three rules are where that
 * intention becomes falsifiable.
 *
 * ## They need two packages, and they say so when they only have one
 *
 * All three compare against the package as it was opened. Asked to validate a
 * file with no history - `pptx-studio validate deck.pptx`, a deck somebody
 * dropped on a page - there is nothing to compare against and they do not run.
 * The report lists them under `skipped` with that reason, because a
 * preservation rule that reports nothing because it had nothing to compare
 * against looks exactly like one that found nothing, and those are opposite
 * answers.
 */

/** `ppt/embeddings/…` - OLE2/CFB compound files. Never ours to rewrite. */
function isEmbedding(part: string): boolean {
  return /^\/ppt\/embeddings\//i.test(part);
}

/**
 * A part nobody edited comes back out byte for byte.
 *
 * "Nobody edited" is `PartInfo.fromArchive`, the store's own record of whether a
 * part's bytes still come from the archive it was opened from - not a byte
 * comparison, which could not tell an edit that happened to produce identical
 * bytes from no edit at all. The rule then checks the converse: a part the store
 * says is untouched really is identical.
 *
 * That is not circular, because the two facts come from different places. The
 * flag is set by whoever called `replacePart`; the bytes come from whatever the
 * writer actually emitted. A writer that re-serialised a clean part - the exact
 * failure this rule exists for - would leave the flag alone and change the
 * bytes, and this is where the two stop agreeing.
 */
export function v027UneditedPartsUnchanged(ctx: Context): void {
  const baseline = ctx.baseline;
  if (baseline === null) return;

  for (const part of ctx.parts()) {
    const before = ctx.baselineRead(part);
    if (before === null) continue; // Added this session. Nothing to preserve.

    const embedding = isEmbedding(part);
    if (embedding && ctx.edited(part)) {
      ctx.add(
        'V027',
        partLocation(part),
        'an OLE embedding was replaced. ppt/embeddings/*.bin are OLE2/CFB compound files and any ' +
          'rewrite of one is a guaranteed repair prompt - we render their preview and never ' +
          'touch the object.',
      );
      continue;
    }
    if (!embedding && ctx.edited(part)) continue; // Edited on purpose.

    const after = ctx.read(part);
    if (after === null) continue;
    if (sameBytes(before, after)) continue;

    ctx.add(
      'V027',
      partLocation(part),
      'nobody edited this part and its bytes changed anyway (' +
        String(before.byteLength) +
        ' bytes in, ' +
        String(after.byteLength) +
        ' out). A part that was not edited is streamed through, never re-serialised: the ' +
        'moment that stops being true, every feature we cannot parse is a feature we can lose.',
    );
  }
}

function sameBytes(a: Uint8Array, b: Uint8Array): boolean {
  if (a.byteLength !== b.byteLength) return false;
  for (let i = 0; i < a.byteLength; i++) if (a[i] !== b[i]) return false;
  return true;
}

/** Local names of the containers held opaque, with the namespace-free test that finds them. */
function isOpaqueContainer(element: XElement): boolean {
  // `mc:AlternateContent` by local name, because the `mc` prefix is only a
  // convention; `a:ext` and `p:ext` likewise. The namespace is not checked
  // because the point is to be *broader* than the schemas we know - an
  // extension container from a vocabulary we have never read is exactly the
  // thing this rule protects.
  return element.local === 'AlternateContent' || element.local === 'ext';
}

/** Every opaque container in a document, as its exact source text. */
function opaqueSources(document: XDocument): { element: XElement; source: string }[] {
  const out: { element: XElement; source: string }[] = [];
  for (const element of descendantElements(document.root)) {
    if (!isOpaqueContainer(element)) continue;
    out.push({ element, source: sourceOf(document, element) });
  }
  return out;
}

/**
 * No `mc:AlternateContent` branch and no extension was rebuilt.
 *
 * The check is a multiset test, not a positional one: every opaque container in
 * the part we are about to write must appear, character for character,
 * somewhere in the part as it was opened. That formulation is deliberate.
 *
 * A positional comparison - same XPath, same text - would be exact and would
 * also fire every time a shape was inserted above one, because the index in the
 * path shifts. The multiset test cannot be fooled in the direction that matters
 * (text nobody wrote before is text we wrote) and is immune to the reordering
 * that is a normal, permitted edit.
 *
 * **Deletion is not this rule's business.** An extension can leave because the
 * shape holding it was deleted, and that is a document operation the user asked
 * for. What must never happen is markup *appearing* here that we composed:
 * `mc:Choice/@Requires` names a prefix rather than a URI, so a rewritten branch
 * is how ignorable extension markup becomes a hard error - and `a34-extlst`
 * measured that PowerPoint carries an unknown `a:ext/@uri` through untouched,
 * so an extension we rebuild is one the file would otherwise have kept for ever.
 */
export function v028OpaqueContainersUnchanged(ctx: Context): void {
  if (ctx.baseline === null) return;

  for (const part of ctx.parts()) {
    if (isRelationshipPartName(part) || !ctx.edited(part)) continue;
    const after = ctx.document(part);
    const before = ctx.baselineDocument(part);
    if (after === null || before === null) continue;

    const known = new Set(opaqueSources(before).map((entry) => entry.source));
    for (const entry of opaqueSources(after)) {
      if (known.has(entry.source)) continue;
      const uri = attributeValue(entry.element, 'uri');
      ctx.add(
        'V028',
        elementLocation(part, entry.element),
        '<' +
          entry.element.qname +
          '>' +
          (uri === undefined ? '' : ' uri="' + uri + '"') +
          ' is not markup this part arrived with. These containers are carried through ' +
          'byte for byte and never composed: we do not know what is in them, which is the ' +
          'whole reason they are held opaque.',
      );
    }
  }
}

/** An `a:fld` and the two things about it that must survive. */
interface Field {
  readonly id: string;
  readonly text: string;
  readonly element: XElement;
}

function fieldsOf(root: XElement): Field[] {
  const out: Field[] = [];
  for (const element of descendantElements(root)) {
    if (element.local !== 'fld') continue;
    out.push({
      id: attributeValue(element, 'id') ?? '',
      text: textContent(element),
      element,
    });
  }
  return out;
}

/** `a:t` elements whose `xml:space` was declared, keyed by path. */
function spacePreserved(root: XElement): Map<string, XElement> {
  const out = new Map<string, XElement>();
  for (const element of descendantElements(root)) {
    if (element.local !== 't') continue;
    if (attributeValue(element, 'xml:space') !== 'preserve') continue;
    out.set(xpathOf(element), element);
  }
  return out;
}

/** The path of every `a:t`, so "is this one still there" is a set lookup. */
function textPaths(root: XElement): Set<string> {
  const out = new Set<string>();
  for (const element of descendantElements(root)) {
    if (element.local === 't') out.add(xpathOf(element));
  }
  return out;
}

/**
 * Text and field identity survive an edit.
 *
 * Three small, separately silent losses.
 *
 * **`xml:space="preserve"` is preserved, not required.** The distinction is
 * measured: 120 `<a:t>` elements in the corpus carry leading or trailing
 * whitespace *without* the attribute and PowerPoint round-trips every one of
 * them with the whitespace intact, so a rule that demanded the attribute would
 * fire on real files that work. What is not allowed is dropping one that was
 * there, because then the attribute's absence starts meaning something it did
 * not mean before.
 *
 * **`a:fld/@id` is a required `ST_Guid`** identifying the field across saves.
 * Regenerating one can make PowerPoint repair the file, and a repair is the one
 * outcome this whole package exists to prevent.
 *
 * **A field's cached `a:t` is its only rendering** anywhere we cannot evaluate
 * the field ourselves. Discarding it turns a date placeholder into an empty box
 * on every consumer that is not PowerPoint - including, for anything we have
 * not implemented, us.
 */
export function v029TextAndFieldIdentity(ctx: Context): void {
  if (ctx.baseline === null) return;

  for (const part of ctx.parts()) {
    if (isRelationshipPartName(part) || !ctx.edited(part)) continue;
    const after = ctx.document(part);
    const before = ctx.baselineDocument(part);
    if (after === null || before === null) continue;

    // xml:space, by path. A path that no longer exists is a deletion and not
    // this rule's business; a path that still exists and has lost the attribute
    // is a rewrite that changed what the text means.
    //
    // Both sets are built once. The obvious shape - scan the new document for
    // each old path - is quadratic in the number of `<a:t>` elements, and a
    // slide with a thousand runs is an ordinary slide.
    const afterSpace = spacePreserved(after.root);
    const afterPaths = textPaths(after.root);
    for (const [path, element] of spacePreserved(before.root)) {
      if (afterSpace.has(path)) continue;
      if (!afterPaths.has(path)) continue;
      ctx.add(
        'V029',
        { part, xpath: path, offset: element.start },
        'this <' +
          element.qname +
          '> arrived with xml:space="preserve" and no longer has it. The attribute is not ' +
          'required - 120 elements in our corpus carry edge whitespace without it - but dropping ' +
          'one that was written changes what the text is.',
      );
    }

    const afterFields = fieldsOf(after.root);
    const beforeById = new Map(fieldsOf(before.root).map((field) => [field.id, field]));
    const beforeIds = new Set(beforeById.keys());

    for (const field of afterFields) {
      if (beforeIds.has(field.id)) {
        const original = beforeById.get(field.id)!;
        if (original.text !== '' && field.text.trim() === '') {
          ctx.add(
            'V029',
            elementLocation(part, field.element),
            'field ' +
              field.id +
              ' has lost its cached text ("' +
              original.text +
              '"). The cache is what renders anywhere the field cannot be evaluated, which ' +
              'includes every consumer that is not PowerPoint.',
          );
        }
        continue;
      }
      ctx.add(
        'V029',
        elementLocation(part, field.element),
        'field id="' +
          field.id +
          '" is not one this part arrived with. a:fld/@id is a required ST_Guid identifying the ' +
          'field across saves; regenerating one can make PowerPoint repair the file. The ids ' +
          'that were here: ' +
          (beforeIds.size === 0 ? 'none' : [...beforeIds].join(', ')) +
          '.',
      );
    }
  }
}
