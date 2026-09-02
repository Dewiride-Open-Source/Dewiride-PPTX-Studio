import {
  CONTENT_TYPE,
  isRelationshipPartName,
  normalizePartName,
  REL_TYPE,
} from '@pptx-studio/opc';
import {
  attributeValue,
  descendantElements,
  namespaceOf,
  NS,
  type XElement,
} from '@pptx-studio/xml';
import type { Context } from './context.js';
import { attributeLocation, elementLocation } from './location.js';
import { bySource, readRelsParts, resolveTarget } from './rels.js';

/**
 * `V018` … `V021`: four identifier spaces, and never one allocator.
 *
 * A presentation has four kinds of id and they do not share a rule between
 * them. Slide ids start at 256 and stop at 2147483647. Master and layout ids
 * start at 2147483648, and - the part no schema says - come out of **one**
 * counter shared between them. Shape ids are unique inside a part and free to
 * repeat across parts. Placeholder indices are not identifiers at all; they are
 * a join key against another part.
 *
 * Every one of those is a way to break a file that looks like a way to be
 * tidy. Renumbering shape ids to be unique across the deck is the obvious
 * example: it is more consistent, it is what a database would do, and it
 * detaches every `p:custDataLst`, VML `@spid` and animation target that names
 * the old number.
 */

const SLIDE_ID_MIN = 256;
const SLIDE_ID_MAX = 2147483647;
const SHEET_ID_MIN = 2147483648;
const SHEET_ID_MAX = 4294967295;

/** An `@id` read as a base-10 integer, or `null` when it is not one. */
function integerAttribute(element: XElement, name: string): number | null {
  const raw = attributeValue(element, name);
  if (raw === undefined || !/^\d+$/.test(raw.trim())) return null;
  return Number(raw.trim());
}

/** `p:sldId/@id` is 256…2147483647 and unique. */
export function v018SlideIds(ctx: Context): void {
  for (const part of ctx.parts()) {
    const document = ctx.document(part);
    if (document === null) continue;
    const seen = new Map<number, XElement>();
    for (const element of descendantElements(document.root)) {
      if (element.local !== 'sldId' || namespaceOf(element) !== NS.p) continue;
      const raw = attributeValue(element, 'id');
      const value = integerAttribute(element, 'id');
      if (value === null) {
        ctx.add(
          'V018',
          attributeLocation(part, element, 'id'),
          raw === undefined
            ? '<p:sldId> has no @id.'
            : '@id="' + raw + '" is not a base-10 unsigned integer.',
        );
        continue;
      }
      if (value < SLIDE_ID_MIN || value > SLIDE_ID_MAX) {
        ctx.add(
          'V018',
          attributeLocation(part, element, 'id'),
          '@id=' +
            String(value) +
            ' is outside ST_SlideId, which is ' +
            String(SLIDE_ID_MIN) +
            '…' +
            String(SLIDE_ID_MAX) +
            '. The floor is the part that gets missed: nothing about a first slide suggests its ' +
            'id should start at 256.',
        );
      }
      if (seen.has(value)) {
        ctx.add(
          'V018',
          attributeLocation(part, element, 'id'),
          '@id=' + String(value) + ' is used twice in this slide list.',
        );
      } else {
        seen.set(value, element);
      }
    }
  }
}

/**
 * Master and layout ids share one number space.
 *
 * The uniqueness check runs across the whole package rather than per part,
 * because `p:sldMasterId` lives in `ppt/presentation.xml` and `p:sldLayoutId`
 * lives in each master's own `p:sldLayoutIdLst` - so a collision between the
 * two is invisible to anything that validates one part at a time. That is
 * exactly how `a12-masters` was written wrong: a chassis numbering masters and
 * layouts from two counters is indistinguishable from correct with one master
 * and collides on the second.
 */
export function v019SheetIds(ctx: Context): void {
  const seen = new Map<number, string>();

  for (const part of ctx.parts()) {
    const document = ctx.document(part);
    if (document === null) continue;
    for (const element of descendantElements(document.root)) {
      const isMaster = element.local === 'sldMasterId';
      const isLayout = element.local === 'sldLayoutId';
      if ((!isMaster && !isLayout) || namespaceOf(element) !== NS.p) continue;

      const raw = attributeValue(element, 'id');
      const value = integerAttribute(element, 'id');
      if (value === null) {
        // `@id` is optional on both in the schema. PowerPoint always writes
        // one, and a list entry without one cannot collide with anything, so
        // this is only reported when the attribute is present and unreadable.
        if (raw !== undefined) {
          ctx.add(
            'V019',
            attributeLocation(part, element, 'id'),
            '@id="' + raw + '" is not a base-10 unsigned integer.',
          );
        }
        continue;
      }
      if (value < SHEET_ID_MIN || value > SHEET_ID_MAX) {
        ctx.add(
          'V019',
          attributeLocation(part, element, 'id'),
          '@id=' +
            String(value) +
            ' is outside ' +
            (isMaster ? 'ST_SlideMasterId' : 'ST_SlideLayoutId') +
            ', which is ' +
            String(SHEET_ID_MIN) +
            ' and up.',
        );
        continue;
      }
      const previous = seen.get(value);
      if (previous !== undefined) {
        ctx.add(
          'V019',
          attributeLocation(part, element, 'id'),
          '@id=' +
            String(value) +
            ' is already used by ' +
            previous +
            '. Master and layout ids are one number space, not two: PowerPoint allocates them ' +
            'from a single running counter - master, its layouts, next master, its layouts - and ' +
            'refuses a package that does not.',
        );
      } else {
        seen.set(value, '<' + element.qname + '> in ' + part);
      }
    }
  }
}

/**
 * The `mc:Choice` / `mc:Fallback` ancestors of an element, outermost first.
 *
 * Two shapes in different branches of the same `mc:AlternateContent` never
 * exist in the same document: a consumer picks one branch and the other is not
 * there. So they may carry the same `@id`, and PowerPoint's own writer does
 * exactly that - `a22-chartex` has a `p:graphicFrame` with `id="10"` in the
 * `mc:Choice` and the `p:pic` that stands in for it, also `id="10"`, in the
 * `mc:Fallback`. Reporting that as a duplicate would fire on files PowerPoint
 * wrote, which is the corpus's whole job to prevent.
 */
function branchPath(element: XElement): XElement[] {
  const path: XElement[] = [];
  for (let node = element.parent; node !== undefined; node = node.parent) {
    if (node.local === 'Choice' || node.local === 'Fallback') path.push(node);
  }
  return path.reverse();
}

/**
 * True when two elements can never both be present.
 *
 * They cannot when their branch paths diverge at some depth *into two branches
 * of the same `mc:AlternateContent`*. Diverging into branches of two different
 * `mc:AlternateContent` elements is not exclusion - both are chosen, and both
 * sets of shapes are on the slide together.
 */
function mutuallyExclusive(a: readonly XElement[], b: readonly XElement[]): boolean {
  const depth = Math.min(a.length, b.length);
  for (let i = 0; i < depth; i++) {
    if (a[i] === b[i]) continue;
    return a[i]!.parent === b[i]!.parent;
  }
  return false;
}

/**
 * `p:cNvPr/@id` is unique within its part and reads as a positive signed
 * 32-bit integer.
 *
 * PresentationML only, which is narrower than "every element called `cNvPr`"
 * and had to be. A SmartArt drawing part writes `dsp:cNvPr id="0"` on every
 * shape it contains - PowerPoint generates those parts and regenerates them on
 * any diagram interaction, so whatever the ids there are for, it is not
 * identity. `a23-smartart` is where that turned up.
 */
export function v020ShapeIds(ctx: Context): void {
  for (const part of ctx.parts()) {
    if (isRelationshipPartName(part)) continue;
    const document = ctx.document(part);
    if (document === null) continue;

    const seen = new Map<number, { label: string; branch: XElement[] }[]>();
    for (const element of descendantElements(document.root)) {
      if (element.local !== 'cNvPr' || namespaceOf(element) !== NS.p) continue;
      const raw = attributeValue(element, 'id');
      if (raw === undefined) {
        ctx.add('V020', elementLocation(part, element), '<' + element.qname + '> has no @id.');
        continue;
      }
      const value = integerAttribute(element, 'id');
      if (value === null) {
        ctx.add(
          'V020',
          attributeLocation(part, element, 'id'),
          '@id="' + raw + '" is not a base-10 unsigned integer.',
        );
        continue;
      }
      // Measured, not read. `ST_DrawingElementId` is `xsd:unsignedInt`, so the
      // schema allows the whole 32-bit range - but 2147483648…4294967294 are
      // whole-package refusals and 4294967295 opens, because PowerPoint reads
      // the attribute as a *signed* 32-bit integer and keeps 0xFFFFFFFF as a
      // sentinel it renumbers away on save. `a39-large-ids` is the deck.
      if (value > SLIDE_ID_MAX && value !== SHEET_ID_MAX) {
        ctx.add(
          'V020',
          attributeLocation(part, element, 'id'),
          '@id=' +
            String(value) +
            ' is in 2147483648…4294967294, which PowerPoint refuses outright. ST_DrawingElementId ' +
            'is xsd:unsignedInt and every value up to 2147483647 opens; PowerPoint reads this one ' +
            'signed. (4294967295 also opens - it is minus one, which PowerPoint keeps as a ' +
            'sentinel and renumbers away on save.)',
        );
      }
      const branch = branchPath(element);
      const others = seen.get(value) ?? [];
      const clash = others.find((other) => !mutuallyExclusive(other.branch, branch));
      if (clash !== undefined) {
        ctx.add(
          'V020',
          attributeLocation(part, element, 'id'),
          '@id=' +
            String(value) +
            ' is already used by <' +
            clash.label +
            '> in this part. Shape ids are unique within a part and free to repeat across parts - ' +
            'renumbering them to be unique across the deck is tidier and detaches every VML ' +
            '@spid and animation target that names the old number.',
        );
      }
      others.push({
        label: element.qname + ' ' + (attributeValue(element, 'name') ?? ''),
        branch,
      });
      seen.set(value, others);
    }
  }
}

/** A placeholder as the matcher sees it. `null` type means the attribute was absent. */
interface Placeholder {
  readonly type: string;
  readonly idx: number;
  readonly element: XElement;
}

/**
 * `CT_Placeholder` defaults, materialised.
 *
 * `@type` defaults to `body` and `@idx` to `0` in the ECMA-376 schema, and
 * `ctrTitle` normalises to `title` because they are the same slot with
 * different centring. Worth flagging while it is fresh: the plan says
 * `<p:ph/>` means `type="obj"` in sub-phase 2.9 and `type ??= 'body'` in 7.1,
 * and those cannot both be right. The schema says `body`, so that is what this
 * rule uses - and it is one of the reasons the rule is a warning rather than
 * fatal. Sub-phase 7.1 settles it against the 121-case matrix.
 */
function placeholdersOf(root: XElement): Placeholder[] {
  const out: Placeholder[] = [];
  for (const element of descendantElements(root)) {
    if (element.local !== 'ph' || namespaceOf(element) !== NS.p) continue;
    const rawType = attributeValue(element, 'type') ?? 'body';
    const type = rawType === 'ctrTitle' ? 'title' : rawType;
    const idx = Number(attributeValue(element, 'idx') ?? '0');
    out.push({ type, idx: Number.isFinite(idx) ? idx : 0, element });
  }
  return out;
}

/** Types matched on type alone, ignoring `@idx`. See tier 3. */
const TYPE_ONLY = new Set(['sldNum', 'dt', 'ftr', 'hdr']);

/** The five tiers, in order. Returns the matching layout placeholder, or `null`. */
function matchInLayout(slide: Placeholder, layout: readonly Placeholder[]): Placeholder | null {
  const exact = layout.find((ph) => ph.type === slide.type && ph.idx === slide.idx);
  if (exact !== undefined) return exact;
  if (slide.type === 'title') {
    const anyTitle = layout.find((ph) => ph.type === 'title');
    if (anyTitle !== undefined) return anyTitle;
  }
  if (TYPE_ONLY.has(slide.type)) {
    const byType = layout.find((ph) => ph.type === slide.type);
    if (byType !== undefined) return byType;
  }
  if (slide.type === 'body') {
    const bodies = layout.filter((ph) => ph.type === 'body');
    if (bodies.length === 1) return bodies[0]!;
  }
  return null;
}

/**
 * Every slide placeholder finds a counterpart in its layout.
 *
 * A warning, and the reason matters: PowerPoint opens the file. A `p:ph` that
 * matches nothing inherits nothing - no geometry, no text style, no prompt -
 * and the shape is drawn wherever the renderer can work out, which is a deck
 * that is silently wrong rather than a deck that is refused. Refusing an export
 * over it would block a user from saving a file that already opened fine
 * everywhere, which is the failure mode a validator has to avoid most.
 */
export function v021PlaceholderIndices(ctx: Context): void {
  const relsBySource = bySource(readRelsParts(ctx));

  for (const part of ctx.parts()) {
    if (ctx.contentType(part) !== CONTENT_TYPE.slide) continue;
    const document = ctx.document(part);
    if (document === null) continue;

    const rels = relsBySource.get(normalizePartName(part));
    const layoutRel = (rels?.relationships ?? []).find((rel) => rel.type === REL_TYPE.slideLayout);
    if (layoutRel === undefined) continue; // `V009` has already said so.
    const layoutPart = resolveTarget(rels!.source, layoutRel.target).part;
    if (layoutPart === null) continue;
    const layoutDocument = ctx.document(layoutPart);
    if (layoutDocument === null) continue;

    const layout = placeholdersOf(layoutDocument.root);
    for (const slidePh of placeholdersOf(document.root)) {
      if (matchInLayout(slidePh, layout) !== null) continue;
      ctx.add(
        'V021',
        elementLocation(part, slidePh.element),
        'placeholder (type=' +
          slidePh.type +
          ', idx=' +
          String(slidePh.idx) +
          ') matches nothing in ' +
          layoutPart +
          ', which offers ' +
          (layout.length === 0
            ? 'none'
            : layout.map((ph) => ph.type + '/' + String(ph.idx)).join(', ')) +
          '. An unmatched placeholder inherits no geometry and no text style.',
      );
    }
  }
}
