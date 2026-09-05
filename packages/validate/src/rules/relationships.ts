import {
  CONTENT_TYPE,
  isRelationshipPartName,
  isValidRelationshipId,
  normalizePartName,
  REL_TYPE,
} from '@pptx-studio/opc';
import {
  attributeNamespaceOf,
  descendantElements,
  NS,
  type XAttribute,
  type XElement,
} from '@pptx-studio/xml';
import type { Context } from '../context.js';
import { attributeLocation, elementLocation, partLocation } from '../report/location.js';
import { bySource, readRelsParts, resolveTarget, type RelsPart } from '../rels.js';

/**
 * `V006` … `V009`: the relationship graph.
 *
 * The graph is where a `.pptx` keeps almost everything that is not text. A
 * slide does not name its layout, a picture does not name its image, and a
 * chart does not name its workbook: each of them names an `rId`, and the `rId`
 * means something only inside one `.rels` part. Three consequences run through
 * all four rules below.
 *
 * **Ids are scoped, not global.** `rId3` in `slide1.xml.rels` and `rId3` in
 * `slide2.xml.rels` are unrelated. A global id registry is not an optimisation,
 * it is a bug, and it is a bug that works fine on every one-slide test deck.
 *
 * **Dangling is fatal; orphaned is harmless.** A reference to a relationship
 * that is not there, or a relationship to a part that is not there, breaks the
 * file. A relationship nothing references does not. That asymmetry is what
 * decides the direction media garbage collection is allowed to be aggressive
 * in, and getting it backwards is how a "clean-up" pass deletes the layouts a
 * user was about to switch to.
 *
 * **Targets resolve against the source part's folder.** Not the package root,
 * and not the `.rels` file's own folder.
 */

/**
 * Attributes in the relationship namespace that name a relationship id.
 *
 * All of them do, in practice - `r:id`, `r:embed`, `r:link`, `r:pict`,
 * `r:dm`, `r:lo`, `r:qs`, `r:cs`, `r:href` and the rest. Rather than list the
 * ones we know, this rule matches on the *namespace*, so an attribute from a
 * schema we have not read yet is checked rather than skipped. The two things
 * that must not be treated as ids are handled explicitly below.
 */
function isRelationshipReference(owner: XElement, attr: XAttribute): boolean {
  if (attr.prefix === '' || attr.prefix === 'xmlns') return false;
  // `r:` is conventional and not guaranteed, so the prefix is resolved rather
  // than compared. A part is free to bind the relationship namespace to any
  // prefix, and matching on the string `r` would silently stop checking one
  // that did.
  return attributeNamespaceOf(owner, attr) === NS.r;
}

/** Every `r:*` reference resolves in its own part's `.rels`. */
export function v006ReferencesResolve(ctx: Context): void {
  const relsBySource = bySource(readRelsParts(ctx));

  for (const part of ctx.parts()) {
    if (isRelationshipPartName(part)) continue;
    const document = ctx.document(part);
    if (document === null) continue;

    const rels = relsBySource.get(normalizePartName(part));
    const known = new Set((rels?.relationships ?? []).map((rel) => rel.id));

    for (const element of descendantElements(document.root)) {
      for (const attr of element.attributes) {
        if (!isRelationshipReference(element, attr)) continue;
        // An empty value is not a dangling reference, it is the absence of one.
        // `<a:hlinkClick r:id="" action="ppaction://noaction"/>` is what
        // PowerPoint writes for an action with no target, and it is common.
        if (attr.value === '') continue;
        if (known.has(attr.value)) continue;
        ctx.add(
          'V006',
          attributeLocation(part, element, attr.qname),
          rels === undefined
            ? '<' +
                element.qname +
                ' ' +
                attr.qname +
                '="' +
                attr.value +
                '"> refers to a relationship, but this part has no .rels at all.'
            : '<' +
                element.qname +
                ' ' +
                attr.qname +
                '="' +
                attr.value +
                '"> names a relationship that ' +
                rels.partName +
                ' does not declare. Ids are scoped to one .rels part; this one has ' +
                (known.size === 0 ? 'none' : [...known].sort().join(', ')) +
                '.',
        );
      }
    }
  }
}

/** Relationship ids are unique within a `.rels` and are `xsd:ID`s. */
export function v007RelationshipIds(ctx: Context): void {
  for (const rels of readRelsParts(ctx)) {
    const seen = new Map<string, number>();
    for (const rel of rels.relationships) {
      if (rel.id === '') {
        ctx.add('V007', elementLocation(rels.partName, rel.element), 'a <Relationship> has no Id.');
        continue;
      }
      if (!isValidRelationshipId(rel.id)) {
        ctx.add(
          'V007',
          elementLocation(rels.partName, rel.element),
          'Id="' +
            rel.id +
            '" is not an xsd:ID. An xsd:ID is an XML NCName: it cannot begin with a digit and ' +
            'cannot contain a space or a colon.',
        );
      }
      const count = (seen.get(rel.id) ?? 0) + 1;
      seen.set(rel.id, count);
      if (count === 2) {
        ctx.add(
          'V007',
          elementLocation(rels.partName, rel.element),
          'Id="' +
            rel.id +
            '" appears twice. Which target a reference to it means would be undecidable.',
        );
      }
    }
  }
}

/** Internal targets resolve, stay inside the package, and name a part that exists. */
export function v008TargetsResolve(ctx: Context): void {
  for (const rels of readRelsParts(ctx)) {
    for (const rel of rels.relationships) {
      if (rel.targetMode === 'External') continue;
      if (rel.targetMode !== 'Internal') {
        ctx.add(
          'V008',
          elementLocation(rels.partName, rel.element),
          'TargetMode="' + rel.targetMode + '" is neither Internal nor External.',
        );
        continue;
      }
      if (rel.target === '') {
        ctx.add(
          'V008',
          elementLocation(rels.partName, rel.element),
          rel.id + ' has an empty Target.',
        );
        continue;
      }
      const resolved = resolveTarget(rels.source, rel.target);
      if (resolved.part === null) {
        ctx.add(
          'V008',
          elementLocation(rels.partName, rel.element),
          rel.id +
            ' -> "' +
            rel.target +
            '" does not resolve against ' +
            rels.source +
            ': ' +
            (resolved.failure ?? 'unknown reason'),
        );
        continue;
      }
      if (!ctx.store.has(resolved.part)) {
        ctx.add(
          'V008',
          elementLocation(rels.partName, rel.element),
          rel.id +
            ' -> ' +
            resolved.part +
            ' names a part that is not in the package. A dangling relationship is fatal where an ' +
            'orphaned one is harmless, which is why garbage collection may only ever be ' +
            'aggressive in the other direction.',
        );
      }
    }
  }
}

/** The chartex content type. Not in `CONTENT_TYPE`: it is Microsoft's, not ECMA's. */
const CHART_EX_CONTENT_TYPE = 'application/vnd.ms-office.chartex+xml';

/**
 * The ECMA `customXml` relationship type.
 *
 * A `p:contentPart` reached by *this* type is a whole-package refusal, while
 * the Microsoft 2010 type of the same name opens - and so do two relationship
 * types that make no sense at all. PowerPoint validates that one type rather
 * than types in general, which is a strange enough finding that it is worth
 * writing the constant down beside the rule that uses it.
 */
const REL_CUSTOM_XML_ECMA =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml';

/** Elements whose `r:` attribute means "this is a media object's target". */
const MEDIA_REFERENCE_ELEMENTS = new Set(['audioFile', 'videoFile', 'media']);

/**
 * The edges a part must have, and the ones it must not.
 *
 * Four findings, none of them in any schema. See the rule's `why` in
 * `rules.ts` for the bisections; the code below is deliberately a table read
 * against the package rather than four separate walks.
 */
export function v009RequiredEdges(ctx: Context): void {
  const relsParts = readRelsParts(ctx);
  const relsBySource = bySource(relsParts);

  for (const part of ctx.parts()) {
    if (isRelationshipPartName(part)) continue;
    const contentType = ctx.contentType(part);
    const rels = relsBySource.get(normalizePartName(part));

    if (contentType === CONTENT_TYPE.slide) {
      requireExactlyOne(ctx, part, rels, REL_TYPE.slideLayout, CONTENT_TYPE.slideLayout, 'layout');
    } else if (contentType === CONTENT_TYPE.slideLayout) {
      requireExactlyOne(ctx, part, rels, REL_TYPE.slideMaster, CONTENT_TYPE.slideMaster, 'master');
    } else if (contentType === CHART_EX_CONTENT_TYPE) {
      // A `cx:chartSpace` with no `.rels` of its own is a whole-package
      // refusal, including one PowerPoint wrote itself. The classic
      // `c:chartSpace` needs neither of these.
      for (const [type, label] of [
        ['http://schemas.microsoft.com/office/2011/relationships/chartStyle', 'chartStyle'],
        [
          'http://schemas.microsoft.com/office/2011/relationships/chartColorStyle',
          'chartColorStyle',
        ],
      ] as const) {
        if ((rels?.relationships ?? []).some((rel) => rel.type === type)) continue;
        ctx.add(
          'V009',
          partLocation(part),
          'a chartex part with no ' +
            label +
            ' relationship. Unlike a classic c:chartSpace, which needs neither, a cx:chartSpace ' +
            'without both is a whole-package refusal.',
        );
      }
    }
  }

  // `p:contentPart` reached by the ECMA customXml type.
  for (const part of ctx.parts()) {
    if (isRelationshipPartName(part)) continue;
    const document = ctx.document(part);
    if (document === null) continue;
    const rels = relsBySource.get(normalizePartName(part));
    if (rels === undefined) continue;
    for (const element of descendantElements(document.root)) {
      if (element.local !== 'contentPart') continue;
      for (const attr of element.attributes) {
        if (attr.local !== 'id' || attr.prefix === '') continue;
        const relationship = rels.relationships.find((rel) => rel.id === attr.value);
        if (relationship?.type !== REL_CUSTOM_XML_ECMA) continue;
        ctx.add(
          'V009',
          attributeLocation(part, element, attr.qname),
          'a <' +
            element.qname +
            '> reached by the ECMA customXml relationship type is a whole-package refusal. The ' +
            'Microsoft 2010 type of the same name opens; PowerPoint validates this one type ' +
            'rather than types in general.',
        );
      }
    }
  }

  checkSharedMedia(ctx, relsBySource);
}

function requireExactlyOne(
  ctx: Context,
  part: string,
  rels: RelsPart | undefined,
  type: string,
  targetContentType: string,
  label: string,
): void {
  const matching = (rels?.relationships ?? []).filter((rel) => rel.type === type);
  if (matching.length !== 1) {
    ctx.add(
      'V009',
      partLocation(part),
      'has ' +
        String(matching.length) +
        ' ' +
        label +
        ' relationship(s); it needs exactly one. The binding lives only in the .rels part - ' +
        'nothing in the part itself names its ' +
        label +
        ' - so zero and two are both whole-package refusals.',
    );
    return;
  }
  const rel = matching[0]!;
  const resolved = resolveTarget(rels!.source, rel.target);
  if (resolved.part === null || !ctx.store.has(resolved.part)) {
    // Reported by `V008` as a dangling target; the extra thing this rule knows
    // is that it is *this* edge, which is the one that cannot be missing.
    ctx.add(
      'V009',
      elementLocation(rels!.partName, rel.element),
      'the ' + label + ' relationship points at a part that is not in the package.',
    );
    return;
  }
  const actual = ctx.contentType(resolved.part);
  if (actual !== targetContentType) {
    ctx.add(
      'V009',
      elementLocation(rels!.partName, rel.element),
      'the ' +
        label +
        ' relationship resolves to ' +
        resolved.part +
        ', which is typed ' +
        (actual === undefined ? '(nothing)' : '"' + actual + '"') +
        ' rather than "' +
        targetContentType +
        '". One pointing at a master instead was built as a single change and refused.',
    );
  }
}

/**
 * A media part that is both a transition sound and a media object's target.
 *
 * The only refusal we have measured whose message is "PowerPoint could not open
 * the file" rather than "the file or directory is corrupted and unreadable" -
 * which is worth recording, because the message is the only signal there is.
 * Two copies of the same bytes open; one shared copy does not.
 */
function checkSharedMedia(ctx: Context, relsBySource: ReadonlyMap<string, RelsPart>): void {
  const asTransitionSound = new Map<string, string>();
  const asMediaObject = new Map<string, string>();

  for (const part of ctx.parts()) {
    if (isRelationshipPartName(part)) continue;
    const document = ctx.document(part);
    if (document === null) continue;
    const rels = relsBySource.get(normalizePartName(part));
    if (rels === undefined) continue;

    for (const element of descendantElements(document.root)) {
      const isSound = element.local === 'snd';
      const isMedia = MEDIA_REFERENCE_ELEMENTS.has(element.local);
      if (!isSound && !isMedia) continue;
      for (const attr of element.attributes) {
        if (attr.prefix === '' || attr.value === '') continue;
        if (attr.local !== 'embed' && attr.local !== 'link') continue;
        const rel = rels.relationships.find((r) => r.id === attr.value);
        if (rel === undefined || rel.targetMode === 'External') continue;
        const target = resolveTarget(rels.source, rel.target).part;
        if (target === null) continue;
        (isSound ? asTransitionSound : asMediaObject).set(normalizePartName(target), part);
      }
    }
  }

  for (const [target, soundOwner] of asTransitionSound) {
    const mediaOwner = asMediaObject.get(target);
    if (mediaOwner === undefined) continue;
    ctx.add(
      'V009',
      partLocation(target),
      'is both a transition sound (from ' +
        soundOwner +
        ') and a media object target (from ' +
        mediaOwner +
        '). Two copies of the same bytes open; one shared copy is a refusal.',
    );
  }
}
