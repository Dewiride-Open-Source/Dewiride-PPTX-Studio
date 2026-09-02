/**
 * The twenty-nine rules, and the evidence for each.
 *
 * ## Why a table and not twenty-nine functions
 *
 * The functions are in the files beside this one. What lives here is the part a
 * person reads: what each rule claims, how sure we are, and *how we know*. A
 * validator whose rules exist only as code is a validator nobody can audit, and
 * this one has to be audited, because roughly half of it enforces things no
 * schema says.
 *
 * That is the fact that shapes the whole package. ECMA-376 is a description of
 * a file format; PowerPoint is an implementation that refuses files the
 * description permits. Sub-phase 0.7 and the corpus bisections in
 * `tools/corpus/ROSTER.md` found nineteen such refusals by building a package
 * with one change in it and watching PowerPoint decline to open it - no
 * diagnostic, no log, no part named, just "PowerPoint could not open the file"
 * or `0x80070570`. Every rule below whose `evidence` says `measured` came from
 * that loop and from nowhere else.
 *
 * So each rule carries three things beyond its check:
 *
 * - `evidence` - `schema` (ECMA-376 says so), `measured` (we watched PowerPoint
 *   refuse it), or `both`.
 * - `why` - the sentence that justifies the rule to somebody who is about to
 *   delete it because it fired on their file.
 * - `severity` - `fatal` refuses an export; `warning` is reported and does not.
 *
 * ## On being wrong in the safe direction
 *
 * A false positive here blocks an export a user wanted. A false negative hands
 * them a file PowerPoint will not open, with no way to find out why. Those are
 * not symmetric, but the first is not free either - a validator that fires on
 * good files gets turned off, and then it catches nothing.
 *
 * The resolution is `origin`, which lives on the finding rather than here: a
 * fatal a caller *introduced* refuses the export, and the identical fatal that
 * was already in the file when it was opened is reported and does not. That is
 * the same split `PartStore.write` already makes for dangling relationships,
 * and for the same reason - refusing to re-export a file we did not break makes
 * the file unopenable in this editor and does not fix anything.
 */

export type RuleCategory =
  'package' | 'relationships' | 'order' | 'required' | 'ids' | 'refused' | 'preservation';

export type Severity = 'fatal' | 'warning';

/** Where the rule comes from. See the header. */
export type Evidence = 'schema' | 'measured' | 'both';

export interface Rule {
  readonly id: RuleId;
  readonly category: RuleCategory;
  readonly severity: Severity;
  readonly evidence: Evidence;
  /** One line, imperative, in the voice of the thing that must be true. */
  readonly title: string;
  /** Why the rule exists, for whoever is about to argue with it. */
  readonly why: string;
  /**
   * True when the rule needs the package as it was opened as well as the
   * package about to be written.
   *
   * Six rules do. They are skipped rather than passed when no baseline is
   * given, and the report says so - a preservation rule that silently reports
   * nothing because it had nothing to compare against is worse than one that
   * did not run, because it looks like a pass.
   */
  readonly needsBaseline?: true;
}

/**
 * The shape each entry is checked against.
 *
 * `RULES` is `as const satisfies readonly RuleShape[]` rather than typed
 * `readonly Rule[]`, so that `RuleId` below can be the union of the twenty-nine
 * literal ids written here instead of `string`. Annotating the array would
 * widen `id`, and then a caller could ask for a rule that does not exist and be
 * told so only at run time.
 */
interface RuleShape {
  readonly id: string;
  readonly category: RuleCategory;
  readonly severity: Severity;
  readonly evidence: Evidence;
  readonly title: string;
  readonly why: string;
  readonly needsBaseline?: true;
}

export const RULES = [
  // --- package ------------------------------------------------------------

  {
    id: 'V001',
    category: 'package',
    severity: 'fatal',
    evidence: 'both',
    title: 'every part resolves to a content type',
    why:
      'OPC has no default type. A part whose extension matches no `Default` and ' +
      'which carries no `Override` has no type at all, and a consumer has no way ' +
      'to know what it is holding. PowerPoint refuses such a package outright.',
  },
  {
    id: 'V002',
    category: 'package',
    severity: 'fatal',
    evidence: 'both',
    title: 'the content-type map is internally consistent',
    why:
      'Every `Override` names a part that exists; no extension is declared ' +
      '`Default` twice; no part is `Override`-ed twice. And the one Office ' +
      'omission that is famous for it: whenever any `.fntdata` part exists there ' +
      'must be a `<Default Extension="fntdata" ContentType="application/x-fontdata"/>`. ' +
      'Leaving it out is the canonical "PowerPoint found a problem with content" ' +
      'bug and it is what sub-phase 8.7 exists to get right.',
  },
  {
    id: 'V003',
    category: 'package',
    severity: 'fatal',
    evidence: 'schema',
    title: 'the archive carries no directory entries and no ZIP64 record',
    why:
      'A directory entry is not a part - no OPC part name may end in a slash - so ' +
      'writing one puts a thing in the package that nothing can name. ZIP64 is ' +
      'refused for a narrower reason: PowerPoint reads it, but our writer emits ' +
      'ZIP32 only, so a ZIP64 record in something we produced means a length ' +
      'field went somewhere it should not have. Both are about what *we* hand ' +
      'over; a package that arrives with either is reported and not refused.',
  },
  {
    id: 'V004',
    category: 'package',
    severity: 'fatal',
    evidence: 'both',
    title: 'part names are ASCII, well formed, and free of pointless percent-escapes',
    why:
      'The OPC grammar, plus the one thing it does not say. RFC 3986 §6.2.2.2 ' +
      'normalisation is enforced by PowerPoint rather than merely recommended: a ' +
      'percent-escape of an *unreserved* character is a whole-package refusal ' +
      'with `0x808D1005`. Measured nine for nine on one-name packages - `%2D`, ' +
      '`%41`, `%5F` and `%7E` refused; `%20`, `%23`, `%24`, `%2C` and `%3A` open. ' +
      '`@pptx-studio/opc` reports this as a warning (`M1.8`) because it must open ' +
      'what it is given; here it is fatal, because we are about to write it.',
  },
  {
    id: 'V005',
    category: 'package',
    severity: 'fatal',
    evidence: 'both',
    title: 'there is exactly one main presentation part, of a matching content type',
    why:
      '`_rels/.rels` must carry exactly one `officeDocument` relationship and it ' +
      'must resolve to a part whose content type is one of the four PresentationML ' +
      'main-part types. And the pairing PowerPoint enforces outside the schema: it ' +
      'checks the *file extension* against that content type and refuses the pair ' +
      'when they disagree, which is why a macro-enabled deck has to be written ' +
      '`.pptm`. `a32-macros` is the corpus deck that measured it.',
  },

  // --- relationships ------------------------------------------------------

  {
    id: 'V006',
    category: 'relationships',
    severity: 'fatal',
    evidence: 'both',
    title: 'every relationship reference resolves in its own part’s `.rels`',
    why:
      '`r:id`, `r:embed`, `r:link`, `r:pict`, `r:dm`, `r:lo`, `r:qs`, `r:cs` and ' +
      'every other attribute in the relationship namespace. Relationship ids are ' +
      '`xsd:ID` scoped to one `.rels` part, so the lookup is always against the ' +
      'referring part’s own collection and never against a global registry - a ' +
      'global one is not a shortcut, it is a bug. A dangling reference is fatal ' +
      'where an orphan relationship is harmless, which is the asymmetry that ' +
      'decides which direction media garbage collection may be aggressive in.',
  },
  {
    id: 'V007',
    category: 'relationships',
    severity: 'fatal',
    evidence: 'schema',
    title: 'relationship ids are unique within a `.rels` and match the `xsd:ID` grammar',
    why:
      '`Relationship/@Id` is `xsd:ID`: an XML Name, unique in its document. Two ' +
      'relationships sharing an id in one `.rels` makes every reference to it ' +
      'ambiguous, and which one a consumer picks is not defined anywhere.',
  },
  {
    id: 'V008',
    category: 'relationships',
    severity: 'fatal',
    evidence: 'both',
    title: 'internal targets resolve inside the package, relative to the source part’s folder',
    why:
      'A target is resolved against the folder of the part that *owns the `.rels`*, ' +
      'not against the package root and not against the `.rels` part’s own folder. ' +
      'Getting that wrong moves every relative target one directory up and is the ' +
      'classic reason a rebuilt deck loses all its images at once. `TargetMode=' +
      '"External"` is passed through untouched and not resolved at all.',
  },
  {
    id: 'V009',
    category: 'relationships',
    severity: 'fatal',
    evidence: 'measured',
    title: 'the relationship edges a part must have, and the ones it must not',
    why:
      'Four findings, none of them in any schema, all found by bisection. ' +
      '(1) A slide needs **exactly one** `slideLayout` relationship resolving to a ' +
      'layout part - zero, two, one pointing at a master and one pointing at a ' +
      'missing part were each built as a single change and each refused; the ' +
      'binding lives only in the rels part, which is why "change layout" is a ' +
      'relationship rewrite. (2) A layout needs exactly one `slideMaster`. ' +
      '(3) A `cx:chartSpace` part with no `.rels` of its own is a whole-package ' +
      'refusal - it needs a `chartStyle` and a `chartColorStyle`, where the classic ' +
      '`c:chartSpace` needs neither. (4) A media part may not be both a media ' +
      'object’s target and a transition’s `p:snd`; two copies of the same bytes ' +
      'open, one shared copy does not.',
  },

  // --- element order ------------------------------------------------------

  {
    id: 'V010',
    category: 'order',
    severity: 'fatal',
    evidence: 'both',
    title: 'children appear in schema-sequence order',
    why:
      'OOXML complex types are `xsd:sequence` almost everywhere, and PowerPoint ' +
      'enforces it. Microsoft’s own Open XML SDK shipped a regression that merely ' +
      'swapped two elements in `slideMaster1.xml` and PowerPoint refused the file. ' +
      'The ranks come from `@pptx-studio/xml`’s table, generated from the ECMA-376 ' +
      'Transitional schemas and checked against 194 148 elements PowerPoint wrote. ' +
      'This rule also covers the two the plan calls out by name, because the table ' +
      'already ranks them: `a:rPr` before `a:t`, and `a:endParaRPr` last in `a:p`.',
  },
  {
    id: 'V011',
    category: 'order',
    severity: 'fatal',
    evidence: 'schema',
    title: '`extLst` is last among its siblings, and every `a:ext` carries a `@uri`',
    why:
      'Its own rule rather than a case of `V010`, because it has to hold in the ' +
      'places the generated table cannot rank - extension markup, `xsd:any` ' +
      'wildcards, and the vocabularies outside the four schemas we generate from. ' +
      '`extLst` is always last in every type that has one, and an `a:ext` without ' +
      'a `@uri` is an extension nobody can identify, which makes it unpreservable: ' +
      'the whole `extLst` contract is "carry it through untouched, keyed by uri".',
  },
  {
    id: 'V012',
    category: 'order',
    severity: 'fatal',
    evidence: 'both',
    title: 'no child appears in a parent whose content model has no place for it',
    why:
      'Ordering catches a child in the wrong *position*; this catches one in the ' +
      'wrong *parent*, which the ordering rule cannot see because an unrankable ' +
      'child is skipped there rather than reported. The case that made it a rule ' +
      'is measured: `a:ahXY` or `a:cxn` as a direct child of `a:custGeom`, without ' +
      'its `a:ahLst`/`a:cxnLst` wrapper, is a whole-package refusal. Only parents ' +
      'the table knows and whose model has no `xsd:any` are checked, and markup in ' +
      'a namespace the table does not cover is left alone.',
  },

  // --- required -----------------------------------------------------------

  {
    id: 'V013',
    category: 'required',
    severity: 'fatal',
    evidence: 'schema',
    title: '`p:presentation` carries `p:notesSz`',
    why:
      'The asymmetry that catches people: in `CT_Presentation`, `p:sldSz` is ' +
      '`[0..1]` and `p:notesSz` is `[1..1]`. A generator that treats the two the ' +
      'same way, or that omits both because the deck has no notes, writes a ' +
      'presentation part that is missing a required child.',
  },
  {
    id: 'V014',
    category: 'required',
    severity: 'fatal',
    evidence: 'schema',
    title: '`p:clrMap` carries all twelve attributes',
    why:
      '`CT_ColorMapping` has twelve attributes - `bg1`, `tx1`, `bg2`, `tx2`, ' +
      '`accent1`…`accent6`, `hlink`, `folHlink` - and every one is required. There ' +
      'is no default and no partial map: eleven of twelve is not a map with one ' +
      'slot missing, it is an invalid element. This is also the element that ' +
      'decides what `bg1` and `tx1` resolve to, so a wrong one is a deck that ' +
      'renders in the wrong colours everywhere at once.',
  },
  {
    id: 'V015',
    category: 'required',
    severity: 'fatal',
    evidence: 'schema',
    title: '`p:spTree` begins with `p:nvGrpSpPr` then `p:grpSpPr`',
    why:
      'A shape tree is a group shape, and `CT_GroupShape` requires both of them ' +
      'first, in that order, before any child shape. An empty slide still has ' +
      'them. This is separate from `V010` because the failure is a missing ' +
      'required child rather than a misordered optional one, and the two need ' +
      'different messages: "add this" against "move this".',
  },
  {
    id: 'V016',
    category: 'required',
    severity: 'fatal',
    evidence: 'schema',
    title: 'every text body has `a:bodyPr` and at least one `a:p`',
    why:
      '`CT_TextBody` is `bodyPr` then optional `lstStyle` then `a:p` at ' +
      '`[1..unbounded]`. A shape whose text was deleted down to nothing is the ' +
      'way this gets broken in an editor: the last paragraph goes, and what is ' +
      'left is a `p:txBody` with no `a:p` in it. An empty paragraph is not the ' +
      'same as no paragraph, and `a:endParaRPr` is where its height comes from.',
  },
  {
    id: 'V017',
    category: 'required',
    severity: 'fatal',
    evidence: 'schema',
    title: '`p:graphicFrame` carries `p:xfrm` and `a:graphic`',
    why:
      'And note the namespace: it is `p:xfrm`, PresentationML, not the `a:xfrm` ' +
      'that every other shape uses. A graphic frame has no placeholder-inherited ' +
      'geometry path, so this is also the reason sub-phase 7.4 gives a rebound ' +
      'chart or table a *copy* of the target layout’s transform instead of ' +
      'deleting its own the way it does for every other shape - delete it there ' +
      'and the frame collapses to zero size.',
  },

  // --- ids ----------------------------------------------------------------

  {
    id: 'V018',
    category: 'ids',
    severity: 'fatal',
    evidence: 'schema',
    title: '`p:sldId/@id` is 256…2147483647 and unique',
    why:
      '`ST_SlideId` is not zero-based and not unbounded. The floor of 256 is the ' +
      'part people get wrong, because nothing about a first slide suggests its id ' +
      'should start there. Four id spaces exist in a presentation and they are not ' +
      'one allocator; this is the first of them.',
  },
  {
    id: 'V019',
    category: 'ids',
    severity: 'fatal',
    evidence: 'both',
    title: 'master and layout ids are ≥ 2147483648 and unique **across both lists**',
    why:
      '`ST_SlideMasterId` and `ST_SlideLayoutId` are both "2147483648 and up" and ' +
      'no schema says they may not overlap. PowerPoint allocates them from one ' +
      'running counter - master, its layouts, next master, its layouts - and ' +
      'refuses a package that does not. Found by bisection when `a12-masters` ' +
      'would not open: two counters are indistinguishable from correct with one ' +
      'master and collide on the second, so **every** two-master package was ' +
      'refused.',
  },
  {
    id: 'V020',
    category: 'ids',
    severity: 'fatal',
    evidence: 'both',
    title: '`p:cNvPr/@id` is unique within its part and ≤ 2147483647',
    why:
      'Unique *within a part*, and free to repeat across parts - a shape id is not ' +
      'a document-wide identity and treating it as one is how a duplicate-slide ' +
      'implementation ends up renumbering things it should not. The range is ' +
      'measured rather than read: `ST_DrawingElementId` is `xsd:unsignedInt`, but ' +
      'every value from 0 to 2147483647 opens, 2147483648…4294967294 are ' +
      'whole-package refusals, and 4294967295 opens - PowerPoint reads the ' +
      'attribute as a signed 32-bit integer and keeps `0xFFFFFFFF` as a sentinel ' +
      'it renumbers away on save. `a39-large-ids` is the corpus deck.',
  },
  {
    id: 'V021',
    category: 'ids',
    severity: 'warning',
    evidence: 'schema',
    title: 'a slide placeholder’s `(type, idx)` has a counterpart in its layout',
    why:
      'A `p:ph` that matches nothing in the layout inherits nothing: no geometry, ' +
      'no text style, no prompt. PowerPoint opens the file and draws the shape at ' +
      'whatever it can work out, which is why this is a warning and not fatal - ' +
      'the deck is not refused, it is silently wrong. Matching follows the five ' +
      'tiers sub-phase 7.1 hardens, and the asymmetry that matters is checked ' +
      'here too: slide→layout matches on `(type, idx)` where layout→master ' +
      'matches on **type only**.',
  },

  // --- refused ------------------------------------------------------------

  {
    id: 'V022',
    category: 'refused',
    severity: 'fatal',
    evidence: 'measured',
    title: '`p:ph/@type` is not `hdr` or `sldImg` outside a notes or handout part',
    why:
      'A whole-package refusal, on a slide layout and on a slide alike, either one ' +
      'alone with no other change. The other seven content types were built as ' +
      'one-type packages in the same bisection and every one opens: `obj`, ' +
      '`chart`, `tbl`, `clipArt`, `dgm`, `media`, `pic`. Nothing in the schema ' +
      'says so - `CT_Placeholder` is one complex type shared by masters, layouts, ' +
      'slides, notes slides and handout masters, and `ST_PlaceholderType` is one ' +
      'enumeration holding all sixteen values. The restriction is real and ' +
      'unwritten: those two belong to the notes and handout families.',
  },
  {
    id: 'V023',
    category: 'refused',
    severity: 'fatal',
    evidence: 'measured',
    title: 'every geometry guide named is a guide that is defined',
    why:
      'A whole-package refusal, not a repair and not a dropped shape. It applies ' +
      'both to an `a:gd` formula naming an adjust value no `a:avLst` defines and ' +
      'to an `a:pt` coordinate naming a guide no `a:gdLst` defines. ' +
      '`ST_GeomGuideName` is an unconstrained token, so this is schema-legal ' +
      'markup refused for a reason no schema states - and it is the rule that ' +
      'sub-phase 2.5 needs, because dragging an adjust handle is exactly the ' +
      'operation that can write a formula referring to a guide that is no longer ' +
      'there.',
  },
  {
    id: 'V024',
    category: 'refused',
    severity: 'fatal',
    evidence: 'both',
    title: '`c:tx` holds `c:strRef` or `c:v`, never `c:strLit`',
    why:
      '`CT_SerTx` is a choice of exactly those two. The trap is that `c:cat` and ' +
      '`c:val`, two elements away in the same series, both accept the literal ' +
      'forms - so the markup looks like something that ought to work, and it is a ' +
      'whole-package refusal.',
  },
  {
    id: 'V025',
    category: 'refused',
    severity: 'fatal',
    evidence: 'measured',
    title: 'no `p:control`',
    why:
      'A whole-package refusal in all eight forms tried: bare, name-only, with and ' +
      'without `r:id`, with a `p:pic` preview, with an ActiveX part and its `.bin`, ' +
      'and inside a macro-enabled package. An empty `p:controls` is accepted, ' +
      'which places the refusal precisely on the child element. We do not author ' +
      'ActiveX controls and never will; the rule exists so that a file that has ' +
      'one is not handed back with it re-serialised into a form PowerPoint likes ' +
      'less than the original.',
  },
  {
    id: 'V026',
    category: 'refused',
    severity: 'fatal',
    evidence: 'measured',
    title: 'a `cs:chartStyle` carries all thirty-one of its entries',
    why:
      'Four entries refused; thirty-one opened. And the shape of the finding is ' +
      'the interesting half: a chart with **no** chart-style relationship at all is ' +
      'fine, so this is not "the part is required" - it is "if the part exists it ' +
      'must be complete". A subset is worse than an absence, which is the opposite ' +
      'of what a partial-styling model would predict.',
  },

  // --- preservation -------------------------------------------------------

  {
    id: 'V027',
    category: 'preservation',
    severity: 'fatal',
    evidence: 'schema',
    needsBaseline: true,
    title: 'a part nobody edited comes back out byte-for-byte',
    why:
      'The load-bearing one, and the reason the architecture is what it is. ' +
      'Preserving charts, SmartArt, animations, OLE and macros is unachievable in ' +
      'any design where the writer has to *understand* a feature to emit it, so ' +
      'preservation is the default state rather than a feature - and this rule is ' +
      'where that stops being an intention. `ppt/embeddings/*.bin` is called out ' +
      'inside it: those are OLE2/CFB compound files and any rewrite of one is a ' +
      'guaranteed repair prompt.',
  },
  {
    id: 'V028',
    category: 'preservation',
    severity: 'fatal',
    evidence: 'schema',
    needsBaseline: true,
    title: 'no `mc:AlternateContent` branch and no `extLst` was rebuilt',
    why:
      'Two things we hold opaque, for one reason: we do not know what is in them. ' +
      'An `mc:AlternateContent` branch we do not understand must arrive and leave ' +
      'identical, because `mc:Choice/@Requires` names a *prefix* and rewriting the ' +
      'branch is how ignorable extension markup becomes a hard error. An `extLst` ' +
      'is an ordered list keyed by `@uri` and never a typed model that is rebuilt ' +
      'from fields - `a34-extlst` measured that PowerPoint carries an unknown ' +
      '`a:ext/@uri` through untouched, so anything we drop there is something the ' +
      'file would otherwise have kept forever.',
  },
  {
    id: 'V029',
    category: 'preservation',
    severity: 'fatal',
    evidence: 'both',
    needsBaseline: true,
    title: 'text and field identity survive: `xml:space`, `a:fld/@id`, cached field text',
    why:
      'Three small things that are each silently destructive. An `a:t` with ' +
      'leading or trailing whitespace needs `xml:space="preserve"` if it was ' +
      'written with it, because dropping the attribute changes what the text *is*. ' +
      '`a:fld/@id` is a required `ST_Guid` and regenerating one can make PowerPoint ' +
      'repair the file. And a field’s cached `a:t` is the only thing that renders ' +
      'when we cannot evaluate the field ourselves, so discarding it turns a date ' +
      'placeholder into an empty box on every consumer that is not PowerPoint.',
  },
] as const satisfies readonly RuleShape[];

export type RuleId = (typeof RULES)[number]['id'];

const BY_ID = new Map<string, Rule>(RULES.map((rule) => [rule.id, rule as Rule]));

export function ruleById(id: string): Rule | undefined {
  return BY_ID.get(id);
}

export const RULE_IDS: readonly RuleId[] = RULES.map((rule) => rule.id);

/** Rules that need the package as it was opened. See `Rule.needsBaseline`. */
export const BASELINE_RULES: readonly RuleId[] = RULES.filter(
  (rule) => 'needsBaseline' in rule,
).map((rule) => rule.id);
