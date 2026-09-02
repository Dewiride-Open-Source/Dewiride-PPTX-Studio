/**
 * `C-LEX`'s declared inventory: every lexical form the committed corpus
 * contains, and every form this project has measured somewhere else and the
 * corpus does not.
 *
 * `lexical.ts` is the mechanism - it scans bytes and reports what is there.
 * This file is the reviewed statement of what ought to be there and why, in the
 * same relationship `gen/decks/*.ts` has to `gen/build-probes.ts` and
 * `written/decks.ts` has to its builder. `lexical.test.ts` asserts the two
 * against each other in both directions.
 *
 * ## What a row means
 *
 * `producers` is pinned rather than computed, for the reason a census is
 * pinned: a change in coverage should be a diff somebody reads, not a number
 * that quietly moves. Deck counts are deliberately not pinned - which decks
 * happen to carry a form is churn, and how many distinct serializers wrote it
 * is the rule.
 *
 * A row exists when the form is observed in the corpus **or** when this project
 * has measured it elsewhere and can cite where. Without that filter the table
 * becomes a wish list: every construct XML 1.0 admits would qualify, and a rule
 * that lists everything imaginable says nothing about what matters.
 *
 * ## The two measurements everything here cites
 *
 * **ADR 0004** scanned all 2834 XML parts of 37 real Office packages at the
 * byte level - the nine `.potx` templates that ship with Office 16 and 28
 * sample decks. That is the corpus's notion of what is normal, and none of it
 * is committed here, which is why it has to be cited rather than checked.
 *
 * **E8** measured one PowerPoint 16.0.20326 save across 38 parts, and lives in
 * `corpus/ground-truth/powerpoint-conventions.json`. It is narrower and it is
 * ours; where the two disagree the wider one is describing the wild and the
 * narrower one is describing this build.
 */

/**
 * The four serializers that have written a byte of this corpus.
 *
 * `GEN` and `ZIP` are separate because `tools/corpus/gen` hands its entries to
 * `writeZip` and decides none of the header fields itself. Two decks whose
 * containers both came out of `writeZip` are one piece of evidence about ZIP
 * conventions, whichever generator called it.
 *
 * `PPT` is spelled once and used by two collections on purpose: Tier C's XML
 * came out of Tier B unchanged, so a form found in both has one producer and
 * not two. The collision is what enforces the caveat.
 */
export const SERIALIZER = {
  GEN: 'tools/corpus/gen',
  ZIP: 'tools/ground-truth/zip.ts',
  PPT: 'Microsoft PowerPoint 16.0.20326',
  OPC: 'packages/opc',
} as const;

const { GEN, ZIP, PPT, OPC } = SERIALIZER;

/** A form this project knows about, and where it stands in the corpus. */
export interface LexicalForm {
  readonly dimension: string;
  readonly form: string;
  /** One line: what the form is. */
  readonly what: string;
  /** Evidence the form occurs outside this repository, or the counter-evidence. */
  readonly wild: string;
  /** Serializers that emit it here, sorted. Empty means no deck exercises it. */
  readonly producers: readonly string[];
  /** Required below two producers: why it stands there, and what would close it. */
  readonly gap?: string;
}

/**
 * The gap shared by the four named entity references, which stand or fall
 * together because one deck would close all four.
 */
function entityGap(spelling: string, wildCount: number): string {
  return (
    'Every entity reference in this corpus was written by `tools/corpus/gen`. The nine ' +
    'PowerPoint-authored decks contain **no ampersand at all** across their 387 XML parts, so our ' +
    'escaping is only ever checked against our own escaper - which is the exact failure `C-LEX` ' +
    'exists to name. It is not a limit of PowerPoint: E8 measured it writing `&amp;`, `&lt;` and ' +
    '`&gt;` inside an `a:t` and `' +
    spelling +
    '` is one of the ' +
    String(wildCount) +
    ' occurrences ADR 0004 counted in real files, and a follow-up on 2026-08-27 caught it writing ' +
    '`name="a &quot;quoted&quot; name"` for a renamed shape. The decks both measurements used ' +
    'were never committed. One more Tier B deck - a title and a shape name carrying `& < > "` - ' +
    'closes all four forms at once, and is the cheapest gap in this table to close.'
  );
}

/**
 * Why several of these gaps cannot be closed by authoring another deck.
 *
 * ADR 0004 opened 62 hand-built lexical variants in the installed PowerPoint,
 * then re-saved the ones it accepted and diffed the slide part. PowerPoint
 * accepts far more than it writes, and what it does on save is the constraint
 * that matters here: comments and processing instructions are **discarded**, a
 * CDATA section is **rewritten as plain text**, `&#72;` is **resolved to `H`**,
 * a single-quoted attribute is **rewritten double**, `<p:spPr></p:spPr>` is
 * **collapsed**, a byte order mark is **stripped**.
 *
 * So for those forms our second producer is not merely silent - it is incapable.
 * No deck PowerPoint writes can ever carry one, and the gap closes only with a
 * producer this corpus does not have. Saying "awaiting a second producer" of
 * them would imply a deck could be authored that never can be.
 */
function normalisedAway(fate: string): string {
  return (
    'Structurally unreachable from our second producer rather than merely missing: PowerPoint ' +
    fate +
    ' on save (ADR 0004 opened 62 lexical variants and diffed the re-saved slide part), so no ' +
    'deck it writes can ever carry one. Closing this needs a producer the corpus does not have, ' +
    'and the form is here because `packages/xml` must still preserve it - a deck that has been ' +
    'through PowerPoint is no longer a fixture for `C-LEX`.'
  );
}

/** The three container forms that belong to a unit fixture rather than a deck. */
const CONTAINER_FIXTURE_GAP =
  'Recorded here to say where it *is* covered rather than to ask for a deck. ADR 0002 found that ' +
  'Office writes none of the four ZIP features hardest to handle - no data descriptors, no ' +
  'directory entries, no UTF-8-flagged names, no ZIP64 - so a corpus of Office output could never ' +
  'have exercised them, and all four live in `packages/opc/src/zip-reader.test.ts` instead. ' +
  '`tools/ground-truth/zip.ts` cannot write the first two either: bit 3 moves the sizes into a ' +
  'trailing descriptor it does not emit, so setting it would produce an archive that lies about ' +
  'itself. A committed deck is the wrong home for these.';

export const LEXICAL_FORMS: readonly LexicalForm[] = [
  // ------------------------------------------------------ the byte order mark
  {
    dimension: 'xml.bom',
    form: 'absent',
    what: 'the part begins at `<`',
    wild: '1797 of ADR 0004’s 2834 real parts carry no BOM',
    producers: [PPT, GEN],
  },
  {
    dimension: 'xml.bom',
    form: 'present',
    what: 'the part begins `EF BB BF`',
    wild:
      '**1037 of 2834** real parts carry one, and one package mixes both: `Pitchbook.potx` has ' +
      'parts with no BOM and CRLF, and parts with a BOM and LF (ADR 0004)',
    producers: [],
    gap:
      'No producer available here writes a BOM. PowerPoint 16.0.20326 wrote none across E8’s ' +
      '38 parts, and neither of our writers emits one - `conventions.test.ts` requires that of ' +
      'them. Unrepresented is not untested: the tokenizer keeps a BOM rather than swallowing it ' +
      '(`ignoreBOM: true`, an option named backwards, ADR 0004) and `packages/xml` pins that with ' +
      'its own fixture. Closing it needs a deck we did not author, and not one PowerPoint wrote ' +
      'either: ADR 0004 added a BOM to a part that had none, and PowerPoint opened the file and ' +
      'then stripped the BOM on save.',
  },

  // ------------------------------------------------------------- declarations
  {
    dimension: 'xml.declaration',
    form: 'present',
    what: 'the part opens with an XML declaration',
    wild: 'all 2834 real parts have one - ADR 0004’s 2161 + 656 + 17',
    producers: [PPT, GEN],
  },
  {
    dimension: 'xml.declaration',
    form: 'absent',
    what: 'the part opens at its root element',
    wild: '0 of 2834 (ADR 0004)',
    producers: [PPT],
    gap:
      'One producer, and it is the one that writes a declaration everywhere else: PowerPoint’s ' +
      'chart component omits it. Four of the corpus’s 1422 XML parts, all in `b05-chart` - ' +
      '`ppt/charts/style1.xml`, `style2.xml`, `colors1.xml` and `colors2.xml` begin at ' +
      '`<cs:chartStyle` and `<cs:colorStyle`. ADR 0004 scanned 2834 real parts and found no such ' +
      'part; E8 scanned 38 and found none either. What matters is not a second producer but that ' +
      'nothing here adds one back, which `packages/opc` gets for free by never re-serializing a ' +
      'part it did not edit.',
  },
  {
    dimension: 'xml.declarationEncoding',
    form: 'UTF-8',
    what: '`encoding="UTF-8"`, upper case',
    wild: 'one of the three declarations ADR 0004 counted',
    producers: [PPT, GEN],
  },
  {
    dimension: 'xml.declarationEncoding',
    form: 'utf-8',
    what: '`encoding="utf-8"`, lower case',
    wild: 'two of the three declarations ADR 0004 counted spell it this way',
    producers: [],
    gap:
      'The spelling most real declarations use, and neither producer here writes it. This is the ' +
      'form behind ADR 0004’s conclusion that a serializer emitting a canonical declaration ' +
      'loses byte fidelity on most of the corpus. `packages/xml` keeps the declaration as source ' +
      'bytes and never rebuilds it, so the behaviour is decided - but no committed deck exercises ' +
      'the decision.',
  },
  {
    dimension: 'xml.declarationStandalone',
    form: 'yes',
    what: '`standalone="yes"`',
    wild: 'two of the three declarations ADR 0004 counted',
    producers: [PPT, GEN],
  },
  {
    dimension: 'xml.declarationStandalone',
    form: 'omitted',
    what: 'no `standalone` pseudo-attribute',
    wild: 'one of the three declarations ADR 0004 counted',
    producers: [],
    gap: 'The same family as the lower-case encoding, and one deck closes both.',
  },
  {
    dimension: 'xml.afterDeclaration',
    form: 'crlf',
    what: 'CRLF between the declaration and the root element',
    wild: '656 of 2834 (ADR 0004)',
    producers: [PPT, GEN],
  },
  {
    dimension: 'xml.afterDeclaration',
    form: 'none',
    what: 'the root element follows the declaration immediately',
    wild: '**2161 of 2834** - the majority form (ADR 0004)',
    producers: [],
    gap:
      'The commonest arrangement in the wild, and absent here: all 1418 declared parts in this ' +
      'corpus are followed by CRLF, because both producers write it. ADR 0005 measured the one ' +
      'consequence - 656 real parts differ under a *forced* rebuild, because CRLF and LF are the ' +
      'same value after XML 1.0 §2.11 and a text node rebuilt from its value cannot know ' +
      'which it came from - so the behaviour is measured on real files and unrepresented in the ' +
      'corpus.',
  },
  {
    dimension: 'xml.afterDeclaration',
    form: 'lf',
    what: 'a bare LF between the declaration and the root element',
    wild: '17 of 2834 (ADR 0004)',
    producers: [],
    gap: 'The rarest of ADR 0004’s three, and the one a CRLF-writing producer cannot reach.',
  },
  {
    dimension: 'xml.lineBreaks',
    form: 'declaration-break',
    what: 'exactly one newline in the part, the CRLF after the declaration',
    wild: 'the 656 parts of ADR 0004’s corpus that carry a CRLF there',
    producers: [PPT, GEN],
  },
  {
    dimension: 'xml.lineBreaks',
    form: 'one-line',
    what: 'no newline anywhere in the part',
    wild: 'the 2161 real parts with nothing after the declaration, so far as ADR 0004 reaches',
    producers: [PPT],
    gap:
      'One producer, and weaker than it looks: the four parts that are one line are exactly the ' +
      'four that omit the declaration, so the corpus cannot separate the two facts. Neither is ' +
      'independent evidence for the other until some deck carries one without the other.',
  },
  {
    dimension: 'xml.lineBreaks',
    form: 'multi-line',
    what: 'a pretty-printed part',
    wild: 'not measured; ADR 0004 found no mixed content and no indented part among the 2834',
    producers: [],
    gap:
      'No producer here writes one and none was found in the wild, but every byte of that ' +
      'whitespace would be significant to a gate that re-emits parts unchanged. The row exists to ' +
      'make the absence a statement rather than an oversight.',
  },

  // ------------------------------------------------ element and attribute lexis
  {
    dimension: 'xml.selfClosing',
    form: 'tight',
    what: '`<a:off x="0" y="0"/>`',
    wild: '27 955 of ADR 0004’s 98 777 self-closing tags',
    producers: [PPT, GEN],
  },
  {
    dimension: 'xml.selfClosing',
    form: 'spaced',
    what: '`<a:off x="0" y="0" />`, with whitespace before `/>`',
    wild: '**70 822 of 98 777** - the dominant real-world form (ADR 0004)',
    producers: [GEN],
    gap:
      'The largest gap in this table: the form most real files use has one producer here, and it ' +
      'is ours. E8’s PowerPoint build wrote 0 spaced of 1305, so `a36-spaced-tags` is the ' +
      'only witness and it was built to be one. A second producer exists and is not committed: ' +
      'ADR 0009 records that `ppt/charts/chartEx1.xml`, in a deck PowerPoint wrote on 2026-08-27, ' +
      'spaced 12 of its 12 while the other 3046 tags across the same sixty parts spaced none. ' +
      '`b05-chart` carries `c:` charts and no ChartEx part, so a tenth Tier B deck with a ChartEx ' +
      'chart would close this with a genuinely independent producer.',
  },
  {
    dimension: 'xml.emptyElement',
    form: 'self-closed',
    what: '`<x/>`',
    wild: '98 777 self-closing tags across 2834 parts (ADR 0004)',
    producers: [PPT, GEN],
  },
  {
    dimension: 'xml.emptyElement',
    form: 'paired',
    what: '`<x></x>` - an empty element written long',
    wild:
      '33, all in `docProps` (ADR 0004), and measured again here independently: PowerPoint writes ' +
      '`<dc:title></dc:title>`, `<dc:creator></dc:creator>`, ' +
      '`<cp:lastModifiedBy></cp:lastModifiedBy>` and `<Company></Company>` into every deck it ' +
      'saves, and nothing else',
    producers: [PPT, GEN],
  },
  {
    dimension: 'xml.attrQuote',
    form: 'double',
    what: 'an attribute value delimited by a double quote',
    wild: '170 019 of 170 019 (ADR 0004)',
    producers: [PPT, GEN],
  },
  {
    dimension: 'xml.attrQuote',
    form: 'single',
    what: 'an attribute value delimited by an apostrophe',
    wild: '**0 of 170 019** (ADR 0004). Legal in production [10]; emitted by no producer measured',
    producers: [GEN],
    gap:
      'No wild evidence at all, and saying so is the point. `a36-spaced-tags` writes one because ' +
      'ADR 0005 chose to record the quote character per **attribute** rather than per part, and ' +
      'this is the only fixture that can fail if that regresses - the one form here a serializer ' +
      'cannot reproduce by remembering a single boolean. And PowerPoint can never be the second: ' +
      'ADR 0004 measured it opening a single-quoted attribute and rewriting it double on save, so ' +
      'this is unreachable rather than merely missing.',
  },
  {
    dimension: 'xml.attrEquals',
    form: 'tight',
    what: '`name="value"`',
    wild: 'the form every one of ADR 0004’s 170 019 attributes uses',
    producers: [PPT, GEN],
  },
  {
    dimension: 'xml.attrEquals',
    form: 'spaced',
    what: '`name = "value"` - whitespace either side of the equals sign',
    wild: 'not counted by ADR 0004; production [41] admits it, `Eq ::= S? "=" S?`',
    producers: [GEN],
    gap:
      'The same standing as the single quote: legal, unobserved in the wild, exercised by `a36`. ' +
      'PowerPoint opens it (ADR 0004) and almost certainly normalises it away, though this form was ' +
      'not separately diffed on resave the way the quote character and the self-closing spacing ' +
      'were.',
  },
  {
    dimension: 'xml.tagClose',
    form: 'tight',
    what: 'a tag ending `>` with no whitespace before it',
    wild: 'the ordinary form throughout ADR 0004’s corpus',
    producers: [PPT, GEN],
  },
  {
    dimension: 'xml.tagClose',
    form: 'spaced-start',
    what: 'whitespace before `>` on a start tag',
    wild: 'not counted by ADR 0004; production [40] admits it',
    producers: [GEN],
    gap:
      'Legal, unobserved in the wild, exercised by `a36`. PowerPoint opens it and normalises the ' +
      'lexical form of any part it rewrites, so it is not a form a tenth Tier B deck could supply.',
  },
  {
    dimension: 'xml.tagClose',
    form: 'spaced-end',
    what: 'whitespace before `>` on an end tag',
    wild: 'not counted by ADR 0004; production [42] admits it',
    producers: [GEN],
    gap:
      'Legal, unobserved in the wild, exercised by `a36`. PowerPoint opens it and normalises the ' +
      'lexical form of any part it rewrites, so it is not a form a tenth Tier B deck could supply.',
  },

  // ---------------------------------------------------------------- references
  {
    dimension: 'xml.entity',
    form: 'amp',
    what: 'the ampersand reference',
    wild: '3 of the 39 references ADR 0004 counted',
    producers: [GEN],
    gap: entityGap('&amp;', 3),
  },
  {
    dimension: 'xml.entity',
    form: 'lt',
    what: 'the less-than reference',
    wild: '5 of 39 (ADR 0004)',
    producers: [GEN],
    gap: entityGap('&lt;', 5),
  },
  {
    dimension: 'xml.entity',
    form: 'gt',
    what: 'the greater-than reference',
    wild:
      '5 of 39 (ADR 0004). ADR 0005 escapes it unconditionally, having measured 0 literal `>` ' +
      'against 5 references',
    producers: [GEN],
    gap: entityGap('&gt;', 5),
  },
  {
    dimension: 'xml.entity',
    form: 'quot',
    what: 'the double-quote reference',
    wild: '26 of 39 - the commonest (ADR 0004)',
    producers: [GEN],
    gap: entityGap('&quot;', 26),
  },
  {
    dimension: 'xml.entity',
    form: 'decimal',
    what: 'a decimal character reference',
    wild: '**0 of 2834 parts** (ADR 0004)',
    producers: [GEN],
    gap:
      'No wild evidence: not one numeric character reference across 2834 real parts. `a29-math` ' +
      'and `a40-unicode` write them because the tokenizer must decode them and §4.1 makes ' +
      'them legal wherever a character is, not because a producer was seen to. PowerPoint cannot ' +
      'be the second producer: ADR 0004 measured it resolving `&#72;` to a literal `H` on save.',
  },
  {
    dimension: 'xml.entity',
    form: 'hex-lower',
    what: 'a hexadecimal character reference',
    wild: '0 of 2834 (ADR 0004)',
    producers: [GEN],
    gap:
      'As the decimal form, and unreachable for the same reason. `a40-unicode` carries the astral, ' +
      'combining and bidi cases, which is where the hexadecimal spelling earns its keep - and where ' +
      'ADR 0009 found that a `&#13;` does not survive PowerPoint at all, coming back as a paragraph ' +
      'break.',
  },
  {
    dimension: 'xml.entity',
    form: 'hex-upper-ill-formed',
    what: 'an upper-case `X` after the hash - not a character reference at all',
    wild:
      'never legal: production [66] spells the prefix as a terminal in a case-sensitive grammar, ' +
      'so this is ill-formed rather than rare',
    producers: [],
    gap:
      'A tripwire rather than a gap. `packages/xml` refuses it, and ADR 0004 records that the ' +
      'refusal has still not been checked against PowerPoint. This row exists so a deck growing ' +
      'one fails `C-LEX` loudly instead of being catalogued as a convention; the home for an ' +
      'ill-formed fixture is `corpus/reject/`, not here.',
  },

  // ------------------------------------------------- markup beyond elements
  {
    dimension: 'xml.markup',
    form: 'comment',
    what: 'an XML comment',
    wild: '0 of 2834 (ADR 0004). PowerPoint opens a file with one and discards it on save',
    producers: [GEN],
    gap: normalisedAway('discards comments'),
  },
  {
    dimension: 'xml.markup',
    form: 'cdata',
    what: 'a CDATA section',
    wild: '0 of 2834 (ADR 0004). PowerPoint opens one and rewrites it as plain text on save',
    producers: [GEN],
    gap: normalisedAway('rewrites a CDATA section as plain text'),
  },
  {
    dimension: 'xml.markup',
    form: 'processing-instruction',
    what: 'a processing instruction after the declaration',
    wild: '0 of 2834 (ADR 0004). PowerPoint opens one and discards it on save',
    producers: [],
    gap:
      normalisedAway('discards processing instructions') +
      ' Unlike the comment and the CDATA section it is not exercised at all: `a34-extlst` carries ' +
      'those two and no deck carries this one, so the serializer path for it is pinned only by ' +
      "`packages/xml`'s own tests.",
  },
  {
    dimension: 'xml.markup',
    form: 'doctype',
    what: 'a document type declaration',
    wild:
      '0 of 2834 (ADR 0004), and the one construct PowerPoint **refuses**: `<!DOCTYPE p:sld>` ' +
      'fails to open with `0x80070570`',
    producers: [],
    gap:
      'Not a gap and never will be one. The tokenizer rejects DOCTYPE outright, which is what ' +
      'removes the XXE and billion-laughs surface entirely, and PowerPoint refuses the file too. ' +
      'The row exists so that a deck growing one fails here rather than being catalogued as a ' +
      'convention; a file we refuse to parse belongs in `corpus/reject/` under `C-REJECT`.',
  },

  // ---------------------------------------------------------------- the container
  {
    dimension: 'zip.versionMadeBy',
    form: '20',
    what: 'ZIP 2.0, host 0',
    wild: 'the value both of our writers emit',
    producers: [OPC, ZIP],
  },
  {
    dimension: 'zip.versionMadeBy',
    form: '45',
    what: 'ZIP 4.5, host 0',
    wild: 'PowerPoint 16.0.20326, on every entry of every deck it wrote here',
    producers: [PPT],
    gap:
      'One producer, in the field nothing reads. E8 did not record it, and `packages/opc`’s ' +
      'own header table called 20 the value "as PowerPoint writes" for two sub-phases, until Tier ' +
      'C put the two side by side. A second producer means a non-Microsoft authoring tool, which ' +
      'this corpus has no licensed source for.',
  },
  {
    dimension: 'zip.versionNeeded',
    form: '10',
    what: 'no compression needed - a stored entry',
    wild: 'PowerPoint writes it on stored entries, and so does `packages/opc`',
    producers: [PPT, OPC],
  },
  {
    dimension: 'zip.versionNeeded',
    form: '20',
    what: 'deflate needed',
    wild:
      'PowerPoint on deflated entries, and both of our writers. `tools/ground-truth/zip.ts` writes ' +
      'it on stored entries too, where PowerPoint writes 10, which is why Tier A contributes ' +
      'nothing to the `10` row. Deliberately not fixed: changing it would re-pin all 41 Tier A ' +
      'hashes for a field no reader consults',
    producers: [PPT, OPC, ZIP],
  },
  {
    dimension: 'zip.method',
    form: 'stored',
    what: 'method 0',
    wild:
      'PowerPoint stores its thumbnail and media; Tier A stores everything, so its hashes ' +
      'do not depend on which zlib built them',
    producers: [PPT, OPC, ZIP],
  },
  {
    dimension: 'zip.method',
    form: 'deflate',
    what: 'method 8',
    wild: 'PowerPoint deflates every XML part',
    producers: [PPT, OPC, ZIP],
  },
  {
    dimension: 'zip.deflateLevelHint',
    form: 'super-fast',
    what: 'general-purpose bits 1 and 2 both set',
    wild: 'PowerPoint sets it on every deflated entry and on none of its stored ones',
    producers: [PPT, ZIP],
  },
  {
    dimension: 'zip.deflateLevelHint',
    form: 'normal',
    what: 'bits 1 and 2 clear',
    wild: 'what `packages/opc` writes, deflating at level 6',
    producers: [OPC, ZIP],
  },
  {
    dimension: 'zip.deflateLevelHint',
    form: 'maximum',
    what: 'bit 1 set',
    wild: 'any archiver that sets it; not measured on a real deck',
    producers: [ZIP],
    gap:
      '`a35-zip-shapes` sets it so that a reader which acted on the level bits would be caught. ' +
      'Nothing in this project reads them, which is exactly why a fixture has to carry them.',
  },
  {
    dimension: 'zip.deflateLevelHint',
    form: 'fast',
    what: 'bit 2 set',
    wild: 'any archiver that sets it; not measured on a real deck',
    producers: [ZIP],
    gap: 'As `maximum`, and in the same deck.',
  },
  {
    dimension: 'zip.utf8NameFlag',
    form: 'absent',
    what: 'bit 11 clear - the entry name is CP437',
    wild: 'ADR 0002: no UTF-8-flagged name across 37 real packages',
    producers: [PPT, OPC, ZIP],
  },
  {
    dimension: 'zip.utf8NameFlag',
    form: 'present',
    what: 'bit 11 set - the entry name is UTF-8',
    wild: 'ADR 0002: Office writes none',
    producers: [ZIP],
    gap:
      'One producer, and it can only ever be one: an OPC part name is ASCII by grammar, so a ' +
      'conformant producer never needs the flag. `a35-zip-shapes` sets it because a reader must ' +
      'still not be confused by it.',
  },
  {
    dimension: 'zip.localExtra',
    form: 'none',
    what: 'no local extra field',
    wild: 'PowerPoint writes one on five entries of a package and none on the rest',
    producers: [PPT, OPC, ZIP],
  },
  {
    dimension: 'zip.localExtra',
    form: 'growth-hint',
    what: 'Microsoft’s `0xA220` growth hint',
    wild: 'PowerPoint writes 520, 520, 264, 264 and 264 bytes of it on the first five entries',
    producers: [PPT, ZIP],
  },
  {
    dimension: 'zip.centralExtra',
    form: 'none',
    what: 'no central-directory extra field',
    wild: 'no producer measured here writes one',
    producers: [PPT, OPC, ZIP],
  },
  {
    dimension: 'zip.entryComment',
    form: 'none',
    what: 'no per-entry comment',
    wild: 'no producer measured here writes one',
    producers: [PPT, OPC, ZIP],
  },
  {
    dimension: 'zip.dosDateTime',
    form: '1980-01-01T00:00',
    what: 'the DOS epoch',
    wild: 'PowerPoint writes it on every entry, and so do both of our writers',
    producers: [PPT, OPC, ZIP],
  },
  {
    dimension: 'zip.firstEntry',
    form: 'content-types',
    what: '`[Content_Types].xml` is entry 0',
    wild: 'required by OPC and honoured by every producer measured',
    producers: [PPT, OPC, ZIP],
  },
  {
    dimension: 'zip.directoryEntries',
    form: 'absent',
    what: 'no zero-length entry whose name ends in a slash',
    wild: 'ADR 0002: Office writes none across 37 real packages',
    producers: [PPT, OPC, ZIP],
  },
  {
    dimension: 'zip.directoryEntries',
    form: 'present',
    what: 'a zero-length entry whose name ends in a slash',
    wild: 'ADR 0002: not written by Office',
    producers: [],
    gap: CONTAINER_FIXTURE_GAP,
  },
  {
    dimension: 'zip.dataDescriptor',
    form: 'absent',
    what: 'bit 3 clear - the sizes are in the local header',
    wild: 'ADR 0002: Office writes none',
    producers: [PPT, OPC, ZIP],
  },
  {
    dimension: 'zip.dataDescriptor',
    form: 'present',
    what: 'bit 3 set - the sizes trail the payload',
    wild: 'ADR 0002: not written by Office; written by streaming producers generally',
    producers: [],
    gap: CONTAINER_FIXTURE_GAP,
  },
  {
    dimension: 'zip.zip64',
    form: 'absent',
    what: 'no ZIP64 end-of-central-directory locator',
    wild: 'ADR 0002: Office writes none',
    producers: [PPT, OPC, ZIP],
  },
  {
    dimension: 'zip.zip64',
    form: 'present',
    what: 'a ZIP64 end-of-central-directory record',
    wild: 'ADR 0002: other producers emit one even for small archives',
    producers: [],
    gap: CONTAINER_FIXTURE_GAP,
  },
];
