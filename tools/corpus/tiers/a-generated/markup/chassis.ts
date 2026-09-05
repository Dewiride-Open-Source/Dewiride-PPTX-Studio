/**
 * The package chassis for Tier A corpus decks.
 *
 * ## Why this is a third builder and not a reuse of one of the two
 *
 * The repository already hand-writes PresentationML in two places, and neither
 * is the right host for forty feature probes:
 *
 * - `tools/bench/deck.ts` is a *recipe interpreter* streaming to a file
 *   descriptor, built for three hundred slides and two hundred megabytes. Its
 *   vocabulary is cadences - "a chart every twenty-fifth slide". A probe deck
 *   wants the opposite: one slide saying exactly one thing. Expressing forty
 *   probes as recipe fields would add roughly sixty booleans to `DeckRecipe`
 *   and a `writeDeck` nobody can read, and the probes could still only say what
 *   the loop already knows how to say.
 * - `tools/ground-truth/lib/pptx.ts` is closer in shape but is a sub-phase 0.7
 *   artifact, and two committed measurements - `color-transforms.json` and
 *   `eot-headers.json` - were taken from decks it built. Changing what it emits
 *   would make those measurements describe a deck that no longer exists.
 *
 * So this one is written fresh and is the **canonical** builder from here on.
 *
 * The obvious hazard in having three is three divergent notions of correct
 * markup, and sharing code is the weak answer to it - it makes the three agree
 * without making any of them right. `conventions.test.ts` is the strong answer:
 * it asserts every builder in the repository matches the byte conventions
 * measured from PowerPoint itself in `corpus/ground-truth/powerpoint-conventions.json`.
 * That checks the property that actually matters, holds for a builder written
 * tomorrow, and would have caught a divergence that shared code cannot.
 *
 * ## Entries are stored, not deflated
 *
 * ADR 0009 committed the bytes rather than the recipes because a recipe's hash
 * is the hash of a *build*: `node:zlib` and `fflate` both move, and a routine
 * upgrade would invalidate every pinned hash with no source change. Committing
 * the bytes fixes the legal half of that - the hash is of the reviewed object -
 * but not `C-REGEN`, which rebuilds each deck and compares. Deflate would put
 * zlib's version in that comparison forever.
 *
 * Storing removes it. These bytes are a pure function of our own XML, so
 * `C-REGEN` is stable for the life of the project. The decks are tens of
 * kilobytes against a 512 KiB per-file cap.
 *
 * What that costs is real and is not left to be an accident: no Tier A deck
 * then exercises DEFLATE at all. `a35-zip-shapes` is the deck that does, on
 * purpose - mixed methods, the general-purpose flag bits, and the 520-byte
 * `0xa220` growth hint PowerPoint writes - and it is the only corpus deck whose
 * bytes depend on zlib.
 */

import { growthHint, writeZip, type ZipEntry } from '../../../../ground-truth/lib/zip.ts';

// ---------------------------------------------------------------- namespaces

export const NS_P = 'http://schemas.openxmlformats.org/presentationml/2006/main';
export const NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
export const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const NS_CT = 'http://schemas.openxmlformats.org/package/2006/content-types';
const NS_REL = 'http://schemas.openxmlformats.org/package/2006/relationships';

/**
 * The base of every ECMA-376 relationship type URI.
 *
 * `REL + 'chart'` is the whole of a chart relationship's `@Type`. The
 * extension types Microsoft added later - chartEx, ink, media - are **not**
 * under this base and are written out in full by the decks that need them,
 * which keeps the difference visible rather than hidden behind a helper.
 */
export const REL = NS_R + '/';
const CT_PML = 'application/vnd.openxmlformats-officedocument.presentationml.';

/**
 * The declaration, and the CRLF after it.
 *
 * Measured, not chosen. Every one of the 38 XML parts of a PowerPoint-saved
 * deck begins with exactly these bytes - no BOM, this spelling, then CRLF and
 * nothing else on the line.
 */
export const DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';

const NS_DECLS = `xmlns:a="${NS_A}" xmlns:r="${NS_R}" xmlns:p="${NS_P}"`;

export const SLIDE_WIDTH = 12192000;
export const SLIDE_HEIGHT = 6858000;
const NOTES_WIDTH = 6858000;
const NOTES_HEIGHT = 9144000;

const encoder = new TextEncoder();

/**
 * Escape text for an XML text node.
 *
 * PowerPoint's answer to a stray `<` inside an `a:t` is to refuse the whole
 * package - "the file or directory is corrupted and unreadable", no part name.
 * Sub-phase 0.7 found that the hard way.
 *
 * `>` is escaped as well. XML does not require it in character data and
 * PowerPoint escapes it anyway, which `powerpoint-conventions.json` records;
 * matching that costs nothing and keeps our output inside the set of forms a
 * real producer emits.
 */
export function escapeXml(text: string): string {
  return text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

/**
 * Escape a value going into a double-quoted attribute.
 *
 * `escapeXml` is for character data and leaves `"` alone, which is correct
 * there and silently catastrophic in an attribute: a shape named
 * `vert="vert270"` becomes `name="vert="vert270""`, the parser stops at the
 * second quote, and the part is malformed from that byte on. `a10` and `a11`
 * were both written that way first and the census caught it - PowerPoint would
 * have refused the package with no part name at all.
 *
 * The first version of this function **threw** on a quote rather than escaping
 * it. `corpus/ground-truth/powerpoint-conventions.json` recorded the three
 * entities PowerPoint was measured writing, `conventions.test.ts` holds every
 * producer here to that set, and E8's deck happened to contain no attribute
 * with a quote in it - so its absence was not evidence, and emitting a fourth
 * entity on that basis would have been guessing.
 *
 * It is no longer a guess. On 2026-08-27 a deck was authored through PowerPoint
 * COM with a shape renamed `a "quoted" name`, and build 16.0.20326 wrote
 *
 * ```xml
 * <p:cNvPr id="3" name="a &quot;quoted&quot; name">
 * ```
 *
 * with `&quot;` and nothing else - two occurrences, and no other entity
 * anywhere in the package. So `&quot;` is a convention PowerPoint has, the
 * measurement is recorded alongside the other three, and `a19-decorative`
 * exercises it in the one attribute a user types free text into.
 */
export function escapeAttribute(value: string): string {
  return escapeXml(value).replace(/"/g, '&quot;');
}

// --------------------------------------------------- non-visual drawing props

/**
 * `p:cNvPr` - the one element every shape, group, picture, connector and frame
 * begins with, and the whole of a slide's accessibility surface.
 *
 * `@descr` is the alt text; `@title` is the separate short title the Alt Text
 * pane offers and screen readers announce first. Both are plain attributes and
 * both are easy to drop on a round trip because nothing renders them.
 *
 * `a:extLst` here is where `adec:decorative` lives - see `a19-decorative`. The
 * child order is `a:hlinkClick, a:hlinkHover, a:extLst`, so the extension list
 * is last, as always.
 */
export interface DrawingProps {
  readonly id: number;
  readonly name: string;
  /** `@descr`. Alt text. */
  readonly descr?: string;
  /** `@title`. Announced before `@descr`, and a different field in the UI. */
  readonly title?: string;
  /** `@hidden`. A shape that is in the tree and not drawn. */
  readonly hidden?: boolean;
  /**
   * `a:hlinkClick` markup, already built.
   *
   * A hyperlink is not the only thing that lives here. A media shape carries
   * `<a:hlinkClick r:id="" action="ppaction://media"/>` - a required `r:id`
   * with the empty string in it, because the `@action` is the whole content.
   */
  readonly hlinkClick?: string;
  /** `a:extLst` markup, already built. */
  readonly extLst?: string;
}

/** `p:cNvPr`. Attributes in schema order: id, name, descr, hidden, title. */
export function cNvPrXml(props: DrawingProps): string {
  const attributes =
    ` id="${String(props.id)}" name="${escapeAttribute(props.name)}"` +
    (props.descr === undefined ? '' : ` descr="${escapeAttribute(props.descr)}"`) +
    (props.hidden === true ? ' hidden="1"' : '') +
    (props.title === undefined ? '' : ` title="${escapeAttribute(props.title)}"`);
  // Child order: hlinkClick, hlinkHover, extLst.
  const children = (props.hlinkClick ?? '') + (props.extLst ?? '');
  return children === ''
    ? `<p:cNvPr${attributes}/>`
    : `<p:cNvPr${attributes}>${children}</p:cNvPr>`;
}

// ------------------------------------------------------------------ the spec

export interface ProbeRel {
  readonly id: string;
  readonly type: string;
  readonly target: string;
  readonly external?: boolean;
}

/** An extra part: media, an embedding, anything the probe's markup points at. */
export interface ProbePart {
  /** Archive-relative, e.g. `ppt/media/image1.png`. */
  readonly name: string;
  readonly bytes: Uint8Array | string;
  /**
   * How `[Content_Types].xml` should type it. A `Default` keys on the
   * extension and is what PowerPoint writes for media; an `Override` names the
   * part. Extensions already covered - `xml`, `rels` - need neither.
   */
  readonly contentType?:
    | { readonly kind: 'default'; readonly extension: string; readonly type: string }
    | { readonly kind: 'override'; readonly type: string };
}

export interface ProbeSlide {
  /** The `p:spTree` children after `p:grpSpPr`. Already-built markup. */
  readonly body: string;
  /** Shown in the title placeholder, so a human opening the deck can navigate. */
  readonly title: string;
  /** Relationships beyond `rId1`, which is always the layout. */
  readonly rels?: readonly ProbeRel[];
  /**
   * Markup between `</p:cSld>` and `</p:sld>`, after `p:clrMapOvr`.
   * `p:transition` then `p:timing`, in that order - `CT_Slide` is a sequence.
   */
  readonly tail?: string;
  /**
   * Markup inside `p:cSld`, after `p:spTree`.
   *
   * `CT_CommonSlideData` is `bg, spTree, custDataLst, controls, extLst`, so
   * this is where a slide's own `p:extLst` goes - and it is a different
   * `extLst` from `tail`'s, one level in. PowerPoint writes `p14:creationId`
   * here on every slide it creates.
   */
  readonly cSldTail?: string;
  /**
   * Index into the package's layouts, counted across every master in order.
   * With the default pair, 0 is Title Only and 1 is Blank.
   *
   * The binding lives **only** here, in `slideN.xml.rels`. A slide part does
   * not name its own layout anywhere in its own markup, which is why changing
   * layout is a relationship rewrite rather than an edit to the slide - and
   * with several masters it is also the only thing that says which master a
   * slide inherits from, by way of the layout it points at.
   */
  readonly layout?: number;
  /**
   * `p:clrMapOvr`. Defaults to `<a:masterClrMapping/>`, which means "use the
   * master's".
   *
   * The other branch is `a:overrideClrMapping`, which carries all twelve
   * attributes again and replaces the master's map for this slide alone. A
   * renderer that reads `p:clrMap` off the master and stops has no way to
   * notice, and every `a:schemeClr val="bg1"` on the slide resolves to the
   * wrong colour.
   */
  readonly clrMapOvr?: string;
  /**
   * Extra attributes on `<p:sld>` itself, namespace declarations included.
   *
   * `mc:Ignorable` is the one that matters and the one that can only go here:
   * it names **prefixes**, so it has to sit on an element where those prefixes
   * are in scope, and PowerPoint writes it on the part root of every slide
   * carrying 2010-or-later markup. See `a37-mce`.
   */
  readonly rootAttributes?: string;
  /**
   * The notes slide's `p:spTree` children, after `p:grpSpPr`.
   *
   * Present means the package grows a `ppt/notesSlides/notesSlideN.xml` bound
   * to this slide and to the notes master, so `notesMaster` has to be declared
   * as well.
   */
  readonly notes?: string;
}

/**
 * One slide layout.
 *
 * The chassis ships two - Title Only and Blank - and every deck up to
 * `a06-lines` uses them unchanged. A deck that declares `layouts` replaces the
 * pair outright rather than extending it, so what a placeholder deck inherits
 * from is entirely its own and nothing arrives from scaffolding it did not ask
 * for.
 */
export interface ProbeLayout {
  /** `ST_SlideLayoutType` for `@type`. */
  readonly type: string;
  /** `p:cSld/@name`, which is what a layout picker shows a user. */
  readonly name: string;
  /** `p:spTree` children after `p:grpSpPr` - the layout's own placeholders. */
  readonly shapes?: string;
  /**
   * Whether `shapes` includes a title placeholder.
   *
   * Declared rather than sniffed out of the markup: a slide bound to this
   * layout gets its `ProbeSlide.title` written into a title placeholder only
   * when there is one here to inherit from, and a title `p:ph` on a slide whose
   * layout has none is exactly the orphan case `a02` is the probe for. Making
   * that a deliberate choice keeps it out of the chassis's hands.
   */
  readonly hasTitle?: boolean;
  /**
   * `p:transition`.
   *
   * A transition is not a slide-only thing: `CT_SlideLayout` and
   * `CT_SlideMaster` both hold one, and a slide bound to this layout inherits
   * it. That is what lets `a16-transitions` carry all twenty-one variants in a
   * three-slide deck.
   */
  readonly transition?: string;
  /** `p:hf`. `CT_SlideLayout` puts it after `p:timing`, before `p:extLst`. */
  readonly hf?: string;
}

/**
 * One slide master, with the layouts that belong to it.
 *
 * A package may hold several, and `a12-masters` is the deck that says so. Two
 * facts make that more than a count.
 *
 * **Each master carries its own theme part.** `ppt/theme/themeN.xml` is reached
 * from `slideMasterN.xml.rels`, never from the presentation, so two masters can
 * hold two entirely different colour schemes and a shape that reads
 * `a:schemeClr val="accent1"` resolves differently on either side of the deck.
 * A renderer with one global theme renders half of a multi-master deck in the
 * wrong colours, and nothing in the markup looks wrong.
 *
 * **Each master carries its own `p:clrMap`.** The map is what turns `bg1` into
 * `lt1` or into `dk1`, so an inverted master is one attribute per slot and not
 * a different theme at all.
 */
export interface ProbeMaster {
  /** Replaces the fixed title-and-body pair on this master. */
  readonly placeholders?: string;
  /** Extra `p:spTree` children, after the placeholders. */
  readonly shapes?: string;
  /** `p:transition`. `CT_SlideMaster`: after `p:sldLayoutIdLst`, before `p:hf`. */
  readonly transition?: string;
  /** `p:hf`. `CT_SlideMaster`: after `p:timing`, before `p:txStyles`. */
  readonly hf?: string;
  /** Replaces this master's whole `p:txStyles` block. */
  readonly textStyles?: string;
  /** The twelve `p:clrMap` attributes, written out. Defaults to the identity map. */
  readonly clrMap?: string;
  /** This master's layouts. Numbered across the package in master order. */
  readonly layouts: readonly ProbeLayout[];
  /** `a:clrScheme` children for this master's theme. Defaults to the shared scheme. */
  readonly colours?: string;
  /** `a:theme/@name` and the scheme names inside it. */
  readonly themeName?: string;
  /**
   * `a:majorFont/a:latin/@typeface` on this master's theme, which is what
   * `+mj-lt` resolves to. Defaults to Calibri Light.
   */
  readonly majorLatin?: string;
  /** `a:minorFont/a:latin/@typeface` - what `+mn-lt` resolves to. */
  readonly minorLatin?: string;
  /** `a:objectDefaults` on this master's theme. */
  readonly objectDefaults?: string;
}

/**
 * The notes master.
 *
 * Its `p:spTree` is not a slide master's. PowerPoint writes six placeholders -
 * `hdr`, `dt`, `sldImg`, `body`, `ftr`, `sldNum` - and two of those are the
 * pair a slide layout is refused for carrying. `CT_NotesMaster` is
 * `cSld, clrMap, hf, notesStyle, extLst`, which is a different sequence from
 * every other sheet: there is no layout list and no `p:txStyles`, and the one
 * text-style block it does have is `p:notesStyle`.
 */
export interface ProbeSheetMaster {
  readonly shapes?: string;
  readonly hf?: string;
  readonly clrMap?: string;
  /** `p:notesStyle`, on the notes master only. */
  readonly notesStyle?: string;
}

export interface ProbePackage {
  readonly title: string;
  readonly slides: readonly ProbeSlide[];
  readonly parts?: readonly ProbePart[];
  /**
   * Several masters, each with its own layouts, theme and colour map.
   *
   * Mutually exclusive with the flat `layouts` / `masterPlaceholders` /
   * `masterShapes` / `masterHf` / `textStyles` / `themeObjectDefaults` fields,
   * which are the one-master spelling of the same thing and are what every
   * deck through `a11-autofit` uses.
   */
  readonly masters?: readonly ProbeMaster[];
  /**
   * The notes master, and with it `p:notesMasterIdLst` and a theme of its own.
   *
   * Required by any slide that declares `notes`: a notes slide relates to the
   * notes master, and PowerPoint refuses a package where that edge dangles.
   */
  readonly notesMaster?: ProbeSheetMaster;
  /**
   * The handout master. PowerPoint writes one only when a user has opened the
   * handout master view and changed something, so most real decks have none.
   */
  readonly handoutMaster?: ProbeSheetMaster;
  /** Replaces the default Title Only / Blank pair outright. */
  readonly layouts?: readonly ProbeLayout[];
  /**
   * Replaces the master's two fixed placeholders - title, and body at idx 1.
   *
   * A deck needs this only when the master tier of the text cascade is what it
   * is probing, because the fixed pair carry no `a:lstStyle` and a source with
   * nothing in it cannot be told apart from a source that was never consulted.
   */
  readonly masterPlaceholders?: string;
  /** Extra `p:spTree` children on the master, after its placeholders. */
  readonly masterShapes?: string;
  /**
   * `a:objectDefaults` on the theme - `a:spDef`, `a:lnDef`, `a:txDef`.
   *
   * `CT_OfficeStyleSheet` is themeElements, objectDefaults, extraClrSchemeLst,
   * custClrLst, extLst, so this goes immediately after the theme elements. It
   * is the eighth of the ten sources in the text cascade and the only one that
   * lives in the theme rather than in a sheet.
   */
  readonly themeObjectDefaults?: string;
  /** `p:hf` on the master. `CT_SlideMaster`: after `p:timing`, before `p:txStyles`. */
  readonly masterHf?: string;
  /**
   * Replaces the master's whole `p:txStyles` block.
   *
   * Two of the ten sources in the text cascade live in here - the bucket styles
   * a placeholder inherits from - so a deck probing the cascade has to be able
   * to say what they are rather than inherit the chassis's opinion.
   */
  readonly textStyles?: string;
  /** Relationships on `ppt/presentation.xml` beyond the ones written here. */
  readonly presentationRels?: readonly ProbeRel[];
  /** Markup for `p:presentation`'s tail, after `p:notesSz`. Sequence order. */
  readonly presentationTail?: string;
  /**
   * Extra attributes on `p:presentation` itself.
   *
   * `@embedTrueTypeFonts` is the one that matters and the one that is easy to
   * lose: omit it and PowerPoint ignores a perfectly formed `p:embeddedFontLst`
   * entirely, with no diagnostic of any kind. See `a31-embedded-fonts`.
   */
  readonly presentationAttributes?: string;
  readonly slideWidth?: number;
  readonly slideHeight?: number;
  /**
   * `p:sldSz/@type`, an `ST_SlideSizeType`.
   *
   * A hint about the intended delivery medium, not a size: `@cx` and `@cy` are
   * the truth and the two can disagree. Defaults to `screen16x9` when the
   * extent is the default one and to nothing at all otherwise, which is what
   * PowerPoint does for a size it has no name for.
   */
  readonly slideSizeType?: string;
  /**
   * `p:notesSz`, which is **not** derived from `p:sldSz`.
   *
   * It is required where `p:sldSz` is optional, and it is the size of the notes
   * and handout sheet, so a portrait notes page under a landscape deck is the
   * normal case rather than an oddity. Defaults to A4-ish portrait.
   */
  readonly notesWidth?: number;
  readonly notesHeight?: number;
  /**
   * Deflate instead of storing. Only `a35-zip-shapes` sets this, and only
   * because it is the deck about compression. See the file comment.
   */
  readonly deflate?: boolean;
  /**
   * Per-entry archive shape, keyed by the entry name exactly as written.
   *
   * Only `a35-zip-shapes` uses it, and only for entries it can name: the
   * chassis decides most part names, so this is a lookup rather than a
   * callback, which keeps what the deck is claiming visible in the deck.
   *
   * A key that matches no entry is an error rather than a no-op. A deck that
   * says it puts a growth hint on `[Content_Types].xml` and silently does not
   * is worse than one that fails to build.
   */
  readonly zip?: Readonly<Record<string, ZipShape>>;
  /**
   * Relationships on `_rels/.rels`, beyond the three written here.
   *
   * The package-relationship part, not a part's own. The OPC thumbnail lives
   * here and nowhere else - `docProps/thumbnail.jpeg` is reached by a
   * `metadata/thumbnail` relationship from the package root, which is why it is
   * the one part in a PowerPoint deck that no part references. See
   * `a33-thumbnail`.
   */
  readonly packageRels?: readonly ProbeRel[];
  /**
   * Type this package as macro-enabled: `p:presentation`'s content type becomes
   * the `ms-powerpoint.presentation.macroEnabled` one.
   *
   * The content type is the whole of it. There is no attribute, no element and
   * no flag anywhere in the markup that says a deck holds macros - only the
   * content type of the main part, and the presence of a `vbaProject.bin`. See
   * `a32-macros`, which is also the deck that has to be written `.pptm`,
   * because PowerPoint checks the file extension against the content type and
   * refuses the pair when they disagree.
   */
  readonly macroEnabled?: boolean;
  /**
   * `p:sldId/@id` for each slide, replacing the default 256, 257, 258, ...
   *
   * `ST_SlideId` is 256 to 2147483647 - not 0-based, not unbounded, and a
   * different space from `ST_SlideMasterId`. `a39-large-ids` is the deck that
   * sits at both ends of it.
   */
  readonly slideIds?: readonly number[];
  /**
   * The first `p:sldMasterId/@id`, from which layouts are numbered upward.
   *
   * Defaults to 2147483648, the bottom of `ST_SlideMasterId`. One counter
   * serves masters and layouts together - see `slideMasterXml` for why that is
   * load-bearing rather than tidy - so raising this raises every layout id with
   * it. `a39-large-ids` sets it near the top of the range.
   */
  readonly firstSheetId?: number;
}

/** How one ZIP entry is written. See `ProbePackage.zip`. */
export interface ZipShape {
  /** Store rather than deflate, overriding `ProbePackage.deflate`. */
  readonly store?: boolean;
  /** The general-purpose bit flag. PowerPoint writes `0x0006` on every entry. */
  readonly flags?: number;
  /** Total bytes of the `0xa220` growth hint in the local header. */
  readonly growthHint?: number;
}

// --------------------------------------------------------------- fixed parts

/**
 * A `.rels` part.
 *
 * Exported because from `a20-tables` on the probes bring parts of their own -
 * a chart with its colour and style parts, a diagram with four, an OLE object
 * with an embedding - and every one of those parts has relationships that
 * resolve **against its own folder**, not the slide's. A deck declares those as
 * ordinary `parts` entries whose bytes this builds, so the rId numbering stays
 * where a reviewer can see it.
 */
export function relsXml(entries: readonly ProbeRel[]): string {
  return (
    DECLARATION +
    `<Relationships xmlns="${NS_REL}">` +
    entries
      .map(
        (entry) =>
          `<Relationship Id="${entry.id}" Type="${entry.type}" Target="${entry.target}"` +
          (entry.external === true ? ' TargetMode="External"' : '') +
          '/>',
      )
      .join('') +
    '</Relationships>'
  );
}

/** An empty `p:spTree` prologue: `nvGrpSpPr` then `grpSpPr`, always, in that order. */
export function spTreeHead(): string {
  return (
    '<p:spTree>' +
    '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
    '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>' +
    '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
  );
}

/**
 * The theme.
 *
 * Its `effectStyleLst` is deliberately **empty**, unlike the benchmark deck's,
 * which carries an `a:outerShdw` nobody asked for. A probe deck's census should
 * report the feature the probe is about and not the scaffolding's opinions, so
 * the only `a:outerShdw` in `a04-effects` is the one `a04-effects` writes.
 *
 * The two gradient entries in `fillStyleLst` cannot go the same way: `fillRef`
 * and `bgRef` index into these lists 1-based, and a list shorter than three is
 * an out-of-range reference on the first slide that uses it. So every probe
 * deck reports `gradientFill` at least 2, and every `features` map says so.
 */
function themeXml(options: {
  readonly name?: string;
  readonly colours?: string;
  readonly objectDefaults?: string;
  readonly majorLatin?: string;
  readonly minorLatin?: string;
}): string {
  const name = options.name ?? 'PPTX Studio Corpus';
  const objectDefaults = options.objectDefaults ?? '';
  const scheme = [
    '<a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>',
    '<a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>',
    '<a:dk2><a:srgbClr val="44546A"/></a:dk2>',
    '<a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>',
    '<a:accent1><a:srgbClr val="4472C4"/></a:accent1>',
    '<a:accent2><a:srgbClr val="ED7D31"/></a:accent2>',
    '<a:accent3><a:srgbClr val="A5A5A5"/></a:accent3>',
    '<a:accent4><a:srgbClr val="FFC000"/></a:accent4>',
    '<a:accent5><a:srgbClr val="5B9BD5"/></a:accent5>',
    '<a:accent6><a:srgbClr val="70AD47"/></a:accent6>',
    '<a:hlink><a:srgbClr val="0563C1"/></a:hlink>',
    '<a:folHlink><a:srgbClr val="954F72"/></a:folHlink>',
  ].join('');

  const fill = (i: number): string =>
    i === 0
      ? '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
      : '<a:gradFill rotWithShape="1"><a:gsLst>' +
        `<a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="${String(60000 - i * 10000)}"/></a:schemeClr></a:gs>` +
        `<a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="${String(100000 - i * 20000)}"/></a:schemeClr></a:gs>` +
        '</a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill>';

  const line = (w: number): string =>
    `<a:ln w="${String(w)}" cap="flat" cmpd="sng" algn="ctr">` +
    '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln>';

  return (
    DECLARATION +
    `<a:theme xmlns:a="${NS_A}" name="${escapeAttribute(name)}">` +
    '<a:themeElements>' +
    `<a:clrScheme name="${escapeAttribute(name)}">${options.colours ?? scheme}</a:clrScheme>` +
    `<a:fontScheme name="${escapeAttribute(name)}">` +
    `<a:majorFont><a:latin typeface="${escapeAttribute(options.majorLatin ?? 'Calibri Light')}"/>` +
    '<a:ea typeface=""/><a:cs typeface=""/></a:majorFont>' +
    `<a:minorFont><a:latin typeface="${escapeAttribute(options.minorLatin ?? 'Calibri')}"/>` +
    '<a:ea typeface=""/><a:cs typeface=""/></a:minorFont>' +
    '</a:fontScheme>' +
    `<a:fmtScheme name="${escapeAttribute(name)}">` +
    `<a:fillStyleLst>${fill(0)}${fill(1)}${fill(2)}</a:fillStyleLst>` +
    `<a:lnStyleLst>${line(6350)}${line(12700)}${line(19050)}</a:lnStyleLst>` +
    '<a:effectStyleLst>' +
    '<a:effectStyle><a:effectLst/></a:effectStyle>' +
    '<a:effectStyle><a:effectLst/></a:effectStyle>' +
    '<a:effectStyle><a:effectLst/></a:effectStyle>' +
    '</a:effectStyleLst>' +
    '<a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill>' +
    '<a:solidFill><a:schemeClr val="phClr"><a:tint val="95000"/></a:schemeClr></a:solidFill>' +
    '<a:solidFill><a:schemeClr val="phClr"><a:shade val="90000"/></a:schemeClr></a:solidFill>' +
    '</a:bgFillStyleLst>' +
    '</a:fmtScheme>' +
    '</a:themeElements>' +
    objectDefaults +
    '</a:theme>'
  );
}

const EMPTY_PARA = '<a:p><a:endParaRPr lang="en-GB"/></a:p>';

export interface PlaceholderSpec extends DrawingProps {
  /** `ST_PlaceholderType`. Omitted is not the same as `obj`: see the doc below. */
  readonly type?: string;
  readonly idx?: number;
  /** `ST_Direction`: `horz` or `vert`. */
  readonly orient?: string;
  /** `ST_PlaceholderSize`: `full`, `half` or `quarter`. */
  readonly size?: string;
  readonly hasCustomPrompt?: boolean;
  /** Omitted means no `a:xfrm` at all, which is the fact inheritance runs on. */
  readonly x?: number;
  readonly y?: number;
  readonly cx?: number;
  readonly cy?: number;
  /** `a:p` children of the `p:txBody`. Defaults to one empty paragraph. */
  readonly body?: string;
  /** `a:bodyPr` markup, when the default empty one will not do. */
  readonly bodyPr?: string;
  /** `a:lstStyle` markup - the shape-level tier of the text cascade. */
  readonly lstStyle?: string;
  /** `a:spLocks` attributes, when the default `noGrp="1"` is not enough. */
  readonly locks?: string;
  /** No `p:txBody` at all - which is what a notes slide's `sldImg` has. */
  readonly noTextBody?: boolean;
}

/**
 * A placeholder shape.
 *
 * No `a:prstGeom`, so the census's `presetGeom` count stays the probe's own.
 *
 * Two absences here are load-bearing rather than shorthand. A `p:ph` with no
 * `@type` is **not** untyped - it means `type="obj" idx="0"`, and the matcher
 * has to materialize that before comparing or half the built-in layouts stop
 * matching. And omitting `x`/`y`/`cx`/`cy` writes no `a:xfrm` at all, which is
 * the fact that makes Change Layout work: a shape that never had explicit
 * geometry can take the new layout's, and one that has an `a:xfrm` cannot.
 */
export function placeholderXml(spec: PlaceholderSpec): string {
  const attributes =
    (spec.type === undefined ? '' : ` type="${spec.type}"`) +
    (spec.orient === undefined ? '' : ` orient="${spec.orient}"`) +
    (spec.size === undefined ? '' : ` sz="${spec.size}"`) +
    (spec.idx === undefined ? '' : ` idx="${String(spec.idx)}"`) +
    (spec.hasCustomPrompt === true ? ' hasCustomPrompt="1"' : '');
  const xfrm =
    spec.x === undefined
      ? ''
      : `<a:xfrm><a:off x="${String(spec.x)}" y="${String(spec.y ?? 0)}"/>` +
        `<a:ext cx="${String(spec.cx ?? 0)}" cy="${String(spec.cy ?? 0)}"/></a:xfrm>`;
  return (
    '<p:sp><p:nvSpPr>' +
    cNvPrXml(spec) +
    `<p:cNvSpPr><a:spLocks ${spec.locks ?? 'noGrp="1"'}/></p:cNvSpPr>` +
    `<p:nvPr><p:ph${attributes}/></p:nvPr>` +
    '</p:nvSpPr>' +
    `<p:spPr>${xfrm}</p:spPr>` +
    (spec.noTextBody === true
      ? ''
      : '<p:txBody>' +
        (spec.bodyPr ?? '<a:bodyPr/>') +
        (spec.lstStyle ?? '<a:lstStyle/>') +
        (spec.body ?? EMPTY_PARA) +
        '</p:txBody>') +
    '</p:sp>'
  );
}

function textStyleXml(kind: 'title' | 'body' | 'other'): string {
  const level = (n: number, size: number): string =>
    `<a:lvl${String(n)}pPr marL="${String(n === 1 ? 0 : 342900 * (n - 1))}" indent="${String(n === 1 ? 0 : -342900)}">` +
    `<a:defRPr sz="${String(size)}" kern="1200">` +
    '<a:solidFill><a:schemeClr val="tx1"/></a:solidFill>' +
    `<a:latin typeface="+m${kind === 'title' ? 'j' : 'n'}-lt"/></a:defRPr></a:lvl${String(n)}pPr>`;
  const sizes = kind === 'title' ? [2400] : [1800, 1600, 1400, 1200, 1200];
  return (
    `<p:${kind}Style>` +
    sizes.map((size, index) => level(index + 1, size)).join('') +
    `</p:${kind}Style>`
  );
}

/** The title placeholder's geometry, shared by the master and the Title Only layout. */
export const TITLE_BOX = { x: 457200, y: 274638, cx: 11277600, cy: 700088 } as const;

/**
 * The default pair: Title Only for the probes, Blank for anything that wants
 * nothing. Every deck through `a06-lines` uses exactly this.
 */
const DEFAULT_LAYOUTS: readonly ProbeLayout[] = [
  {
    type: 'titleOnly',
    name: 'Title Only',
    hasTitle: true,
    shapes: placeholderXml({ id: 2, name: 'Title 1', type: 'title', ...TITLE_BOX }),
  },
  { type: 'blank', name: 'Blank' },
];

/** The identity map: all twelve slots, each pointing at the obvious one. */
export const IDENTITY_CLR_MAP =
  'bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2"' +
  ' accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6"' +
  ' hlink="hlink" folHlink="folHlink"';

/**
 * One master.
 *
 * `layoutIds` are this master's `p:sldLayoutId/@id` values, allocated by the
 * caller from a **single package-wide counter shared with the masters**. That
 * sharing is not decoration. `ST_SlideMasterId` and `ST_SlideLayoutId` are the
 * same range - 2147483648 and up, disjoint from `p:sldId`'s 256…2147483647 -
 * and PowerPoint allocates from one sequence: master, its layouts, next master,
 * its layouts. A first attempt here numbered masters and layouts from two
 * counters, which is right for one master and collides on the second, and
 * **any** package with two masters was then a whole-package refusal. See
 * `ROSTER.md`'s "What PowerPoint refuses".
 *
 * The part numbering is separate again: layouts are `slideLayout1.xml` upward
 * across the package, while each master's own `.rels` numbers them `rId1`
 * upward, because rIds are `xsd:ID` scoped to one `.rels` part.
 */
function slideMasterXml(master: ProbeMaster, layoutIds: readonly number[]): string {
  return (
    DECLARATION +
    `<p:sldMaster ${NS_DECLS}>` +
    '<p:cSld><p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg>' +
    spTreeHead() +
    (master.placeholders ??
      placeholderXml({ id: 2, name: 'Title Placeholder 1', type: 'title', ...TITLE_BOX }) +
        placeholderXml({
          id: 3,
          name: 'Text Placeholder 2',
          type: 'body',
          idx: 1,
          x: 457200,
          y: 1097280,
          cx: 11277600,
          cy: 5303520,
        })) +
    (master.shapes ?? '') +
    '</p:spTree></p:cSld>' +
    // All twelve. Eleven is a repair prompt.
    `<p:clrMap ${master.clrMap ?? IDENTITY_CLR_MAP}/>` +
    '<p:sldLayoutIdLst>' +
    master.layouts
      .map(
        (_layout, index) =>
          `<p:sldLayoutId id="${String(layoutIds[index] ?? 0)}"` +
          ` r:id="rId${String(index + 1)}"/>`,
      )
      .join('') +
    '</p:sldLayoutIdLst>' +
    // `CT_SlideMaster` is cSld, clrMap, sldLayoutIdLst, transition, timing, hf,
    // txStyles, extLst - so the transition and `p:hf` sit here, between the
    // layout list and the text styles, in that order and nowhere else.
    (master.transition ?? '') +
    (master.hf ?? '') +
    (master.textStyles ??
      `<p:txStyles>${textStyleXml('title')}${textStyleXml('body')}${textStyleXml('other')}</p:txStyles>`) +
    '</p:sldMaster>'
  );
}

/**
 * `p:notesMaster`. Sequence: `cSld, clrMap, hf, notesStyle, extLst`.
 *
 * No layout list, no `p:txStyles`. The default placeholders are the six
 * PowerPoint 16.0.20326 writes, at the geometry it writes them at, measured
 * from a deck it saved rather than invented: `hdr` with no `@idx` at all, `dt`
 * at 1, `sldImg` at 2, `body` at 3, `ftr` at 4, `sldNum` at 5.
 */
function notesMasterXml(master: ProbeSheetMaster): string {
  return (
    DECLARATION +
    `<p:notesMaster ${NS_DECLS}>` +
    '<p:cSld><p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg>' +
    spTreeHead() +
    (master.shapes ?? '') +
    '</p:spTree></p:cSld>' +
    `<p:clrMap ${master.clrMap ?? IDENTITY_CLR_MAP}/>` +
    (master.hf ?? '') +
    (master.notesStyle ??
      '<p:notesStyle><a:lvl1pPr marL="0" algn="l" defTabSz="914400" rtl="0"' +
        ' eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1200" kern="1200">' +
        '<a:solidFill><a:schemeClr val="tx1"/></a:solidFill>' +
        '<a:latin typeface="+mn-lt"/></a:defRPr></a:lvl1pPr></p:notesStyle>') +
    '</p:notesMaster>'
  );
}

/** `p:handoutMaster`. Sequence: `cSld, clrMap, hf, extLst` - no text styles at all. */
function handoutMasterXml(master: ProbeSheetMaster): string {
  return (
    DECLARATION +
    `<p:handoutMaster ${NS_DECLS}>` +
    '<p:cSld><p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg>' +
    spTreeHead() +
    (master.shapes ?? '') +
    '</p:spTree></p:cSld>' +
    `<p:clrMap ${master.clrMap ?? IDENTITY_CLR_MAP}/>` +
    (master.hf ?? '') +
    '</p:handoutMaster>'
  );
}

/** `p:notes`. Sequence: `cSld, clrMapOvr, extLst`, and no `p:transition`. */
function notesSlideXml(body: string): string {
  return (
    DECLARATION +
    `<p:notes ${NS_DECLS}>` +
    '<p:cSld>' +
    spTreeHead() +
    body +
    '</p:spTree></p:cSld>' +
    '<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>' +
    '</p:notes>'
  );
}

function slideLayoutXml(layout: ProbeLayout): string {
  return (
    DECLARATION +
    `<p:sldLayout ${NS_DECLS} type="${layout.type}" preserve="1">` +
    `<p:cSld name="${escapeAttribute(layout.name)}">` +
    spTreeHead() +
    (layout.shapes ?? '') +
    '</p:spTree></p:cSld>' +
    '<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>' +
    (layout.transition ?? '') +
    (layout.hf ?? '') +
    '</p:sldLayout>'
  );
}

// ------------------------------------------------------------------- writing

export interface BuiltPackage {
  readonly bytes: Uint8Array;
  readonly entries: number;
}

export function buildProbePackage(spec: ProbePackage): BuiltPackage {
  const slideWidth = spec.slideWidth ?? SLIDE_WIDTH;
  const slideHeight = spec.slideHeight ?? SLIDE_HEIGHT;
  const notesWidth = spec.notesWidth ?? NOTES_WIDTH;
  const notesHeight = spec.notesHeight ?? NOTES_HEIGHT;
  const extraParts = spec.parts ?? [];

  // The flat one-master fields and `masters` are two spellings of the same
  // thing, and every deck through `a11-autofit` uses the flat one. Folding it
  // into a one-element list here rather than branching all the way down is
  // what keeps those decks byte-identical through this change.
  const masters: readonly ProbeMaster[] = spec.masters ?? [
    {
      ...(spec.masterPlaceholders === undefined ? {} : { placeholders: spec.masterPlaceholders }),
      ...(spec.masterShapes === undefined ? {} : { shapes: spec.masterShapes }),
      ...(spec.masterHf === undefined ? {} : { hf: spec.masterHf }),
      ...(spec.textStyles === undefined ? {} : { textStyles: spec.textStyles }),
      ...(spec.themeObjectDefaults === undefined
        ? {}
        : { objectDefaults: spec.themeObjectDefaults }),
      layouts: spec.layouts ?? DEFAULT_LAYOUTS,
    },
  ];

  // Layouts are numbered across the whole package, not per master, because the
  // part names are: three masters with two layouts each write slideLayout1 to
  // slideLayout6. `ProbeSlide.layout` indexes into this.
  const layouts = masters.flatMap((master, masterIndex) =>
    master.layouts.map((layout) => ({ layout, masterNumber: masterIndex + 1 })),
  );

  // Themes are numbered after the slide masters: one per master, then the
  // notes master's, then the handout master's - which is the order PowerPoint
  // writes them in and the reason a real deck with notes has a `theme2.xml`
  // nobody's slides ever read.
  const notesThemeNumber = spec.notesMaster === undefined ? null : masters.length + 1;
  const handoutThemeNumber =
    spec.handoutMaster === undefined ? null : masters.length + (notesThemeNumber === null ? 1 : 2);
  const themeCount =
    masters.length + (notesThemeNumber === null ? 0 : 1) + (handoutThemeNumber === null ? 0 : 1);

  // One counter for masters and layouts together. See `slideMasterXml`.
  const masterIds: number[] = [];
  const layoutIdsByMaster: number[][] = [];
  let sheetId = spec.firstSheetId ?? 2147483648;
  for (const master of masters) {
    masterIds.push(sheetId);
    sheetId += 1;
    const ids: number[] = [];
    for (let index = 0; index < master.layouts.length; index++) {
      ids.push(sheetId);
      sheetId += 1;
    }
    layoutIdsByMaster.push(ids);
  }

  const notedSlides = spec.slides.filter((slide) => slide.notes !== undefined);
  if (notedSlides.length > 0 && spec.notesMaster === undefined) {
    throw new Error('a slide declares notes but the package declares no notesMaster');
  }

  // --- content types --------------------------------------------------------
  // Defaults are de-duplicated by extension: two PNG parts is one Default, and
  // a second `<Default Extension="png"/>` is a package PowerPoint refuses.
  const defaults = new Map<string, string>([
    ['rels', 'application/vnd.openxmlformats-package.relationships+xml'],
    ['xml', 'application/xml'],
  ]);
  const overrides: string[] = [];
  const override = (partName: string, contentType: string): void => {
    overrides.push(`<Override PartName="${partName}" ContentType="${contentType}"/>`);
  };

  // The only thing in the whole package that says a deck holds macros.
  override(
    '/ppt/presentation.xml',
    spec.macroEnabled === true
      ? 'application/vnd.ms-powerpoint.presentation.macroEnabled.main+xml'
      : CT_PML + 'presentation.main+xml',
  );
  override('/ppt/presProps.xml', CT_PML + 'presProps+xml');
  override('/ppt/viewProps.xml', CT_PML + 'viewProps+xml');
  override('/ppt/tableStyles.xml', CT_PML + 'tableStyles+xml');
  for (let i = 0; i < themeCount; i++) {
    override(
      `/ppt/theme/theme${String(i + 1)}.xml`,
      'application/vnd.openxmlformats-officedocument.theme+xml',
    );
  }
  for (let i = 0; i < masters.length; i++) {
    override(`/ppt/slideMasters/slideMaster${String(i + 1)}.xml`, CT_PML + 'slideMaster+xml');
  }
  if (spec.notesMaster !== undefined) {
    override('/ppt/notesMasters/notesMaster1.xml', CT_PML + 'notesMaster+xml');
  }
  if (spec.handoutMaster !== undefined) {
    override('/ppt/handoutMasters/handoutMaster1.xml', CT_PML + 'handoutMaster+xml');
  }
  for (let i = 0; i < layouts.length; i++) {
    override(`/ppt/slideLayouts/slideLayout${String(i + 1)}.xml`, CT_PML + 'slideLayout+xml');
  }
  for (let i = 0; i < spec.slides.length; i++) {
    override(`/ppt/slides/slide${String(i + 1)}.xml`, CT_PML + 'slide+xml');
  }
  for (let i = 0; i < notedSlides.length; i++) {
    override(`/ppt/notesSlides/notesSlide${String(i + 1)}.xml`, CT_PML + 'notesSlide+xml');
  }
  for (const part of extraParts) {
    const ct = part.contentType;
    if (ct === undefined) continue;
    if (ct.kind === 'default') defaults.set(ct.extension, ct.type);
    else override('/' + part.name, ct.type);
  }
  override('/docProps/core.xml', 'application/vnd.openxmlformats-package.core-properties+xml');
  override(
    '/docProps/app.xml',
    'application/vnd.openxmlformats-officedocument.extended-properties+xml',
  );

  const contentTypes =
    DECLARATION +
    `<Types xmlns="${NS_CT}">` +
    [...defaults]
      .map(([extension, type]) => `<Default Extension="${extension}" ContentType="${type}"/>`)
      .join('') +
    overrides.join('') +
    '</Types>';

  // --- presentation.xml -----------------------------------------------------
  const presRels: ProbeRel[] = [];
  let rid = 0;
  const nextRid = (): string => 'rId' + String(++rid);

  const masterRids = masters.map((_master, index) => {
    const id = nextRid();
    presRels.push({
      id,
      type: REL + 'slideMaster',
      target: `slideMasters/slideMaster${String(index + 1)}.xml`,
    });
    return id;
  });
  let notesMasterRid: string | null = null;
  if (spec.notesMaster !== undefined) {
    notesMasterRid = nextRid();
    presRels.push({
      id: notesMasterRid,
      type: REL + 'notesMaster',
      target: 'notesMasters/notesMaster1.xml',
    });
  }
  let handoutMasterRid: string | null = null;
  if (spec.handoutMaster !== undefined) {
    handoutMasterRid = nextRid();
    presRels.push({
      id: handoutMasterRid,
      type: REL + 'handoutMaster',
      target: 'handoutMasters/handoutMaster1.xml',
    });
  }
  const slideRids = spec.slides.map((_, i) => {
    const id = nextRid();
    presRels.push({ id, type: REL + 'slide', target: `slides/slide${String(i + 1)}.xml` });
    return id;
  });
  presRels.push({ id: nextRid(), type: REL + 'presProps', target: 'presProps.xml' });
  presRels.push({ id: nextRid(), type: REL + 'viewProps', target: 'viewProps.xml' });
  presRels.push({ id: nextRid(), type: REL + 'theme', target: 'theme/theme1.xml' });
  presRels.push({ id: nextRid(), type: REL + 'tableStyles', target: 'tableStyles.xml' });
  presRels.push(...(spec.presentationRels ?? []));

  const sizeType =
    spec.slideSizeType ??
    (slideWidth === SLIDE_WIDTH && slideHeight === SLIDE_HEIGHT ? 'screen16x9' : undefined);

  const presentation =
    DECLARATION +
    `<p:presentation ${NS_DECLS}${spec.presentationAttributes ?? ''}>` +
    '<p:sldMasterIdLst>' +
    masterRids
      .map((id, index) => `<p:sldMasterId id="${String(masterIds[index] ?? 0)}" r:id="${id}"/>`)
      .join('') +
    '</p:sldMasterIdLst>' +
    // `CT_Presentation` is a sequence: sldMasterIdLst, notesMasterIdLst,
    // handoutMasterIdLst, sldIdLst, sldSz, notesSz, … - so both of these sit
    // before the slide list, which is the one place a tail option cannot put
    // them. `p:notesMasterId` has no `@id`; only the slide masters do.
    (notesMasterRid === null
      ? ''
      : `<p:notesMasterIdLst><p:notesMasterId r:id="${notesMasterRid}"/></p:notesMasterIdLst>`) +
    (handoutMasterRid === null
      ? ''
      : '<p:handoutMasterIdLst>' +
        `<p:handoutMasterId r:id="${handoutMasterRid}"/>` +
        '</p:handoutMasterIdLst>') +
    '<p:sldIdLst>' +
    slideRids
      .map((id, i) => `<p:sldId id="${String(spec.slideIds?.[i] ?? 256 + i)}" r:id="${id}"/>`)
      .join('') +
    '</p:sldIdLst>' +
    `<p:sldSz cx="${String(slideWidth)}" cy="${String(slideHeight)}"` +
    (sizeType === undefined ? '' : ` type="${sizeType}"`) +
    '/>' +
    `<p:notesSz cx="${String(notesWidth)}" cy="${String(notesHeight)}"/>` +
    (spec.presentationTail ?? '') +
    '</p:presentation>';

  // --- docProps -------------------------------------------------------------
  // No author, no company, no machine name. Sub-phase 1.1's privacy rule
  // applies to decks a real application authored; holding our own generator to
  // it as well means the rule is one sentence rather than an exception list.
  const core =
    DECLARATION +
    '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"' +
    ' xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/"' +
    ' xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">' +
    `<dc:title>${escapeXml(spec.title)}</dc:title>` +
    '<dc:creator>PPTX Studio corpus generator</dc:creator>' +
    '<cp:lastModifiedBy>PPTX Studio corpus generator</cp:lastModifiedBy>' +
    '<dcterms:created xsi:type="dcterms:W3CDTF">2026-08-27T00:00:00Z</dcterms:created>' +
    '<dcterms:modified xsi:type="dcterms:W3CDTF">2026-08-27T00:00:00Z</dcterms:modified>' +
    '</cp:coreProperties>';

  const app =
    DECLARATION +
    '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"' +
    ' xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">' +
    '<Application>PPTX Studio corpus generator</Application>' +
    `<Slides>${String(spec.slides.length)}</Slides>` +
    '<ScaleCrop>false</ScaleCrop>' +
    '<AppVersion>0.0000</AppVersion>' +
    '</Properties>';

  // --- assembly -------------------------------------------------------------
  const entries: ZipEntry[] = [];
  const shapes = new Map(Object.entries(spec.zip ?? {}));
  const shaped = new Set<string>();
  const add = (name: string, content: string | Uint8Array): void => {
    const shape = shapes.get(name);
    if (shape !== undefined) shaped.add(name);
    entries.push({
      name,
      bytes: typeof content === 'string' ? encoder.encode(content) : content,
      store: shape?.store ?? spec.deflate !== true,
      ...(shape?.flags === undefined ? {} : { flags: shape.flags }),
      ...(shape?.growthHint === undefined ? {} : { extra: growthHint(shape.growthHint) }),
    });
  };

  // `[Content_Types].xml` first: an OPC rule, not a ZIP one.
  add('[Content_Types].xml', contentTypes);
  add(
    '_rels/.rels',
    relsXml([
      { id: 'rId1', type: REL + 'officeDocument', target: 'ppt/presentation.xml' },
      {
        id: 'rId2',
        type: 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties',
        target: 'docProps/core.xml',
      },
      { id: 'rId3', type: REL + 'extended-properties', target: 'docProps/app.xml' },
      ...(spec.packageRels ?? []),
    ]),
  );
  add('ppt/presentation.xml', presentation);
  add('ppt/_rels/presentation.xml.rels', relsXml(presRels));
  add('ppt/presProps.xml', DECLARATION + `<p:presentationPr ${NS_DECLS}/>`);
  add('ppt/viewProps.xml', DECLARATION + `<p:viewPr ${NS_DECLS}/>`);
  add(
    'ppt/tableStyles.xml',
    DECLARATION + `<a:tblStyleLst xmlns:a="${NS_A}" def="{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"/>`,
  );
  let layoutNumber = 0;
  masters.forEach((master, index) => {
    const m = String(index + 1);
    const firstLayoutNumber = layoutNumber + 1;
    add(
      `ppt/theme/theme${m}.xml`,
      themeXml({
        ...(master.themeName === undefined ? {} : { name: master.themeName }),
        ...(master.colours === undefined ? {} : { colours: master.colours }),
        ...(master.objectDefaults === undefined ? {} : { objectDefaults: master.objectDefaults }),
        ...(master.majorLatin === undefined ? {} : { majorLatin: master.majorLatin }),
        ...(master.minorLatin === undefined ? {} : { minorLatin: master.minorLatin }),
      }),
    );
    add(
      `ppt/slideMasters/slideMaster${m}.xml`,
      slideMasterXml(master, layoutIdsByMaster[index] ?? []),
    );
    add(
      `ppt/slideMasters/_rels/slideMaster${m}.xml.rels`,
      relsXml([
        ...master.layouts.map((_layout, offset) => ({
          id: 'rId' + String(offset + 1),
          type: REL + 'slideLayout',
          target: `../slideLayouts/slideLayout${String(firstLayoutNumber + offset)}.xml`,
        })),
        {
          id: 'rId' + String(master.layouts.length + 1),
          type: REL + 'theme',
          target: `../theme/theme${m}.xml`,
        },
      ]),
    );
    for (const layout of master.layouts) {
      layoutNumber += 1;
      const n = String(layoutNumber);
      add(`ppt/slideLayouts/slideLayout${n}.xml`, slideLayoutXml(layout));
      add(
        `ppt/slideLayouts/_rels/slideLayout${n}.xml.rels`,
        relsXml([
          {
            id: 'rId1',
            type: REL + 'slideMaster',
            target: `../slideMasters/slideMaster${m}.xml`,
          },
        ]),
      );
    }
  });

  if (spec.notesMaster !== undefined && notesThemeNumber !== null) {
    const t = String(notesThemeNumber);
    add(`ppt/theme/theme${t}.xml`, themeXml({ name: 'PPTX Studio Corpus Notes' }));
    add('ppt/notesMasters/notesMaster1.xml', notesMasterXml(spec.notesMaster));
    add(
      'ppt/notesMasters/_rels/notesMaster1.xml.rels',
      relsXml([{ id: 'rId1', type: REL + 'theme', target: `../theme/theme${t}.xml` }]),
    );
  }
  if (spec.handoutMaster !== undefined && handoutThemeNumber !== null) {
    const t = String(handoutThemeNumber);
    add(`ppt/theme/theme${t}.xml`, themeXml({ name: 'PPTX Studio Corpus Handout' }));
    add('ppt/handoutMasters/handoutMaster1.xml', handoutMasterXml(spec.handoutMaster));
    add(
      'ppt/handoutMasters/_rels/handoutMaster1.xml.rels',
      relsXml([{ id: 'rId1', type: REL + 'theme', target: `../theme/theme${t}.xml` }]),
    );
  }

  let notesSlideNumber = 0;
  spec.slides.forEach((slide, index) => {
    const n = String(index + 1);
    const layoutIndex = slide.layout ?? 0;
    const bound = layouts[layoutIndex];
    if (bound === undefined) {
      throw new Error(`slide ${n} binds to layout ${String(layoutIndex)}, which does not exist`);
    }
    const title =
      bound.layout.hasTitle === true
        ? placeholderXml({
            id: 2,
            name: 'Title ' + n,
            type: 'title',
            ...TITLE_BOX,
            body: `<a:p><a:r><a:rPr lang="en-GB" dirty="0"/><a:t>${escapeXml(slide.title)}</a:t></a:r></a:p>`,
          })
        : '';
    add(
      `ppt/slides/slide${n}.xml`,
      DECLARATION +
        `<p:sld ${NS_DECLS}${slide.rootAttributes ?? ''}>` +
        '<p:cSld>' +
        spTreeHead() +
        title +
        slide.body +
        '</p:spTree>' +
        (slide.cSldTail ?? '') +
        '</p:cSld>' +
        `<p:clrMapOvr>${slide.clrMapOvr ?? '<a:masterClrMapping/>'}</p:clrMapOvr>` +
        (slide.tail ?? '') +
        '</p:sld>',
    );

    // Exactly one, always. Measured on 2026-08-28: PowerPoint refuses a package
    // where a slide has none, has two, has one pointing at a master, or has one
    // pointing at a part that is not there. `a38-degenerate` says so at length.
    const slideRels: ProbeRel[] = [
      {
        id: 'rId1',
        type: REL + 'slideLayout',
        target: `../slideLayouts/slideLayout${String(layoutIndex + 1)}.xml`,
      },
      ...(slide.rels ?? []),
    ];
    if (slide.notes !== undefined) {
      notesSlideNumber += 1;
      slideRels.push({
        id: 'rId' + String(slideRels.length + 1),
        type: REL + 'notesSlide',
        target: `../notesSlides/notesSlide${String(notesSlideNumber)}.xml`,
      });
    }
    add(`ppt/slides/_rels/slide${n}.xml.rels`, relsXml(slideRels));

    if (slide.notes !== undefined) {
      const k = String(notesSlideNumber);
      add(`ppt/notesSlides/notesSlide${k}.xml`, notesSlideXml(slide.notes));
      // Both edges, and the second is the one that is easy to forget: a notes
      // slide relates back to the slide it annotates as well as forward to the
      // notes master, and PowerPoint uses that edge - not the part number - to
      // decide which slide a notes page belongs to.
      add(
        `ppt/notesSlides/_rels/notesSlide${k}.xml.rels`,
        relsXml([
          { id: 'rId1', type: REL + 'notesMaster', target: '../notesMasters/notesMaster1.xml' },
          { id: 'rId2', type: REL + 'slide', target: `../slides/slide${n}.xml` },
        ]),
      );
    }
  });

  for (const part of extraParts) add(part.name, part.bytes);
  add('docProps/core.xml', core);
  add('docProps/app.xml', app);

  const unused = [...shapes.keys()].filter((name) => !shaped.has(name));
  if (unused.length > 0) {
    throw new Error('zip shape names no entry: ' + unused.join(', '));
  }

  return { bytes: writeZip(entries), entries: entries.length };
}
