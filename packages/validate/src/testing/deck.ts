import { PartStore, storedEntry, writeZip } from '@pptx-studio/opc';

/**
 * A minimal deck that passes all twenty-nine rules, and the seams to break it.
 *
 * Test-only, and never reachable from `src/index.ts` so it cannot ship.
 *
 * ## Why this exists rather than a corpus deck
 *
 * The core packages test in real Chromium, not jsdom, deliberately - a Node API
 * leak stays invisible under jsdom until a user opens a tab. A browser test has
 * no filesystem, so the fifty-two corpus decks are not reachable from here at
 * all. They are reachable from `tools/`, and that is where the other half of
 * this sub-phase's verification lives: `tools/corpus/suites/validate.test.ts` runs
 * every committed deck through the validator and asserts no fatal finding, on
 * the strength of the one thing the corpus is: **every deck in it opens in
 * PowerPoint**, so any rule that fires on one is a rule that is wrong.
 *
 * The two halves check opposite things and neither substitutes for the other.
 * The corpus proves the rules do not fire on good files. This proves each rule
 * fires on the one bad file it is about - which cannot be shown with real decks,
 * because the corpus contains no dangling relationship, no duplicate id, no
 * missing content type and no malformed part name anywhere in it.
 *
 * ## The deck
 *
 * A presentation, one master, one layout, one slide, no theme. Nothing in the
 * twenty-nine mentions a theme, and leaving it out keeps every fixture below
 * readable. Every part is written schema-ordered by hand rather than by a
 * generator, because a generator that produced the fixtures *and* satisfied the
 * ordering rule would be checking itself.
 */

const encoder = new TextEncoder();

const DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';

const NS_P = 'http://schemas.openxmlformats.org/presentationml/2006/main';
const NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

const ROOTS = ' xmlns:a="' + NS_A + '" xmlns:r="' + NS_R + '" xmlns:p="' + NS_P + '"';

/** The identity colour map: all twelve attributes, which is the only legal number. */
export const IDENTITY_CLR_MAP =
  'bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" ' +
  'accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" ' +
  'hlink="hlink" folHlink="folHlink"';

/** A placeholder shape, schema-ordered. `ph` is the `<p:ph .../>` markup, or ''. */
export function shape(id: number, name: string, ph: string, body = ''): string {
  return (
    '<p:sp>' +
    '<p:nvSpPr>' +
    '<p:cNvPr id="' +
    String(id) +
    '" name="' +
    name +
    '"/>' +
    '<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>' +
    '<p:nvPr>' +
    ph +
    '</p:nvPr>' +
    '</p:nvSpPr>' +
    '<p:spPr>' +
    '<a:xfrm><a:off x="838200" y="365125"/><a:ext cx="10515600" cy="1325563"/></a:xfrm>' +
    '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>' +
    body +
    '</p:spPr>' +
    '<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></p:txBody>' +
    '</p:sp>'
  );
}

function spTree(shapes: string): string {
  return (
    '<p:spTree>' +
    '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
    '<p:grpSpPr/>' +
    shapes +
    '</p:spTree>'
  );
}

/** One `<Relationship>` line. */
export function rel(id: string, type: string, target: string, mode?: string): string {
  return (
    '<Relationship Id="' +
    id +
    '" Type="' +
    type +
    '" Target="' +
    target +
    '"' +
    (mode === undefined ? '' : ' TargetMode="' + mode + '"') +
    '/>'
  );
}

export function relsPart(body: string): string {
  return (
    DECLARATION +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    body +
    '</Relationships>'
  );
}

const OD = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/';
const PML = 'application/vnd.openxmlformats-officedocument.presentationml.';

/**
 * The parts of a deck that passes, keyed by ZIP entry name.
 *
 * A plain record so a fixture can replace one part by spreading over it, which
 * is how every test below breaks exactly one thing.
 */
export function minimalDeck(): Record<string, string> {
  return {
    '[Content_Types].xml':
      DECLARATION +
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/ppt/presentation.xml" ContentType="' +
      PML +
      'presentation.main+xml"/>' +
      '<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="' +
      PML +
      'slideMaster+xml"/>' +
      '<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="' +
      PML +
      'slideLayout+xml"/>' +
      '<Override PartName="/ppt/slides/slide1.xml" ContentType="' +
      PML +
      'slide+xml"/>' +
      '</Types>',

    '_rels/.rels': relsPart(rel('rId1', OD + 'officeDocument', 'ppt/presentation.xml')),

    'ppt/presentation.xml':
      DECLARATION +
      '<p:presentation' +
      ROOTS +
      '>' +
      '<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>' +
      '<p:sldIdLst><p:sldId id="256" r:id="rId2"/></p:sldIdLst>' +
      '<p:sldSz cx="12192000" cy="6858000"/>' +
      '<p:notesSz cx="6858000" cy="9144000"/>' +
      '</p:presentation>',

    'ppt/_rels/presentation.xml.rels': relsPart(
      rel('rId1', OD + 'slideMaster', 'slideMasters/slideMaster1.xml') +
        rel('rId2', OD + 'slide', 'slides/slide1.xml'),
    ),

    'ppt/slideMasters/slideMaster1.xml':
      DECLARATION +
      '<p:sldMaster' +
      ROOTS +
      '>' +
      '<p:cSld>' +
      spTree(shape(2, 'Title Placeholder 1', '<p:ph type="title"/>')) +
      '</p:cSld>' +
      '<p:clrMap ' +
      IDENTITY_CLR_MAP +
      '/>' +
      '<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst>' +
      '</p:sldMaster>',

    'ppt/slideMasters/_rels/slideMaster1.xml.rels': relsPart(
      rel('rId1', OD + 'slideLayout', '../slideLayouts/slideLayout1.xml'),
    ),

    'ppt/slideLayouts/slideLayout1.xml':
      DECLARATION +
      '<p:sldLayout' +
      ROOTS +
      ' type="title" preserve="1">' +
      '<p:cSld name="Title Slide">' +
      spTree(shape(2, 'Title 1', '<p:ph type="title"/>')) +
      '</p:cSld>' +
      '<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>' +
      '</p:sldLayout>',

    'ppt/slideLayouts/_rels/slideLayout1.xml.rels': relsPart(
      rel('rId1', OD + 'slideMaster', '../slideMasters/slideMaster1.xml'),
    ),

    'ppt/slides/slide1.xml':
      DECLARATION +
      '<p:sld' +
      ROOTS +
      '>' +
      '<p:cSld>' +
      spTree(shape(2, 'Title 1', '<p:ph type="title"/>')) +
      '</p:cSld>' +
      '<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>' +
      '</p:sld>',

    'ppt/slides/_rels/slide1.xml.rels': relsPart(
      rel('rId1', OD + 'slideLayout', '../slideLayouts/slideLayout1.xml'),
    ),
  };
}

/** Part names in the order Office writes them, so the fixtures look like real files. */
const ORDER = [
  '[Content_Types].xml',
  '_rels/.rels',
  'ppt/presentation.xml',
  'ppt/_rels/presentation.xml.rels',
  'ppt/slideMasters/slideMaster1.xml',
  'ppt/slideMasters/_rels/slideMaster1.xml.rels',
  'ppt/slideLayouts/slideLayout1.xml',
  'ppt/slideLayouts/_rels/slideLayout1.xml.rels',
  'ppt/slides/slide1.xml',
  'ppt/slides/_rels/slide1.xml.rels',
];

export interface DeckOverrides {
  /** Replace or add parts. A value of `null` removes the part. */
  readonly parts?: Readonly<Record<string, string | Uint8Array | null>>;
  /** Extra ZIP entries appended verbatim - directory entries, for instance. */
  readonly extraEntries?: readonly { name: string; bytes: Uint8Array }[];
}

/** The archive bytes for a deck, minimal by default. */
export function deckBytes(overrides: DeckOverrides = {}): Uint8Array {
  const parts: Record<string, string | Uint8Array | null> = {
    ...minimalDeck(),
    ...(overrides.parts ?? {}),
  };

  const names = [
    ...ORDER.filter((name) => name in parts),
    ...Object.keys(parts).filter((name) => !ORDER.includes(name)),
  ];

  const entries = names
    .filter((name) => parts[name] !== null && parts[name] !== undefined)
    .map((name) => {
      const value = parts[name]!;
      return storedEntry(name, typeof value === 'string' ? encoder.encode(value) : value);
    });

  for (const extra of overrides.extraEntries ?? []) {
    entries.push(storedEntry(extra.name, extra.bytes));
  }
  return writeZip(entries);
}

/** The deck as a `PartStore`, opened from its own bytes. */
export function deckStore(overrides: DeckOverrides = {}): PartStore {
  return PartStore.open(deckBytes(overrides));
}

/** Both, for the rules that want the archive as well as the store. */
export function deck(overrides: DeckOverrides = {}): {
  bytes: Uint8Array;
  store: PartStore;
} {
  const bytes = deckBytes(overrides);
  return { bytes, store: PartStore.open(bytes) };
}
