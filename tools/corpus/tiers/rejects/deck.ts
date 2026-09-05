import { writeZip } from '../../../ground-truth/lib/zip.ts';

/**
 * The smallest package that opens, and the seams to break it.
 *
 * `corpus/reject/` is the mirror image of the rest of the corpus. Every deck in
 * `corpus/decks`, `corpus/authored` and `corpus/written` is a file PowerPoint
 * **opens**, and the gate on them is that the validator says nothing. Every
 * package here is one PowerPoint **refuses**, and the gate is that the
 * validator says the right thing.
 *
 * ## Why it is not built on the Tier A chassis
 *
 * `tools/corpus/tiers/a-generated/markup/chassis.ts` builds a good deck - a master with placeholders,
 * a layout, a theme with two gradient fills so `fillRef/@idx` resolves,
 * `docProps`, the lot. That is exactly right for a probe deck, whose job is to
 * carry one feature into PowerPoint surrounded by enough context to be
 * realistic.
 *
 * It is the wrong shape for a refusal. What a refusal fixture is *for* is the
 * smallest markup that reproduces it, because that is what `cli bisect` reduces
 * towards in sub-phase 1.5 and what a person reads when they want to know
 * whether a rule is still true. Twenty-two parts of correct scaffolding around
 * one wrong attribute makes the fixture harder to read and the claim harder to
 * check. So this builds ten parts and no theme: a presentation, one master, one
 * layout, one slide, and the four relationship parts that bind them.
 *
 * Every byte is `store`d rather than deflated, for the reason the rest of the
 * corpus stores its entries: a deflated archive's bytes depend on which zlib
 * built it, so a pinned SHA-256 would need re-pinning after a routine Node
 * upgrade with no source change.
 */

const encoder = new TextEncoder();

const DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';

const NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const NS_P = 'http://schemas.openxmlformats.org/presentationml/2006/main';

const ROOTS = ` xmlns:a="${NS_A}" xmlns:r="${NS_R}" xmlns:p="${NS_P}"`;

export const OD = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/';
export const PML = 'application/vnd.openxmlformats-officedocument.presentationml.';

/** The twelve `p:clrMap` attributes. Eleven is not a map with a gap; it is invalid. */
export const IDENTITY_CLR_MAP =
  'bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" ' +
  'accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" ' +
  'hlink="hlink" folHlink="folHlink"';

/** A shape, schema-ordered. `ph` is `<p:ph .../>` markup or ''; `spPr` extends the geometry. */
export function shape(
  id: number,
  name: string,
  ph: string,
  geometry = '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>',
): string {
  return (
    '<p:sp>' +
    '<p:nvSpPr>' +
    `<p:cNvPr id="${String(id)}" name="${name}"/>` +
    '<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>' +
    `<p:nvPr>${ph}</p:nvPr>` +
    '</p:nvSpPr>' +
    '<p:spPr>' +
    '<a:xfrm><a:off x="838200" y="365125"/><a:ext cx="10515600" cy="1325563"/></a:xfrm>' +
    geometry +
    '</p:spPr>' +
    '<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></p:txBody>' +
    '</p:sp>'
  );
}

export function spTree(shapes: string): string {
  return (
    '<p:spTree>' +
    '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
    '<p:grpSpPr/>' +
    shapes +
    '</p:spTree>'
  );
}

export function rel(id: string, type: string, target: string, mode?: string): string {
  return (
    `<Relationship Id="${id}" Type="${type}" Target="${target}"` +
    (mode === undefined ? '' : ` TargetMode="${mode}"`) +
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

/** The parts of a package that opens, keyed by ZIP entry name. */
export function basePackage(): Record<string, string> {
  return {
    '[Content_Types].xml':
      DECLARATION +
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ' +
      'ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      `<Override PartName="/ppt/presentation.xml" ContentType="${PML}presentation.main+xml"/>` +
      `<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="${PML}slideMaster+xml"/>` +
      `<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="${PML}slideLayout+xml"/>` +
      `<Override PartName="/ppt/slides/slide1.xml" ContentType="${PML}slide+xml"/>` +
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
      `<p:clrMap ${IDENTITY_CLR_MAP}/>` +
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

/** Entry order, as Office writes it. Anything new is appended in insertion order. */
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

/** A part map with overrides applied. `null` removes a part. */
export type PartOverrides = Readonly<Record<string, string | null>>;

export function packageBytes(overrides: PartOverrides = {}): Uint8Array {
  const parts: Record<string, string | null> = { ...basePackage(), ...overrides };
  const names = [
    ...ORDER.filter((name) => name in parts),
    ...Object.keys(parts).filter((name) => !ORDER.includes(name)),
  ];
  return writeZip(
    names
      .filter((name) => parts[name] !== null)
      .map((name) => ({ name, bytes: encoder.encode(parts[name]!), store: true })),
  );
}
