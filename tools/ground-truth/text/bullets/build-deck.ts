/**
 * Experiment T5, step 2 - write the probe decks.
 *
 * ```
 * node tools/ground-truth/text/bullets/build-deck.ts <work-dir>
 * ```
 *
 * One package per deck, plus `bullet-inputs.json` describing what every shape
 * was asked and which marker names it.
 *
 * ## The marker
 *
 * An EMF is a list of drawing calls with no shape names in it, so every probe
 * paragraph ends with a run of the form `#12#` and the attribution rule is that
 * everything drawn strictly between two markers belongs to the probe whose
 * marker follows. That works because GDI is asked for a bullet before it is
 * asked for the paragraph it belongs to, which the authoring step confirmed on
 * forty-one shapes in a row.
 *
 * `#` is asserted to appear in no probe string and in no bullet character, and
 * the marker run pins `a:latin`, `a:ea` and `a:cs` to one face so that the
 * marker itself can never split into two records and be mistaken for two.
 *
 * ## What is held still
 *
 * `kern="0"` and no `a:lnSpc` everywhere, as in T2, T3 and T4. `marL` and
 * `indent` are written as an explicit `0` on every probe that is not asking
 * about them, because the schema's defaults for those two are 347663 and
 * -342900 and inheriting either would move the text by a quarter inch under a
 * question about something else.
 *
 * Insets are zero so that `Paragraphs(i).BoundLeft` minus the shape's own left
 * is the bullet's advance with nothing else in it.
 */

import { mkdirSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { deflateSync } from 'node:zlib';

import { allProbes, CASCADE_MASTER_BULLET, type Bullet, type Para, type Probe } from './probes.ts';
import {
  buildSheetPackage,
  emu,
  SCHEME_ONE,
  shape,
  type LayoutSpec,
  type MasterSpec,
  type SlideSpec,
} from '../../lib/sheet-pptx.ts';

const outDir = process.argv[2];
if (outDir === undefined)
  throw new Error('usage: tools/ground-truth/text/bullets/build-deck.ts <out-dir>');
mkdirSync(outDir, { recursive: true });

/* -------------------------------------------------------------------------- */
/* markup                                                                     */
/* -------------------------------------------------------------------------- */

function escapeXml(text: string): string {
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

/** A colour that is either six hex digits or a scheme slot name. */
function colorXml(value: string): string {
  return /^[0-9A-Fa-f]{6}$/.test(value)
    ? `<a:srgbClr val="${value}"/>`
    : `<a:schemeClr val="${value}"/>`;
}

/**
 * The bullet elements, in `CT_TextParagraphProperties` order.
 *
 * The order is not cosmetic: `a:buClrTx`/`a:buClr`, then `a:buSzTx`/`a:buSzPct`/
 * `a:buSzPts`, then `a:buFontTx`/`a:buFont`, then the bullet itself. Every
 * complex type in OOXML is an `xsd:sequence` and PowerPoint refuses a file that
 * gets one wrong - and this particular sequence has four exclusive groups in a
 * row, which is exactly the shape a hand-written emitter transposes.
 */
function bulletXml(bullet: Bullet): string {
  const parts: string[] = [];
  if (bullet.clrTx === true) parts.push('<a:buClrTx/>');
  else if (bullet.clr !== undefined) parts.push(`<a:buClr>${colorXml(bullet.clr)}</a:buClr>`);

  if (bullet.szTx === true) parts.push('<a:buSzTx/>');
  else if (bullet.szPct !== undefined) parts.push(`<a:buSzPct val="${String(bullet.szPct)}"/>`);
  else if (bullet.szPts !== undefined) parts.push(`<a:buSzPts val="${String(bullet.szPts)}"/>`);

  if (bullet.font === 'tx') parts.push('<a:buFontTx/>');
  else if (bullet.font !== undefined)
    parts.push(`<a:buFont typeface="${escapeXml(bullet.font)}"/>`);

  if (bullet.kind === 'none') parts.push('<a:buNone/>');
  else if (bullet.kind === 'char') {
    if (bullet.char === undefined) throw new Error('a char bullet with no char');
    parts.push(`<a:buChar char="${escapeXml(bullet.char)}"/>`);
  } else if (bullet.kind === 'autonum') {
    if (bullet.scheme === undefined) throw new Error('an autonum bullet with no scheme');
    const startAt = bullet.startAt === undefined ? '' : ` startAt="${String(bullet.startAt)}"`;
    parts.push(`<a:buAutoNum type="${bullet.scheme}"${startAt}/>`);
  } else if (bullet.kind === 'blip') {
    if (bullet.blip === undefined) throw new Error('a blip bullet with no media');
    parts.push(`<a:buBlip><a:blip r:embed="${blipRelId(bullet.blip)}"/></a:buBlip>`);
  }
  return parts.join('');
}

/** The three media files, and the fixed rId each is given on every slide. */
const BLIP_MEDIA = [
  { id: 'bullet-square', w: 32, h: 32, rgb: [0xc0, 0x00, 0x00] as const },
  { id: 'bullet-wide', w: 64, h: 16, rgb: [0x00, 0x70, 0xc0] as const },
  { id: 'bullet-tall', w: 16, h: 64, rgb: [0x00, 0xa0, 0x40] as const },
];

function blipRelId(media: string): string {
  const at = BLIP_MEDIA.findIndex((m) => m.id === media);
  if (at < 0) throw new Error(`no media named ${media}`);
  return `rId${String(at + 2)}`;
}

/**
 * Reject anything XML 1.0 cannot carry, and anything the marker rule needs.
 *
 * Not defensive tidiness. A stray U+0001 in one probe string got as far as
 * PowerPoint, which answered by refusing the whole package - no part name, no
 * character, just the one sentence it ever says - and the probe it belonged to
 * looked structurally identical to its thirteen neighbours that opened fine.
 * A control character in an `a:t` has to fail here, loudly, or it fails there,
 * silently.
 */
function assertProbeText(text: string, what: string): void {
  if (text.includes('#')) {
    throw new Error(`${what} ${JSON.stringify(text)} contains a marker delimiter`);
  }
  for (const char of text) {
    const code = char.codePointAt(0) ?? 0;
    const legal =
      code === 0x09 || code === 0x0a || code === 0x0d || (code >= 0x20 && code !== 0x7f);
    if (!legal) {
      throw new Error(
        `${what} ${JSON.stringify(text)} holds U+${code.toString(16).toUpperCase().padStart(4, '0')}, which XML 1.0 forbids`,
      );
    }
  }
}

function rPrXml(p: Para, tag: 'a:rPr' | 'a:defRPr' | 'a:endParaRPr', sz: number): string {
  const faces =
    `<a:latin typeface="${escapeXml(p.face)}"/>` +
    (p.ea === undefined ? '' : `<a:ea typeface="${escapeXml(p.ea)}"/>`) +
    (p.cs === undefined ? '' : `<a:cs typeface="${escapeXml(p.cs)}"/>`) +
    (p.sym === undefined ? '' : `<a:sym typeface="${escapeXml(p.sym)}"/>`);
  const fill = p.runClr === undefined ? '' : `<a:solidFill>${colorXml(p.runClr)}</a:solidFill>`;
  return (
    `<${tag} lang="${p.lang ?? 'en-US'}" sz="${String(sz)}" kern="0" dirty="0">` +
    fill +
    faces +
    `</${tag}>`
  );
}

/**
 * The marker run.
 *
 * All three script slots name the same face, so no `#` or digit can be pushed
 * into a second text record by the script-run rule this experiment is also
 * measuring - a marker that split in two would make every probe after it read
 * one probe late.
 */
function markerXml(marker: number): string {
  return (
    '<a:r><a:rPr lang="en-US" sz="1200" kern="0" dirty="0">' +
    '<a:latin typeface="Arial"/><a:ea typeface="Arial"/><a:cs typeface="Arial"/>' +
    '</a:rPr>' +
    `<a:t>#${String(marker)}#</a:t></a:r>`
  );
}

function paraXml(p: Para, marker: number): string {
  const attrs =
    (p.level === undefined ? '' : ` lvl="${String(p.level)}"`) +
    (p.marL === undefined ? '' : ` marL="${String(emu(p.marL, 'marL'))}"`) +
    (p.indent === undefined ? '' : ` indent="${String(emu(p.indent, 'indent'))}"`) +
    (p.defTabSz === undefined ? '' : ` defTabSz="${String(emu(p.defTabSz, 'defTabSz'))}"`) +
    (p.algn === undefined ? '' : ` algn="${p.algn}"`) +
    (p.rtl === undefined ? '' : ` rtl="${p.rtl ? '1' : '0'}"`);

  const defRPr = p.defSz === undefined ? '' : rPrXml(p, 'a:defRPr', p.defSz);
  const pPr = `<a:pPr${attrs}>${bulletXml(p.bullet)}${defRPr}</a:pPr>`;

  const body: string[] = [];
  if (p.field !== undefined) {
    const [type, cached] = p.field;
    const typeAttr = type === '' ? '' : ` type="${escapeXml(type)}"`;
    body.push(
      `<a:fld id="{B9CFA1B7-1A46-4F4C-9E2B-${String(marker).padStart(12, '0')}}"${typeAttr}>` +
        rPrXml(p, 'a:rPr', p.sz) +
        `<a:t>${escapeXml(cached)}</a:t></a:fld>`,
    );
  }
  const text = p.text ?? 'X';
  assertProbeText(text, 'probe text');
  if (p.bullet.char !== undefined) assertProbeText(p.bullet.char, 'buChar/@char');
  if (text !== '') {
    const runs =
      p.hardBreak === true
        ? [text.slice(0, text.length >> 1), text.slice(text.length >> 1)]
        : [text];
    runs.forEach((run, i) => {
      if (i > 0) body.push(`<a:br>${rPrXml(p, 'a:rPr', p.sz)}</a:br>`);
      body.push(`<a:r>${rPrXml(p, 'a:rPr', p.sz)}<a:t>${escapeXml(run)}</a:t></a:r>`);
    });
  }
  if (p.secondRun !== undefined) {
    assertProbeText(p.secondRun.text, 'second run text');
    body.push(
      `<a:r>${rPrXml(p, 'a:rPr', p.secondRun.sz)}<a:t>${escapeXml(p.secondRun.text)}</a:t></a:r>`,
    );
  }
  body.push(markerXml(marker));
  return `<a:p>${pPr}${body.join('')}</a:p>`;
}

function lstStyleXml(bullet: Bullet): string {
  return `<a:lstStyle><a:lvl1pPr marL="0" indent="0">${bulletXml(bullet)}</a:lvl1pPr></a:lstStyle>`;
}

/* -------------------------------------------------------------------------- */
/* a solid PNG, for the picture bullets                                       */
/* -------------------------------------------------------------------------- */

function crc32(bytes: Uint8Array): number {
  let crc = 0xffffffff;
  for (const byte of bytes) {
    crc ^= byte;
    for (let bit = 0; bit < 8; bit++) crc = (crc >>> 1) ^ (0xedb88320 & -(crc & 1));
  }
  return (crc ^ 0xffffffff) >>> 0;
}

function pngChunk(type: string, data: Uint8Array): Uint8Array {
  const out = new Uint8Array(12 + data.length);
  const view = new DataView(out.buffer);
  view.setUint32(0, data.length, false);
  for (let i = 0; i < 4; i++) out[4 + i] = type.charCodeAt(i);
  out.set(data, 8);
  view.setUint32(8 + data.length, crc32(out.subarray(4, 8 + data.length)), false);
  return out;
}

/**
 * A solid rectangle, which is all a picture bullet has to be.
 *
 * Three aspect ratios rather than one colour swatch: a square source cannot
 * tell "scaled to the line box" from "scaled to the em" from "letterboxed into
 * a square", and those are three different left edges for the text after it.
 */
function solidPng(width: number, height: number, rgb: readonly number[]): Uint8Array {
  const stride = width * 3 + 1;
  const raw = new Uint8Array(stride * height);
  for (let y = 0; y < height; y++) {
    const row = y * stride;
    raw[row] = 0;
    for (let x = 0; x < width; x++) {
      raw[row + 1 + x * 3] = rgb[0] ?? 0;
      raw[row + 2 + x * 3] = rgb[1] ?? 0;
      raw[row + 3 + x * 3] = rgb[2] ?? 0;
    }
  }
  const ihdr = new Uint8Array(13);
  const header = new DataView(ihdr.buffer);
  header.setUint32(0, width, false);
  header.setUint32(4, height, false);
  ihdr[8] = 8;
  ihdr[9] = 2;
  const deflated = deflateSync(raw, { level: 9 });
  const parts = [
    new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
    pngChunk('IHDR', ihdr),
    pngChunk('IDAT', new Uint8Array(deflated.buffer, deflated.byteOffset, deflated.byteLength)),
    pngChunk('IEND', new Uint8Array(0)),
  ];
  const total = parts.reduce((sum, part) => sum + part.length, 0);
  const png = new Uint8Array(total);
  let at = 0;
  for (const part of parts) {
    png.set(part, at);
    at += part.length;
  }
  return png;
}

/* -------------------------------------------------------------------------- */
/* packing                                                                    */
/* -------------------------------------------------------------------------- */

const SLIDE_W = 960;
const SLIDE_H = 540;
const COLUMNS = [20, 490];
const TOP = 12;
const BOTTOM = 528;
const GAP = 4;

interface Placed {
  readonly probe: Probe;
  readonly slide: number;
  readonly x: number;
  readonly y: number;
  readonly markers: readonly number[];
}

/**
 * Two columns, top to bottom, then a new slide.
 *
 * Nothing is measured across shapes, so the packing only has to keep probes
 * apart. It does have to keep them *properly* apart: C4's lesson was that a
 * reading of the neighbouring shape looks like a plausible number rather than
 * an error, so every shape gets its stated height plus a gap and none of them
 * paint outside their own rectangle.
 */
function pack(probes: readonly Probe[], onePerSlide: boolean, startMarker: number): Placed[] {
  // A deck holding anything wider than a column gets one column. Two columns
  // would put the wide probe's tail past the slide's right edge, where it is
  // not drawn at all - and a probe whose marker was clipped welds itself onto
  // its neighbour instead of failing.
  const columns = probes.some((p) => p.width > 460) ? [COLUMNS[0] ?? 20] : COLUMNS;
  const out: Placed[] = [];
  let slide = 1;
  let column = 0;
  let y = TOP;
  let marker = startMarker;
  for (const probe of probes) {
    if (onePerSlide) {
      if (out.length > 0) slide++;
      column = 0;
      y = TOP;
    } else if (y + probe.height > BOTTOM) {
      column++;
      y = TOP;
      if (column >= columns.length) {
        column = 0;
        slide++;
      }
    }
    const markers = probe.paragraphs.map(() => marker++);
    out.push({ probe, slide, x: columns[column] ?? 20, y, markers });
    y += probe.height + GAP;
  }
  return out;
}

/* -------------------------------------------------------------------------- */
/* decks                                                                      */
/* -------------------------------------------------------------------------- */

const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

/** The master's `p:bodyStyle`, which only the `cascade-master` deck states. */
function bodyStyleXml(bullet: Bullet): string {
  const levels: string[] = [];
  for (let i = 1; i <= 9; i++) {
    levels.push(
      `<a:lvl${String(i)}pPr marL="0" indent="0" algn="l" defTabSz="914400">` +
        (i === 1 ? bulletXml(bullet) : '<a:buNone/>') +
        `<a:defRPr sz="2400" kern="0"/></a:lvl${String(i)}pPr>`,
    );
  }
  return (
    '<p:txStyles><p:titleStyle><a:lvl1pPr><a:defRPr sz="2400"/></a:lvl1pPr></p:titleStyle>' +
    `<p:bodyStyle>${levels.join('')}</p:bodyStyle>` +
    '<p:otherStyle><a:lvl1pPr><a:defRPr sz="1800"/></a:lvl1pPr></p:otherStyle></p:txStyles>'
  );
}

const probes = allProbes();
const byDeck = new Map<string, Probe[]>();
for (const probe of probes) {
  const list = byDeck.get(probe.deck);
  if (list === undefined) byDeck.set(probe.deck, [probe]);
  else list.push(probe);
}

interface DeckInput {
  readonly deck: string;
  readonly file: string;
  readonly slides: number;
  readonly shapes: readonly {
    readonly id: string;
    readonly family: string;
    readonly slide: number;
    readonly left: number;
    readonly top: number;
    readonly markers: readonly number[];
    readonly asks: string;
    readonly readChars: boolean;
  }[];
}

const inputs: DeckInput[] = [];
let nextMarker = 1;

for (const [deck, list] of byDeck) {
  const usesMaster = deck === 'cascade-master';
  const usesBlip = deck === 'blip';
  const onePerSlide = deck === 'field-slidenum' || usesMaster;

  const placed = pack(list, onePerSlide, nextMarker);
  nextMarker += placed.reduce((sum, p) => sum + p.markers.length, 0);
  const slideCount = placed.reduce((max, p) => Math.max(max, p.slide), 1);

  const slides: SlideSpec[] = [];
  for (let n = 1; n <= slideCount; n++) {
    const shapes = placed
      .filter((p) => p.slide === n)
      .map((p, i) =>
        shape({
          id: i + 2,
          name: p.probe.id,
          rect: { x: p.x, y: p.y, w: p.probe.width, h: p.probe.height },
          ...(usesMaster ? { ph: 'type="body" idx="1"' } : {}),
          bodyPr:
            `<a:bodyPr wrap="${p.probe.wrap ? 'square' : 'none'}" ` +
            'lIns="0" tIns="0" rIns="0" bIns="0" anchor="t"><a:noAutofit/></a:bodyPr>',
          ...(p.probe.family === 'cascade' && deck === 'cascade'
            ? { lstStyle: lstStyleXml(CASCADE_MASTER_BULLET) }
            : {}),
          paragraphs: p.probe.paragraphs.map((para, k) => paraXml(para, p.markers[k] ?? -1)),
        }),
      );
    slides.push({
      layout: 0,
      name: `${deck}-${String(n)}`,
      shapes,
      ...(usesBlip
        ? {
            rels: [
              {
                id: 'rId1',
                type: `${REL}/slideLayout`,
                target: '../slideLayouts/slideLayout1.xml',
              },
              ...BLIP_MEDIA.map((m, k) => ({
                id: `rId${String(k + 2)}`,
                type: `${REL}/image`,
                target: `../media/${m.id}.png`,
              })),
            ],
          }
        : {}),
    });
  }

  const master: MasterSpec = {
    theme: 0,
    ...(usesMaster ? { txStyles: bodyStyleXml(CASCADE_MASTER_BULLET) } : {}),
  };
  const layout: LayoutSpec = {
    master: 0,
    type: 'obj',
    name: 'Probe',
    ...(usesMaster
      ? {
          shapes: [
            shape({
              id: 2,
              name: 'body-ph',
              rect: { x: 20, y: 12, w: 460, h: 44 },
              ph: 'type="body" idx="1"',
              bodyPr: '<a:bodyPr lIns="0" tIns="0" rIns="0" bIns="0"/>',
            }),
          ],
        }
      : {}),
  };

  const bytes = buildSheetPackage({
    themes: [{ scheme: SCHEME_ONE, majorLatin: 'Georgia', minorLatin: 'Tahoma' }],
    masters: [master],
    layouts: [layout],
    slides,
    ...(usesBlip
      ? {
          media: BLIP_MEDIA.map((m) => ({ name: `${m.id}.png`, bytes: solidPng(m.w, m.h, m.rgb) })),
        }
      : {}),
  });

  const file = `bullets-${deck}.pptx`;
  writeFileSync(join(outDir, file), bytes);
  inputs.push({
    deck,
    file,
    slides: slideCount,
    shapes: placed.map((p) => ({
      id: p.probe.id,
      family: p.probe.family,
      slide: p.slide,
      left: p.x,
      top: p.y,
      markers: p.markers,
      asks: p.probe.asks,
      readChars: p.probe.readChars === true,
    })),
  });
  console.log(
    `${file.padEnd(34)} ${String(slideCount).padStart(3)} slide(s)  ${String(list.length).padStart(4)} probe(s)  ${String(bytes.length).padStart(7)} bytes`,
  );
}

writeFileSync(
  join(outDir, 'bullet-inputs.json'),
  JSON.stringify(
    {
      slideWidth: SLIDE_W,
      slideHeight: SLIDE_H,
      decks: inputs,
    },
    null,
    2,
  ),
);

const shapes = inputs.reduce((sum, d) => sum + d.shapes.length, 0);
const slides = inputs.reduce((sum, d) => sum + d.slides, 0);
console.log(`\n${String(inputs.length)} decks, ${String(slides)} slides, ${String(shapes)} probes`);
console.log(`wrote ${join(outDir, 'bullet-inputs.json')}`);
