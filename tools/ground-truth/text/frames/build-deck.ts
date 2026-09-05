/**
 * Experiment T6, step 2 - write the probe decks.
 *
 * ```
 * node tools/ground-truth/text/frames/build-deck.ts <work-dir>
 * ```
 *
 * One package per family, one more per hostile probe, plus `frame-inputs.json`.
 * Probes read from the EMF get a slide each, since an EMF names no shapes.
 */

import { mkdirSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';

import {
  allProbes,
  inheritProbes,
  type Frame,
  type Para,
  type Probe,
  type Space,
} from './probes.ts';
import {
  buildSheetPackage,
  emu,
  SCHEME_ONE,
  shape,
  SLIDE_HEIGHT_PT,
  SLIDE_WIDTH_PT,
  type LayoutSpec,
  type MasterSpec,
  type SlideSpec,
} from '../../lib/sheet-pptx.ts';

const arg = process.argv[2];
if (arg === undefined) throw new Error('usage: build-deck.ts <out-dir>');
const outDir: string = arg;
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

/** Reject anything XML 1.0 cannot carry: PowerPoint refuses the package and names nothing. */
function assertProbeText(text: string, what: string): void {
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

/** Points to EMU, for a value that may be fractional in points but not in EMU. */
function ins(points: number, what: string): string {
  return String(emu(points, what));
}

function spaceXml(tag: string, space: Space): string {
  const inner =
    space.kind === 'percent'
      ? `<a:spcPct val="${String(space.value)}"/>`
      : `<a:spcPts val="${String(space.value)}"/>`;
  return `<${tag}>${inner}</${tag}>`;
}

/**
 * `a:bodyPr`, with every attribute the probe states and nothing it does not.
 *
 * `null` writes no attribute and `0` writes `="0"`, which is the whole inset
 * question.
 */
function bodyPrXml(f: Frame): string {
  const attrs: string[] = [];
  const push = (name: string, value: string | undefined): void => {
    if (value !== undefined) attrs.push(`${name}="${value}"`);
  };
  push('rot', f.rot === undefined ? undefined : String(Math.round(f.rot * 60000)));
  push(
    'spcFirstLastPara',
    f.spcFirstLastPara === undefined ? undefined : f.spcFirstLastPara ? '1' : '0',
  );
  push('vertOverflow', f.vertOverflow);
  push('horzOverflow', f.horzOverflow);
  push('vert', f.vert);
  push('wrap', f.wrap);
  for (const edge of ['lIns', 'tIns', 'rIns', 'bIns'] as const) {
    const value = f[edge];
    if (value === undefined || value === null) continue;
    push(edge, ins(value, edge));
  }
  push('numCol', f.numCol === undefined ? undefined : String(f.numCol));
  push('spcCol', f.spcCol === undefined ? undefined : ins(f.spcCol, 'spcCol'));
  push('rtlCol', f.rtlCol === undefined ? undefined : f.rtlCol ? '1' : '0');
  push('anchor', f.anchor);
  push('anchorCtr', f.anchorCtr === undefined ? undefined : f.anchorCtr ? '1' : '0');
  push('upright', f.upright === undefined ? undefined : f.upright ? '1' : '0');
  push('compatLnSpc', f.compatLnSpc === undefined ? undefined : f.compatLnSpc ? '1' : '0');
  push('fromWordArt', f.fromWordArt === undefined ? undefined : f.fromWordArt ? '1' : '0');

  const autofit =
    f.autofit === undefined
      ? ''
      : f.autofit === 'none'
        ? '<a:noAutofit/>'
        : f.autofit === 'norm'
          ? '<a:normAutofit/>'
          : '<a:spAutoFit/>';
  const open = `<a:bodyPr${attrs.length > 0 ? ` ${attrs.join(' ')}` : ''}`;
  return autofit === '' ? `${open}/>` : `${open}>${autofit}</a:bodyPr>`;
}

const SZ_DEFAULT = 1800;
const FACE_DEFAULT = 'Arial';

function rPrXml(p: Para, tag: 'a:rPr' | 'a:defRPr' | 'a:endParaRPr', sz: number): string {
  const face = p.face ?? FACE_DEFAULT;
  return (
    `<${tag} lang="en-US" sz="${String(sz)}" kern="0" dirty="0">` +
    `<a:latin typeface="${escapeXml(face)}"/>` +
    `<a:ea typeface="${escapeXml(face)}"/>` +
    `<a:cs typeface="${escapeXml(face)}"/>` +
    `</${tag}>`
  );
}

/**
 * `a:pPr`, in the order `CT_TextParagraphProperties` sequences its children.
 *
 * `a:buFont` precedes `a:buChar`; PowerPoint repairs a file that reverses them.
 */
function pPrXml(p: Para): string {
  const attrs =
    ` marL="${ins(p.marL ?? 0, 'marL')}" indent="${ins(p.indent ?? 0, 'indent')}"` +
    (p.algn === undefined ? '' : ` algn="${p.algn}"`);

  const children: string[] = [];
  if (p.lnSpc !== undefined) children.push(spaceXml('a:lnSpc', p.lnSpc));
  if (p.spcBef !== undefined) children.push(spaceXml('a:spcBef', p.spcBef));
  if (p.spcAft !== undefined) children.push(spaceXml('a:spcAft', p.spcAft));

  const bullet = p.bullet;
  if (bullet !== undefined) {
    if (bullet.kind === 'char') {
      if (bullet.char === undefined) throw new Error('a char bullet with no character');
      assertProbeText(bullet.char, 'buChar/@char');
      if (bullet.font !== undefined) {
        children.push(`<a:buFont typeface="${escapeXml(bullet.font)}"/>`);
      }
      children.push(`<a:buChar char="${escapeXml(bullet.char)}"/>`);
    } else if (bullet.kind === 'autonum') {
      if (bullet.scheme === undefined) throw new Error('an autonum bullet with no scheme');
      children.push(`<a:buAutoNum type="${bullet.scheme}"/>`);
    } else {
      children.push('<a:buNone/>');
    }
  }

  if (p.empty === true || p.sz !== undefined) {
    children.push(rPrXml(p, 'a:defRPr', p.sz ?? SZ_DEFAULT));
  }
  return `<a:pPr${attrs}>${children.join('')}</a:pPr>`;
}

function paraXml(p: Para): string {
  const sz = p.sz ?? SZ_DEFAULT;
  const body: string[] = [];

  if (p.empty !== true) {
    assertProbeText(p.text, 'probe text');
    const lines = p.lines ?? 1;
    if (!Number.isInteger(lines) || lines < 1) {
      throw new Error(`a paragraph cannot hold ${String(lines)} lines`);
    }
    for (let i = 0; i < lines; i++) {
      if (i > 0) {
        // The break carries its own run properties, which is how a file says
        // how tall the line it starts should be.
        body.push(`<a:br>${rPrXml(p, 'a:rPr', p.breakSz ?? sz)}</a:br>`);
      }
      const text = lines === 1 ? p.text : `${p.text}L${String(i + 1)}`;
      body.push(`<a:r>${rPrXml(p, 'a:rPr', sz)}<a:t>${escapeXml(text)}</a:t></a:r>`);
    }
  }

  if (p.endSz !== undefined || p.empty === true) {
    body.push(rPrXml(p, 'a:endParaRPr', p.endSz ?? sz));
  }
  return `<a:p>${pPrXml(p)}${body.join('')}</a:p>`;
}

/* -------------------------------------------------------------------------- */
/* placement                                                                  */
/* -------------------------------------------------------------------------- */

const MARGIN = 12;
const GAP = 10;

interface Placed {
  readonly probe: Probe;
  readonly slide: number;
  readonly x: number;
  readonly y: number;
}

/** One grid per box size, and a slide of its own for anything read from the EMF. */
function place(probes: readonly Probe[]): Placed[] {
  const placed: Placed[] = [];
  let slide = 0;

  // Well down the slide: an overflowing bottom-anchored block starts above its
  // shape, and PowerPoint stops drawing at the slide edge.
  const ISOLATED_Y = 150;
  for (const probe of probes.filter((p) => p.isolate === true)) {
    slide++;
    placed.push({ probe, slide, x: MARGIN, y: ISOLATED_Y });
  }

  const packed = probes.filter((p) => p.isolate !== true);
  const sizes = [...new Set(packed.map((p) => `${String(p.w)}x${String(p.h)}`))];
  for (const size of sizes) {
    const group = packed.filter((p) => `${String(p.w)}x${String(p.h)}` === size);
    const first = group[0];
    if (first === undefined) continue;
    const cols = Math.max(1, Math.floor((SLIDE_WIDTH_PT - MARGIN) / (first.w + GAP)));
    const rows = Math.max(1, Math.floor((SLIDE_HEIGHT_PT - MARGIN) / (first.h + GAP)));
    const perSlide = cols * rows;
    group.forEach((probe, i) => {
      if (i % perSlide === 0) slide++;
      const cell = i % perSlide;
      placed.push({
        probe,
        slide,
        x: MARGIN + (cell % cols) * (probe.w + GAP),
        y: MARGIN + Math.floor(cell / cols) * (probe.h + GAP),
      });
    });
  }
  return placed;
}

/* -------------------------------------------------------------------------- */
/* packages                                                                   */
/* -------------------------------------------------------------------------- */

interface DeckRecord {
  readonly deck: string;
  readonly file: string;
  readonly shapes: readonly {
    readonly id: string;
    readonly slide: number;
    readonly family: string;
    readonly rect: { x: number; y: number; w: number; h: number };
    readonly isolate: boolean;
    readonly frame: Frame;
    readonly paragraphs: number;
    readonly lines: readonly number[];
    readonly vars: Readonly<Record<string, string | number | boolean | null>>;
  }[];
}

const decks: DeckRecord[] = [];

/** A shape's outline, so a human opening the deck can see the box being measured. */
const OUTLINE = '<a:ln w="6350"><a:solidFill><a:srgbClr val="C0C0C0"/></a:solidFill></a:ln>';
const NO_FILL = '<a:noFill/>';

function writeDeck(name: string, probes: readonly Probe[]): void {
  if (probes.length === 0) return;
  const placed = place(probes);
  const slideCount = Math.max(...placed.map((p) => p.slide));

  const slides: SlideSpec[] = [];
  for (let s = 1; s <= slideCount; s++) {
    const here = placed.filter((p) => p.slide === s);
    slides.push({
      layout: 0,
      name: `${name} ${String(s)}`,
      shapes: here.map((entry, i) =>
        shape({
          id: 100 + i,
          name: entry.probe.id,
          rect: { x: entry.x, y: entry.y, w: entry.probe.w, h: entry.probe.h },
          fill: NO_FILL,
          line: OUTLINE,
          bodyPr: bodyPrXml(entry.probe.frame),
          paragraphs: entry.probe.paras.map(paraXml),
        }),
      ),
    });
  }

  const masters: MasterSpec[] = [{ theme: 0 }];
  const layouts: LayoutSpec[] = [{ master: 0, name: 'Probe' }];
  const bytes = buildSheetPackage({
    themes: [{ scheme: SCHEME_ONE }],
    masters,
    layouts,
    slides,
  });
  const file = `${name}.pptx`;
  writeFileSync(join(outDir, file), bytes);

  decks.push({
    deck: name,
    file,
    shapes: placed.map((entry) => ({
      id: entry.probe.id,
      slide: entry.slide,
      family: entry.probe.family,
      rect: { x: entry.x, y: entry.y, w: entry.probe.w, h: entry.probe.h },
      isolate: entry.probe.isolate === true,
      frame: entry.probe.frame,
      paragraphs: entry.probe.paras.length,
      lines: entry.probe.paras.map((p) => (p.empty === true ? 1 : (p.lines ?? 1))),
      vars: entry.probe.vars,
    })),
  });
  console.log(
    `${file.padEnd(28)} ${String(probes.length).padStart(4)} probe(s)  ${String(slideCount).padStart(3)} slide(s)`,
  );
}

/* -------------------------------------------------------------------------- */
/* the inheritance package                                                    */
/* -------------------------------------------------------------------------- */

/**
 * One package per inheritance probe: two probes sharing a master cannot state
 * two different masters.
 */
function writeInheritDecks(): void {
  for (const probe of inheritProbes()) {
    const body = (frame: Frame, id: number): string =>
      shape({
        id,
        name: probe.id,
        ph: 'type="body" idx="1"',
        ...(id === 100 ? { rect: { x: MARGIN, y: MARGIN, w: probe.w, h: probe.h } } : {}),
        fill: NO_FILL,
        line: OUTLINE,
        bodyPr: bodyPrXml(frame),
        paragraphs: probe.paras.map(paraXml),
      });

    const bytes = buildSheetPackage({
      themes: [{ scheme: SCHEME_ONE }],
      // The master's placeholder states the geometry, so the layout's and the
      // slide's can inherit it and the probe measures one thing at a time.
      masters: [{ theme: 0, shapes: [body(probe.masterFrame, 100)] }],
      layouts: [{ master: 0, name: 'Probe', shapes: [body(probe.layoutFrame, 101)] }],
      slides: [{ layout: 0, name: probe.id, shapes: [body(probe.slideFrame, 102)] }],
    });
    const file = `inherit-${probe.property}-${probe.at === 'slide' && probe.vars['at'] === 'none' ? 'none' : probe.at}.pptx`;
    writeFileSync(join(outDir, file), bytes);
    decks.push({
      deck: probe.id,
      file,
      shapes: [
        {
          id: probe.id,
          slide: 1,
          family: 'inherit',
          rect: { x: MARGIN, y: MARGIN, w: probe.w, h: probe.h },
          isolate: false,
          frame: probe.slideFrame,
          paragraphs: probe.paras.length,
          lines: probe.paras.map((p) => p.lines ?? 1),
          vars: probe.vars,
        },
      ],
    });
    console.log(`${file.padEnd(28)}    1 probe     1 slide`);
  }
}

/* -------------------------------------------------------------------------- */
/* main                                                                       */
/* -------------------------------------------------------------------------- */

const probes = allProbes();
const families = [...new Set(probes.map((p) => p.family))];

for (const family of families) {
  const group = probes.filter((p) => p.family === family);
  const wellBehaved = group.filter((p) => p.hostile !== true);
  writeDeck(family, wellBehaved);
  for (const probe of group.filter((p) => p.hostile === true)) {
    writeDeck(probe.id, [probe]);
  }
}
writeInheritDecks();

const inputs = {
  experiment: 'T6',
  subPhase: '3.6',
  question: 'anchors, insets, vertical text, a:br and the height of an empty paragraph',
  slide: { widthPt: SLIDE_WIDTH_PT, heightPt: SLIDE_HEIGHT_PT },
  decks,
};
writeFileSync(join(outDir, 'frame-inputs.json'), `${JSON.stringify(inputs, null, 2)}\n`);

const shapeCount = decks.reduce((n, d) => n + d.shapes.length, 0);
console.log(
  `\n${String(decks.length)} package(s), ${String(shapeCount)} probe(s)\nwrote ${join(outDir, 'frame-inputs.json')}`,
);
