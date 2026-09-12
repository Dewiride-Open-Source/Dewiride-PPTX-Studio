/**
 * One builder per kind of slide, each pure in the slide's index and its own generator.
 *
 * Ids: the chassis writes the title placeholder as 2 and the body placeholder here is 3, so a
 * slide's own shapes count from 10.
 */

import { placeholderXml, type ProbeSlide } from '../../../markup/chassis.ts';
import {
  connector,
  grid,
  group,
  line,
  picture,
  prstGeom,
  scheme,
  shape,
  solidFill,
} from '../../../markup/shapes.ts';
import {
  bodyPr,
  br,
  buAutoNum,
  buChar,
  buFont,
  field,
  normAutofit,
  para,
  run,
  textLine,
  txBody,
  type RunProps,
} from '../../../markup/text.ts';

import {
  CHAPTER_TITLES,
  DECK_SUBTITLE,
  DECK_TITLE,
  between,
  figure,
  lcg,
  name,
  phrase,
  pick,
  role,
  seedOf,
  sentence,
  slideTitle,
} from './content.ts';
import { BARS_RID, MEDIA_RELS, SKYLINE_RID } from './media.ts';
import { chapterOf, kindAt, type Kind } from './schedule.ts';

/** The four layouts `deck.ts` declares, by index. */
export const LAYOUT = { titleOnly: 0, blank: 1, titleAndContent: 2, section: 3 } as const;

/** The content area under the title: what `grid()` in the chassis lays cells over. */
const CONTENT = { x: 457200, y: 1188720, cx: 11277600, cy: 5181600 } as const;

/** Every content slide carries its number in the same corner; `a:fld/@id` is a literal per slide. */
function slideNumber(index: number): string {
  const id = `{A4600000-0000-4000-8000-${String(index + 1).padStart(12, '0')}}`;
  return shape({
    id: 9,
    name: `Slide Number ${String(index + 1)}`,
    x: 10515600,
    y: 6356350,
    cx: 1219200,
    cy: 365125,
    textBody: txBody({
      bodyPr: bodyPr({ wrap: 'square', anchor: 'ctr' }),
      paras: para({
        props: { algn: 'r' },
        content: field({ id, type: 'slidenum', text: String(index + 1), props: { sz: 1000 } }),
        endProps: { sz: 1000 },
      }),
    }),
  });
}

interface BoxSpec {
  readonly id: number;
  readonly name: string;
  readonly x: number;
  readonly y: number;
  readonly cx: number;
  readonly cy: number;
  readonly paras: string;
  readonly body?: Parameters<typeof bodyPr>[0];
  readonly fill?: string;
  readonly rotation?: number;
}

/** A text box: a rectangle with no fill and no outline whose only content is text. */
function textBox(spec: BoxSpec): string {
  return shape({
    id: spec.id,
    name: spec.name,
    x: spec.x,
    y: spec.y,
    cx: spec.cx,
    cy: spec.cy,
    ...(spec.fill === undefined ? {} : { fill: spec.fill }),
    ...(spec.rotation === undefined ? {} : { rotation: spec.rotation }),
    textBody: txBody({
      bodyPr: bodyPr({ wrap: 'square', ...spec.body }),
      paras: spec.paras,
    }),
  });
}

const gradLin = (from: string, to: string, ang = 5400000): string =>
  '<a:gradFill rotWithShape="1"><a:gsLst>' +
  `<a:gs pos="0">${from}</a:gs><a:gs pos="100000">${to}</a:gs>` +
  `</a:gsLst><a:lin ang="${String(ang)}" scaled="0"/></a:gradFill>`;

const gradPath = (from: string, to: string): string =>
  '<a:gradFill rotWithShape="1"><a:gsLst>' +
  `<a:gs pos="0">${from}</a:gs><a:gs pos="100000">${to}</a:gs>` +
  '</a:gsLst><a:path path="circle"><a:fillToRect l="50000" t="50000" r="50000" b="50000"/></a:path></a:gradFill>';

const patt = (prst: string, fg: string, bg: string): string =>
  `<a:pattFill prst="${prst}"><a:fgClr>${fg}</a:fgClr><a:bgClr>${bg}</a:bgClr></a:pattFill>`;

const SHADOW =
  '<a:effectLst><a:outerShdw blurRad="50800" dist="38100" dir="2700000" algn="tl" rotWithShape="0">' +
  '<a:srgbClr val="000000"><a:alpha val="40000"/></a:srgbClr></a:outerShdw></a:effectLst>';

const GLOW =
  '<a:effectLst><a:glow rad="63500"><a:schemeClr val="accent1"><a:alpha val="40000"/></a:schemeClr></a:glow></a:effectLst>';

const ACCENTS = ['accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6'] as const;

const accent = (n: number): string => ACCENTS[n % ACCENTS.length]!;

const bgPr = (fill: string): string => `<p:bg><p:bgPr>${fill}<a:effectLst/></p:bgPr></p:bg>`;

const bgRef = (idx: number, color: string): string =>
  `<p:bg><p:bgRef idx="${String(idx)}">${color}</p:bgRef></p:bg>`;

/* -------------------------------------------------------------------------- */
/* the kinds                                                                  */
/* -------------------------------------------------------------------------- */

function cover(): ProbeSlide {
  return {
    title: DECK_TITLE,
    layout: LAYOUT.blank,
    background: bgPr(solidFill(scheme('dk2'))),
    body:
      shape({
        id: 10,
        name: 'Accent panel',
        x: 8686800,
        y: -457200,
        cx: 2743200,
        cy: 7772400,
        rotation: 900000,
        fill: solidFill(scheme('accent1', '<a:alpha val="70000"/>')),
        caption: false,
      }) +
      textBox({
        id: 11,
        name: 'Deck title',
        x: 914400,
        y: 2286000,
        cx: 7315200,
        cy: 1143000,
        body: { anchor: 'b' },
        paras: textLine(DECK_TITLE, { sz: 4400, b: true, fill: solidFill(scheme('lt1')) }),
      }) +
      textBox({
        id: 12,
        name: 'Deck subtitle',
        x: 914400,
        y: 3474720,
        cx: 7315200,
        cy: 640080,
        paras: textLine(DECK_SUBTITLE, { sz: 2400, fill: solidFill(scheme('lt1')) }),
      }) +
      textBox({
        id: 13,
        name: 'Audience',
        x: 914400,
        y: 5486400,
        cx: 7315200,
        cy: 457200,
        paras: textLine('Prepared for the board', {
          sz: 1400,
          fill: solidFill(scheme('lt1', '<a:alpha val="80000"/>')),
        }),
      }),
  };
}

function agenda(index: number): ProbeSlide {
  const items = CHAPTER_TITLES.map((chapter) =>
    para({
      props: {
        marL: 457200,
        indent: -457200,
        bullet: buAutoNum('arabicPeriod'),
        buFont: buFont('+mj-lt'),
      },
      content: run(chapter, { sz: 2000 }),
    }),
  ).join('');
  return {
    title: 'Agenda',
    layout: LAYOUT.titleAndContent,
    body:
      placeholderXml({ id: 3, name: 'Content Placeholder 2', idx: 1, body: items }) +
      slideNumber(index),
  };
}

function section(index: number): ProbeSlide {
  const chapter = chapterOf(index);
  const route = chapter % 3;
  const dark = route === 0;
  const text = (props: RunProps): RunProps => ({
    ...props,
    fill: solidFill(scheme(dark ? 'lt1' : 'tx1')),
  });
  const background =
    route === 0
      ? bgRef(1003, scheme(accent(chapter)))
      : route === 1
        ? bgPr(gradLin(scheme(accent(chapter), '<a:tint val="30000"/>'), scheme('lt1'), 2700000))
        : undefined;
  return {
    title: CHAPTER_TITLES[chapter] ?? 'Appendix',
    layout: LAYOUT.section,
    ...(background === undefined ? {} : { background }),
    body:
      textBox({
        id: 10,
        name: 'Chapter number',
        x: 914400,
        y: 1828800,
        cx: 3657600,
        cy: 1828800,
        body: { anchor: 'b' },
        paras: textLine(String(chapter + 1).padStart(2, '0'), text({ sz: 6000, b: true })),
      }) +
      textBox({
        id: 11,
        name: 'Chapter title',
        x: 914400,
        y: 3657600,
        cx: 9144000,
        cy: 914400,
        paras: textLine(CHAPTER_TITLES[chapter] ?? 'Appendix', text({ sz: 3600 })),
      }),
  };
}

function bullets(index: number, nth: number): ProbeSlide {
  const rng = lcg(seedOf(index));
  const chapter = chapterOf(index);
  const paras: string[] = [];
  const count = between(rng, 4, 6);
  for (let i = 0; i < count; i++) {
    paras.push(para({ content: run(sentence(rng), { sz: 2000 }) }));
    if (rng() < 0.4) {
      const subs = between(rng, 1, 3);
      for (let k = 0; k < subs; k++) {
        paras.push(para({ props: { lvl: 1 }, content: run(phrase(rng), { sz: 1800 }) }));
      }
    }
  }
  // Every fourth bullets slide has been shrunk to fit, at a step of PowerPoint's own ladder.
  const shrunk = nth % 4 === 3;
  return {
    title: slideTitle(chapter, rng),
    layout: LAYOUT.titleAndContent,
    body:
      placeholderXml({
        id: 3,
        name: 'Content Placeholder 2',
        idx: 1,
        ...(shrunk ? { bodyPr: bodyPr({ autofit: normAutofit(92500, 10000) }) } : {}),
        body: paras.join(''),
      }) + slideNumber(index),
  };
}

function twoColumn(index: number): ProbeSlide {
  const rng = lcg(seedOf(index));
  const chapter = chapterOf(index);
  const left =
    para({
      content:
        run(phrase(rng), { sz: 1600, b: true }) +
        run(' — ', { sz: 1600 }) +
        run(sentence(rng), { sz: 1600 }),
    }) +
    para({
      content:
        run('Read across: ', { sz: 1400, i: true }) +
        run(sentence(rng), { sz: 1400 }) +
        run(' Footnote', { sz: 1400, u: 'sng' }) +
        run('1', { sz: 1400, baseline: 30000 }),
    }) +
    para({
      content:
        run(pick(rng, ['Watch', 'Note', 'Risk']), {
          sz: 1200,
          cap: 'all',
          spc: 100,
          fill: solidFill(scheme('accent2')),
        }) +
        br({ sz: 1200 }) +
        run(sentence(rng), { sz: 1400 }),
    });
  const right =
    para({
      props: { algn: 'just' },
      content: run(sentence(rng) + ' ' + sentence(rng), { sz: 1400 }),
    }) +
    [0, 1, 2]
      .map(() =>
        para({
          props: { marL: 285750, indent: -285750, bullet: buChar('•'), buFont: buFont('Arial') },
          content: run(phrase(rng), { sz: 1400 }),
        }),
      )
      .join('');
  return {
    title: slideTitle(chapter, rng),
    layout: LAYOUT.titleOnly,
    body:
      textBox({
        id: 10,
        name: 'Left column',
        x: CONTENT.x,
        y: CONTENT.y,
        cx: 5486400,
        cy: CONTENT.cy,
        paras: left,
      }) +
      textBox({
        id: 11,
        name: 'Right column',
        x: CONTENT.x + 5791200,
        y: CONTENT.y,
        cx: 5486400,
        cy: CONTENT.cy,
        paras: right,
      }) +
      slideNumber(index),
  };
}

function pictureCaption(index: number, nth: number): ProbeSlide {
  const rng = lcg(seedOf(index));
  const chapter = chapterOf(index);
  const cropped = nth % 3 === 2;
  const relId = nth % 2 === 0 ? SKYLINE_RID : BARS_RID;
  return {
    title: slideTitle(chapter, rng),
    layout: LAYOUT.titleOnly,
    rels: MEDIA_RELS,
    body:
      picture({
        id: 10,
        name: nth % 2 === 0 ? 'Site photograph' : 'Chart screenshot',
        description: nth % 2 === 0 ? 'Hills under a morning sky' : 'Five bars, two colours',
        relId,
        x: CONTENT.x,
        y: CONTENT.y,
        cx: 7315200,
        cy: 4114800,
        line: line({ width: 12700, fill: solidFill(scheme('tx1')) }),
        ...(cropped ? { srcRect: { l: 10000, t: 5000, r: 20000, b: 25000 } } : {}),
      }) +
      textBox({
        id: 11,
        name: 'Caption',
        x: CONTENT.x,
        y: CONTENT.y + 4206240,
        cx: 7315200,
        cy: 457200,
        paras: textLine(`Figure ${String(nth + 1)}. ${phrase(rng)}.`, { sz: 1200, i: true }),
      }) +
      textBox({
        id: 12,
        name: 'Aside',
        x: CONTENT.x + 7620000,
        y: CONTENT.y,
        cx: 3657600,
        cy: CONTENT.cy,
        paras:
          para({ content: run(sentence(rng), { sz: 1600 }) }) +
          para({ content: run(sentence(rng), { sz: 1600 }) }),
      }) +
      slideNumber(index),
  };
}

const PRESETS = [
  'rect',
  'roundRect',
  'ellipse',
  'triangle',
  'diamond',
  'hexagon',
  'star5',
  'rightArrow',
  'chevron',
  'cloud',
  'heart',
  'octagon',
] as const;

function shapeGrid(index: number, nth: number): ProbeSlide {
  const rng = lcg(seedOf(index));
  const chapter = chapterOf(index);
  const cell = grid(4, 3, 91440);
  const label = (text: string): string =>
    txBody({
      bodyPr: bodyPr({ wrap: 'square', anchor: 'ctr' }),
      paras: textLine(text, { sz: 1200, fill: solidFill(scheme('tx1')) }, { algn: 'ctr' }),
    });
  const fillOf = (k: number): string => {
    switch ((k + nth) % 4) {
      case 0:
        return solidFill(scheme(accent(k), '<a:lumMod val="60000"/><a:lumOff val="40000"/>'));
      case 1:
        return gradLin(scheme(accent(k), '<a:tint val="40000"/>'), scheme(accent(k)));
      case 2:
        return gradPath(scheme('lt1'), scheme(accent(k), '<a:tint val="60000"/>'));
      default:
        return patt(['pct25', 'dkUpDiag', 'smGrid'][k % 3]!, scheme(accent(k)), scheme('lt1'));
    }
  };
  const lineOf = (k: number): string | undefined => {
    switch ((k + nth) % 4) {
      case 0:
        return line({ width: 12700, fill: solidFill(scheme('tx1')) });
      case 1:
        return line({
          width: 28575,
          fill: solidFill(scheme(accent(k), '<a:shade val="50000"/>')),
          dash: '<a:prstDash val="dash"/>',
        });
      case 2:
        return line({
          width: 9525,
          fill: solidFill(scheme('tx1')),
          dash: '<a:prstDash val="sysDot"/>',
        });
      default:
        return undefined;
    }
  };
  const shapes: string[] = [];
  for (let k = 0; k < 11; k++) {
    const c = cell(k);
    const preset = PRESETS[(k + nth) % PRESETS.length]!;
    const outline = lineOf(k);
    shapes.push(
      shape({
        id: 10 + k,
        name: `${preset} ${String(k + 1)}`,
        x: c.x,
        y: c.y,
        cx: c.cx,
        cy: c.cy,
        geometry: prstGeom(preset),
        fill: fillOf(k),
        ...(outline === undefined ? {} : { line: outline }),
        ...(k % 3 === 2 ? { rotation: 900000 * ((k + nth) % 4) } : {}),
        ...(k === 7 ? { flipH: true } : {}),
        ...(k === 1 ? { effects: SHADOW } : k === 5 ? { effects: GLOW } : {}),
        textBody: label(`${figure(rng)} ${pick(rng, ['YoY', 'QoQ', 'target', 'actual'])}`),
      }),
    );
  }
  // The twelfth cell is a group of three, turned as one.
  const g = cell(11);
  const third = Math.floor(g.cx / 3);
  shapes.push(
    group({
      id: 30,
      name: 'Grouped trio',
      x: g.x,
      y: g.y,
      cx: g.cx,
      cy: g.cy,
      rotation: 1200000,
      children: [0, 1, 2]
        .map((k) =>
          shape({
            id: 31 + k,
            name: `Leaf ${String(k + 1)}`,
            x: g.x + k * third,
            y: g.y + (k % 2) * Math.floor(g.cy / 4),
            cx: third - 45720,
            cy: Math.floor((g.cy * 3) / 4),
            geometry: prstGeom(['ellipse', 'roundRect', 'diamond'][k]!),
            fill: solidFill(scheme(accent(k + nth))),
            caption: false,
          }),
        )
        .join(''),
    }),
  );
  // A connector with a head, from the first cell to the one beneath it.
  const from = cell(0);
  const to = cell(4);
  shapes.push(
    connector({
      id: 40,
      name: 'Straight Arrow Connector 40',
      x: from.x + Math.floor(from.cx / 2),
      y: from.y + from.cy,
      cx: 0,
      cy: to.y - (from.y + from.cy),
      geometry: prstGeom('straightConnector1'),
      line: line({
        width: 19050,
        fill: solidFill(scheme('tx1')),
        tailEnd: '<a:tailEnd type="triangle"/>',
      }),
    }),
  );
  return {
    title: slideTitle(chapter, rng),
    layout: LAYOUT.titleOnly,
    body: shapes.join('') + slideNumber(index),
  };
}

function quote(index: number): ProbeSlide {
  const rng = lcg(seedOf(index));
  const who = name(rng);
  return {
    title: 'Quote',
    layout: LAYOUT.blank,
    body:
      connector({
        id: 10,
        name: 'Rule',
        x: 1828800,
        y: 1600200,
        cx: 8534400,
        cy: 0,
        geometry: prstGeom('line'),
        line: line({ width: 25400, fill: solidFill(scheme('accent1')) }),
      }) +
      textBox({
        id: 11,
        name: 'Quotation',
        x: 1371600,
        y: 1828800,
        cx: 9448800,
        cy: 2743200,
        body: { anchor: 'ctr', lIns: 360000, tIns: 360000, rIns: 360000, bIns: 360000 },
        paras: textLine(
          `“${sentence(rng)} ${sentence(rng)}”`,
          { sz: 2800, i: true },
          { algn: 'ctr' },
        ),
      }) +
      textBox({
        id: 12,
        name: 'Attribution',
        x: 1371600,
        y: 4663440,
        cx: 9448800,
        cy: 640080,
        body: { anchor: 'b' },
        paras: textLine(
          `${who}, ${role(rng)}`,
          { sz: 1400, fill: solidFill(scheme('tx2')) },
          { algn: 'r' },
        ),
      }) +
      slideNumber(index),
  };
}

function barDiagram(index: number): ProbeSlide {
  const rng = lcg(seedOf(index));
  const chapter = chapterOf(index);
  const bars = 8;
  const slot = Math.floor(CONTENT.cx / bars);
  const barWidth = Math.floor(slot * 0.6);
  const baseY = CONTENT.y + 4114800;
  const shapes: string[] = [];
  for (let k = 0; k < bars; k++) {
    const height = between(rng, 6, 38) * 91440;
    const x = CONTENT.x + k * slot + Math.floor((slot - barWidth) / 2);
    shapes.push(
      shape({
        id: 10 + k,
        name: `Bar ${String(k + 1)}`,
        x,
        y: baseY - height,
        cx: barWidth,
        cy: height,
        fill: gradLin(scheme('accent1'), scheme('accent1', '<a:shade val="60000"/>')),
        caption: false,
      }),
      textBox({
        id: 20 + k,
        name: `Value ${String(k + 1)}`,
        x,
        y: baseY - height - 365760,
        cx: barWidth,
        cy: 320040,
        paras: textLine(figure(rng), { sz: 1000 }, { algn: 'ctr' }),
      }),
      textBox({
        id: 30 + k,
        name: `Label ${String(k + 1)}`,
        x,
        y: baseY + 45720,
        cx: barWidth,
        cy: 320040,
        paras: textLine(
          `Q${String((k % 4) + 1)} ${String(25 + Math.floor(k / 4))}`,
          { sz: 1000 },
          { algn: 'ctr' },
        ),
      }),
    );
  }
  shapes.push(
    connector({
      id: 40,
      name: 'Baseline',
      x: CONTENT.x,
      y: baseY,
      cx: CONTENT.cx,
      cy: 0,
      geometry: prstGeom('line'),
      line: line({ width: 12700, fill: solidFill(scheme('tx1')) }),
    }),
    textBox({
      id: 41,
      name: 'Axis title',
      x: CONTENT.x,
      y: baseY + 411480,
      cx: CONTENT.cx,
      cy: 365760,
      paras: textLine(`${phrase(rng)}, by quarter`, { sz: 1200, i: true }, { algn: 'ctr' }),
    }),
  );
  return {
    title: slideTitle(chapter, rng),
    layout: LAYOUT.titleOnly,
    body: shapes.join('') + slideNumber(index),
  };
}

function closing(): ProbeSlide {
  return {
    title: 'Thank you',
    layout: LAYOUT.blank,
    background: bgPr(solidFill(scheme('dk2'))),
    body:
      textBox({
        id: 10,
        name: 'Thank you',
        x: 9144000,
        y: 914400,
        cx: 1828800,
        cy: 5029200,
        body: { vert: 'vert270', anchor: 'ctr' },
        paras: textLine(
          'Thank you',
          { sz: 4000, b: true, fill: solidFill(scheme('lt1')) },
          { algn: 'ctr' },
        ),
      }) +
      textBox({
        id: 11,
        name: 'Contact',
        x: 914400,
        y: 5486400,
        cx: 7315200,
        cy: 457200,
        paras: textLine('northwind.example — analytics@northwind.example', {
          sz: 1400,
          fill: solidFill(scheme('lt1')),
        }),
      }),
  };
}

/** The hundred slides, built in order. */
export function hundredSlides(): ProbeSlide[] {
  const seen: Partial<Record<Kind, number>> = {};
  const slides: ProbeSlide[] = [];
  for (let index = 0; index < 100; index++) {
    const kind = kindAt(index);
    const nth = seen[kind] ?? 0;
    seen[kind] = nth + 1;
    switch (kind) {
      case 'cover':
        slides.push(cover());
        break;
      case 'agenda':
        slides.push(agenda(index));
        break;
      case 'section':
        slides.push(section(index));
        break;
      case 'bullets':
        slides.push(bullets(index, nth));
        break;
      case 'twoColumn':
        slides.push(twoColumn(index));
        break;
      case 'pictureCaption':
        slides.push(pictureCaption(index, nth));
        break;
      case 'shapeGrid':
        slides.push(shapeGrid(index, nth));
        break;
      case 'quote':
        slides.push(quote(index));
        break;
      case 'barDiagram':
        slides.push(barDiagram(index));
        break;
      case 'closing':
        slides.push(closing());
        break;
    }
  }
  return slides;
}
