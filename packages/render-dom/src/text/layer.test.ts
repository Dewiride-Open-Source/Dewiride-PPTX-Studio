import {
  layoutText,
  textNodes,
  type ResolvedFrame,
  type ResolvedParagraph,
  type ResolvedRun,
  type ResolvedText,
  type TextBlock,
} from '@pptx-studio/render-svg';
import { afterEach, describe, expect, it } from 'vitest';

import { RenderDomError } from '../errors.js';
import { createNode } from '../mount.js';

import { mountTextLayer, type LayerBlock } from './layer.js';

/* -------------------------------------------------------------------------- */
/* one layout, two renderers                                                  */
/* -------------------------------------------------------------------------- */

const FRAME: ResolvedFrame = {
  anchor: 't',
  anchorCtr: false,
  vertical: 'horz',
  wrap: 'square',
  insets: { left: 0, top: 0, right: 0, bottom: 0 },
  columns: 1,
  columnSpacing: 0,
  rtlColumns: false,
  bodyRotation: 0,
  upright: false,
  vertOverflow: 'overflow',
  fontScale: 1,
  lineSpaceReduction: 0,
};

function run(text: string, over: Partial<ResolvedRun> = {}): ResolvedRun {
  return {
    text,
    font: { family: 'Arial', sz: 3200 },
    color: null,
    highlight: null,
    underline: undefined,
    strike: undefined,
    caps: undefined,
    baseline: 0,
    hardBreak: false,
    ...over,
  };
}

function paragraph(
  runs: readonly ResolvedRun[],
  over: Partial<ResolvedParagraph> = {},
): ResolvedParagraph {
  return {
    level: 0,
    align: 'l',
    marginLeft: 0,
    indent: 0,
    lineSpacing: { kind: 'percent', value: 100000 },
    spaceBefore: { kind: 'points', value: 0 },
    spaceAfter: { kind: 'points', value: 0 },
    runs,
    endRun: runs[0] ?? run(''),
    ...over,
  };
}

function laid(
  paragraphs: readonly ResolvedParagraph[],
  frame: Partial<ResolvedFrame> = {},
): TextBlock {
  const text: ResolvedText = { frame: { ...FRAME, ...frame }, paragraphs };
  return layoutText(text, {
    widthPt: 600,
    heightPt: 240,
    rot: 0,
    flipH: false,
    flipV: false,
    rulesFor: () => ({
      underline: { offset: 0.103, thickness: 0.075 },
      strike: { offset: -0.2595, thickness: 0.05 },
    }),
  });
}

function entry(block: TextBlock, over: Partial<LayerBlock> = {}): LayerBlock {
  return { block, leftPt: 120, topPt: 120, widthPt: 600, heightPt: 240, cNvPrId: 7, ...over };
}

const hosts: HTMLElement[] = [];

function host(): HTMLElement {
  const el = document.createElement('div');
  el.style.position = 'absolute';
  el.style.left = '0';
  el.style.top = '0';
  document.body.appendChild(el);
  hosts.push(el);
  return el;
}

afterEach(() => {
  for (const el of hosts.splice(0)) el.remove();
});

/* -------------------------------------------------------------------------- */

describe('mountTextLayer', () => {
  it('builds one element per shape and one per line', () => {
    const block = laid([paragraph([run('Alpha')]), paragraph([run('bravo')])]);
    const layer = mountTextLayer(host(), [entry(block)], { widthPt: 960, heightPt: 540 });
    const shape = layer.element(7);
    expect(shape).not.toBeNull();
    expect(shape?.children.length).toBe(2);
    expect(layer.element(8)).toBeNull();
  });

  it('is inert to the pointer unless asked otherwise', () => {
    const block = laid([paragraph([run('Alpha')])]);
    const el = host();
    expect(
      mountTextLayer(el, [entry(block)], { widthPt: 960, heightPt: 540 }).root.style.pointerEvents,
    ).toBe('none');
    expect(
      mountTextLayer(el, [entry(block)], { widthPt: 960, heightPt: 540 }, { interactive: true })
        .root.style.pointerEvents,
    ).toBe('auto');
  });

  it('zooms with one transform and never a re-layout', () => {
    const block = laid([paragraph([run('Alpha')])]);
    const layer = mountTextLayer(host(), [entry(block)], { widthPt: 960, heightPt: 540 });
    const before = layer.element(7)?.getBoundingClientRect().width ?? 0;
    layer.resize(2);
    expect(layer.root.style.transform).toBe('scale(2)');
    expect(layer.element(7)?.getBoundingClientRect().width).toBeCloseTo(before * 2, 6);
  });

  it('turns the shape rather than the glyphs', () => {
    const block = laid([paragraph([run('Alpha')])]);
    const flipped = layoutText(
      { frame: FRAME, paragraphs: [paragraph([run('Alpha')])] },
      {
        widthPt: 600,
        heightPt: 240,
        rot: 0,
        flipH: true,
        flipV: true,
        rulesFor: () => ({
          underline: { offset: 0.103, thickness: 0.075 },
          strike: { offset: -0.2595, thickness: 0.05 },
        }),
      },
    );
    const layer = mountTextLayer(host(), [entry(block), entry(flipped, { cNvPrId: 8 })], {
      widthPt: 960,
      heightPt: 540,
    });
    expect(layer.element(7)?.style.transform).toBe('');
    expect(layer.element(8)?.style.transform).toBe('rotate(180deg)');
    expect(layer.element(8)?.style.transform).not.toContain('scale');
  });

  it('holds the text as real nodes, which is the point of the layer', () => {
    const block = laid([paragraph([run('Alpha bravo')])]);
    const layer = mountTextLayer(host(), [entry(block)], { widthPt: 960, heightPt: 540 });
    expect(layer.element(7)?.textContent).toBe('Alpha bravo');
    const range = document.createRange();
    range.selectNodeContents(layer.element(7) as Node);
    expect(range.toString()).toBe('Alpha bravo');
  });

  it('lifts a superscript and lowers a subscript', () => {
    const block = laid([
      paragraph([run('x'), run('2', { baseline: 30000 }), run('y', { baseline: -25000 })]),
    ]);
    const layer = mountTextLayer(host(), [entry(block)], { widthPt: 960, heightPt: 540 });
    const spans = [...(layer.element(7)?.firstElementChild?.children ?? [])] as HTMLElement[];
    expect(Number.parseFloat(spans[1]?.style.top ?? '0')).toBeLessThan(0);
    expect(Number.parseFloat(spans[2]?.style.top ?? '0')).toBeGreaterThan(0);
    expect(spans[0]?.style.top).toBe('');
  });

  it("writes a run's colour, rather than inheriting one", () => {
    const block = laid([paragraph([run('a', { color: { r: 192 / 255, g: 0, b: 0, a: 1 } })])]);
    const layer = mountTextLayer(host(), [entry(block)], { widthPt: 960, heightPt: 540 });
    const span = layer.element(7)?.firstElementChild?.firstElementChild as HTMLElement;
    expect(span.style.color).toBe('rgb(192, 0, 0)');
  });

  it('unmounts without leaving anything behind', () => {
    const el = host();
    const layer = mountTextLayer(el, [entry(laid([paragraph([run('a')])]))], {
      widthPt: 960,
      heightPt: 540,
    });
    layer.unmount();
    expect(el.children.length).toBe(0);
  });

  it('refuses a host that is not an element', () => {
    expect(() =>
      mountTextLayer(null as unknown as Element, [], { widthPt: 960, heightPt: 540 }),
    ).toThrow(RenderDomError);
  });

  it('draws a rule as its own element, not as text-decoration', () => {
    // Chromium's own `text-decoration` is 0 of 16 against PowerPoint on both
    // the offset and the thickness, so the rule is geometry.
    const block = laid([paragraph([run('Alpha', { underline: 'sng' })])]);
    const layer = mountTextLayer(host(), [entry(block)], { widthPt: 960, heightPt: 540 });
    const line = layer.element(7)?.firstElementChild as HTMLElement;
    const rule = line.lastElementChild as HTMLElement;
    expect(rule.tagName).toBe('DIV');
    expect(line.style.textDecoration).toBe('');
    expect(Number.parseFloat(rule.style.height)).toBeCloseTo(0.075 * 32, 3);
  });
});

/* -------------------------------------------------------------------------- */
/* the sub-phase's own verification                                            */
/* -------------------------------------------------------------------------- */

/**
 * Where the browser actually put a line's baseline, in CSS pixels.
 *
 * A zero-height inline block sits with its bottom margin edge on the baseline
 * and adds no advance, so appending one reads the baseline back without moving
 * anything.
 */
function browserBaseline(line: HTMLElement): number {
  const marker = document.createElement('span');
  marker.style.display = 'inline-block';
  marker.style.width = '0';
  marker.style.height = '0';
  line.appendChild(marker);
  const top = marker.getBoundingClientRect().top - line.getBoundingClientRect().top;
  marker.remove();
  return top;
}

/**
 * How far the HTML layer's baseline may sit from the layout's, in points.
 *
 * One CSS pixel, which is one point in this space. It is not zero because
 * Chromium rounds a font's ascent to a whole pixel before laying a line out.
 */
const BASELINE_DRIFT_PT = 1;

describe('the two renderers agree on line boxes', () => {
  const cases: readonly { name: string; paragraphs: readonly ResolvedParagraph[] }[] = [
    { name: 'one line', paragraphs: [paragraph([run('Alpha bravo')])] },
    {
      name: 'two paragraphs of different sizes',
      paragraphs: [
        paragraph([run('Alpha', { font: { family: 'Arial', sz: 1800 } })]),
        paragraph([run('bravo', { font: { family: 'Arial', sz: 4400 } })]),
      ],
    },
    {
      name: 'a line holding two sizes',
      paragraphs: [
        paragraph([
          run('Alpha', { font: { family: 'Arial', sz: 1800 } }),
          run('bravo', { font: { family: 'Arial', sz: 4400 } }),
        ]),
      ],
    },
    {
      name: 'a wrapped paragraph',
      paragraphs: [
        paragraph([run('Alpha bravo charlie delta echo foxtrot golf hotel india juliet kilo')]),
      ],
    },
    {
      name: 'a face that is not the default',
      paragraphs: [paragraph([run('Alpha', { font: { family: 'Georgia', sz: 3600 } })])],
    },
  ];

  it.each(cases)('$name', ({ paragraphs }) => {
    const block = laid(paragraphs);
    expect(block.lines.length).toBeGreaterThan(0);

    // The SVG renderer places the line once, at `x` and `y` on a `<text>`.
    const svg = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
    svg.append(
      createNode(document, textNodes(block, { x: 0, y: 0, cx: 7620000, cy: 3048000 })[0]!),
    );
    const texts = [...svg.querySelectorAll('text')];
    expect(texts.length).toBe(block.lines.length);

    const layer = mountTextLayer(host(), [entry(block, { leftPt: 0, topPt: 0 })], {
      widthPt: 960,
      heightPt: 540,
    });
    const shape = layer.element(7);
    const lines = [...(shape?.children ?? [])] as HTMLElement[];
    expect(lines.length).toBe(block.lines.length);

    block.lines.forEach((line, index) => {
      const svgLine = texts[index];
      expect(Number(svgLine?.getAttribute('y')), `svg line ${String(index)}`).toBeCloseTo(
        line.baselinePt,
        3,
      );
      expect(Number(svgLine?.getAttribute('x')), `svg line ${String(index)}`).toBeCloseTo(
        line.leftPt,
        3,
      );
      // One CSS pixel is one point here, and half of one is the tolerance: the
      // browser lays out on device pixels and the layout does not.
      const html = lines[index];
      expect(html, `html line ${String(index)}`).toBeDefined();
      const top =
        (html?.getBoundingClientRect().top ?? 0) - (shape?.getBoundingClientRect().top ?? 0);
      expect(top, `html top ${String(index)}`).toBeCloseTo(line.topPt, 1);
      // One CSS pixel is one point here. Chromium rounds a font's ascent to a
      // whole pixel before it lays a line out, so no `line-height` closes this
      // gap; the SVG emitter, checked exactly above, is the one that exports.
      const drift = Math.abs(browserBaseline(html as HTMLElement) + top - line.baselinePt);
      expect(drift, `html baseline ${String(index)}`).toBeLessThanOrEqual(BASELINE_DRIFT_PT);
    });
  });
});

/* -------------------------------------------------------------------------- */
/* upright East Asian glyphs, in both renderers                               */
/* -------------------------------------------------------------------------- */

describe('an eaVert frame, in the HTML layer', () => {
  const CJK = '日本語';

  function upright(vertical: ResolvedFrame['vertical']): TextBlock {
    return laid([paragraph([run(CJK, { font: { family: 'Yu Gothic', sz: 3200 } })])], {
      vertical,
      insets: { left: 0, top: 0, right: 0, bottom: 0 },
      wrap: 'none',
    });
  }

  it('draws one element per glyph, not one per piece', () => {
    const block = upright('eaVert');
    const layer = mountTextLayer(host(), [entry(block)], { widthPt: 960, heightPt: 540 });
    const line = layer.element(7)?.children[0];
    expect(line?.children.length).toBe([...CJK].length);
    expect([...(line?.children ?? [])].map((el) => el.textContent)).toStrictEqual([...CJK]);
  });

  it('turns each glyph a quarter back out of the line', () => {
    const layer = mountTextLayer(host(), [entry(upright('eaVert'))], {
      widthPt: 960,
      heightPt: 540,
    });
    const glyphs = [...(layer.element(7)?.children[0]?.children ?? [])] as HTMLElement[];
    for (const glyph of glyphs) expect(glyph.style.transform).toBe('rotate(-90deg)');
  });

  it('puts each glyph pen where the SVG emitter puts it', () => {
    const block = upright('eaVert');
    const glyphs = block.lines[0]?.pieces[0]?.upright ?? [];
    expect(glyphs.length).toBe([...CJK].length);
    const layer = mountTextLayer(host(), [entry(block)], { widthPt: 960, heightPt: 540 });
    const elements = [...(layer.element(7)?.children[0]?.children ?? [])] as HTMLElement[];
    glyphs.forEach((glyph, index) => {
      const el = elements[index];
      const size = (block.lines[0]?.pieces[0]?.font.sz ?? 0) / 100;
      expect(el?.style.left).toBe(`${String(Math.round(glyph.alongPt * 1000) / 1000)}px`);
      expect(el?.style.top).toBe(`${String(Math.round((glyph.acrossPt - size) * 1000) / 1000)}px`);
      expect(el?.style.transformOrigin).toBe(`0px ${String(size)}px`);
    });
  });

  it('leaves a vert frame flowing, one element per piece', () => {
    const layer = mountTextLayer(host(), [entry(upright('vert'))], {
      widthPt: 960,
      heightPt: 540,
    });
    const line = layer.element(7)?.children[0];
    expect(line?.children.length).toBe(1);
    expect(line?.children[0]?.textContent).toBe(CJK);
  });
});
