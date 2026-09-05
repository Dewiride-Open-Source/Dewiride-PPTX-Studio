import { describe, expect, it } from 'vitest';

import fixture from '../../../../corpus/ground-truth/frames.json' with { type: 'json' };
import { TextError } from '../errors.js';
import { lastLineHeight, MEASURED_FACE_METRICS } from '../lines/autofit.js';
import { blockHeight, lineAdvance } from '../lines/line-model.js';
import {
  anchorFraction,
  blockOrigin,
  columnBox,
  contentBox,
  DEFAULT_INSETS,
  drawnLines,
  ELLIPSIS,
  emptyParagraphHeight,
  frameAxes,
  MAX_COLUMNS,
  offsetAlong,
  VERTICAL_AXES,
  type Anchor,
  type Insets,
  type VerticalText,
} from './frame.js';

/* -------------------------------------------------------------------------- */
/* the fixture                                                                */
/* -------------------------------------------------------------------------- */

interface Measured {
  readonly dLeft: number | null;
  readonly dTop: number | null;
  readonly blockW: number | null;
  readonly blockH: number | null;
  readonly frame?: readonly (number | null)[];
  readonly margins?: readonly (number | null)[];
  readonly lineLefts: number | readonly (number | null)[];
  readonly lineTops: number | readonly (number | null)[];
  readonly paraTops: number | readonly (number | null)[];
  readonly paraLefts: number | readonly (number | null)[];
}

interface Probe {
  readonly id: string;
  readonly family: string;
  readonly vars: Readonly<Record<string, string | number | boolean | null>>;
  readonly frame: Readonly<Record<string, unknown>>;
  readonly box: readonly [number, number];
  readonly repaired?: boolean;
  readonly measured: Measured | null;
  readonly drew?: readonly string[];
  readonly verticalFace?: boolean;
}

const probes = fixture.probes as unknown as readonly Probe[];
const questions = fixture.questions as unknown as readonly {
  id: string;
  rows: number;
  winner: string;
  candidates: { name: string; score: number }[];
}[];

/** Thousandths of a point back to points. */
const pt = (v: number | null | undefined): number => (v ?? 0) / 1000;

const unpack = (v: number | readonly (number | null)[], n: number): number[] =>
  typeof v === 'number' ? Array.from({ length: n }, () => v / 1000) : v.map((x) => pt(x));

const inFamily = (family: string): readonly Probe[] =>
  probes.filter((p) => p.family === family && p.repaired !== true && p.measured !== null);

const byId = (id: string): Probe => {
  const probe = probes.find((p) => p.id === id);
  if (probe === undefined) throw new Error(`no probe ${id}`);
  return probe;
};

const measuredOf = (probe: Probe): Measured => {
  if (probe.measured === null) throw new Error(`${probe.id} has no reading`);
  return probe.measured;
};

/** The probe table states insets in points, and `null` means the attribute is absent. */
const insetsOf = (probe: Probe): Insets => {
  const edge = (name: string, fallback: number): number => {
    const stated = probe.frame[name];
    return stated === undefined || stated === null ? fallback : Number(stated);
  };
  return {
    left: edge('lIns', DEFAULT_INSETS.left),
    top: edge('tIns', DEFAULT_INSETS.top),
    right: edge('rIns', DEFAULT_INSETS.right),
    bottom: edge('bIns', DEFAULT_INSETS.bottom),
  };
};

/** What the analysis found, so a test names the score it is defending. */
const scoreOf = (id: string): string => {
  const q = questions.find((x) => x.id === id);
  if (q === undefined) throw new Error(`no question ${id}`);
  const rivals = q.candidates.filter((c) => c.name !== q.winner).map((c) => String(c.score));
  return `${String(q.rows)}/${String(q.rows)}, rivals ${rivals.join(' ')}`;
};

describe('the fixture', () => {
  it('holds every probe and every question T6 scored', () => {
    expect(probes.length).toBe(601);
    expect(questions.length).toBe(21);
    for (const q of questions) {
      const winner = q.candidates.find((c) => c.name === q.winner);
      expect(winner?.score, q.id).toBe(q.rows);
    }
  });

  it('names the three packages PowerPoint refused to read as written', () => {
    expect([...fixture.repairedDecks].sort((a, b) => a.localeCompare(b))).toEqual([
      'col-0',
      'col-17',
      'col-negspc',
    ]);
  });
});

/* -------------------------------------------------------------------------- */
/* insets                                                                     */
/* -------------------------------------------------------------------------- */

describe('the insets', () => {
  it(`defaults to 7.2/3.6/7.2/3.6 points, per attribute (${scoreOf('inset-applies')})`, () => {
    expect(DEFAULT_INSETS).toEqual({ left: 7.2, top: 3.6, right: 7.2, bottom: 3.6 });
    const absent = byId('ins-absent-all');
    expect(pt(measuredOf(absent).dLeft)).toBeCloseTo(DEFAULT_INSETS.left, 3);
    expect(pt(measuredOf(absent).dTop)).toBeCloseTo(DEFAULT_INSETS.top, 3);
  });

  it('defaults each edge on its own, not the element as a whole', () => {
    const onlyLeft = byId('ins-absent-l');
    expect(pt(measuredOf(onlyLeft).dLeft)).toBeCloseTo(DEFAULT_INSETS.left, 3);
    expect(pt(measuredOf(onlyLeft).dTop)).toBeCloseTo(0, 3);
  });

  it('puts the block corner at the stated inset, at every value measured', () => {
    for (const probe of inFamily('inset')) {
      const edge = String(probe.vars['edge']);
      if (edge !== 'l' && edge !== 't') continue;
      const insets = insetsOf(probe);
      const box = contentBox({ widthPt: probe.box[0], heightPt: probe.box[1] }, insets);
      const m = measuredOf(probe);
      expect(box.leftPt, probe.id).toBeCloseTo(pt(m.dLeft), 2);
      expect(box.topPt, probe.id).toBeCloseTo(pt(m.dTop), 2);
    }
  });

  it('honours a negative inset rather than clamping it to zero', () => {
    const probe = byId('ins-neg-l');
    expect(probe.repaired).not.toBe(true);
    expect(pt(measuredOf(probe).dLeft)).toBeCloseTo(-18, 2);
    const box = contentBox({ widthPt: probe.box[0], heightPt: probe.box[1] }, insetsOf(probe));
    expect(box.leftPt).toBeCloseTo(-18, 2);
  });

  it(`collapses to the midpoint of the two edges when they exceed the frame (${scoreOf('inset-collapse')})`, () => {
    for (const id of ['ins-huge-lr', 'ins-huge-asym-lr']) {
      const probe = byId(id);
      const box = contentBox({ widthPt: probe.box[0], heightPt: probe.box[1] }, insetsOf(probe));
      expect(box.leftPt, id).toBeCloseTo(pt(measuredOf(probe).dLeft), 2);
      expect(box.widthPt, id).toBe(0);
    }
  });

  it('is not the middle of the frame, which the symmetric probe alone cannot tell apart', () => {
    const probe = byId('ins-huge-asym-lr');
    const box = contentBox({ widthPt: probe.box[0], heightPt: probe.box[1] }, insetsOf(probe));
    expect(box.leftPt).not.toBeCloseTo(probe.box[0] / 2, 1);
    expect(box.leftPt).toBeCloseTo(185, 2);
  });
});

/* -------------------------------------------------------------------------- */
/* anchoring                                                                  */
/* -------------------------------------------------------------------------- */

describe('@anchor', () => {
  it(`puts the block inside the inset box, not the frame (${scoreOf('anchor')})`, () => {
    for (const probe of [...inFamily('anchor'), ...inFamily('anchor-spacing')]) {
      const m = measuredOf(probe);
      const insets = insetsOf(probe);
      const box = contentBox({ widthPt: probe.box[0], heightPt: probe.box[1] }, insets);
      const origin = blockOrigin({
        content: box,
        axes: frameAxes(),
        anchor: probe.vars['anchor'] as Anchor,
        blockAcrossPt: pt(m.blockH),
        blockAlongPt: 0,
      });
      expect(origin.topPt, probe.id).toBeCloseTo(pt(m.dTop), 2);
    }
  });

  it('would be wrong on every skewed-inset probe if the insets were ignored', () => {
    const wrong = inFamily('anchor')
      .filter((p) => p.vars['insets'] === 'skew')
      .filter((p) => {
        const m = measuredOf(p);
        const naive = anchorFraction(p.vars['anchor'] as Anchor) * (p.box[1] - pt(m.blockH));
        return Math.abs(naive - pt(m.dTop)) > 0.25;
      });
    expect(wrong.length).toBe(70);
  });

  it(`lays just and dist out as bottom, stretching nothing (${scoreOf('just-and-dist-stretch')})`, () => {
    expect(anchorFraction('just')).toBe(anchorFraction('b'));
    expect(anchorFraction('dist')).toBe(anchorFraction('b'));
    for (const probe of inFamily('anchor')) {
      const anchor = String(probe.vars['anchor']);
      if (anchor !== 'just' && anchor !== 'dist') continue;
      const twin = byId(probe.id.replace(/^anc-(just|dist)-/, 'anc-b-'));
      const m = measuredOf(probe);
      const t = measuredOf(twin);
      expect(pt(m.dTop), probe.id).toBeCloseTo(pt(t.dTop), 3);
      const lines = unpack(m.lineTops, 6);
      expect(lines, probe.id).toEqual(unpack(t.lineTops, 6));
    }
  });

  it('rejects an anchor outside ST_TextAnchoringType', () => {
    expect(() => anchorFraction('middle' as Anchor)).toThrow(TextError);
  });
});

describe('@anchorCtr', () => {
  it(`centres the block along the line axis; absent is off (${scoreOf('anchor-ctr')})`, () => {
    const off = measuredOf(byId('actr-0-l-square'));
    const absent = measuredOf(byId('actr-absent-l-square'));
    const on = measuredOf(byId('actr-1-l-square'));
    expect(unpack(absent.lineLefts, 3)).toEqual(unpack(off.lineLefts, 3));
    expect(unpack(on.lineLefts, 3)).not.toEqual(unpack(off.lineLefts, 3));
  });

  it('centres the widest line, leaving the others aligned inside it', () => {
    const on = measuredOf(byId('actr-1-r-square'));
    const lefts = unpack(on.lineLefts, 3);
    const widest = Math.min(...lefts);
    expect(widest).toBeCloseTo(53, 1);
    // Right-aligned inside the block, so the short lines sit further right.
    expect(lefts[0]).toBeGreaterThan(widest);
    expect(lefts[2]).toBeGreaterThan(widest);
  });

  it('is not the same as centring each line, which the equal-width case hides', () => {
    const anchored = unpack(measuredOf(byId('actr-1-l-square')).lineLefts, 3);
    const centred = unpack(measuredOf(byId('actr-0-ctr-square')).lineLefts, 3);
    expect(anchored).not.toEqual(centred);
  });

  it('moves the block origin along the line axis, and only when it is on', () => {
    const probe = byId('actr-1-l-square');
    const content = contentBox({ widthPt: probe.box[0], heightPt: probe.box[1] }, insetsOf(probe));
    const widest = Math.min(...unpack(measuredOf(byId('actr-1-l-square')).lineLefts, 3));
    const blockAlongPt = content.widthPt - 2 * widest;
    const on = blockOrigin({
      content,
      axes: frameAxes(),
      anchor: 't',
      anchorCtr: true,
      blockAcrossPt: 0,
      blockAlongPt,
    });
    const off = blockOrigin({
      content,
      axes: frameAxes(),
      anchor: 't',
      blockAcrossPt: 0,
      blockAlongPt,
    });
    expect(on.leftPt).toBeCloseTo(widest, 2);
    expect(off.leftPt).toBeCloseTo(content.leftPt, 6);
  });

  it('moves the block along the turned axis when the text is vertical', () => {
    const content = contentBox(
      { widthPt: 220, heightPt: 220 },
      { left: 0, top: 0, right: 0, bottom: 0 },
    );
    const vertical = blockOrigin({
      content,
      axes: frameAxes('vert'),
      anchor: 't',
      anchorCtr: true,
      blockAcrossPt: 21.6,
      blockAlongPt: 134,
    });
    expect(vertical.topPt).toBeCloseTo(43, 1);
    const measured = unpack(measuredOf(byId('vactr-vert-1')).lineTops, 2);
    expect(measured[0]).toBeCloseTo(43, 1);
  });
});

/* -------------------------------------------------------------------------- */
/* the block's height                                                         */
/* -------------------------------------------------------------------------- */

describe('the height of a block', () => {
  const arial = MEASURED_FACE_METRICS['Arial'];
  if (arial === undefined) throw new Error('Arial has no measured metrics');

  it(`keeps a trimmed last line unless spcFirstLastPara (${scoreOf('block-height')})`, () => {
    for (const probe of inFamily('last-line')) {
      const sz = Number(probe.vars['sz']);
      const lines = Number(probe.vars['lines']);
      const pct = Number(probe.vars['pct']);
      const flp = probe.vars['spcFirstLastPara'] === true;
      const advance = lineAdvance(sz, { kind: 'percent', value: pct });
      const predicted = blockHeight(lines, advance, lastLineHeight(advance, sz / 100, arial), flp);
      expect(predicted, probe.id).toBeCloseTo(pt(measuredOf(probe).blockH), 2);
    }
  });

  it('is the same either way at single spacing, which is why the sweep goes to 300%', () => {
    const advance = lineAdvance(1800);
    const trimmed = blockHeight(3, advance, lastLineHeight(advance, 18, arial), false);
    expect(trimmed).toBeCloseTo(blockHeight(3, advance, 0, true), 6);
    const wide = lineAdvance(1800, { kind: 'percent', value: 200000 });
    expect(blockHeight(3, wide, lastLineHeight(wide, 18, arial), false)).not.toBeCloseTo(
      blockHeight(3, wide, 0, true),
      1,
    );
  });

  it('counts a paragraph boundary as costing nothing on its own', () => {
    const brLines = measuredOf(byId('ctl-n3-sz1800')).blockH;
    const paragraphs = measuredOf(byId('ctl-paras3')).blockH;
    expect(paragraphs).toBe(brLines);
  });

  it('starts the block exactly on the frame edge', () => {
    for (const probe of inFamily('control')) {
      expect(pt(measuredOf(probe).dTop), probe.id).toBeCloseTo(0, 3);
      expect(pt(measuredOf(probe).dLeft), probe.id).toBeCloseTo(0, 3);
    }
  });
});

describe('an empty paragraph', () => {
  it(`is 1.2 x its a:endParaRPr size (${scoreOf('empty-paragraph-height')})`, () => {
    for (const probe of inFamily('end-para')) {
      if (probe.vars['where'] !== 'middle') continue;
      const kind = String(probe.vars['kind']);
      const stated = Number(probe.vars['endSz']);
      const sz = kind === 'absent' || kind === 'empty-run' ? 1800 : stated;
      const lnSpc = Number(probe.vars['lnSpc'] ?? 100000) / 100000;
      const tops = unpack(measuredOf(probe).paraTops, 3);
      const gap = (tops[2] ?? 0) - (tops[1] ?? 0);
      expect(emptyParagraphHeight(sz, lnSpc), probe.id).toBeCloseTo(gap, 2);
    }
  });

  it('takes a:endParaRPr over a:defRPr when the two disagree, both ways round', () => {
    const small = unpack(measuredOf(byId('end-conflict-800-4000')).paraTops, 3);
    const big = unpack(measuredOf(byId('end-conflict-4000-800')).paraTops, 3);
    expect((small[2] ?? 0) - (small[1] ?? 0)).toBeCloseTo(emptyParagraphHeight(800), 2);
    expect((big[2] ?? 0) - (big[1] ?? 0)).toBeCloseTo(emptyParagraphHeight(4000), 2);
  });

  it('rejects a size the format cannot hold', () => {
    expect(() => emptyParagraphHeight(0)).toThrow(TextError);
    expect(() => emptyParagraphHeight(1800, -1)).toThrow(TextError);
  });
});

/* -------------------------------------------------------------------------- */
/* vertical text                                                              */
/* -------------------------------------------------------------------------- */

describe('vertical text', () => {
  it(`swaps the two axes, each with its own start edge (${scoreOf('vert-anchor-axis')})`, () => {
    for (const probe of inFamily('vert-anchor')) {
      if (!probe.id.startsWith('vanc-')) continue;
      const m = measuredOf(probe);
      const axes = frameAxes(probe.vars['vert'] as VerticalText);
      const across = axes.stacking === 'x' ? probe.box[0] : probe.box[1];
      const thickness = axes.stacking === 'x' ? pt(m.blockW) : pt(m.blockH);
      const predicted = offsetAlong(
        axes.anchorFrom,
        anchorFraction(probe.vars['anchor'] as Anchor),
        across,
        thickness,
      );
      const measured = axes.stacking === 'x' ? pt(m.dLeft) : pt(m.dTop);
      expect(predicted, probe.id).toBeCloseTo(measured, 2);
    }
  });

  it('does not agree with itself: vert starts right, vert270 and mongolianVert left', () => {
    expect(VERTICAL_AXES.vert.anchorFrom).toBe('right');
    expect(VERTICAL_AXES.eaVert.anchorFrom).toBe('right');
    expect(VERTICAL_AXES.vert270.anchorFrom).toBe('left');
    expect(VERTICAL_AXES.mongolianVert.anchorFrom).toBe('left');
    expect(VERTICAL_AXES.vert270.alignFrom).toBe('bottom');
    expect(VERTICAL_AXES.mongolianVert.alignFrom).toBe('top');
  });

  it(`leaves the insets on their own physical edges (${scoreOf('vert-insets')})`, () => {
    const skew: Insets = { left: 13, top: 20, right: 3, bottom: 5 };
    const box = contentBox({ widthPt: 200, heightPt: 200 }, skew);
    for (const probe of inFamily('vert')) {
      if (!probe.id.startsWith('vert-ins-')) continue;
      const m = measuredOf(probe);
      const axes = frameAxes(probe.vars['vert'] as VerticalText);
      // The block's own start edge, whichever one the axes say it runs from.
      const measured = {
        left: pt(m.dLeft),
        top: pt(m.dTop),
        right: 200 - pt(m.dLeft) - pt(m.blockW),
        bottom: 200 - pt(m.dTop) - pt(m.blockH),
      };
      expect(measured[axes.anchorFrom], `${probe.id} anchor`).toBeCloseTo(skew[axes.anchorFrom], 2);
      expect(measured[axes.alignFrom], `${probe.id} algn`).toBeCloseTo(skew[axes.alignFrom], 2);
    }
    expect(box.leftPt).toBeCloseTo(skew.left, 3);
    expect(box.topPt).toBeCloseTo(skew.top, 3);
  });

  it(`asks for the @-prefixed face only for eaVert and mongolianVert (${scoreOf('vertical-face')})`, () => {
    for (const probe of inFamily('vert')) {
      if (probe.vars['script'] !== 'cjk' || probe.verticalFace === undefined) continue;
      const vert = String(probe.vars['vert']);
      const expected = vert === 'eaVert' || vert === 'mongolianVert';
      expect(probe.verticalFace, probe.id).toBe(expected);
      if (vert !== 'absent') expect(frameAxes(vert as VerticalText).verticalFace).toBe(expected);
    }
  });

  it('draws one record per character for the two wordArt types', () => {
    expect(frameAxes('wordArtVert').stackedGlyphs).toBe(true);
    expect(frameAxes('wordArtVertRtl').stackedGlyphs).toBe(true);
    expect(byId('vert-wordArtVert-latin-u0').drew).toEqual(['W', 'x', 'y', 'z']);
    expect(byId('vert-vert-latin-u0').drew).toEqual(['Wxyz']);
  });

  it('rejects a @vert outside ST_TextVerticalType', () => {
    expect(() => frameAxes('sideways' as VerticalText)).toThrow(TextError);
  });
});

/* -------------------------------------------------------------------------- */
/* overflow                                                                   */
/* -------------------------------------------------------------------------- */

describe('@vertOverflow', () => {
  const lines = Array.from({ length: 5 }, (_, i) => ({ topPt: i * 21.6, heightPt: 21.6 }));
  const available = 46;

  it(`draws everything when absent or overflow (${scoreOf('vert-overflow')})`, () => {
    expect(drawnLines(lines, available)).toEqual({ count: 5, ellipsis: false });
    expect(drawnLines(lines, available, 'overflow')).toEqual({ count: 5, ellipsis: false });
    expect(byId('ovf-v-absent-t').drew).toHaveLength(5);
  });

  it('replaces the whole first line that does not fit with a lone ellipsis', () => {
    expect(drawnLines(lines, available, 'ellipsis')).toEqual({ count: 2, ellipsis: true });
    expect(byId('ovf-v-ellipsis-t').drew).toEqual(['Line one', 'Line two', ELLIPSIS]);
  });

  it('does not append the ellipsis to the last line that fits', () => {
    const drew = byId('ovf-v-ellipsis-t').drew ?? [];
    expect(drew[1]).toBe('Line two');
    expect(drew[1]).not.toContain(ELLIPSIS);
  });

  it('stops laying out on clip rather than drawing and masking', () => {
    expect(drawnLines(lines, available, 'clip')).toEqual({ count: 2, ellipsis: false });
    expect(byId('ovf-v-clip-t').drew).toEqual(['Line one', 'Line two']);
  });

  it('drops a line whose bottom lands exactly on the edge', () => {
    expect(byId('ovf-fit-exact-clip').drew).toEqual(['Line one']);
    expect(byId('ovf-fit-over-clip').drew).toEqual(['Line one', 'Line two']);
    expect(drawnLines(lines, 43.2, 'clip')).toEqual({ count: 1, ellipsis: false });
    expect(drawnLines(lines, 44, 'clip')).toEqual({ count: 2, ellipsis: false });
  });

  it(`changes what is drawn and never the block (${scoreOf('overflow-layout')})`, () => {
    const heights = new Set(
      inFamily('overflow')
        .filter((p) => p.vars['axis'] === 'vert')
        .map((p) => measuredOf(p).blockH),
    );
    expect(heights.size).toBe(1);
  });

  it('rejects an overflow outside ST_TextVertOverflowType', () => {
    expect(() => drawnLines(lines, available, 'scroll' as never)).toThrow(TextError);
  });
});

/* -------------------------------------------------------------------------- */
/* a:br                                                                       */
/* -------------------------------------------------------------------------- */

describe('a:br', () => {
  it(`starts no bullet and advances no autonumber (${scoreOf('break-bullets')})`, () => {
    expect((byId('br-char-br').drew ?? []).filter((t) => t === '•')).toHaveLength(1);
    expect((byId('br-char-2p').drew ?? []).filter((t) => t === '•')).toHaveLength(2);
    expect((byId('br-num-br3').drew ?? []).filter((t) => /^\d+\.$/.test(t))).toHaveLength(1);
    expect(byId('br-num-2p').drew).toContain('2.');
    expect(byId('br-num-br').drew).not.toContain('2.');
  });

  it(`pays no spcBef or spcAft (${scoreOf('break-spacing')})`, () => {
    expect(measuredOf(byId('br-spc-br')).blockH).toBe(measuredOf(byId('br-brsz-none')).blockH);
    expect(pt(measuredOf(byId('br-spc-2p')).blockH)).toBeCloseTo(73.2, 2);
  });

  it(`does not reset the first-line indent (${scoreOf('break-indent')})`, () => {
    const positive = unpack(measuredOf(byId('br-indent-pos-br')).lineLefts, 2);
    expect(positive[0]).toBeCloseTo(56, 2);
    expect(positive[1]).toBeCloseTo(20, 2);
  });

  it('takes the height of the line from the following run, not from its own a:rPr', () => {
    const heights = ['br-brsz-small', 'br-brsz-big', 'br-brsz-none'].map(
      (id) => measuredOf(byId(id)).blockH,
    );
    expect(new Set(heights).size).toBe(1);
  });
});

/* -------------------------------------------------------------------------- */
/* columns                                                                    */
/* -------------------------------------------------------------------------- */

describe('columns', () => {
  it(`are (frame - (n-1) x spcCol) / n wide (${scoreOf('column-pitch')})`, () => {
    for (const probe of inFamily('columns')) {
      const numCol = Number(probe.vars['numCol']);
      if (!Number.isInteger(numCol) || numCol < 2) continue;
      const spcCol = probe.vars['spcCol'] === 'absent' ? 0 : Number(probe.vars['spcCol']);
      const rtlCol = probe.vars['rtlCol'] === '1';
      const content = contentBox(
        { widthPt: probe.box[0], heightPt: probe.box[1] },
        insetsOf(probe),
      );
      const lefts = unpack(measuredOf(probe).paraLefts, 8);
      const columns = [...new Set(lefts.map((x) => Math.round(x * 1000) / 1000))];
      expect(columns.length, `${probe.id} filled one column`).toBeGreaterThan(1);
      columns.forEach((left, index) => {
        const box = columnBox({ content, index, numCol, spcColPt: spcCol, rtlCol });
        expect(box.leftPt, `${probe.id} column ${String(index)}`).toBeCloseTo(left, 2);
      });
    }
  });

  it('fills from the right when rtlCol is on', () => {
    const content = contentBox(
      { widthPt: 300, heightPt: 110 },
      { left: 0, top: 0, right: 0, bottom: 0 },
    );
    expect(columnBox({ content, index: 0, numCol: 3, rtlCol: true }).leftPt).toBeCloseTo(200, 3);
    expect(columnBox({ content, index: 0, numCol: 3 }).leftPt).toBeCloseTo(0, 3);
  });

  it('refuses the counts PowerPoint repairs', () => {
    const content = contentBox({ widthPt: 300, heightPt: 110 }, DEFAULT_INSETS);
    expect(() => columnBox({ content, index: 0, numCol: 0 })).toThrow(TextError);
    expect(() => columnBox({ content, index: 0, numCol: MAX_COLUMNS + 1 })).toThrow(TextError);
    expect(() => columnBox({ content, index: 0, numCol: 2, spcColPt: -18 })).toThrow(TextError);
    expect(MAX_COLUMNS).toBe(16);
  });

  it('records that PowerPoint repaired each of those three packages', () => {
    for (const id of ['col-0', 'col-17', 'col-negspc']) {
      expect(byId(id).repaired, id).toBe(true);
    }
  });
});
