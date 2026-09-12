import { describe, expect, it } from 'vitest';

import { parseSheet, parseTheme, type Sheet } from '@pptx-studio/model';
import { turnedInsets } from '@pptx-studio/text';
import { parseXmlString } from '@pptx-studio/xml';

import fixture from '../../../../corpus/ground-truth/text-rendering.json' with { type: 'json' };
import frames from '../../../../corpus/ground-truth/frames.json' with { type: 'json' };
import columns from '../../../../corpus/ground-truth/wordart-columns.json' with { type: 'json' };
import { layoutSheet, type Placed } from '../layout.js';
import { RenderError } from '../errors.js';
import { serializeSvg, type SvgElement, type SvgNode } from '../node.js';

import type { FaceBox, RunFont } from '@pptx-studio/text';

import { createTextEngine, textBlockOf } from './draw.js';
import { textNodes } from './emit.js';
import {
  SHIFT_SIZE_RATIO,
  SMALL_CAPS_RATIO,
  alignOffset,
  capStretches,
  layoutText,
  stretches,
  strutHeight,
  textTurn,
  type TextBlock,
} from './layout.js';
import {
  resolveText,
  type ResolvedFrame,
  type ResolvedParagraph,
  type ResolvedRun,
  type ResolvedText,
} from './resolve.js';

/* -------------------------------------------------------------------------- */
/* the fixture                                                                */
/* -------------------------------------------------------------------------- */

const flipRows = fixture.powerpoint.flip.rows as readonly {
  probe: string;
  rot: number;
  flipH: boolean;
  flipV: boolean;
  extraTurn: number;
  mirrored: boolean;
}[];

const rotationRows = fixture.powerpoint.rotation as readonly {
  probe: string;
  rot: number;
  upright: boolean;
  bodyRot: number;
  angle: number;
}[];

const uprightRows = fixture.powerpoint.uprightBox as readonly {
  probe: string;
  rot: number;
  boxLeft: number;
}[];

const alignRows = fixture.powerpoint.align as readonly {
  probe: string;
  algn: string;
  line: number;
  lastLine: boolean;
  offset: number;
  width: number;
  naturalWidth: number;
  anchorWidth: number;
}[];

const baselineRows = fixture.powerpoint.baseline.rows as readonly {
  probe: string;
  face: string;
  size: number;
  sizes: readonly number[];
  lineHeight: number;
  drop: number;
}[];

const capsRows = fixture.powerpoint.caps as readonly {
  probe: string;
  cap: string;
  statedSize: number;
  drawn: readonly { text: string; size: number }[];
}[];

const shiftRows = fixture.powerpoint.shift as readonly {
  probe: string;
  percent: number;
  statedSize: number;
  drawnSize: number;
  rise: number;
  sizeRatio: number;
}[];

const runRows = fixture.powerpoint.runs as readonly {
  probe: string;
  records: number;
  advances: readonly number[];
}[];

/* -------------------------------------------------------------------------- */
/* literals, so a layout can be asserted without a package                    */
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
    // Thousandths of a percent, as `a:spcPct/@val` writes them.
    lineSpacing: { kind: 'percent', value: 100000 },
    spaceBefore: { kind: 'points', value: 0 },
    spaceAfter: { kind: 'points', value: 0 },
    runs,
    endRun: runs[0] ?? run(''),
    ...over,
  };
}

function body(
  paragraphs: readonly ResolvedParagraph[],
  frame: Partial<ResolvedFrame> = {},
): ResolvedText {
  return { frame: { ...FRAME, ...frame }, paragraphs };
}

/** A measurer that charges one point per character, so a width is countable. */
const flatMeasurer = {
  calls: [] as string[],
  measure(text: string, font: { sz: number }): { width: number } {
    this.calls.push(text);
    return { width: (text.length * font.sz) / 100 };
  },
};

/**
 * A face box with a share of exactly four fifths, for arithmetic anyone can do.
 *
 * The eighth of an em for the ideographic baseline is likewise a round number,
 * so an upright glyph's baseline lands at seven eighths of its cell.
 */
const flatFaceBox = {
  box: (): { ascent: number; descent: number; ideographic: number } => ({
    ascent: 0.8,
    descent: 0.2,
    ideographic: 0.125,
  }),
};

const flatRules = {
  underline: { offset: 0.1, thickness: 0.05 },
  strike: { offset: -0.25, thickness: 0.05 },
};

function laid(text: ResolvedText, over: Partial<Parameters<typeof layoutText>[1]> = {}): TextBlock {
  flatMeasurer.calls = [];
  return layoutText(text, {
    widthPt: 600,
    heightPt: 240,
    rot: 0,
    flipH: false,
    flipV: false,
    measurer: flatMeasurer,
    faceBox: flatFaceBox,
    rulesFor: () => flatRules,
    ...over,
  });
}

/* -------------------------------------------------------------------------- */
/* the turn                                                                   */
/* -------------------------------------------------------------------------- */

const TURN_FRAME = { widthPt: 400, heightPt: 160, rot: 0, flipH: false, flipV: false };

function turnOf(over: Partial<typeof TURN_FRAME>, frame: Partial<ResolvedFrame> = {}): number {
  return textTurn({ ...FRAME, ...frame }, { ...TURN_FRAME, ...over }).turnDeg;
}

describe('what a flip does to text', () => {
  it(`is what PowerPoint did on all ${String(flipRows.length)} probes`, () => {
    for (const row of flipRows) {
      const base = turnOf({ rot: row.rot });
      const mine = turnOf({ rot: row.rot, flipH: row.flipH, flipV: row.flipV });
      expect((mine - base + 360) % 360, row.probe).toBe(row.extraTurn);
    }
  });

  it('never mirrors the glyphs, which is what every row measured', () => {
    expect(flipRows.every((row) => !row.mirrored)).toBe(true);
  });

  it('is not the counter-flip on both axes the plan assumed', () => {
    // Counter-flipping both would leave flipV alone; PowerPoint turns it half
    // a revolution, and the two readings differ on ten of the twenty rows.
    const wrong = flipRows.filter(
      (row) => ((row.flipH ? 180 : 0) + (row.flipV ? 180 : 0)) % 360 !== row.extraTurn,
    );
    expect(wrong.length).toBe(10);
  });

  it('ignores flipH outright', () => {
    expect(turnOf({ flipH: true })).toBe(turnOf({}));
    expect(turnOf({ rot: 90, flipH: true })).toBe(turnOf({ rot: 90 }));
  });

  it('turns flipV half a revolution', () => {
    expect(turnOf({ flipV: true })).toBe(180);
    expect(turnOf({ rot: 90, flipV: true })).toBe(270);
    expect(turnOf({ rot: 270, flipV: true, flipH: true })).toBe(90);
  });
});

describe('what a rotation does to text', () => {
  it(`is what PowerPoint did on all ${String(rotationRows.length)} probes`, () => {
    for (const row of rotationRows) {
      const angle = turnOf({ rot: row.rot }, { upright: row.upright, bodyRotation: row.bodyRot });
      expect(angle, row.probe).toBe(row.angle);
    }
  });

  it('is not the sum whatever @upright says', () => {
    const sum = (row: { rot: number; bodyRot: number }): number =>
      (((row.rot + row.bodyRot) % 360) + 360) % 360;
    const wrong = rotationRows.filter((row) => sum(row) !== row.angle);
    expect(wrong.length).toBe(4);
    expect(wrong.every((row) => row.upright)).toBe(true);
  });

  it('adds a:bodyPr/@rot to the shape rotation', () => {
    expect(turnOf({ rot: 30 }, { bodyRotation: 45 })).toBe(75);
  });
});

describe('the box @upright lays text out in', () => {
  it(`is where PowerPoint put it on all ${String(uprightRows.length)} probes`, () => {
    for (const row of uprightRows) {
      const { boxPt } = textTurn({ ...FRAME, upright: true }, { ...TURN_FRAME, rot: row.rot });
      // The probes put a 400 x 160 frame at x=280, and the fixture records the
      // box's left edge on the slide.
      expect(280 + boxPt.leftPt, row.probe).toBeCloseTo(row.boxLeft, 3);
    }
  });

  it('swaps the extents only on a quadrant', () => {
    expect(textTurn({ ...FRAME, upright: true }, { ...TURN_FRAME, rot: 90 }).boxPt.widthPt).toBe(
      160,
    );
    expect(textTurn({ ...FRAME, upright: true }, { ...TURN_FRAME, rot: 30 }).boxPt.widthPt).toBe(
      400,
    );
    expect(textTurn({ ...FRAME, upright: true }, { ...TURN_FRAME, rot: 180 }).boxPt.widthPt).toBe(
      400,
    );
  });

  it('leaves the frame alone when the text is not upright', () => {
    expect(textTurn(FRAME, { ...TURN_FRAME, rot: 90 }).boxPt.widthPt).toBe(400);
  });
});

/* -------------------------------------------------------------------------- */
/* alignment                                                                  */
/* -------------------------------------------------------------------------- */

describe('where a line starts', () => {
  it(`is where PowerPoint put it on all ${String(alignRows.length)} rows`, () => {
    for (const row of alignRows) {
      const stretched = stretches(row.algn, row.lastLine);
      const offset = stretched ? 0 : alignOffset(row.algn, 600, row.anchorWidth);
      expect(Math.abs(offset - row.offset), row.probe).toBeLessThanOrEqual(0.01);
    }
  });

  it('does not count the spaces a line ends with', () => {
    const wrong = alignRows.filter((row) => {
      if (stretches(row.algn, row.lastLine)) return false;
      return Math.abs(alignOffset(row.algn, 600, row.naturalWidth) - row.offset) > 0.01;
    });
    expect(wrong.length).toBeGreaterThan(0);
  });

  it('leaves every justifying alignment at the left edge', () => {
    for (const algn of ['l', 'just', 'justLow', 'dist', 'thaiDist']) {
      expect(alignOffset(algn, 600, 100)).toBe(0);
    }
  });

  it('centres and right-aligns from the anchored width', () => {
    expect(alignOffset('ctr', 600, 100)).toBe(250);
    expect(alignOffset('r', 600, 100)).toBe(500);
  });
});

describe('which lines stretch', () => {
  it('is dist and only dist on a last line', () => {
    expect(stretches('dist', true)).toBe(true);
    expect(stretches('just', true)).toBe(false);
    expect(stretches('justLow', true)).toBe(false);
    expect(stretches('thaiDist', true)).toBe(false);
  });

  it('is every justifying alignment on a line that is not the last', () => {
    for (const algn of ['just', 'justLow', 'dist', 'thaiDist']) {
      expect(stretches(algn, false), algn).toBe(true);
    }
    expect(stretches('ctr', false)).toBe(false);
  });

  it('matches the width PowerPoint reported on every align row', () => {
    for (const row of alignRows) {
      const grew = Math.abs(row.width - row.naturalWidth) > 0.5;
      expect(grew, row.probe).toBe(stretches(row.algn, row.lastLine));
    }
  });
});

/* -------------------------------------------------------------------------- */
/* caps and the baseline shift                                                */
/* -------------------------------------------------------------------------- */

describe('a:rPr/@cap', () => {
  it('draws what PowerPoint drew, at the sizes it drew it', () => {
    for (const row of capsRows) {
      const stretchesOf = capStretches('Hamburg fox', row.cap);
      const drawn = stretchesOf
        .map((s) => s.text)
        .join('')
        .replace(/ /g, '');
      const measured = row.drawn
        .map((d) => d.text)
        .join('')
        .replace(/ /g, '');
      expect(drawn, row.probe).toBe(measured);
      for (const s of stretchesOf) {
        const size = s.sizeRatio * row.statedSize;
        expect(
          row.drawn.some((d) => Math.abs(d.size - size) <= 0.15),
          `${row.probe} at ${String(size)}pt`,
        ).toBe(true);
      }
    }
  });

  it('reduces only the letters that were lowercase', () => {
    expect(capStretches('Ham', 'small')).toEqual([
      { text: 'H', sizeRatio: 1 },
      { text: 'AM', sizeRatio: SMALL_CAPS_RATIO },
    ]);
  });

  it('leaves the size alone for all and for none', () => {
    expect(capStretches('Ham', 'all')).toEqual([{ text: 'HAM', sizeRatio: 1 }]);
    expect(capStretches('Ham', 'none')).toEqual([{ text: 'Ham', sizeRatio: 1 }]);
    expect(capStretches('Ham', undefined)).toEqual([{ text: 'Ham', sizeRatio: 1 }]);
  });

  it('is four fifths, not three quarters', () => {
    expect(SMALL_CAPS_RATIO).toBe(0.8);
  });
});

describe('a:rPr/@baseline', () => {
  it('draws a shifted run at two thirds of the size it asked for', () => {
    for (const row of shiftRows) {
      expect(Math.abs(row.sizeRatio - SHIFT_SIZE_RATIO), row.probe).toBeLessThan(0.01);
    }
  });

  it('rises by that percentage of the size the run asked for', () => {
    const block = laid(body([paragraph([run('xx'), run('xx', { baseline: 30000 })])]));
    const piece = block.lines[0]?.pieces[1];
    expect(piece?.risePt).toBeCloseTo(0.3 * 32, 6);
    expect(piece?.font.sz).toBeCloseTo(3200 * SHIFT_SIZE_RATIO, 6);
  });

  it('is not a fraction of the line box', () => {
    for (const row of shiftRows) {
      const asLineBox = (row.percent / 100) * 1.2 * row.statedSize;
      expect(Math.abs(asLineBox - row.rise), row.probe).toBeGreaterThan(0.2);
    }
  });
});

/* -------------------------------------------------------------------------- */
/* the layout                                                                 */
/* -------------------------------------------------------------------------- */

describe('layoutText', () => {
  it('gives the line box to the largest run, whatever the order', () => {
    const large = run('xxxx', { font: { family: 'Arial', sz: 4400 } });
    const small = run('xxxx', { font: { family: 'Arial', sz: 1800 } });
    for (const runs of [
      [small, large],
      [large, small],
      [small, large, small],
    ]) {
      expect(laid(body([paragraph(runs)])).lines[0]?.heightPt).toBeCloseTo(1.2 * 44, 6);
    }
  });

  it('puts the baseline at the face box share of the line box', () => {
    const line = laid(body([paragraph([run('xxxx')])])).lines[0];
    expect(line?.baselinePt).toBeCloseTo(1.2 * 32 * 0.8, 6);
  });

  it('stacks the next line one of this line boxes below', () => {
    const block = laid(
      body([
        paragraph([run('xxxx', { font: { family: 'Arial', sz: 1800 } })]),
        paragraph([run('xxxx', { font: { family: 'Arial', sz: 4400 } })]),
      ]),
    );
    const [first, second] = block.lines;
    expect((second?.topPt ?? 0) - (first?.topPt ?? 0)).toBeCloseTo(1.2 * 18, 6);
  });

  it('measures a line as one string, so a kern crosses a run boundary', () => {
    // Two runs with the same formatting are one drawing call in PowerPoint and
    // the same advances as the whole string - 6 of 6.
    const whole = runRows.find((row) => row.probe === 'runs-whole-unset');
    const split = runRows.find((row) => row.probe === 'runs-split-unset');
    expect(split?.advances).toEqual(whole?.advances);
    expect(split?.records).toBe(1);
  });

  it('carries a run boundary into the pieces without moving anything', () => {
    const one = laid(body([paragraph([run('AVAWA')])]));
    const two = laid(body([paragraph([run('AV'), run('AWA')])]));
    expect(two.lines[0]?.pieces.length).toBe(2);
    expect(two.lines[0]?.widthPt).toBeCloseTo(one.lines[0]?.widthPt ?? 0, 10);
    expect(two.lines[0]?.pieces[1]?.leftPt).toBeCloseTo(2 * 32, 10);
  });

  it('anchors a centred line without its trailing spaces', () => {
    const bare = laid(body([paragraph([run('ab')], { align: 'ctr' })]));
    const spaced = laid(body([paragraph([run('ab ')], { align: 'ctr' })]));
    expect(spaced.lines[0]?.leftPt).toBeCloseTo(bare.lines[0]?.leftPt ?? 0, 10);
  });

  it('stretches a justified line with word-spacing and not with positions', () => {
    const block = laid(body([paragraph([run('a b c')], { align: 'dist' })]));
    const line = block.lines[0];
    expect(line?.wordSpacingPt).toBeGreaterThan(0);
    expect(line?.pieces[0]?.leftPt).toBe(0);
  });

  it('leaves a justified line with no space alone', () => {
    const block = laid(body([paragraph([run('abc')], { align: 'dist' })]));
    expect(block.lines[0]?.wordSpacingPt).toBe(0);
  });

  it('never gives a justified line a negative word spacing', () => {
    // A line wider than its column has nothing to share out, and pulling the
    // words together would be a stretch in the other direction.
    const block = laid(
      body([paragraph([run('a b', { font: { family: 'Arial', sz: 90000 } })], { align: 'dist' })]),
    );
    const line = block.lines[0];
    const ink = (line?.pieces ?? []).reduce((sum, piece) => sum + piece.widthPt, 0);
    expect(ink).toBeGreaterThan(600);
    expect(line?.wordSpacingPt).toBe(0);
  });

  it('places a rule over the characters it covers and not the trailing space', () => {
    const block = laid(body([paragraph([run('ab ', { underline: 'sng' })])]));
    const rule = block.lines[0]?.pieces[0]?.rules[0];
    expect(rule?.leftPt).toBe(0);
    expect(rule?.widthPt).toBeCloseTo(2 * 32, 10);
    expect(rule?.topPt).toBeCloseTo(0.1 * 32, 10);
  });

  it('places a rule per word for u="words"', () => {
    const block = laid(body([paragraph([run('ab cd', { underline: 'words' })])]));
    const rules = block.lines[0]?.pieces[0]?.rules ?? [];
    expect(rules.length).toBe(2);
    expect(rules[1]?.leftPt).toBeCloseTo(3 * 32, 10);
  });

  it('draws every ST_TextVerticalType there is', () => {
    for (const vert of [
      'horz',
      'vert',
      'vert270',
      'wordArtVert',
      'eaVert',
      'mongolianVert',
      'wordArtVertRtl',
    ]) {
      const block = laid(
        body([paragraph([run('a')])], { vertical: vert as ResolvedFrame['vertical'] }),
      );
      expect(block.lines.length, vert).toBe(1);
    }
  });

  it('gives a vertical frame the swapped extents and a quarter turn', () => {
    const block = laid(body([paragraph([run('a')])], { vertical: 'vert' }));
    expect(block.turnDeg).toBe(90);
    expect(block.boxPt.widthPt).toBe(240);
  });

  it('anchors a block at the bottom when the frame says so', () => {
    const top = laid(body([paragraph([run('a')])], { anchor: 't' }));
    const bottom = laid(body([paragraph([run('a')])], { anchor: 'b' }));
    expect(bottom.lines[0]?.topPt).toBeCloseTo(240 - 1.2 * 32, 6);
    expect(top.lines[0]?.topPt).toBe(0);
  });

  it('applies the stored autofit scale rather than re-running the ladder', () => {
    const block = laid(body([paragraph([run('xxxx')])], { fontScale: 0.5 }));
    expect(block.lines[0]?.heightPt).toBeCloseTo(1.2 * 16, 6);
    expect(block.lines[0]?.pieces[0]?.font.sz).toBeCloseTo(1600, 6);
  });
});

describe('strutHeight', () => {
  it("lands a browser's baseline where the layout put it", () => {
    const face = { ascent: 0.9, descent: 0.25, ideographic: 0.12 };
    const size = 40;
    const drop = 1.2 * size * (face.ascent / (face.ascent + face.descent));
    const strut = strutHeight(drop, face, size);
    const browser = (strut - (face.ascent + face.descent) * size) / 2 + face.ascent * size;
    expect(browser).toBeCloseTo(drop, 10);
  });

  it('never asks for a negative height', () => {
    expect(strutHeight(0, { ascent: 0.9, descent: 0.2, ideographic: 0.12 }, 40)).toBe(0);
  });

  it('solves every baseline the fixture recorded', () => {
    const boxes = fixture.browser.faces as Record<string, { ascent: number; descent: number }>;
    for (const row of baselineRows) {
      const box = boxes[row.face];
      if (box === undefined) continue;
      const strut = strutHeight(row.drop, { ...box, ideographic: 0 }, row.size);
      const browser = (strut - (box.ascent + box.descent) * row.size) / 2 + box.ascent * row.size;
      expect(browser, row.probe).toBeCloseTo(row.drop, 6);
    }
  });
});

/* -------------------------------------------------------------------------- */
/* the resolver, against markup rather than a literal                         */
/* -------------------------------------------------------------------------- */

const NS =
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"';

const SP_TREE_HEAD =
  '<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
  '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>' +
  '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>';

/** A theme naming both collections, which every real slide can reach. */
const THEME_XML =
  `<a:theme ${NS} name="T"><a:themeElements><a:clrScheme name="T">` +
  [
    'dk1',
    'lt1',
    'dk2',
    'lt2',
    'accent1',
    'accent2',
    'accent3',
    'accent4',
    'accent5',
    'accent6',
    'hlink',
    'folHlink',
  ]
    .map((slot) => `<a:${slot}><a:srgbClr val="123456"/></a:${slot}>`)
    .join('') +
  '</a:clrScheme><a:fontScheme name="T">' +
  '<a:majorFont><a:latin typeface="Georgia"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>' +
  '<a:minorFont><a:latin typeface="Verdana"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont>' +
  '</a:fontScheme><a:fmtScheme name="T"><a:fillStyleLst/><a:lnStyleLst/><a:effectStyleLst/>' +
  '<a:bgFillStyleLst/></a:fmtScheme></a:themeElements></a:theme>';

const CLR_MAP =
  '<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" ' +
  'accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" ' +
  'folHlink="folHlink"/>';

/** The master a slide inherits from, which is where the theme hangs. */
function masterSheet(): Sheet {
  const xml = `<p:sldMaster ${NS}><p:cSld name="master">${SP_TREE_HEAD}</p:spTree></p:cSld>${CLR_MAP}</p:sldMaster>`;
  const theme = parseTheme(parseXmlString(THEME_XML).root, '/ppt/theme/theme1.xml');
  return { ...parseSheet(parseXmlString(xml).root, '/ppt/slideMaster1.xml'), parent: null, theme };
}

/** One slide holding one text shape, parsed the way a real package is. */
function slideWithText(txBody: string): Sheet {
  const xml =
    `<p:sld ${NS}><p:cSld name="slide">${SP_TREE_HEAD}` +
    '<p:sp><p:nvSpPr><p:cNvPr id="2" name="probe"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>' +
    '<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="7620000" cy="3048000"/></a:xfrm>' +
    '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>' +
    txBody +
    '</p:sp></p:spTree></p:cSld>' +
    '<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>';
  return {
    ...parseSheet(parseXmlString(xml).root, '/ppt/slide.xml'),
    parent: masterSheet(),
    theme: null,
  };
}

describe('resolveText', () => {
  it('reads a run through the cascade and lands its colour', () => {
    // The only test that asks the resolver a question about markup. Everything
    // else drives the layout from literals, which is how a wrong discriminant
    // on `Fill` painted every coloured run black without a test noticing.
    const placed = layoutSheet(
      slideWithText(
        '<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r>' +
          '<a:rPr lang="en-US" sz="2400" b="1" u="sng">' +
          '<a:solidFill><a:srgbClr val="C00000"/></a:solidFill>' +
          '<a:latin typeface="Georgia"/></a:rPr>' +
          '<a:t>Alpha</a:t></a:r></a:p></p:txBody>',
      ),
    )[0];
    expect(placed).toBeDefined();
    const resolved = resolveText(placed as Placed);
    const run = resolved?.paragraphs[0]?.runs[0];
    expect(run?.text).toBe('Alpha');
    expect(run?.font).toMatchObject({ family: 'Georgia', sz: 2400, bold: true });
    expect(run?.underline).toBe('sng');
    expect(run?.color?.r).toBeCloseTo(192 / 255, 6);
    expect(run?.color?.g).toBe(0);
    expect(run?.color?.b).toBe(0);
  });

  it('draws a run no source names a face for in the theme minor', () => {
    // Measured on six probes; ADR 0034. Before it, a paragraph past the last
    // level a master declares resolved to no face and threw at the CSS shorthand.
    const placed = layoutSheet(
      slideWithText(
        '<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Alpha</a:t></a:r></a:p></p:txBody>',
      ),
    )[0];
    expect(resolveText(placed as Placed)?.paragraphs[0]?.runs[0]?.font.family).toBe('Verdana');
  });

  it('gives a shape with no text body nothing at all', () => {
    const placed = layoutSheet(slideWithText(''))[0];
    expect(resolveText(placed as Placed)).toBeNull();
  });

  it("reads the frame's own insets and anchor in points", () => {
    const placed = layoutSheet(
      slideWithText(
        '<p:txBody><a:bodyPr lIns="0" tIns="0" rIns="0" bIns="0" anchor="ctr" wrap="none"/>' +
          '<a:lstStyle/><a:p/></p:txBody>',
      ),
    )[0];
    const frame = resolveText(placed as Placed)?.frame;
    expect(frame?.insets).toEqual({ left: 0, top: 0, right: 0, bottom: 0 });
    expect(frame?.anchor).toBe('ctr');
    expect(frame?.wrap).toBe('none');
  });

  it('carries the stored autofit scale rather than a percentage', () => {
    const placed = layoutSheet(
      slideWithText(
        '<p:txBody><a:bodyPr><a:normAutofit fontScale="62500" lnSpcReduction="20000"/></a:bodyPr>' +
          '<a:lstStyle/><a:p/></p:txBody>',
      ),
    )[0];
    const frame = resolveText(placed as Placed)?.frame;
    expect(frame?.fontScale).toBeCloseTo(0.625, 6);
    expect(frame?.lineSpaceReduction).toBeCloseTo(0.2, 6);
  });
});

/* -------------------------------------------------------------------------- */
/* the emitter                                                                */
/* -------------------------------------------------------------------------- */

const FRAME_EMU = { x: 1524000, y: 1524000, cx: 7620000, cy: 3048000 };

function svgOf(block: TextBlock): string {
  return serializeSvg(textNodes(block, FRAME_EMU)[0] as SvgElement);
}

describe('textNodes', () => {
  it('emits real text elements and never a foreignObject', () => {
    const svg = svgOf(laid(body([paragraph([run('Alpha')])])));
    expect(svg).toContain('<text');
    expect(svg).toContain('<tspan');
    expect(svg).not.toContain('foreignObject');
  });

  it('places the line once and lets the pieces flow', () => {
    // A tspan carrying its own `x` restarts shaping and loses the kern that
    // crosses a run boundary, which T8 measured PowerPoint keeping.
    const svg = svgOf(laid(body([paragraph([run('AV'), run('AWA')])])));
    const spans = svg.match(/<tspan/g) ?? [];
    expect(spans.length).toBe(2);
    expect(svg).not.toMatch(/<tspan[^>]* x=/);
  });

  it('turns the group rather than the glyphs', () => {
    const block = laid(body([paragraph([run('A')])]), { rot: 90 });
    expect(svgOf(block)).toContain('rotate(90)');
  });
  it('never scales negatively, whatever the flip', () => {
    const block = laid(body([paragraph([run('A')])]), { flipH: true, rot: 30 });
    const svg = svgOf(block);
    expect(svg).not.toContain('scale(-');
    expect(svg).toContain('rotate(30)');
  });

  it("writes a run's colour as the hex PowerPoint drew", () => {
    // `Rgba` channels are 0..1. Treating them as bytes paints every coloured
    // run black, which no test using `color: null` would ever notice.
    const red = { r: 192 / 255, g: 0, b: 0, a: 1 };
    const block = laid(body([paragraph([run('a', { color: red })])]));
    expect(svgOf(block)).toContain('fill="#C00000"');
  });

  it('carries a translucent colour as an opacity rather than a fourth channel', () => {
    const block = laid(body([paragraph([run('a', { color: { r: 0, g: 0, b: 1, a: 0.5 } })])]));
    const svg = svgOf(block);
    expect(svg).toContain('fill="#0000FF"');
    expect(svg).toContain('fill-opacity="0.5"');
  });

  it('draws a rule as its own element', () => {
    const svg = svgOf(laid(body([paragraph([run('ab', { underline: 'sng' })])])));
    expect(svg).toContain('<rect');
  });

  it('draws a patterned rule as a dashed stroke', () => {
    const svg = svgOf(laid(body([paragraph([run('ab', { underline: 'dotted' })])])));
    expect(svg).toContain('stroke-dasharray');
  });

  it('keeps the spaces a run holds', () => {
    const svg = svgOf(laid(body([paragraph([run('a  b')])])));
    expect(svg).toContain('xml:space="preserve"');
    expect(svg).toContain('a  b');
  });

  it('emits nothing for a block with no lines', () => {
    expect(
      textNodes(
        { turnDeg: 0, boxPt: { leftPt: 0, topPt: 0, widthPt: 0, heightPt: 0 }, lines: [] },
        FRAME_EMU,
      ),
    ).toEqual([]);
  });

  it('carries a baseline shift as a dy that the next piece undoes', () => {
    const block = laid(body([paragraph([run('x', { baseline: 30000 }), run('y')])]));
    const svg = svgOf(block);
    const shifts = [...svg.matchAll(/dy="(-?[\d.]+)"/g)].map((match) => Number(match[1]));
    expect(shifts.length).toBe(2);
    expect((shifts[0] ?? 0) + (shifts[1] ?? 0)).toBeCloseTo(0, 6);
  });
});

describe('the CSS family a piece is drawn in', () => {
  it("is the run's own family, quoted, when nobody says otherwise", () => {
    // A literal rather than askedFamily(...): the browser renderer measures
    // through this exact string, so the markup must carry this exact string.
    const block = laid(body([paragraph([run('a')])]));
    expect(block.lines[0]?.pieces[0]?.cssFamily).toBe('"Arial"');
    expect(svgOf(block)).toContain('font-family="&quot;Arial&quot;"');
  });

  it('asks the caller per piece, with the weight and style the piece is set in', () => {
    const block = laid(
      body([
        paragraph([
          run('a', { font: { family: 'Arial', sz: 3200, bold: true } }),
          run('b', { font: { family: 'Georgia', sz: 3200, italic: true } }),
        ]),
      ]),
      {
        cssFamilyFor: (font: RunFont) =>
          `${font.family}/${font.bold === true ? 'b' : ''}${font.italic === true ? 'i' : ''}`,
      },
    );
    expect(block.lines[0]?.pieces.map((piece) => piece.cssFamily)).toEqual([
      'Arial/b',
      'Georgia/i',
    ]);
    expect(svgOf(block)).toContain('font-family="Arial/b"');
  });

  it('reaches the layout through the engine a slide is drawn with', () => {
    const placed = layoutSheet(
      slideWithText(
        '<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US">' +
          '<a:latin typeface="Georgia"/></a:rPr><a:t>Alpha</a:t></a:r></a:p></p:txBody>',
      ),
    )[0] as Placed;
    const options = { measurer: flatMeasurer, faceBox: flatFaceBox, rulesFor: () => flatRules };
    const named = createTextEngine({
      ...options,
      cssFamilyFor: (font: RunFont) => `stack:${font.family}`,
    });
    expect(textBlockOf(placed, named)?.lines[0]?.pieces[0]?.cssFamily).toBe('stack:Georgia');
    expect(textBlockOf(placed, createTextEngine(options))?.lines[0]?.pieces[0]?.cssFamily).toBe(
      '"Georgia"',
    );
  });
});

/* -------------------------------------------------------------------------- */
/* the two renderers                                                          */
/* -------------------------------------------------------------------------- */

function textElements(node: SvgNode, out: SvgElement[] = []): SvgElement[] {
  if (node.kind !== 'element') return out;
  if (node.tag === 'text') out.push(node);
  for (const child of node.children) textElements(child, out);
  return out;
}

describe('the two renderers over one layout', () => {
  it('draw the same line boxes, because there is only one set of them', () => {
    const block = laid(body([paragraph([run('Alpha bravo charlie')]), paragraph([run('delta')])]));
    const svg = textNodes(block, FRAME_EMU)[0];
    const lines = textElements(svg as SvgNode);
    expect(lines.length).toBe(block.lines.length);
    block.lines.forEach((line, index) => {
      expect(lines[index]?.attrs['x']).toBe(line.leftPt);
      expect(lines[index]?.attrs['y']).toBe(line.baselinePt);
    });
  });
});

/* -------------------------------------------------------------------------- */
/* astral characters                                                          */
/* -------------------------------------------------------------------------- */

/** A code unit that is half of a surrogate pair and has lost the other half. */
function loneSurrogates(text: string): readonly string[] {
  const lone: string[] = [];
  for (let i = 0; i < text.length; i++) {
    const code = text.charCodeAt(i);
    if (code >= 0xd800 && code <= 0xdbff) {
      const next = text.charCodeAt(i + 1);
      if (next >= 0xdc00 && next <= 0xdfff) i += 1;
      else lone.push(code.toString(16));
    } else if (code >= 0xdc00 && code <= 0xdfff) lone.push(code.toString(16));
  }
  return lone;
}

/** U+1D54F, four of them, which need eight code units and count as four points. */
const ASTRAL = '\u{1D54F}\u{1D54F}\u{1D54F}\u{1D54F}';

describe('text outside the basic multilingual plane', () => {
  it('never emits half of a surrogate pair, which is not well-formed XML', () => {
    const block = laid(body([paragraph([run(`${ASTRAL} alpha bravo charlie delta`)])]), {
      widthPt: 120,
    });
    for (const line of block.lines) {
      for (const piece of line.pieces) expect(loneSurrogates(piece.text)).toEqual([]);
    }
  });

  it('draws every character the run holds, and no character twice', () => {
    const source = `${ASTRAL} alpha bravo charlie delta`;
    const block = laid(body([paragraph([run(source)])]), { widthPt: 120 });
    const drawn = block.lines.flatMap((line) => line.pieces.map((piece) => piece.text)).join('');
    // A break consumes its space, so the spaces are what the two sides differ by.
    expect([...drawn].filter((c) => c !== ' ').join('')).toBe(
      [...source].filter((c) => c !== ' ').join(''),
    );
  });

  it('starts the second run where the first one ends, counted in code points', () => {
    // Two runs, so the offset between cells is load-bearing: read in code units
    // the second run starts past the end of the line and is never drawn.
    const block = laid(body([paragraph([run(ASTRAL), run('beta')])]), { widthPt: 4000 });
    const pieces = block.lines[0]?.pieces ?? [];
    expect(pieces.map((piece) => piece.text)).toEqual([ASTRAL, 'beta']);
  });

  it('measures each piece over the characters it actually draws', () => {
    // `flatMeasurer` charges one point per code unit per hundredth of a point, so
    // the four astral characters are eight units at 32pt and `beta` is four.
    const block = laid(body([paragraph([run(ASTRAL), run('beta')])]), { widthPt: 4000 });
    const pieces = block.lines[0]?.pieces ?? [];
    expect(pieces.map((piece) => piece.widthPt)).toEqual([8 * 32, 4 * 32]);
    expect(block.lines[0]?.widthPt).toBe(12 * 32);
  });

  it('draws the whole run on a line wide enough not to wrap', () => {
    // The line box `wrapText` returns ends at a code point index. Read as a code
    // unit index it stops short by one per astral character, truncating the tail.
    const source = `${ASTRAL} beta`;
    const block = laid(body([paragraph([run(source)])]), { widthPt: 4000 });
    expect(block.lines.length).toBe(1);
    expect(block.lines[0]?.pieces.map((piece) => piece.text).join('')).toBe(source);
  });
});

/* -------------------------------------------------------------------------- */
/* the turned frame, against T6's own readings                                */
/* -------------------------------------------------------------------------- */

interface FrameProbe {
  readonly id: string;
  readonly family: string;
  readonly vars: Readonly<Record<string, unknown>>;
  readonly frame: Readonly<Record<string, unknown>>;
  readonly box: readonly number[];
  readonly measured: {
    readonly dLeft: number | null;
    readonly dTop: number | null;
    readonly frame?: unknown;
    readonly lineLefts?: unknown;
  };
  readonly drew?: readonly string[] | undefined;
}

/** A probe field that is a string, or the default when the probe omits it. */
function word(row: Readonly<Record<string, unknown>>, key: string, fallback: string): string {
  const value = row[key];
  return typeof value === 'string' ? value : fallback;
}

/** A probe field that is a length in points, or zero when the probe omits it. */
function length(row: Readonly<Record<string, unknown>>, key: string): number {
  const value = row[key];
  return typeof value === 'number' ? value : 0;
}

const allFrameProbes = frames.probes as readonly unknown[] as readonly FrameProbe[];

const frameProbes = allFrameProbes.filter(
  (probe) => probe.family === 'vert' && !probe.id.startsWith('vert-wrap-'),
);

/** The probes that wrap, which are the only ones with more than one line. */
const wrapProbes = allFrameProbes.filter((probe) => probe.id.startsWith('vert-wrap-'));

/** The along-axis advance of each script's probe string, from its `horz` row. */
const PROBE_ADVANCE: Readonly<Record<string, number>> = { latin: 49, cjk: 99 };

/** The probe font: 18pt, so its line box is 21.6. */
const PROBE_SIZE = 1800;

/** The insets of the `vert-ins-*` probes, asymmetric on both axes. */
const PROBE_INSETS = { left: 13, top: 20, right: 3, bottom: 5 };

/** Every `ST_TextVerticalType`, all of which this package now draws. */
const DRAWN = [
  'horz',
  'vert',
  'vert270',
  'wordArtVert',
  'eaVert',
  'mongolianVert',
  'wordArtVertRtl',
];

/** A measurer answering with the advance PowerPoint measured for the probe. */
function probeMeasurer(
  advancePt: number,
  text: string,
): { measure: (part: string) => { width: number } } {
  const glyphs = Math.max(1, [...text].length);
  // The probe strings are one script and one size each, so an even share of the
  // measured advance is the same measurement whether asked whole or per glyph.
  return {
    measure: (part: string): { width: number } => ({
      width: (advancePt * [...part].length) / glyphs,
    }),
  };
}

/** The block's near corner on the slide, once the turn has been applied. */
function screenCorner(
  block: TextBlock,
  widthPt: number,
  heightPt: number,
): { leftPt: number; topPt: number } {
  const radians = (block.turnDeg * Math.PI) / 180;
  const cos = Math.round(Math.cos(radians));
  const sin = Math.round(Math.sin(radians));
  const halfX = widthPt / 2;
  const halfY = heightPt / 2;
  const xs: number[] = [];
  const ys: number[] = [];
  for (const line of block.lines) {
    const corners = [
      [line.leftPt, line.topPt],
      [line.leftPt + line.widthPt, line.topPt + line.heightPt],
    ] as const;
    for (const [x, y] of corners) {
      xs.push(halfX + (x - halfX) * cos - (y - halfY) * sin);
      ys.push(halfY + (x - halfX) * sin + (y - halfY) * cos);
    }
  }
  return { leftPt: Math.min(...xs), topPt: Math.min(...ys) };
}

/**
 * A wrapped probe, given an em of advance per glyph so that it wraps at all.
 *
 * The assertion this feeds is which way the lines step, not where they land, so
 * the count need not be PowerPoint's - only more than one.
 */
function laidWrapProbe(probe: FrameProbe): TextBlock {
  const text = 'W'.repeat(20);
  return layoutText(
    body([paragraph([run(text, { font: { family: probeFace(probe), sz: PROBE_SIZE } })])], {
      vertical: word(probe.frame, 'vert', 'horz') as ResolvedFrame['vertical'],
      anchor: 't',
      insets: { left: 0, top: 0, right: 0, bottom: 0 },
      wrap: 'square',
    }),
    {
      widthPt: probe.box[0] ?? 0,
      heightPt: probe.box[1] ?? 0,
      rot: 0,
      flipH: false,
      flipV: false,
      measurer: { measure: (part: string) => ({ width: [...part].length * (PROBE_SIZE / 100) }) },
      faceBox: probeFaceBox(probe),
      rulesFor: () => flatRules,
    },
  );
}

/** The near edge of every line on the slide, in the order the block lists them. */
function screenLefts(block: TextBlock, widthPt: number, heightPt: number): number[] {
  const radians = (block.turnDeg * Math.PI) / 180;
  const cos = Math.round(Math.cos(radians));
  const sin = Math.round(Math.sin(radians));
  const halfX = widthPt / 2;
  const halfY = heightPt / 2;
  return block.lines.map((line) => {
    const xs = [
      [line.leftPt, line.topPt],
      [line.leftPt + line.widthPt, line.topPt + line.heightPt],
    ].map(([x, y]) => halfX + ((x ?? 0) - halfX) * cos - ((y ?? 0) - halfY) * sin);
    return Math.min(...xs);
  });
}

function vertical(text: string, frame: Partial<ResolvedFrame>, advancePt: number): TextBlock {
  return layoutText(
    body([paragraph([run(text, { font: { family: 'Yu Gothic', sz: PROBE_SIZE } })])], {
      insets: { left: 0, top: 0, right: 0, bottom: 0 },
      wrap: 'none',
      ...frame,
    }),
    {
      widthPt: 200,
      heightPt: 200,
      rot: 0,
      flipH: false,
      flipV: false,
      measurer: probeMeasurer(advancePt, text),
      faceBox: flatFaceBox,
      rulesFor: () => flatRules,
    },
  );
}

/** The face a T6 probe was set in, which its deck chose from the probe's script. */
function probeFace(probe: FrameProbe): string {
  const named = probe.vars['face'];
  if (typeof named === 'string') return named;
  return word(probe.vars, 'script', 'latin') === 'cjk' ? 'MS Gothic' : 'Arial';
}

/** T11's browser metrics for that face, so a stacked cell is the real one. */
function probeFaceBox(probe: FrameProbe): { box: () => FaceBox } {
  const metrics = columnFaces.get(probeFace(probe));
  if (metrics === undefined) return flatFaceBox;
  return {
    box: (): FaceBox => ({
      ascent: metrics.ascent,
      descent: metrics.descent,
      ideographic: 0.12,
    }),
  };
}

function laidProbe(probe: FrameProbe, wrap: ResolvedFrame['wrap'] = 'none'): TextBlock {
  const text = (probe.drew ?? []).join('');
  return layoutText(
    body([paragraph([run(text, { font: { family: 'Yu Gothic', sz: PROBE_SIZE } })])], {
      vertical: word(probe.frame, 'vert', 'horz') as ResolvedFrame['vertical'],
      anchor: word(probe.frame, 'anchor', 't') as ResolvedFrame['anchor'],
      insets: {
        left: length(probe.frame, 'lIns'),
        top: length(probe.frame, 'tIns'),
        right: length(probe.frame, 'rIns'),
        bottom: length(probe.frame, 'bIns'),
      },
      wrap,
    }),
    {
      widthPt: probe.box[0] ?? 0,
      heightPt: probe.box[1] ?? 0,
      rot: 0,
      flipH: false,
      flipV: false,
      measurer: probeMeasurer(PROBE_ADVANCE[word(probe.vars, 'script', 'latin')] ?? 49, text),
      faceBox: probeFaceBox(probe),
      rulesFor: () => flatRules,
    },
  );
}

describe('a turned frame, against experiment T6', () => {
  const drawable = frameProbes.filter((probe) => DRAWN.includes(word(probe.frame, 'vert', 'horz')));

  it('has probes for every direction this package draws', () => {
    expect(new Set(drawable.map((probe) => word(probe.frame, 'vert', 'horz'))).size).toBe(
      DRAWN.length,
    );
    expect(drawable.length).toBeGreaterThanOrEqual(15);
  });

  it('puts the block where PowerPoint put it, in all of them', () => {
    const misses: string[] = [];
    for (const probe of drawable) {
      const corner = screenCorner(laidProbe(probe), probe.box[0] ?? 0, probe.box[1] ?? 0);
      const wantLeft = (probe.measured.dLeft ?? 0) / 1000;
      const wantTop = (probe.measured.dTop ?? 0) / 1000;
      if (Math.abs(corner.leftPt - wantLeft) > 0.05 || Math.abs(corner.topPt - wantTop) > 0.05) {
        misses.push(`${probe.id}: ${corner.leftPt.toFixed(2)},${corner.topPt.toFixed(2)}`);
      }
    }
    expect(misses).toStrictEqual([]);
  });

  it('steps its lines the way PowerPoint steps them, in every wrapped direction', () => {
    // The block's own corner is the same whichever end the first line sits at,
    // so only the step between two lines separates them. `vert` walks inward
    // from the far edge and `mongolianVert` outward from zero.
    const wrapped = wrapProbes.filter(
      (probe) =>
        DRAWN.includes(word(probe.frame, 'vert', 'horz')) &&
        Array.isArray(probe.measured.lineLefts) &&
        (probe.measured.lineLefts as number[]).length >= 2,
    );
    expect(wrapped.length).toBeGreaterThanOrEqual(5);
    for (const probe of wrapped) {
      const block = laidWrapProbe(probe);
      expect(block.lines.length, probe.id).toBeGreaterThanOrEqual(2);
      const mine = screenLefts(block, probe.box[0] ?? 0, probe.box[1] ?? 0);
      const measured = probe.measured.lineLefts as number[];
      expect(Math.sign((mine[1] ?? 0) - (mine[0] ?? 0)), probe.id).toBe(
        Math.sign((measured[1] ?? 0) - (measured[0] ?? 0)),
      );
    }
  });

  it('is 17pt out on `vert-ins-vert` if the insets do not turn with the frame', () => {
    // Turning them back the other way first cancels the turn the layout does,
    // which is exactly the rival reading: `lIns` on the laid-out left edge.
    const naive = screenCorner(
      vertical(
        'Wxyz',
        { vertical: 'vert', insets: turnedInsets(PROBE_INSETS, 270) },
        PROBE_ADVANCE['latin'] ?? 49,
      ),
      200,
      200,
    );
    const turned = screenCorner(
      vertical('Wxyz', { vertical: 'vert', insets: PROBE_INSETS }, PROBE_ADVANCE['latin'] ?? 49),
      200,
      200,
    );
    expect(naive.leftPt).toBeCloseTo(158.4, 6);
    expect(turned.leftPt).toBeCloseTo(175.4, 6);
  });

  it('places eaVert exactly where PowerPoint places vert, in every pair', () => {
    // Everything but `frame`, which is the object model's own orientation
    // number and is the one thing the two are meant to differ in.
    const placement = (row: FrameProbe): string => JSON.stringify({ ...row.measured, frame: null });
    const pairs = frameProbes.filter((probe) => word(probe.frame, 'vert', '') === 'eaVert');
    expect(pairs.length).toBeGreaterThanOrEqual(3);
    for (const probe of pairs) {
      const twin = frameProbes.find((row) => row.id === probe.id.replace('eaVert', 'vert'));
      expect(twin, probe.id).toBeDefined();
      expect(placement(probe), probe.id).toBe(placement(twin as FrameProbe));
    }
  });

  it('turns the block a quarter for every direction but the horizontal one', () => {
    for (const vert of [
      'horz',
      'vert',
      'vert270',
      'wordArtVert',
      'eaVert',
      'mongolianVert',
      'wordArtVertRtl',
    ]) {
      const block = vertical('Wxyz', { vertical: vert as ResolvedFrame['vertical'] }, 49);
      const quarter = { horz: 0, vert270: 270 }[vert] ?? 90;
      expect(block.turnDeg, vert).toBe(quarter);
    }
  });
});

describe('upright East Asian glyphs', () => {
  const CJK = '日本語';

  it('stands every East Asian glyph up in an eaVert frame', () => {
    const line = vertical(CJK, { vertical: 'eaVert' }, 54).lines[0];
    expect(line?.pieces).toHaveLength(1);
    expect(line?.pieces[0]?.upright.map((glyph) => glyph.text)).toStrictEqual([...CJK]);
  });

  it('leaves them flowing in a vert frame, which turns the glyphs too', () => {
    expect(vertical(CJK, { vertical: 'vert' }, 54).lines[0]?.pieces[0]?.upright).toStrictEqual([]);
  });

  it('stands them up in a mongolianVert frame as well', () => {
    const line = vertical(CJK, { vertical: 'mongolianVert' }, 54).lines[0];
    expect(line?.pieces[0]?.upright).toHaveLength(3);
  });

  it('advances one glyph cell at a time, and centres each across the line', () => {
    const line = vertical(CJK, { vertical: 'eaVert' }, 54).lines[0];
    const glyphs = line?.pieces[0]?.upright ?? [];
    // A 54pt advance over three glyphs is an 18pt cell, which is the size; the
    // flat face box puts the ideographic baseline an eighth of the em up.
    expect(glyphs.map((glyph) => glyph.alongPt - (line?.leftPt ?? 0))).toStrictEqual([
      15.75, 33.75, 51.75,
    ]);
    for (const glyph of glyphs) expect(glyph.acrossPt).toBeCloseTo((21.6 + 18) / 2, 10);
  });

  it('splits a run at the script boundary and stands up only its own half', () => {
    const line = vertical('日本Ab', { vertical: 'eaVert' }, 72).lines[0];
    expect(line?.pieces.map((piece) => piece.text)).toStrictEqual(['日本', 'Ab']);
    expect(line?.pieces[0]?.upright).toHaveLength(2);
    expect(line?.pieces[1]?.upright).toStrictEqual([]);
  });

  it('never splits a run in a frame that draws no upright glyph', () => {
    const line = vertical('日本Ab', { vertical: 'vert' }, 72).lines[0];
    expect(line?.pieces.map((piece) => piece.text)).toStrictEqual(['日本Ab']);
  });

  it('draws each upright glyph as its own turned <text>', () => {
    const block = vertical(CJK, { vertical: 'eaVert' }, 54);
    const markup = serializeSvg(
      textNodes(block, { x: 0, y: 0, cx: 2540000, cy: 2540000 })[0] as SvgElement,
    );
    expect([...markup.matchAll(/rotate\(-90 /g)]).toHaveLength(3);
    for (const glyph of [...CJK]) expect(markup).toContain(`>${glyph}</text>`);
  });

  it('refuses an underline nobody measured under upright text', () => {
    expect(() => vertical(CJK, { vertical: 'eaVert' }, 54)).not.toThrow();
    expect(() =>
      layoutText(
        body(
          [
            paragraph([
              run(CJK, { font: { family: 'Yu Gothic', sz: PROBE_SIZE }, underline: 'sng' }),
            ]),
          ],
          { vertical: 'eaVert', insets: { left: 0, top: 0, right: 0, bottom: 0 }, wrap: 'none' },
        ),
        {
          widthPt: 200,
          heightPt: 200,
          rot: 0,
          flipH: false,
          flipV: false,
          measurer: probeMeasurer(54, CJK),
          faceBox: flatFaceBox,
          rulesFor: () => flatRules,
        },
      ),
    ).toThrow(RenderError);
  });
});

/* -------------------------------------------------------------------------- */
/* the WordArt column, against experiment T11                                 */
/* -------------------------------------------------------------------------- */

interface ColumnGlyph {
  readonly glyph: string;
  readonly penXPt: number;
  readonly penYPt: number;
}

interface ColumnProbe {
  readonly id: string;
  readonly family: string;
  readonly vert: string;
  readonly text: string;
  readonly paragraphs: number;
  readonly wrap: string;
  readonly secondRun: { readonly text: string; readonly sizePt: number } | null;
  readonly face: string;
  readonly sizePt: number;
  readonly boundWidthPt: number;
  readonly columns: readonly (readonly ColumnGlyph[])[];
}

interface ColumnFace {
  readonly face: string;
  readonly ascent: number;
  readonly descent: number;
  readonly advances: Readonly<Record<string, number>>;
}

/** The frame every T11 probe was drawn in. */
const T11_FRAME = { widthPt: 220, heightPt: 240 };

/**
 * A twelfth of a point: two logical units of the stream the pens were read from.
 *
 * PowerPoint writes a pen as a whole logical unit, so no reading is finer, and a
 * tolerance below this would be asserting the rounding rather than the rule.
 */
const PEN_PT = 2 / columns.unitsPerPoint;

/** What PowerPoint's own pens stray from a uniform stack, as a fraction of the em. */
const PEN_JITTER_EM = columns.penJitterEm;

/** What PowerPoint's own cell strays from seven sixths of the box, per cell. */
const CELL_ROUNDING_EM = columns.cellRoundingEm;

/**
 * How far a pen may sit from where the cell puts it, this many cells down.
 *
 * PowerPoint rounds the cell before it stacks, so the gap grows with the cells
 * above; the fixture measures both that rounding and what its pens stray from a
 * stack of equal cells, and neither is a number this file chose.
 */
function allowedPt(sizePt: number, cellsAbove: number): number {
  return PEN_PT + (CELL_ROUNDING_EM + PEN_JITTER_EM) * sizePt * (cellsAbove + 1);
}

const columnProbes = columns.probes as readonly ColumnProbe[];
const columnFaces = new Map(
  (columns.browser as readonly ColumnFace[]).map((row) => [row.face, row] as const),
);

/** The faces whose own font box is the box PowerPoint measured the column from. */
const FITTING_FACES = new Set(
  (columns.axes as readonly { face: string; columnError: number }[])
    .filter((row) => Math.abs(row.columnError) <= 0.005)
    .map((row) => row.face),
);

function columnMetrics(face: string): ColumnFace {
  const metrics = columnFaces.get(face);
  if (metrics === undefined) throw new Error(`T11 measured no browser metrics for ${face}`);
  return metrics;
}

/** The advance the browser measured, summed over the code points asked for. */
function columnAdvance(metrics: ColumnFace, part: string, sizePt: number): number {
  let width = 0;
  for (const glyph of [...part]) width += (metrics.advances[glyph] ?? 0) * sizePt;
  return width;
}

function wordArtBlock(probe: ColumnProbe): TextBlock {
  const metrics = columnMetrics(probe.face);
  const runs = [run(probe.text, { font: { family: probe.face, sz: probe.sizePt * 100 } })];
  if (probe.secondRun !== null) {
    runs.push(
      run(probe.secondRun.text, {
        font: { family: probe.face, sz: probe.secondRun.sizePt * 100 },
      }),
    );
  }
  const paragraphs = Array.from({ length: probe.paragraphs }, () => paragraph(runs));
  return layoutText(
    body(paragraphs, {
      vertical: probe.vert as ResolvedFrame['vertical'],
      anchor: 't',
      insets: { left: 0, top: 0, right: 0, bottom: 0 },
      wrap: probe.wrap as ResolvedFrame['wrap'],
    }),
    {
      widthPt: T11_FRAME.widthPt,
      heightPt: T11_FRAME.heightPt,
      rot: 0,
      flipH: false,
      flipV: false,
      measurer: {
        // The run's own size, not the probe's: a mixed-size column has two.
        measure: (part: string, font: RunFont): { width: number } => ({
          width: columnAdvance(metrics, part, font.sz / 100),
        }),
      },
      faceBox: {
        box: (): FaceBox => ({
          ascent: metrics.ascent,
          descent: metrics.descent,
          ideographic: 0,
        }),
      },
      rulesFor: () => flatRules,
    },
  );
}

/** A point inside a laid-out block, put on the slide by the block's own turn. */
function screenPoint(block: TextBlock, xPt: number, yPt: number): { xPt: number; yPt: number } {
  const radians = (block.turnDeg * Math.PI) / 180;
  const cos = Math.round(Math.cos(radians));
  const sin = Math.round(Math.sin(radians));
  const halfX = T11_FRAME.widthPt / 2;
  const halfY = T11_FRAME.heightPt / 2;
  return {
    xPt: halfX + (xPt - halfX) * cos - (yPt - halfY) * sin,
    yPt: halfY + (xPt - halfX) * sin + (yPt - halfY) * cos,
  };
}

/**
 * The pen each glyph is drawn from, one array per column.
 *
 * A line of the turned layout is a column on the slide, and PowerPoint draws its
 * columns in paragraph order too, so the two line up without being sorted.
 */
function drawnColumns(block: TextBlock): { glyph: string; xPt: number; yPt: number }[][] {
  return block.lines.map((line) =>
    line.pieces.flatMap((piece) =>
      piece.upright.map((glyph) => ({
        glyph: glyph.text,
        ...screenPoint(
          block,
          line.leftPt + glyph.alongPt - piece.risePt,
          line.topPt + glyph.acrossPt,
        ),
      })),
    ),
  );
}

describe('a WordArt column, against experiment T11', () => {
  const fitting = columnProbes.filter(
    (probe) => probe.vert !== 'horz' && FITTING_FACES.has(probe.face),
  );

  it('has probes on both sides of the font-box boundary, or it proves nothing', () => {
    expect(FITTING_FACES.size).toBeGreaterThanOrEqual(6);
    expect(columnProbes.length - fitting.length).toBeGreaterThanOrEqual(6);
    expect([...FITTING_FACES]).not.toContain('Yu Gothic');
  });

  it('draws every glyph where PowerPoint drew it, on every face whose box is its own', () => {
    let worst = 0;
    for (const probe of fitting) {
      const drawn = drawnColumns(wordArtBlock(probe));
      expect(
        drawn.map((column) => column.map((pen) => pen.glyph).join('')),
        probe.id,
      ).toStrictEqual(probe.columns.map((column) => column.map((glyph) => glyph.glyph).join('')));
      drawn.forEach((column, index) => {
        column.forEach((pen, at) => {
          const want = probe.columns[index]?.[at];
          expect(want, `${probe.id} [${String(index)}][${String(at)}]`).toBeDefined();
          const dx = Math.abs(pen.xPt - (want?.penXPt ?? 0));
          const dy = Math.abs(pen.yPt - (want?.penYPt ?? 0));
          const allowed = allowedPt(probe.sizePt, at);
          worst = Math.max(worst, dx, dy);
          expect(dx, `${probe.id} x ${pen.glyph}`).toBeLessThanOrEqual(allowed);
          expect(dy, `${probe.id} y ${pen.glyph}`).toBeLessThanOrEqual(allowed);
        });
      });
    }
    expect(worst).toBeLessThanOrEqual(allowedPt(28, 9));
  });

  it('breaks the column on the cell and not on the word', () => {
    for (const probe of columnProbes) {
      if (probe.family !== 'break' || !FITTING_FACES.has(probe.face)) continue;
      const block = wordArtBlock(probe);
      expect(block.lines.length, probe.id).toBe(probe.columns.length);
      const drew = block.lines.map((line) =>
        line.pieces.reduce((sum, piece) => sum + piece.upright.length, 0),
      );
      expect(drew, probe.id).toStrictEqual(probe.columns.map((column) => column.length));
    }
  });

  it('stacks `wordArtVert` from the left edge and `wordArtVertRtl` from the right', () => {
    for (const probe of columnProbes) {
      if (probe.columns.length < 2 || !FITTING_FACES.has(probe.face)) continue;
      const drawn = drawnColumns(wordArtBlock(probe));
      const first = drawn[0]?.[0]?.xPt ?? NaN;
      const last = drawn.at(-1)?.[0]?.xPt ?? NaN;
      if (probe.vert === 'wordArtVert') {
        expect(first, probe.id).toBeLessThan(last);
      } else {
        expect(first, probe.id).toBeGreaterThan(last);
      }
    }
  });

  it('is more than a point out if the cell is the horizontal line height instead', () => {
    // 1.2 of the em is what a horizontal line advances by, and it is the reading
    // anyone writes first. The probe that separates them is any face at all.
    const probe = columnProbes.find((row) => row.id === 'cell-Arial-28');
    expect(probe).toBeDefined();
    if (probe === undefined) return;
    const metrics = columnMetrics(probe.face);
    const cellPt = (7 / 6) * (metrics.ascent + metrics.descent) * probe.sizePt;
    const rival = 1.2 * probe.sizePt;
    expect(Math.abs(cellPt - rival)).toBeGreaterThan(1);
  });

  it('turns every stacked glyph back out of the line, in the emitted SVG', () => {
    const probe = columnProbes.find((row) => row.id === 'cell-Arial-18');
    expect(probe).toBeDefined();
    if (probe === undefined) return;
    const frame = { x: 0, y: 0, cx: T11_FRAME.widthPt, cy: T11_FRAME.heightPt };
    const svg = textNodes(wordArtBlock(probe), frame)
      .map((node) => serializeSvg(node))
      .join('');
    expect(svg.match(/rotate\(-90 /g)?.length).toBe([...probe.text].length);
  });
});
