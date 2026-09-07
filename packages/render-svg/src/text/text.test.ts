import { describe, expect, it } from 'vitest';

import { parseSheet, parseTheme, type Sheet } from '@pptx-studio/model';
import { parseXmlString } from '@pptx-studio/xml';

import fixture from '../../../../corpus/ground-truth/text-rendering.json' with { type: 'json' };
import { layoutSheet, type Placed } from '../layout.js';
import { RenderError } from '../errors.js';
import { serializeSvg, type SvgElement, type SvgNode } from '../node.js';

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

/** A face box with a share of exactly four fifths, for arithmetic anyone can do. */
const flatFaceBox = {
  box: (): { ascent: number; descent: number } => ({ ascent: 0.8, descent: 0.2 }),
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

  it('refuses a direction it lays out but cannot draw', () => {
    expect(() => laid(body([paragraph([run('a')])], { vertical: 'eaVert' }))).toThrow(RenderError);
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
    const face = { ascent: 0.9, descent: 0.25 };
    const size = 40;
    const drop = 1.2 * size * (face.ascent / (face.ascent + face.descent));
    const strut = strutHeight(drop, face, size);
    const browser = (strut - (face.ascent + face.descent) * size) / 2 + face.ascent * size;
    expect(browser).toBeCloseTo(drop, 10);
  });

  it('never asks for a negative height', () => {
    expect(strutHeight(0, { ascent: 0.9, descent: 0.2 }, 40)).toBe(0);
  });

  it('solves every baseline the fixture recorded', () => {
    const boxes = fixture.browser.faces as Record<string, { ascent: number; descent: number }>;
    for (const row of baselineRows) {
      const box = boxes[row.face];
      if (box === undefined) continue;
      const strut = strutHeight(row.drop, box, row.size);
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
