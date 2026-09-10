/**
 * Experiment T14 - each gate, shown going red.
 *
 * Every row here is authored rather than measured, so the suite fails when
 * `tools/ground-truth/fonts/real-faces/score.ts` stops deciding what it says it
 * decides. A gate nobody has seen go red is not a gate.
 */

import { describe, expect, it } from 'vitest';

import {
  ADVANCE_QUANTUM,
  BOX_PX,
  BOX_TOLERANCE,
  SHIPPED_BOX,
  glyphsOf,
  widthTolerance,
  type BoxCandidates,
  type BoxPair,
} from './probes.ts';
import { score, type Row, type Run, type RunFace, type Verdict } from './score.ts';

/** 1705/2048 and 615/2048 of a 1000px em, and the whole pixels Chromium reports. */
const EM = 2048;
const READER_BOX = { ascent: 1705, descent: 615 };
const BROWSER_BOX = { ascent: 833, descent: 300 };

const SHORT = 'A'.repeat(10);
const SPREAD = 'A'.repeat(40);
const LONG = 'A'.repeat(1349);

const SAMPLES: Run['samples'] = [
  { id: 'latin-short', script: 'latin', gated: true, asks: 'the per-glyph error', text: SHORT },
  {
    id: 'latin-spread',
    script: 'latin',
    gated: true,
    asks: 'a cmap across its range',
    text: SPREAD,
  },
  { id: 'latin-long', script: 'latin', gated: true, asks: 'whether the error drifts', text: LONG },
  { id: 'arabic', script: 'arabic', gated: false, asks: 'what a missing shaper costs', text: 'اب' },
];

/** A gated row, with the four rivals a whole-pixel browser cannot separate. */
function gatedRow(sample: string, px: number, browser: number, shipped: number): Row {
  return {
    sample,
    px,
    browser,
    readings: {
      shipped,
      exact: shipped + 1e-7,
      advanceRounded: shipped - 1e-7,
      wholeStringTruncated: shipped - 2e-7,
      unkerned: browser + 6,
    },
  };
}

function arabicRow(px: number, browser: number, ratio: number): Row {
  return { sample: 'arabic', px, browser, readings: { shipped: browser * ratio } };
}

/** A face whose three tables agree with themselves, which nearly every real one is. */
function agreeing(pair: BoxPair): BoxCandidates {
  return { hhea: pair, usWin: pair, sTypo: pair, useTypoMetrics: false };
}

function faceWith(
  rows: readonly Row[],
  box = READER_BOX,
  browserBox = BROWSER_BOX,
  candidates: BoxCandidates | null = agreeing(box),
): RunFace {
  return {
    file: '/usr/share/fonts/truetype/probe/Probe.ttf',
    sha256: 'a'.repeat(64),
    family: 'Probe Sans',
    subfamily: 'Regular',
    unitsPerEm: EM,
    metricsSource: 'usWin',
    box: { browser: browserBox, reader: box, ...(candidates === null ? {} : { candidates }) },
    uncovered: SAMPLES.filter((sample) => !rows.some((row) => row.sample === sample.id)).map(
      (sample) => sample.id,
    ),
    rows,
  };
}

function runWith(faces: readonly RunFace[], directories = ['/usr/share/fonts']): Run {
  return {
    experiment: 'T14',
    subPhase: '3.10',
    adr: 'docs/adr/phase-3-text/0042-rendering-without-a-browser.md',
    chromium: '151.0.7922.34',
    image: { platform: 'linux', release: '6.17.0', imageOs: 'ubuntu24', imageVersion: '2026' },
    directories,
    sizes: [16, 1000],
    boxPx: BOX_PX,
    samples: SAMPLES,
    indexed: { files: 1, faces: 1, measured: faces.length },
    skipped: [],
    faces,
  };
}

/** The measurement a green run makes: every gated width inside the tolerance. */
const GREEN = runWith([
  faceWith([
    gatedRow('latin-short', 16, 100, 100),
    gatedRow('latin-spread', 1000, 15291, 15311),
    arabicRow(16, 100, 1.35),
    arabicRow(1000, 6000, 1.36),
  ]),
]);

const kinds = (verdict: Verdict): string[] => verdict.failures.map((failure) => failure.kind);

/**
 * The rows under test, with the two controls every run needs.
 *
 * One gated row the reader gets right and the unkerned rival does not, so the
 * kerning separation holds; one recorded row outside the tolerance, so the
 * missing shaper is priced.
 */
function withGated(...rows: readonly Row[]): Run {
  return runWith([
    faceWith([...rows, gatedRow('latin-short', 16, 100, 100), arabicRow(16, 100, 1.35)]),
  ]);
}

describe('the tolerance', () => {
  it('is half a quantum per glyph, plus the reader step, and nothing else', () => {
    expect(widthTolerance(40)).toBe(40 * (ADVANCE_QUANTUM / 2 + (1 / 65536 + 1 / 131072)));
    expect(widthTolerance(40)).toBeCloseTo(20.000915527, 9);
    expect(widthTolerance(10)).toBeCloseTo(5.000228882, 9);
  });

  it('counts a glyph per code point, not per UTF-16 unit', () => {
    expect(glyphsOf('AV To')).toBe(5);
    expect(glyphsOf('\u{20000}')).toBe(1);
  });
});

describe('the T14 gate, green', () => {
  it('passes when every gated width is inside the tolerance', () => {
    expect(score(GREEN).failures).toEqual([]);
  });

  it('passes the row that sits exactly on half a quantum a glyph', () => {
    // Lato-ThinItalic latin-spread at 1000px: 40 glyphs, every advance a half
    // pixel on a 2000 em, all rounded the same way.
    const [row] = score(withGated(gatedRow('latin-spread', 1000, 15291, 15311))).failures;
    expect(row).toBeUndefined();
  });

  it('passes the row that sits exactly on the tolerance, which is inclusive', () => {
    const edge = 15291 + widthTolerance(40);
    expect(kinds(score(withGated(gatedRow('latin-spread', 1000, 15291, edge))))).toEqual([]);
  });

  it('lets the four readings a whole pixel cannot separate tie at the top', () => {
    const tied = withGated({
      sample: 'latin-short',
      px: 16,
      browser: 100,
      readings: {
        shipped: 100,
        exact: 100,
        advanceRounded: 100,
        wholeStringTruncated: 100,
        unkerned: 106,
      },
    });
    const verdict = score(tied);
    expect(verdict.failures).toEqual([]);
    for (const entry of verdict.scoreboard) {
      expect(entry.inside).toBe(entry.reading === 'unkerned' ? 0 : 2);
    }
  });

  it('scores the face box against the whole pixel the browser rounded it to', () => {
    const { faceBox } = score(GREEN);
    expect(faceBox).toMatchObject({
      agreed: 1,
      of: 1,
      scorable: 1,
      variants: [
        { name: 'exact', fits: 0, of: 1 },
        { name: 'round', fits: 1, of: 1 },
        { name: 'floor', fits: 0, of: 1 },
        { name: 'ceil', fits: 0, of: 1 },
      ],
    });
  });

  it('scores every reading of the box, and separates none on a face that agrees with itself', () => {
    const { readings } = score(GREEN).faceBox;
    expect(readings.map((entry) => entry.reading)).toContain(SHIPPED_BOX);
    for (const entry of readings) {
      expect(entry, entry.reading).toMatchObject({ fits: 1, of: 1, separates: 0 });
    }
  });

  it('reads the box from hhea where the face carries no OS/2, as the reader does', () => {
    const hhea = { ascent: 1705, descent: 615 };
    const bare = runWith([
      faceWith([gatedRow('latin-short', 16, 100, 100), arabicRow(16, 100, 1.35)], READER_BOX, {
        ascent: 833,
        descent: 300,
      }),
    ]).faces.map((face) => ({
      ...face,
      box: { ...face.box, candidates: { hhea, usWin: null, sTypo: null, useTypoMetrics: false } },
    }));
    const verdict = score(runWith(bare));
    expect(verdict.failures).toEqual([]);
    const byName = new Map(verdict.faceBox.readings.map((entry) => [entry.reading, entry]));
    expect(byName.get(SHIPPED_BOX)).toMatchObject({ fits: 1, of: 1 });
    expect(byName.get('OS/2.usWinAscent/usWinDescent')).toMatchObject({ fits: 0, of: 0 });
    expect(byName.get('OS/2.sTypoAscender/sTypoDescender')).toMatchObject({ fits: 0, of: 0 });
  });

  it('reads the box from sTypo where fsSelection bit 7 is set, as the reader does', () => {
    // 1500/500 on a 2048 em is 732.4/244.1 px, and usWin's 1802/401 is 147 px away.
    const bit7 = runWith([
      faceWith(
        [gatedRow('latin-short', 16, 100, 100), arabicRow(16, 100, 1.35)],
        { ascent: 1500, descent: 500 },
        { ascent: 732, descent: 244 },
        {
          hhea: { ascent: 1802, descent: 401 },
          usWin: { ascent: 1802, descent: 401 },
          sTypo: { ascent: 1500, descent: 500 },
          useTypoMetrics: true,
        },
      ),
    ]);
    const verdict = score(bit7);
    expect(verdict.failures).toEqual([]);
    const byName = new Map(verdict.faceBox.readings.map((entry) => [entry.reading, entry]));
    expect(byName.get(SHIPPED_BOX)).toMatchObject({ fits: 1, of: 1, separates: 0 });
    expect(byName.get('OS/2.usWinAscent/usWinDescent')).toMatchObject({ fits: 0, separates: 1 });
    expect(byName.get('hhea.ascender/descender')).toMatchObject({ fits: 0, separates: 1 });
    expect(byName.get('hhea, or sTypo when fsSelection bit 7 is set')).toMatchObject({ fits: 1 });
  });

  it('fits the box that sits exactly on one rounding, which is inclusive', () => {
    // 1667 units on a 2000 em is 833.5 px against the browser's 833.
    const edge = { ascent: 1667, descent: 600 };
    const base = faceWith([gatedRow('latin-short', 16, 100, 100), arabicRow(16, 100, 1.35)]);
    const verdict = score(
      runWith([
        {
          ...base,
          unitsPerEm: 2000,
          box: { browser: { ascent: 833, descent: 300 }, reader: edge, candidates: agreeing(edge) },
        },
      ]),
    );
    expect(verdict.failures).toEqual([]);
    expect(verdict.faceBox.agreed).toBe(1);
    for (const entry of verdict.faceBox.readings) {
      expect(entry, entry.reading).toMatchObject({ fits: 1, of: 1, worstPx: BOX_TOLERANCE });
    }
  });

  it('prices a recorded sample as a ratio, and gates nothing on the number', () => {
    const [arabic] = score(GREEN).shaping;
    expect(arabic).toMatchObject({ sample: 'arabic', faces: 1, of: 2, outsideTolerance: 2 });
    expect(arabic?.medianRatio).toBeCloseTo(1.355, 12);
    expect(arabic?.worstRatio).toBeCloseTo(1.36, 12);
  });

  it('reports a per-script row for the gated script and the recorded one', () => {
    const [latin, arabic] = score(GREEN).byScript;
    expect(latin).toMatchObject({
      script: 'latin',
      gated: true,
      faces: 1,
      comparisons: 2,
      inside: 2,
    });
    expect(latin?.medianRelative).toBeCloseTo(10 / 15291, 12);
    expect(latin?.worstOverTolerance).toBeCloseTo(20 / widthTolerance(40), 12);
    expect(arabic).toMatchObject({ script: 'arabic', gated: false, faces: 1, comparisons: 2 });
    expect(arabic?.medianRelative).toBeCloseTo(0.355, 12);
    expect(arabic?.worstOverTolerance).toBeCloseTo(2160 / widthTolerance(2), 9);
  });
});

describe('the T14 gate, red', () => {
  it('A1 fails when no directory was named', () => {
    expect(kinds(score(runWith(GREEN.faces, [])))).toContain('no-directories');
  });

  it('A1 fails when nothing was measured', () => {
    expect(kinds(score(runWith([])))).toContain('no-faces');
  });

  it('A2 fails when every measured sample was a recorded one', () => {
    const only = runWith([faceWith([arabicRow(16, 100, 1.35)])]);
    expect(kinds(score(only))).toEqual(['no-gated-comparisons']);
  });

  it('A2 fails on a row that carries no reading, or no browser width', () => {
    const absent = runWith([
      faceWith([{ sample: 'latin-short', px: 16, browser: 100, readings: {} }]),
    ]);
    expect(kinds(score(absent))).toContain('unmeasured');
    const nan = runWith([faceWith([gatedRow('latin-short', 16, Number.NaN, 100)])]);
    expect(kinds(score(nan))).toContain('unmeasured');
  });

  it('A3 fails when the browser stops quantising to a whole pixel', () => {
    const fractional = withGated(gatedRow('latin-short', 16, 172.7999725341797, 172.7999725341797));
    const [failure] = score(fractional).failures;
    expect(failure?.kind).toBe('advance-quantum');
    expect(failure?.detail).toContain('172.7999725341797');
  });

  it('A4 fails a hair over the tolerance, naming both widths and the multiple', () => {
    // The narrowest miss the first real run had: Lato-BoldItalic, 1.0063 of it.
    const over = 15291 + widthTolerance(40) * 1.0063;
    const [failure] = score(withGated(gatedRow('latin-spread', 1000, 15291, over))).failures;
    expect(failure?.kind).toBe('width');
    expect(failure?.subject).toBe('/usr/share/fonts/truetype/probe/Probe.ttf');
    expect(failure?.detail).toContain('browser 15291');
    expect(failure?.detail).toContain('1.01 times');
  });

  it('A4 fails a reader that measures narrow, not only one that measures wide', () => {
    const narrow = withGated(gatedRow('latin-spread', 1000, 15291, 15291 - 25));
    expect(kinds(score(narrow))).toEqual(['width']);
  });

  it('A4 fails at three quarters of a pixel a glyph, which is inside one quantum', () => {
    const loose = withGated(gatedRow('latin-spread', 1000, 15291, 15291 + 30));
    expect(kinds(score(loose))).toEqual(['width']);
  });

  it('A4 scales with the glyphs, so the same delta passes long and fails short', () => {
    expect(kinds(score(withGated(gatedRow('latin-long', 1000, 500000, 500600))))).toEqual([]);
    expect(kinds(score(withGated(gatedRow('latin-short', 1000, 500000, 500600))))).toEqual([
      'width',
    ]);
  });

  it('A5 fails when a rival lands inside the tolerance where the shipped reading does not', () => {
    const beaten = runWith([
      faceWith([
        {
          sample: 'latin-spread',
          px: 1000,
          browser: 15291,
          readings: { shipped: 15291 + 30, exact: 15291, unkerned: 15291 + 100 },
        },
        gatedRow('latin-short', 16, 100, 100),
        arabicRow(16, 100, 1.35),
      ]),
    ]);
    const verdict = score(beaten);
    expect(kinds(verdict)).toEqual(['width', 'rival-beats-shipped']);
    expect(verdict.failures[1]?.subject).toBe('exact');
  });

  it('A5 fails when no row tells the kerned reading from the unkerned one', () => {
    const flat = runWith([
      faceWith([
        {
          sample: 'latin-short',
          px: 16,
          browser: 100,
          readings: { shipped: 100, exact: 100, unkerned: 100 },
        },
        arabicRow(16, 100, 1.35),
      ]),
    ]);
    expect(kinds(score(flat))).toEqual(['kerning-unseparated']);
  });

  it('A6 fails when no reading of the box fits the face, and names every one it scored', () => {
    // ipag.ttf: 1802/401 units on a 2048 em against the browser's 880/120.
    const wrong = runWith([
      faceWith(
        [gatedRow('latin-short', 16, 100, 100), arabicRow(16, 100, 1.35)],
        {
          ascent: 1802,
          descent: 401,
        },
        { ascent: 880, descent: 120 },
      ),
    ]);
    const verdict = score(wrong);
    expect(kinds(verdict)).toEqual(['face-box']);
    expect(verdict.faceBox.agreed).toBe(0);
    expect(verdict.failures[0]?.detail).toContain('browser 880/120');
    expect(verdict.failures[0]?.detail).toContain('no reading of the face box fits all 1 face(s)');
    expect(verdict.failures[0]?.detail).toContain('hhea.ascender/descender 0/1');
  });

  it('A6 fails to a rival, not to the face, when one reading fits and the shipped one does not', () => {
    // The reading ADR 0044 leaves open: 1802/246 is one rounding from 880/120,
    // and 1802/401 - what usWin says and the reader takes - is 75.8 px away.
    const split = runWith([
      faceWith(
        [gatedRow('latin-short', 16, 100, 100), arabicRow(16, 100, 1.35)],
        { ascent: 1802, descent: 401 },
        { ascent: 880, descent: 120 },
        {
          hhea: { ascent: 1802, descent: 246 },
          usWin: { ascent: 1802, descent: 401 },
          sTypo: { ascent: 1500, descent: 300 },
          useTypoMetrics: false,
        },
      ),
    ]);
    const verdict = score(split);
    expect(kinds(verdict)).toEqual(['face-box-rival', 'face-box-rival']);
    expect(verdict.failures.map((failure) => failure.subject)).toEqual([
      'hhea.ascender/descender',
      'hhea, or sTypo when fsSelection bit 7 is set',
    ]);
    expect(verdict.failures[0]?.detail).toContain('fits 1/1 face box(es)');
    const byName = new Map(verdict.faceBox.readings.map((entry) => [entry.reading, entry]));
    expect(byName.get(SHIPPED_BOX)).toMatchObject({ fits: 0, of: 1, separates: 0 });
    expect(byName.get('hhea.ascender/descender')).toMatchObject({ fits: 1, of: 1, separates: 1 });
  });

  it('A6 counts a reading that fits every face carrying its table as not fitting the run', () => {
    // One face has no OS/2, so usWin is scored over the other alone and lands
    // 1/1 - which is not the same claim as fitting the run.
    const rows = [gatedRow('latin-short', 16, 100, 100), arabicRow(16, 100, 1.35)];
    const right = faceWith(rows, READER_BOX, BROWSER_BOX, {
      hhea: { ascent: 1500, descent: 500 },
      usWin: READER_BOX,
      sTypo: READER_BOX,
      useTypoMetrics: false,
    });
    const bare = {
      ...faceWith(
        rows,
        { ascent: 1802, descent: 401 },
        { ascent: 880, descent: 120 },
        {
          hhea: { ascent: 1802, descent: 401 },
          usWin: null,
          sTypo: null,
          useTypoMetrics: false,
        },
      ),
      file: '/usr/share/fonts/truetype/probe/Bare.ttf',
      sha256: 'b'.repeat(64),
    };
    const verdict = score(runWith([right, bare]));
    const byName = new Map(verdict.faceBox.readings.map((entry) => [entry.reading, entry]));
    expect(byName.get('OS/2.usWinAscent/usWinDescent')).toMatchObject({ fits: 1, of: 1 });
    expect(byName.get(SHIPPED_BOX)).toMatchObject({ fits: 1, of: 2 });
    expect(kinds(verdict)).toEqual(['face-box']);
  });

  it('A6 fails when a face recorded no candidate pairs, so no reading can be scored', () => {
    const blind = runWith([
      faceWith(
        [gatedRow('latin-short', 16, 100, 100), arabicRow(16, 100, 1.35)],
        READER_BOX,
        BROWSER_BOX,
        null,
      ),
    ]);
    const verdict = score(blind);
    expect(kinds(verdict)).toEqual(['unmeasured']);
    expect(verdict.failures[0]?.subject).toBe('FaceMetrics.candidates');
    expect(verdict.faceBox.scorable).toBe(0);
    expect(verdict.faceBox.agreed).toBe(1);
  });

  it('A6 fails three quarters of a pixel out, which one rounding cannot cost', () => {
    const outside = { ascent: BROWSER_BOX.ascent + 0.75, descent: BROWSER_BOX.descent };
    const off = runWith([
      faceWith([gatedRow('latin-short', 16, 100, 100), arabicRow(16, 100, 1.35)], {
        ascent: (outside.ascent * EM) / BOX_PX,
        descent: (outside.descent * EM) / BOX_PX,
      }),
    ]);
    expect(kinds(score(off))).toEqual(['face-box']);
    expect(BOX_TOLERANCE).toBe(0.5);
  });

  it('A7 fails when a recorded sample measures inside the quantisation it is pricing', () => {
    const invisible = runWith([
      faceWith([gatedRow('latin-short', 16, 100, 100), arabicRow(16, 100, 1)]),
    ]);
    const verdict = score(invisible);
    expect(kinds(verdict)).toEqual(['shaping-invisible']);
    expect(verdict.failures[0]?.subject).toBe('arabic');
  });

  it('A7 passes, and records nothing, when no face on the image covers the script', () => {
    const none = runWith([faceWith([gatedRow('latin-short', 16, 100, 100)])]);
    const verdict = score(none);
    expect(verdict.failures).toEqual([]);
    expect(verdict.shaping).toEqual([
      { sample: 'arabic', faces: 0, medianRatio: 0, worstRatio: 0, outsideTolerance: 0, of: 0 },
    ]);
  });
});
