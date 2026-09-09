/**
 * Experiment T14 - each gate, shown going red.
 *
 * Every row here is authored rather than measured, so the suite fails when
 * `tools/ground-truth/fonts/real-faces/score.ts` stops deciding what it says it
 * decides. A gate nobody has seen go red is not a gate.
 */

import { describe, expect, it } from 'vitest';

import { AGREEMENT, BOX_PX } from './probes.ts';
import { score, type Row, type Run, type RunFace, type Verdict } from './score.ts';

/** 1901/2048 and 483/2048 of a 1000px em: the non-integer box a real face has. */
const EM = 2048;
const READER_BOX = { ascent: 1901, descent: 483 };
const BROWSER_BOX = { ascent: 928.22265625, descent: 235.83984375 };

const SAMPLES: Run['samples'] = [
  {
    id: 'latin-short',
    script: 'latin',
    gated: true,
    shaped: false,
    asks: 'the per-glyph error',
    text: 'HEAD BADGE',
  },
  {
    id: 'arabic',
    script: 'arabic',
    gated: false,
    shaped: true,
    asks: 'how wide an unshaped sum measures',
    text: 'العربية',
  },
];

function latinRow(px: number, browser: number, shipped: number): Row {
  return {
    sample: 'latin-short',
    px,
    browser,
    readings: {
      shipped,
      exact: browser * 1.00000001,
      advanceRounded: browser * 1.00001,
      unkerned: browser * 1.005,
      wholeStringTruncated: browser * 1.0000001,
    },
  };
}

function arabicRow(px: number, browser: number, over: number): Row {
  return {
    sample: 'arabic',
    px,
    browser,
    readings: { shipped: browser * (1 + over) },
  };
}

function faceWith(rows: readonly Row[], box = READER_BOX, browserBox = BROWSER_BOX): RunFace {
  return {
    file: '/usr/share/fonts/truetype/probe/Probe.ttf',
    sha256: 'a'.repeat(64),
    family: 'Probe Sans',
    subfamily: 'Regular',
    unitsPerEm: EM,
    metricsSource: 'usWin',
    box: { browser: browserBox, reader: box },
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
    chromium: '141.0.7390.37',
    image: { platform: 'linux', release: '6.11.0', imageOs: 'ubuntu24', imageVersion: '2025' },
    directories,
    sizes: [16, 32],
    boxPx: BOX_PX,
    agreement: AGREEMENT,
    samples: SAMPLES,
    indexed: { files: 1, faces: 1, measured: faces.length },
    skipped: [],
    faces,
  };
}

/** The measurement a green run makes: the reader exact, every rival worse. */
const GREEN = runWith([
  faceWith([
    latinRow(16, 100, 100),
    latinRow(32, 200, 200),
    arabicRow(16, 100, 0.35),
    arabicRow(32, 200, 0.36),
  ]),
]);

const kinds = (verdict: Verdict): string[] => verdict.failures.map((failure) => failure.kind);

describe('the T14 gate, green', () => {
  it('passes when the reader reproduces every gated width exactly', () => {
    expect(score(GREEN).failures).toEqual([]);
  });

  it('scores the shipped reading alone at the top', () => {
    const board = score(GREEN).scoreboard;
    expect(board.find((entry) => entry.reading === 'shipped')).toMatchObject({ exact: 2, of: 2 });
    for (const rival of board.filter((entry) => entry.reading !== 'shipped')) {
      expect(rival.exact).toBe(0);
    }
  });

  it('names the face box reading a non-integer box separates', () => {
    const { faceBox } = score(GREEN);
    expect(faceBox.answer).toBe('exact');
    expect(faceBox.variants).toEqual([
      { name: 'exact', fits: 1, of: 1 },
      { name: 'round', fits: 0, of: 1 },
      { name: 'floor', fits: 0, of: 1 },
      { name: 'ceil', fits: 0, of: 1 },
    ]);
  });

  it('prices the missing shaper without gating it', () => {
    const [arabic] = score(GREEN).shaping;
    expect(arabic).toMatchObject({ sample: 'arabic', faces: 1 });
    expect(arabic?.worstOver).toBeCloseTo(0.36, 12);
    expect(arabic?.medianOver).toBeCloseTo(0.355, 12);
  });

  it('reports a per-script row for the gated script and the recorded one', () => {
    expect(score(GREEN).byScript).toEqual([
      {
        script: 'latin',
        gated: true,
        faces: 1,
        comparisons: 2,
        exact: 2,
        medianRelative: 0,
        worstRelative: 0,
      },
      {
        script: 'arabic',
        gated: false,
        faces: 1,
        comparisons: 2,
        exact: 0,
        medianRelative: 0.355,
        worstRelative: 0.36,
      },
    ]);
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
    const only = runWith([faceWith([arabicRow(16, 100, 0.35), arabicRow(32, 200, 0.36)])]);
    expect(kinds(score(only))).toEqual(['no-gated-comparisons']);
  });

  it('A3 fails on a width a tenth of a percent out, naming both numbers', () => {
    const off = runWith([
      faceWith([
        latinRow(16, 100, 100.1),
        latinRow(32, 200, 200),
        arabicRow(16, 100, 0.35),
        arabicRow(32, 200, 0.36),
      ]),
    ]);
    const [failure] = score(off).failures;
    expect(failure?.kind).toBe('agreement');
    expect(failure?.subject).toBe('/usr/share/fonts/truetype/probe/Probe.ttf');
    expect(failure?.detail).toContain('browser 100');
    expect(failure?.detail).toContain('reader 100.1');
  });

  it('A3 holds at a residual an order under the gate, and fails an order over it', () => {
    const inside = runWith([faceWith([latinRow(16, 100, 100 * (1 + AGREEMENT / 10))])]);
    const outside = runWith([faceWith([latinRow(16, 100, 100 * (1 + AGREEMENT * 10))])]);
    expect(kinds(score(inside))).not.toContain('agreement');
    expect(kinds(score(outside))).toContain('agreement');
  });

  it('A4 fails when a rival reading reproduces more widths than the shipped one', () => {
    const beaten = runWith([
      faceWith([
        {
          sample: 'latin-short',
          px: 16,
          browser: 100,
          // Inside the agreement gate, so only exact equality separates them.
          readings: { shipped: 100.000005, exact: 100, unkerned: 100.5 },
        },
      ]),
    ]);
    const verdict = score(beaten);
    expect(kinds(verdict)).toEqual(['rival-beats-shipped']);
    expect(verdict.failures[0]?.subject).toBe('exact');
  });

  it('A4 fails on a tie, because the top scorer has to be alone there', () => {
    const tied = runWith([
      faceWith([
        {
          sample: 'latin-short',
          px: 16,
          browser: 100,
          readings: { shipped: 100, exact: 100, unkerned: 100.5 },
        },
      ]),
    ]);
    const verdict = score(tied);
    expect(kinds(verdict)).toEqual(['rival-beats-shipped']);
    expect(verdict.failures[0]?.subject).toBe('exact');
  });

  it('A5 fails when no rounding of the reader box reproduces the browser', () => {
    const wrong = runWith([
      faceWith([latinRow(16, 100, 100)], READER_BOX, { ascent: 900, descent: 235.83984375 }),
    ]);
    const verdict = score(wrong);
    expect(kinds(verdict)).toContain('face-box');
    expect(verdict.faceBox.answer).toBeNull();
    expect(verdict.failures.find((f) => f.kind === 'face-box')?.detail).toContain('browser 900');
  });

  it('A6 fails when a shaped script measures inside the Latin gate', () => {
    const flat = runWith([
      faceWith([latinRow(16, 100, 100), latinRow(32, 200, 200), arabicRow(16, 100, 0)]),
    ]);
    const verdict = score(flat);
    expect(kinds(verdict)).toEqual(['shaping-invisible']);
    expect(verdict.failures[0]?.subject).toBe('arabic');
  });

  it('A6 passes, and records nothing, when no face on the image covers the script', () => {
    const none = runWith([faceWith([latinRow(16, 100, 100), latinRow(32, 200, 200)])]);
    const verdict = score(none);
    expect(verdict.failures).toEqual([]);
    expect(verdict.shaping).toEqual([{ sample: 'arabic', faces: 0, medianOver: 0, worstOver: 0 }]);
  });
});
