/**
 * Experiment T6 - the probe table for anchors, insets and vertical text.
 *
 * Every question is posed as a difference between two probes that differ in one
 * attribute, so the control carries the unknown. See ADR 0032.
 */

/* -------------------------------------------------------------------------- */
/* the vocabulary                                                             */
/* -------------------------------------------------------------------------- */

/** `ST_TextAnchoringType`. */
export type Anchor = 't' | 'ctr' | 'b' | 'just' | 'dist';

/** `ST_TextVerticalType`. `wordArtVertRtl` has no `MsoTextOrientation` spelling. */
export type VertType =
  'horz' | 'vert' | 'vert270' | 'wordArtVert' | 'eaVert' | 'mongolianVert' | 'wordArtVertRtl';

export const VERT_TYPES: readonly VertType[] = [
  'horz',
  'vert',
  'vert270',
  'wordArtVert',
  'eaVert',
  'mongolianVert',
  'wordArtVertRtl',
];

export const ANCHORS: readonly Anchor[] = ['t', 'ctr', 'b', 'just', 'dist'];

/** `a:spcBef` / `a:spcAft` / `a:lnSpc`, as the file states them. */
export interface Space {
  readonly kind: 'percent' | 'points';
  /** Thousandths of a percent, or hundredths of a point. */
  readonly value: number;
}

/** A bullet, only for the families that need one drawn. */
export interface ProbeBullet {
  readonly kind: 'none' | 'char' | 'autonum';
  readonly char?: string | undefined;
  readonly font?: string | undefined;
  readonly scheme?: string | undefined;
}

/** Every attribute of `a:bodyPr` this experiment can write. */
export interface Frame {
  readonly anchor?: Anchor | undefined;
  readonly anchorCtr?: boolean | undefined;
  /** Points. Written as EMU. `null` means "state no attribute at all". */
  readonly lIns?: number | null | undefined;
  readonly tIns?: number | null | undefined;
  readonly rIns?: number | null | undefined;
  readonly bIns?: number | null | undefined;
  readonly vert?: VertType | undefined;
  readonly wrap?: 'square' | 'none' | undefined;
  readonly vertOverflow?: 'overflow' | 'ellipsis' | 'clip' | undefined;
  readonly horzOverflow?: 'overflow' | 'clip' | undefined;
  /** Degrees. Written as 60000ths. */
  readonly rot?: number | undefined;
  readonly upright?: boolean | undefined;
  readonly numCol?: number | undefined;
  /** Points. */
  readonly spcCol?: number | undefined;
  readonly rtlCol?: boolean | undefined;
  readonly spcFirstLastPara?: boolean | undefined;
  readonly autofit?: 'none' | 'norm' | 'spAutoFit' | undefined;
  readonly fromWordArt?: boolean | undefined;
  readonly compatLnSpc?: boolean | undefined;
}

/** One `a:p`. */
export interface Para {
  /**
   * The text. Split into that many lines by `a:br` when `lines` is given, so a
   * paragraph's line count is a property of the file rather than of 3.3's
   * break model.
   */
  readonly text: string;
  /** How many `a:br`-separated lines this paragraph holds. Defaults to 1. */
  readonly lines?: number | undefined;
  /** Hundredths of a point. Defaults to `SZ`. */
  readonly sz?: number | undefined;
  readonly face?: string | undefined;
  readonly algn?: 'l' | 'ctr' | 'r' | 'just' | 'dist' | undefined;
  /** Points. */
  readonly marL?: number | undefined;
  readonly indent?: number | undefined;
  readonly spcBef?: Space | undefined;
  readonly spcAft?: Space | undefined;
  readonly lnSpc?: Space | undefined;
  readonly bullet?: ProbeBullet | undefined;
  /** No runs at all - the empty paragraph whose height `a:endParaRPr` sets. */
  readonly empty?: boolean | undefined;
  /** `a:endParaRPr/@sz`, hundredths of a point. */
  readonly endSz?: number | undefined;
  /** `a:endParaRPr/@sz` is written only when this is true; `endSz` needs it. */
  readonly endFace?: string | undefined;
  /** The size on the `a:br`'s own `a:rPr`, when it differs from the runs'. */
  readonly breakSz?: number | undefined;
}

export interface Rect {
  readonly x: number;
  readonly y: number;
  readonly w: number;
  readonly h: number;
}

export interface Probe {
  readonly id: string;
  readonly family: string;
  readonly frame: Frame;
  readonly paras: readonly Para[];
  /** Box size in points. The builder assigns `x` and `y`. */
  readonly w: number;
  readonly h: number;
  /** A slide of its own, read from the EMF, which names no shapes. */
  readonly isolate?: boolean | undefined;
  /** A package of its own: a repaired deck measures the repair, not the probe. */
  readonly hostile?: boolean | undefined;
  /** What the analysis keys on. */
  readonly vars: Readonly<Record<string, string | number | boolean | null>>;
}

/* -------------------------------------------------------------------------- */
/* held still                                                                 */
/* -------------------------------------------------------------------------- */

/** Hundredths of a point. Arial at 18 has a 21.6pt line box, measured in T2. */
export const SZ = 1800;
/** The line box at `SZ`, in points. 1.2 x the size, measured in T2. */
export const LINE = 21.6;
export const FACE = 'Arial';

/** Zero on every edge, written explicitly. */
export const NO_INSETS: Frame = { lIns: 0, tIns: 0, rIns: 0, bIns: 0 };

/** Asymmetric, so that a centred block distinguishes the inset box from the shape box. */
export const SKEW_INSETS: Frame = { lIns: 13, tIns: 20, rIns: 3, bIns: 5 };

/** The strings. Short, ASCII, and distinct so a reading names its probe. */
const ONE = 'Ax';
const LONG = 'Alpha Bravo Charlie Delta Echo Foxtrot Golf Hotel India Juliet Kilo Lima';
const CJK = '日本語の文';
const CJK_FACE = 'MS Gothic';

const noBullet: ProbeBullet = { kind: 'none' };

function para(text: string, extra: Partial<Para> = {}): Para {
  return { text, algn: 'l', marL: 0, indent: 0, bullet: noBullet, ...extra };
}

/** `n` paragraphs of one line each, numbered so a reading names its line. */
function paras(n: number, extra: Partial<Para> = {}): Para[] {
  return Array.from({ length: n }, (_, i) => para(`P${String(i + 1)}`, extra));
}

/* -------------------------------------------------------------------------- */
/* F1 - control: what the instrument reads when nothing is asked              */
/* -------------------------------------------------------------------------- */

/** Calibration: the block corner, the block height per line count, size and face. */
export function controlProbes(): Probe[] {
  const out: Probe[] = [];
  for (const lines of [1, 2, 3, 4, 5]) {
    for (const sz of [800, 1200, 1800, 3200]) {
      out.push({
        id: `ctl-n${String(lines)}-sz${String(sz)}`,
        family: 'control',
        frame: { ...NO_INSETS, anchor: 't' },
        paras: [para('Ax', { lines, sz })],
        w: 200,
        h: 260,
        vars: { lines, sz, face: FACE, kind: 'lines' },
      });
    }
  }
  for (const face of ['Arial', 'Times New Roman', 'Courier New', 'Verdana', 'Georgia', 'Tahoma']) {
    for (const sz of [1200, 1800, 2800]) {
      out.push({
        id: `ctl-face-${face.replace(/\s+/g, '')}-sz${String(sz)}`,
        family: 'control',
        frame: { ...NO_INSETS, anchor: 't' },
        paras: [para('Ax', { lines: 3, sz, face })],
        w: 200,
        h: 260,
        vars: { lines: 3, sz, face, kind: 'face' },
      });
    }
  }
  // Paragraphs at the same line count as the a:br block above: what a paragraph
  // boundary costs, which every anchor probe then assumes.
  for (const n of [1, 2, 3, 4]) {
    out.push({
      id: `ctl-paras${String(n)}`,
      family: 'control',
      frame: { ...NO_INSETS, anchor: 't' },
      paras: paras(n),
      w: 200,
      h: 260,
      vars: { lines: n, sz: SZ, face: FACE, kind: 'paras' },
    });
  }
  return out;
}

/* -------------------------------------------------------------------------- */
/* F2 - insets                                                                */
/* -------------------------------------------------------------------------- */

const INSET_VALUES = [0, 3.6, 7.2, 14.4, 28.8, 72] as const;

/**
 * One edge at a time, each read by what it can move: `lIns` and `tIns` off the
 * block corner, `rIns` at `algn="r"`, `bIns` at `anchor="b"`.
 */
export function insetProbes(): Probe[] {
  const out: Probe[] = [];
  for (const v of INSET_VALUES) {
    out.push({
      id: `ins-l-${String(v)}`,
      family: 'inset',
      frame: { ...NO_INSETS, lIns: v, anchor: 't' },
      paras: [para(ONE)],
      w: 220,
      h: 120,
      vars: { edge: 'l', value: v, probe: 'left' },
    });
    out.push({
      id: `ins-t-${String(v)}`,
      family: 'inset',
      frame: { ...NO_INSETS, tIns: v, anchor: 't' },
      paras: [para(ONE)],
      w: 220,
      h: 120,
      vars: { edge: 't', value: v, probe: 'top' },
    });
    out.push({
      id: `ins-r-${String(v)}`,
      family: 'inset',
      frame: { ...NO_INSETS, rIns: v, anchor: 't' },
      paras: [para(ONE, { algn: 'r' })],
      w: 220,
      h: 120,
      vars: { edge: 'r', value: v, probe: 'right' },
    });
    out.push({
      id: `ins-b-${String(v)}`,
      family: 'inset',
      frame: { ...NO_INSETS, bIns: v, anchor: 'b' },
      paras: [para(ONE)],
      w: 220,
      h: 120,
      vars: { edge: 'b', value: v, probe: 'bottom' },
    });
  }

  // Every edge absent, and each edge absent on its own with the others zeroed.
  out.push({
    id: 'ins-absent-all',
    family: 'inset',
    frame: { anchor: 't' },
    paras: [para(ONE)],
    w: 220,
    h: 120,
    vars: { edge: 'all', value: -1, probe: 'absent' },
  });
  out.push({
    id: 'ins-absent-all-r',
    family: 'inset',
    frame: { anchor: 't' },
    paras: [para(ONE, { algn: 'r' })],
    w: 220,
    h: 120,
    vars: { edge: 'all', value: -1, probe: 'absent-right' },
  });
  out.push({
    id: 'ins-absent-all-b',
    family: 'inset',
    frame: { anchor: 'b' },
    paras: [para(ONE)],
    w: 220,
    h: 120,
    vars: { edge: 'all', value: -1, probe: 'absent-bottom' },
  });
  for (const edge of ['l', 't', 'r', 'b'] as const) {
    const frame: Record<string, number | null> = { lIns: 0, tIns: 0, rIns: 0, bIns: 0 };
    frame[`${edge}Ins`] = null;
    out.push({
      id: `ins-absent-${edge}`,
      family: 'inset',
      frame: {
        ...(frame as Pick<Frame, 'lIns' | 'tIns' | 'rIns' | 'bIns'>),
        anchor: edge === 'b' ? 'b' : 't',
      },
      paras: [para(ONE, { algn: edge === 'r' ? 'r' : 'l' })],
      w: 220,
      h: 120,
      vars: { edge, value: -1, probe: `absent-${edge}` },
    });
  }

  // The wrap width, which is the other thing `lIns` and `rIns` decide. Read as
  // a line count rather than a position, so it is a reading of a different kind
  // on the same attribute.
  for (const [l, r] of [
    [0, 0],
    [36, 0],
    [0, 36],
    [36, 36],
    [72, 72],
  ] as const) {
    out.push({
      id: `ins-wrap-l${String(l)}-r${String(r)}`,
      family: 'inset',
      frame: { ...NO_INSETS, lIns: l, rIns: r, anchor: 't', wrap: 'square' },
      paras: [para(LONG)],
      w: 300,
      h: 300,
      vars: { edge: 'lr', value: l + r, probe: 'wrap', lIns: l, rIns: r },
    });
  }
  return out;
}

/** Insets no sane file holds, one package each. */
export function hostileInsetProbes(): Probe[] {
  const out: Probe[] = [];
  for (const [edge, value] of [
    ['l', -18],
    ['t', -18],
    ['r', -18],
    ['b', -18],
  ] as const) {
    const frame: Record<string, number> = { lIns: 0, tIns: 0, rIns: 0, bIns: 0 };
    frame[`${edge}Ins`] = value;
    out.push({
      id: `ins-neg-${edge}`,
      family: 'inset-hostile',
      frame: {
        ...(frame as Pick<Frame, 'lIns' | 'tIns' | 'rIns' | 'bIns'>),
        anchor: edge === 'b' ? 'b' : 't',
      },
      paras: [para(ONE, { algn: edge === 'r' ? 'r' : 'l' })],
      w: 220,
      h: 120,
      hostile: true,
      vars: { edge, value, probe: 'negative' },
    });
  }
  out.push({
    id: 'ins-huge-lr',
    family: 'inset-hostile',
    frame: { ...NO_INSETS, lIns: 150, rIns: 150, anchor: 't' },
    paras: [para(ONE)],
    w: 220,
    h: 120,
    hostile: true,
    vars: { edge: 'lr', value: 300, probe: 'wider-than-the-box' },
  });
  out.push({
    id: 'ins-huge-tb',
    family: 'inset-hostile',
    frame: { ...NO_INSETS, tIns: 90, bIns: 90, anchor: 'ctr' },
    paras: [para(ONE)],
    w: 220,
    h: 120,
    hostile: true,
    vars: { edge: 'tb', value: 180, probe: 'taller-than-the-box' },
  });
  // The symmetric pair above put the collapsed box at the midpoint, which two
  // readings both predict: clamp each edge to the centre, or split what is left
  // in proportion to the two insets. These are asymmetric, so only one can fit.
  out.push({
    id: 'ins-huge-asym-lr',
    family: 'inset-hostile',
    frame: { ...NO_INSETS, lIns: 200, rIns: 50, anchor: 't' },
    paras: [para(ONE)],
    w: 220,
    h: 120,
    hostile: true,
    vars: { edge: 'lr', value: 250, probe: 'wider-than-the-box-asymmetric' },
  });
  out.push({
    id: 'ins-huge-asym-tb',
    family: 'inset-hostile',
    frame: { ...NO_INSETS, tIns: 150, bIns: 30, anchor: 'ctr' },
    paras: [para(ONE)],
    w: 220,
    h: 120,
    hostile: true,
    vars: { edge: 'tb', value: 180, probe: 'taller-than-the-box-asymmetric' },
  });
  return out;
}

/* -------------------------------------------------------------------------- */
/* F3 - anchor                                                                */
/* -------------------------------------------------------------------------- */

/** The seven content shapes, chosen so `just` and `dist` cannot agree. */
const ANCHOR_CONTENT: readonly { readonly key: string; readonly paras: readonly Para[] }[] = [
  { key: '1p1l', paras: [para('P1')] },
  { key: '1p2l', paras: [para('P1', { lines: 2 })] },
  { key: '1p3l', paras: [para('P1', { lines: 3 })] },
  { key: '2p1l', paras: paras(2) },
  { key: '3p1l', paras: paras(3) },
  { key: '2p2l', paras: [para('P1', { lines: 2 }), para('P2', { lines: 2 })] },
  {
    key: '3p2l',
    paras: [para('P1', { lines: 2 }), para('P2', { lines: 2 }), para('P3', { lines: 2 })],
  },
];

/**
 * Five anchors x seven content shapes x two heights x two inset sets.
 *
 * One paragraph of three lines and three paragraphs of one line hold the same six
 * line boxes, which is what separates a paragraph rule from a line rule.
 */
export function anchorProbes(): Probe[] {
  const out: Probe[] = [];
  for (const anchor of ANCHORS) {
    for (const content of ANCHOR_CONTENT) {
      for (const h of [110, 260]) {
        for (const [insetKey, insets] of [
          ['zero', NO_INSETS],
          ['skew', SKEW_INSETS],
        ] as const) {
          out.push({
            id: `anc-${anchor}-${content.key}-h${String(h)}-${insetKey}`,
            family: 'anchor',
            frame: { ...insets, anchor },
            paras: content.paras,
            w: 200,
            h,
            vars: { anchor, content: content.key, boxH: h, insets: insetKey },
          });
        }
      }
    }
  }
  return out;
}

/** Anchoring with `spcBef`, `spcAft` and `lnSpc` in play, both `spcFirstLastPara` settings. */
export function anchorSpacingProbes(): Probe[] {
  const out: Probe[] = [];
  const spc: Space = { kind: 'points', value: 1200 };
  const cases: readonly { readonly key: string; readonly paras: readonly Para[] }[] = [
    { key: 'bef', paras: paras(3, { spcBef: spc }) },
    { key: 'aft', paras: paras(3, { spcAft: spc }) },
    { key: 'both', paras: paras(3, { spcBef: spc, spcAft: spc }) },
    { key: 'lnspc150', paras: paras(3, { lnSpc: { kind: 'percent', value: 150000 } }) },
    { key: 'pct-bef', paras: paras(3, { spcBef: { kind: 'percent', value: 50000 } }) },
  ];
  for (const anchor of ANCHORS) {
    for (const c of cases) {
      for (const flp of [false, true]) {
        out.push({
          id: `ancs-${anchor}-${c.key}-${flp ? 'flp1' : 'flp0'}`,
          family: 'anchor-spacing',
          frame: { ...NO_INSETS, anchor, spcFirstLastPara: flp },
          paras: c.paras,
          w: 200,
          h: 300,
          vars: { anchor, spacing: c.key, spcFirstLastPara: flp },
        });
      }
    }
  }
  return out;
}

/** Whether the last line takes a full advance or the trimmed height, per `spcFirstLastPara`. */
export function lastLineProbes(): Probe[] {
  const out: Probe[] = [];
  for (const pct of [100000, 125000, 150000, 200000, 300000]) {
    for (const lines of [1, 2, 3]) {
      for (const flp of [false, true]) {
        for (const sz of [1200, 1800]) {
          out.push({
            id: `lastln-${String(pct)}-n${String(lines)}-${flp ? 'flp1' : 'flp0'}-sz${String(sz)}`,
            family: 'last-line',
            frame: { ...NO_INSETS, anchor: 't', spcFirstLastPara: flp },
            paras: [para('Ax', { lines, sz, lnSpc: { kind: 'percent', value: pct } })],
            w: 200,
            h: 300,
            vars: { pct, lines, spcFirstLastPara: flp, sz },
          });
        }
      }
    }
  }
  return out;
}

/* -------------------------------------------------------------------------- */
/* F4 - anchorCtr                                                             */
/* -------------------------------------------------------------------------- */

/**
 * `@anchorCtr`, which is not text alignment.
 *
 * Every probe carries lines of different widths, which is the only case where
 * centring the block and centring each line differ.
 */
export function anchorCtrProbes(): Probe[] {
  const out: Probe[] = [];
  const content: readonly Para[] = [para('I'), para('Wwwwwwwwww'), para('Mmm')];
  for (const anchorCtr of [null, false, true]) {
    for (const algn of ['l', 'ctr', 'r'] as const) {
      for (const wrap of ['square', 'none'] as const) {
        const frame: Frame = {
          ...NO_INSETS,
          anchor: 't',
          wrap,
          ...(anchorCtr === null ? {} : { anchorCtr }),
        };
        out.push({
          id: `actr-${anchorCtr === null ? 'absent' : anchorCtr ? '1' : '0'}-${algn}-${wrap}`,
          family: 'anchor-ctr',
          frame,
          paras: content.map((p) => ({ ...p, algn })),
          w: 240,
          h: 140,
          vars: { anchorCtr: anchorCtr === null ? 'absent' : String(anchorCtr), algn, wrap },
        });
      }
    }
  }
  return out;
}

/* -------------------------------------------------------------------------- */
/* F5 - vertical text                                                         */
/* -------------------------------------------------------------------------- */

/**
 * Which way the text runs, and which physical edge each inset moves it from.
 *
 * The `@`-prefixed face in the drawing stream separates a rotated line from
 * upright stacked glyphs; the wrap probes say which box dimension bounds a line.
 */
export function vertProbes(): Probe[] {
  const out: Probe[] = [];
  for (const vert of [null, ...VERT_TYPES]) {
    for (const [scriptKey, text, face] of [
      ['latin', 'Wxyz', FACE],
      ['cjk', CJK, CJK_FACE],
    ] as const) {
      for (const upright of [null, true]) {
        const frame: Frame = {
          ...NO_INSETS,
          anchor: 't',
          ...(vert === null ? {} : { vert }),
          ...(upright === null ? {} : { upright }),
        };
        out.push({
          id: `vert-${vert ?? 'absent'}-${scriptKey}-${upright === null ? 'u0' : 'u1'}`,
          family: 'vert',
          frame,
          paras: [para(text, { face })],
          w: 150,
          h: 260,
          isolate: true,
          vars: {
            vert: vert ?? 'absent',
            script: scriptKey,
            upright: upright === null ? 'absent' : '1',
          },
        });
      }
    }
  }
  for (const vert of VERT_TYPES) {
    out.push({
      id: `vert-wrap-${vert}`,
      family: 'vert',
      frame: { ...NO_INSETS, anchor: 't', vert, wrap: 'square' },
      paras: [para(LONG)],
      w: 300,
      h: 100,
      isolate: true,
      vars: { vert, script: 'latin', upright: 'absent', probe: 'wrap' },
    });
  }
  // The insets, in a rotated frame. Which physical edge `lIns` moves the text
  // away from is the whole of the question, and one asymmetric set answers it.
  for (const vert of VERT_TYPES) {
    out.push({
      id: `vert-ins-${vert}`,
      family: 'vert',
      frame: { ...SKEW_INSETS, anchor: 't', vert },
      paras: [para('Wxyz')],
      w: 200,
      h: 200,
      isolate: true,
      vars: { vert, script: 'latin', upright: 'absent', probe: 'insets' },
    });
  }
  return out;
}

/** Which axis `@anchor` and `@algn` govern once the text is turned. */
export function vertAnchorProbes(): Probe[] {
  const out: Probe[] = [];
  for (const vert of ['horz', 'vert', 'vert270', 'eaVert', 'mongolianVert'] as const) {
    for (const anchor of ['t', 'ctr', 'b'] as const) {
      for (const algn of ['l', 'ctr', 'r'] as const) {
        out.push({
          id: `vanc-${vert}-${anchor}-${algn}`,
          family: 'vert-anchor',
          frame: { ...NO_INSETS, anchor, vert },
          paras: [para('Ax', { algn })],
          w: 200,
          h: 200,
          vars: { vert, anchor, algn },
        });
      }
    }
  }
  // `@anchorCtr` centres the block across the stacking axis. If that axis
  // swapped, so does this - and the same probe says which.
  for (const vert of ['horz', 'vert', 'vert270', 'eaVert', 'mongolianVert'] as const) {
    for (const anchorCtr of [false, true]) {
      out.push({
        id: `vactr-${vert}-${anchorCtr ? '1' : '0'}`,
        family: 'vert-anchor',
        frame: { ...NO_INSETS, anchor: 't', vert, anchorCtr },
        paras: [para('I'), para('Wwwwwwwwww')],
        w: 220,
        h: 220,
        vars: { vert, anchor: 't', algn: 'l', anchorCtr: anchorCtr ? '1' : '0' },
      });
    }
  }
  // `wordArtVert` stacks one character per line, and the first pass measured a
  // 23.46pt column for Arial at 18. Whether that is a multiple of the size or a
  // property of the face needs a second size and a second face.
  for (const sz of [1200, 1800, 3200]) {
    for (const face of ['Arial', 'Courier New']) {
      out.push({
        id: `vwa-${String(sz)}-${face.replace(/\s+/g, '')}`,
        family: 'vert-anchor',
        frame: { ...NO_INSETS, anchor: 't', vert: 'wordArtVert' },
        paras: [para('Wxyz', { sz, face })],
        w: 220,
        h: 260,
        vars: { vert: 'wordArtVert', anchor: 't', algn: 'l', sz, face },
      });
    }
  }
  return out;
}

/* -------------------------------------------------------------------------- */
/* F6 - overflow                                                              */
/* -------------------------------------------------------------------------- */

/**
 * `@vertOverflow` and `@horzOverflow`.
 *
 * An ellipsis is a substitution, so the string PowerPoint handed GDI is in the
 * record; clipping is not, and a bounding box reports the same for both.
 */
export function overflowProbes(): Probe[] {
  const out: Probe[] = [];
  const tall: readonly Para[] = [
    para('Line one'),
    para('Line two'),
    para('Line three'),
    para('Line four'),
    para('Line five'),
  ];
  for (const vertOverflow of [null, 'overflow', 'ellipsis', 'clip'] as const) {
    for (const anchor of ['t', 'ctr', 'b'] as const) {
      out.push({
        id: `ovf-v-${vertOverflow ?? 'absent'}-${anchor}`,
        family: 'overflow',
        frame: {
          ...NO_INSETS,
          anchor,
          ...(vertOverflow === null ? {} : { vertOverflow }),
        },
        paras: tall,
        w: 220,
        h: 46,
        isolate: true,
        vars: { axis: 'vert', value: vertOverflow ?? 'absent', anchor },
      });
    }
  }
  // The boundary: a box exactly two line boxes tall, and one on either side of
  // it. Whether a line whose bottom lands exactly on the edge is drawn is a real
  // file and not a corner case. Whole points either side, because a probe box
  // has to be a whole number of EMU.
  for (const [key, h] of [
    ['exact', 2 * LINE],
    ['under', 43],
    ['over', 44],
  ] as const) {
    for (const vertOverflow of ['ellipsis', 'clip'] as const) {
      out.push({
        id: `ovf-fit-${key}-${vertOverflow}`,
        family: 'overflow',
        frame: { ...NO_INSETS, anchor: 't', vertOverflow },
        paras: tall,
        w: 220,
        h,
        isolate: true,
        vars: { axis: 'vert', value: vertOverflow, anchor: 't', fit: key },
      });
    }
  }
  for (const horzOverflow of [null, 'overflow', 'clip'] as const) {
    out.push({
      id: `ovf-h-${horzOverflow ?? 'absent'}`,
      family: 'overflow',
      frame: {
        ...NO_INSETS,
        anchor: 't',
        wrap: 'none',
        ...(horzOverflow === null ? {} : { horzOverflow }),
      },
      paras: [para(LONG)],
      w: 120,
      h: 120,
      isolate: true,
      vars: { axis: 'horz', value: horzOverflow ?? 'absent', anchor: 't' },
    });
  }
  return out;
}

/* -------------------------------------------------------------------------- */
/* F7 - a:br                                                                  */
/* -------------------------------------------------------------------------- */

/**
 * Whether `a:br` starts a bullet, a number, a spacing gap or a first-line indent.
 *
 * Each claim has a two-paragraph control written beside it, because the reading
 * is a difference against what a real paragraph boundary costs.
 */
export function breakProbes(): Probe[] {
  const out: Probe[] = [];
  const spc: Space = { kind: 'points', value: 1500 };
  const char: ProbeBullet = { kind: 'char', char: '•', font: 'Arial' };
  const num: ProbeBullet = { kind: 'autonum', scheme: 'arabicPeriod' };

  const cases: readonly { readonly key: string; readonly paras: readonly Para[] }[] = [
    { key: 'char-br', paras: [para('Aa', { lines: 2, bullet: char })] },
    { key: 'char-2p', paras: [para('Aa', { bullet: char }), para('Aa', { bullet: char })] },
    { key: 'num-br', paras: [para('Aa', { lines: 2, bullet: num })] },
    { key: 'num-2p', paras: [para('Aa', { bullet: num }), para('Aa', { bullet: num })] },
    { key: 'num-br3', paras: [para('Aa', { lines: 3, bullet: num })] },
    {
      key: 'indent-br',
      paras: [para('Aa', { lines: 2, marL: 54, indent: -54, bullet: char })],
    },
    {
      key: 'indent-2p',
      paras: [
        para('Aa', { marL: 54, indent: -54, bullet: char }),
        para('Aa', { marL: 54, indent: -54, bullet: char }),
      ],
    },
    {
      key: 'indent-pos-br',
      paras: [para('Aa', { lines: 2, marL: 20, indent: 36, bullet: char })],
    },
    { key: 'spc-br', paras: [para('Aa', { lines: 2, spcBef: spc, spcAft: spc })] },
    {
      key: 'spc-2p',
      paras: [para('Aa', { spcBef: spc, spcAft: spc }), para('Aa', { spcBef: spc, spcAft: spc })],
    },
    { key: 'brsz-small', paras: [para('Aa', { lines: 2, breakSz: 800 })] },
    { key: 'brsz-big', paras: [para('Aa', { lines: 2, breakSz: 4000 })] },
    { key: 'brsz-none', paras: [para('Aa', { lines: 2 })] },
    {
      key: 'br-lnspc',
      paras: [para('Aa', { lines: 3, lnSpc: { kind: 'percent', value: 200000 } })],
    },
    { key: 'br-run4', paras: [para('Aa', { lines: 4 })] },
    { key: 'br-algn-r', paras: [para('Aa', { lines: 2, algn: 'r' })] },
  ];
  for (const c of cases) {
    out.push({
      id: `br-${c.key}`,
      family: 'break',
      frame: { ...NO_INSETS, anchor: 't' },
      paras: c.paras,
      w: 260,
      h: 260,
      isolate: true,
      vars: { probe: c.key },
    });
  }
  return out;
}

/* -------------------------------------------------------------------------- */
/* F8 - a:endParaRPr                                                          */
/* -------------------------------------------------------------------------- */

/** The height of an empty paragraph, differenced against the two around it. */
export function endParaProbes(): Probe[] {
  const out: Probe[] = [];
  for (const endSz of [800, 1200, 1800, 2800, 4000]) {
    out.push({
      id: `end-mid-${String(endSz)}`,
      family: 'end-para',
      frame: { ...NO_INSETS, anchor: 't' },
      paras: [para('P1'), para('', { empty: true, endSz }), para('P3')],
      w: 220,
      h: 300,
      vars: { where: 'middle', endSz, kind: 'sz' },
    });
  }
  out.push({
    id: 'end-mid-absent',
    family: 'end-para',
    frame: { ...NO_INSETS, anchor: 't' },
    paras: [para('P1'), para('', { empty: true }), para('P3')],
    w: 220,
    h: 300,
    vars: { where: 'middle', endSz: -1, kind: 'absent' },
  });
  // The empty paragraph's own `a:defRPr` size, stated on the pPr rather than on
  // the endParaRPr: two places can say how tall an empty line is and they need
  // not agree.
  for (const sz of [800, 4000]) {
    out.push({
      id: `end-mid-defrpr-${String(sz)}`,
      family: 'end-para',
      frame: { ...NO_INSETS, anchor: 't' },
      paras: [para('P1'), para('', { empty: true, sz }), para('P3')],
      w: 220,
      h: 300,
      vars: { where: 'middle', endSz: sz, kind: 'defRPr' },
    });
  }
  // Trailing and leading, where an empty paragraph changes the block and so
  // moves a centred or bottom-anchored one.
  for (const anchor of ['t', 'ctr', 'b'] as const) {
    for (const endSz of [800, 4000]) {
      out.push({
        id: `end-last-${anchor}-${String(endSz)}`,
        family: 'end-para',
        frame: { ...NO_INSETS, anchor },
        paras: [para('P1'), para('', { empty: true, endSz })],
        w: 220,
        h: 300,
        vars: { where: 'last', endSz, kind: 'sz', anchor },
      });
      out.push({
        id: `end-first-${anchor}-${String(endSz)}`,
        family: 'end-para',
        frame: { ...NO_INSETS, anchor },
        paras: [para('', { empty: true, endSz }), para('P2')],
        w: 220,
        h: 300,
        vars: { where: 'first', endSz, kind: 'sz', anchor },
      });
    }
  }
  // `a:endParaRPr` and `a:defRPr` made to disagree, both ways round, so the
  // reading is which one wins rather than which one was read.
  for (const [endSz, defSz] of [
    [800, 4000],
    [4000, 800],
  ] as const) {
    out.push({
      id: `end-conflict-${String(endSz)}-${String(defSz)}`,
      family: 'end-para',
      frame: { ...NO_INSETS, anchor: 't' },
      paras: [para('P1'), para('', { empty: true, endSz, sz: defSz }), para('P3')],
      w: 220,
      h: 300,
      vars: { where: 'middle', endSz, kind: 'conflict', defSz },
    });
  }
  // A paragraph whose only run holds the empty string is not the same file as a
  // paragraph with no run at all, and implementations routinely conflate them.
  out.push({
    id: 'end-emptyrun',
    family: 'end-para',
    frame: { ...NO_INSETS, anchor: 't' },
    paras: [para('P1'), para(''), para('P3')],
    w: 220,
    h: 300,
    vars: { where: 'middle', endSz: -1, kind: 'empty-run' },
  });
  // An empty paragraph carrying its own line spacing.
  for (const pct of [100000, 200000]) {
    out.push({
      id: `end-lnspc-${String(pct)}`,
      family: 'end-para',
      frame: { ...NO_INSETS, anchor: 't' },
      paras: [
        para('P1'),
        para('', { empty: true, endSz: 1800, lnSpc: { kind: 'percent', value: pct } }),
        para('P3'),
      ],
      w: 220,
      h: 300,
      vars: { where: 'middle', endSz: 1800, kind: 'lnSpc', lnSpc: pct },
    });
  }
  return out;
}

/* -------------------------------------------------------------------------- */
/* F9 - bodyPr/@rot                                                           */
/* -------------------------------------------------------------------------- */

/** `a:bodyPr/@rot` and `@upright`, which are the text's rotation and not the shape's. */
export function rotProbes(): Probe[] {
  const out: Probe[] = [];
  for (const rot of [null, 0, 15, 45, 90, -45, 270]) {
    for (const upright of [null, true]) {
      out.push({
        id: `rot-${rot === null ? 'absent' : String(rot)}-${upright === null ? 'u0' : 'u1'}`,
        family: 'rot',
        frame: {
          ...NO_INSETS,
          anchor: 't',
          ...(rot === null ? {} : { rot }),
          ...(upright === null ? {} : { upright }),
        },
        paras: [para('Wxyz')],
        w: 220,
        h: 220,
        isolate: true,
        vars: { rot: rot === null ? 'absent' : rot, upright: upright === null ? 'absent' : '1' },
      });
    }
  }
  for (const vert of ['vert', 'vert270', 'eaVert'] as const) {
    out.push({
      id: `rot-with-${vert}`,
      family: 'rot',
      frame: { ...NO_INSETS, anchor: 't', rot: 45, vert },
      paras: [para('Wxyz')],
      w: 220,
      h: 220,
      isolate: true,
      vars: { rot: 45, upright: 'absent', vert },
    });
  }
  return out;
}

/* -------------------------------------------------------------------------- */
/* F10 - columns                                                              */
/* -------------------------------------------------------------------------- */

/** `@numCol`, `@spcCol` and `@rtlCol`, read off the `BoundLeft` of the first paragraph in column two. */
export function columnProbes(): Probe[] {
  const out: Probe[] = [];
  const many: readonly Para[] = paras(8);
  for (const numCol of [null, 1, 2, 3]) {
    for (const spcCol of [null, 0, 18]) {
      for (const rtlCol of [null, true]) {
        out.push({
          id: `col-${numCol ?? 'absent'}-${spcCol ?? 'absent'}-${rtlCol === null ? 'l' : 'r'}`,
          family: 'columns',
          frame: {
            ...NO_INSETS,
            anchor: 't',
            ...(numCol === null ? {} : { numCol }),
            ...(spcCol === null ? {} : { spcCol }),
            ...(rtlCol === null ? {} : { rtlCol }),
          },
          paras: many,
          w: 300,
          h: 110,
          vars: {
            numCol: numCol ?? 'absent',
            spcCol: spcCol ?? 'absent',
            rtlCol: rtlCol === null ? 'absent' : '1',
          },
        });
      }
    }
  }
  return out;
}

/** Column counts outside `ST_TextColumnCount`, one package each. */
export function hostileColumnProbes(): Probe[] {
  return [
    {
      id: 'col-17',
      family: 'columns-hostile',
      frame: { ...NO_INSETS, anchor: 't', numCol: 17 },
      paras: paras(4),
      w: 300,
      h: 110,
      hostile: true,
      vars: { numCol: 17, spcCol: 'absent', rtlCol: 'absent' },
    },
    {
      id: 'col-0',
      family: 'columns-hostile',
      frame: { ...NO_INSETS, anchor: 't', numCol: 0 },
      paras: paras(4),
      w: 300,
      h: 110,
      hostile: true,
      vars: { numCol: 0, spcCol: 'absent', rtlCol: 'absent' },
    },
    {
      id: 'col-negspc',
      family: 'columns-hostile',
      frame: { ...NO_INSETS, anchor: 't', numCol: 2, spcCol: -18 },
      paras: paras(4),
      w: 300,
      h: 110,
      hostile: true,
      vars: { numCol: 2, spcCol: -18, rtlCol: 'absent' },
    },
  ];
}

/** An angle outside `ST_Angle`'s useful range, in a package of its own. */
export function hostileRotProbes(): Probe[] {
  return [
    {
      id: 'rot-720',
      family: 'rot-hostile',
      frame: { ...NO_INSETS, anchor: 't', rot: 720 },
      paras: [para('Wxyz')],
      w: 220,
      h: 220,
      isolate: true,
      hostile: true,
      vars: { rot: 720, upright: 'absent' },
    },
  ];
}

/* -------------------------------------------------------------------------- */
/* F11 - wrap                                                                 */
/* -------------------------------------------------------------------------- */

export function wrapProbes(): Probe[] {
  const out: Probe[] = [];
  for (const wrap of [null, 'square', 'none'] as const) {
    for (const algn of ['l', 'ctr', 'r'] as const) {
      out.push({
        id: `wrap-${wrap ?? 'absent'}-${algn}`,
        family: 'wrap',
        frame: { ...NO_INSETS, anchor: 't', ...(wrap === null ? {} : { wrap }) },
        paras: [para(LONG, { algn })],
        w: 200,
        h: 300,
        vars: { wrap: wrap ?? 'absent', algn },
      });
    }
  }
  return out;
}

/* -------------------------------------------------------------------------- */
/* F12 - does a:bodyPr inherit?                                               */
/* -------------------------------------------------------------------------- */

/** One property, stated at one level of the placeholder chain. */
export interface InheritProbe {
  readonly id: string;
  /** Which attribute is under test. */
  readonly property: string;
  /** Where the interesting value is stated. */
  readonly at: 'slide' | 'layout' | 'master';
  /** The `a:bodyPr` for the slide, layout and master placeholders. */
  readonly slideFrame: Frame;
  readonly layoutFrame: Frame;
  readonly masterFrame: Frame;
  readonly paras: readonly Para[];
  readonly w: number;
  readonly h: number;
  readonly vars: Readonly<Record<string, string | number | boolean | null>>;
}

/**
 * Does `a:bodyPr` cascade the way `a:rPr` does?
 *
 * Each property is stated at exactly one level of the chain, with a control that
 * states it nowhere, so the answer separates one level from any level.
 */
export function inheritProbes(): InheritProbe[] {
  const out: InheritProbe[] = [];
  const bare: Frame = {};
  const cases: readonly { readonly property: string; readonly value: Frame }[] = [
    { property: 'anchor', value: { anchor: 'b' } },
    { property: 'lIns', value: { lIns: 54 } },
    { property: 'anchorCtr', value: { anchorCtr: true } },
    { property: 'wrap', value: { wrap: 'none' } },
    { property: 'vert', value: { vert: 'vert270' } },
    { property: 'numCol', value: { numCol: 2 } },
  ];
  for (const c of cases) {
    for (const at of ['slide', 'layout', 'master'] as const) {
      out.push({
        id: `inh-${c.property}-${at}`,
        property: c.property,
        at,
        slideFrame: at === 'slide' ? c.value : bare,
        layoutFrame: at === 'layout' ? c.value : bare,
        masterFrame: at === 'master' ? c.value : bare,
        paras: [para('I'), para('Wwwwwwwwww')],
        w: 240,
        h: 200,
        vars: { property: c.property, at },
      });
    }
  }
  // The control: nothing states anything anywhere. Whatever this reads is the
  // schema default, and every "inherits" claim is a difference against it.
  for (const c of cases) {
    out.push({
      id: `inh-${c.property}-none`,
      property: c.property,
      at: 'slide',
      slideFrame: bare,
      layoutFrame: bare,
      masterFrame: bare,
      paras: [para('I'), para('Wwwwwwwwww')],
      w: 240,
      h: 200,
      vars: { property: c.property, at: 'none' },
    });
  }
  return out;
}

/* -------------------------------------------------------------------------- */
/* everything                                                                 */
/* -------------------------------------------------------------------------- */

export function allProbes(): Probe[] {
  return [
    ...controlProbes(),
    ...insetProbes(),
    ...hostileInsetProbes(),
    ...anchorProbes(),
    ...anchorSpacingProbes(),
    ...lastLineProbes(),
    ...anchorCtrProbes(),
    ...vertProbes(),
    ...vertAnchorProbes(),
    ...overflowProbes(),
    ...breakProbes(),
    ...endParaProbes(),
    ...rotProbes(),
    ...hostileRotProbes(),
    ...columnProbes(),
    ...hostileColumnProbes(),
    ...wrapProbes(),
  ];
}
