/**
 * Experiment T7 - which typeface does a run actually get drawn in?
 *
 * Sub-phase 3.7. Reasoning and results in
 * docs/adr/phase-3-text/0033-font-substitution-and-the-guard.md.
 */

/** Hundredths of a point, as `a:rPr/@sz` writes it. */
export type Hundredths = number;

/** The four typeface slots on `CT_TextCharacterProperties`. */
export type Slot = 'latin' | 'ea' | 'cs' | 'sym';

/** `CT_TextFont`: a name and the three hints that travel with it. */
export interface Face {
  readonly typeface: string;
  readonly panose?: string | undefined;
  /** `(family << 4) | pitch`. */
  readonly pitchFamily?: number | undefined;
  /** Signed, as the attribute is: Shift-JIS is -128. */
  readonly charset?: number | undefined;
}

export type ProbeKind = 'absent' | 'installed' | 'hint' | 'slot' | 'width' | 'theme' | 'alias';

export interface Probe {
  readonly id: string;
  readonly kind: ProbeKind;
  readonly sz: Hundredths;
  readonly text: string;
  /** The slots the run declares. `latin` is always one of them. */
  readonly faces: Readonly<Partial<Record<Slot, Face>>>;
  /** The face this probe must be drawn in, where another probe settles it. */
  readonly control?: string | undefined;
}

/** The `a:fontScheme` a probe package is built with. */
export interface ThemeFonts {
  readonly majorLatin: string;
  readonly minorLatin: string;
  readonly minorEa: string;
  readonly minorCs: string;
}

export const DEFAULT_THEME_FONTS: ThemeFonts = {
  majorLatin: 'Calibri Light',
  minorLatin: 'Calibri',
  minorEa: '',
  minorCs: '',
};

/**
 * Twenty names no machine has, which is also the plan's verification target -
 * "a deck referencing 20 absent fonts reports all 20".
 *
 * The style word is the variable: a mapper that reads the name has to send
 * `Zzz Probe Mincho` somewhere different from `Zzz Probe Mono`.
 */
export const ABSENT_FACES: readonly string[] = [
  'Zzz Probe Plain',
  'Zzz Probe Serif',
  'Zzz Probe Sans',
  'Zzz Probe Sans Serif',
  'Zzz Probe Mono',
  'Zzz Probe Monospace',
  'Zzz Probe Gothic',
  'Zzz Probe Mincho',
  'Zzz Probe Song',
  'Zzz Probe Black',
  'Zzz Probe Light',
  'Zzz Probe Condensed',
  'Zzz Probe Display',
  'Zzz Probe Script',
  'Zzz Probe Text',
  'Zzz Probe Grotesk',
  'Zzz Probe Roman',
  'Zzz Probe Narrow',
  'Zzz Probe Pro',
  'Zzz Probe UI',
];

/** The one absent name every question that is not about names is asked with. */
export const BASE_ABSENT = 'Zzz Probe Plain';

/** Faces this machine has. */
export const INSTALLED_FACES: readonly string[] = [
  'Arial',
  'Times New Roman',
  'Courier New',
  'Georgia',
  'Verdana',
  'Tahoma',
  'Trebuchet MS',
  'Consolas',
  'Calibri',
  'Calibri Light',
  'Cambria',
  'Segoe UI',
  'MS Gothic',
  'Leelawadee UI',
];

/** Names the Windows font mapper resolves to a file with another name. */
export const ALIAS_FACES: readonly string[] = ['Helvetica', 'Times', 'Courier', 'MS Sans Serif'];

/** `(family << 4) | pitch`, spanning every GDI family. */
export const PITCH_FAMILIES: readonly number[] = [0, 2, 17, 18, 34, 49, 50, 66, 82];

/** PANOSE strings copied from the faces whose shape they describe. */
export const PANOSE: readonly { readonly key: string; readonly value: string }[] = [
  { key: 'serif', value: '02020603050405020304' },
  { key: 'sans', value: '020B0604020202020204' },
  { key: 'mono', value: '02070309020205020404' },
  { key: 'script', value: '03010101010101010101' },
  { key: 'zero', value: '00000000000000000000' },
];

/** Signed, as `@charset` is: ANSI, Symbol, Greek, Arabic, Shift-JIS. */
export const CHARSETS: readonly number[] = [0, 2, -95, -78, -128];

export const SZ: Hundredths = 3200;

/** One string per script, so a slot probe carries text only that slot draws. */
export const SCRIPT_TEXT = {
  latin: 'Hamburgefonstiv',
  ea: '日本語テキスト',
  cs: 'สวัสดี',
  cs2: 'مرحبا',
  sym: '',
} as const;

export const WIDTH_TEXT: readonly string[] = ['Hamburgefonstiv', 'iiiii', 'WWWWW', '0123456789'];
export const WIDTH_SIZES: readonly Hundredths[] = [1200, 3200, 5400];

/** The candidates a width comparison scores the absent face against. */
export const WIDTH_FACES: readonly string[] = [
  BASE_ABSENT,
  'Calibri',
  'Arial',
  'Times New Roman',
  'Segoe UI',
];

const face = (typeface: string, rest: Omit<Face, 'typeface'> = {}): Face => ({ typeface, ...rest });

/** Every absent name as Latin text: what does PowerPoint put there? */
export function absentProbes(): Probe[] {
  return ABSENT_FACES.map((typeface, i) => ({
    id: `absent-${String(i).padStart(2, '0')}`,
    kind: 'absent' as const,
    sz: SZ,
    text: SCRIPT_TEXT.latin,
    faces: { latin: face(typeface) },
  }));
}

/** The same reading for installed faces, which is what calibrates the reader. */
export function installedProbes(): Probe[] {
  return INSTALLED_FACES.map((typeface, i) => ({
    id: `installed-${String(i).padStart(2, '0')}`,
    kind: 'installed' as const,
    sz: SZ,
    text: SCRIPT_TEXT.latin,
    faces: { latin: face(typeface) },
    control: typeface,
  }));
}

/**
 * One absent name, every hint varied one at a time.
 *
 * The name is held constant, so a difference in the reading is a difference the
 * hint made.
 */
export function hintProbes(): Probe[] {
  const out: Probe[] = [];
  for (const pitchFamily of PITCH_FAMILIES) {
    out.push({
      id: `hint-pf-${String(pitchFamily)}`,
      kind: 'hint',
      sz: SZ,
      text: SCRIPT_TEXT.latin,
      faces: { latin: face(BASE_ABSENT, { pitchFamily }) },
    });
  }
  for (const { key, value } of PANOSE) {
    out.push({
      id: `hint-panose-${key}`,
      kind: 'hint',
      sz: SZ,
      text: SCRIPT_TEXT.latin,
      faces: { latin: face(BASE_ABSENT, { panose: value }) },
    });
  }
  for (const charset of CHARSETS) {
    out.push({
      id: `hint-charset-${String(charset)}`,
      kind: 'hint',
      sz: SZ,
      text: SCRIPT_TEXT.latin,
      faces: { latin: face(BASE_ABSENT, { charset }) },
    });
  }
  // Every hint agreeing: the strongest push a file can give the mapper.
  out.push({
    id: 'hint-all-mono',
    kind: 'hint',
    sz: SZ,
    text: SCRIPT_TEXT.latin,
    faces: {
      latin: face(BASE_ABSENT, { panose: '02070309020205020404', pitchFamily: 49, charset: 0 }),
    },
  });
  out.push({
    id: 'hint-all-serif',
    kind: 'hint',
    sz: SZ,
    text: SCRIPT_TEXT.latin,
    faces: {
      latin: face(BASE_ABSENT, { panose: '02020603050405020304', pitchFamily: 18, charset: 0 }),
    },
  });
  return out;
}

/** An absent face in each non-Latin slot, with text only that slot draws. */
export function slotProbes(): Probe[] {
  const out: Probe[] = [];
  const charsetOf = { ea: -128, cs: -34, sym: 2 } as const;
  for (const slot of ['ea', 'cs', 'sym'] as const) {
    out.push({
      id: `slot-${slot}-absent`,
      kind: 'slot',
      sz: SZ,
      text: SCRIPT_TEXT[slot],
      faces: { latin: face('Arial'), [slot]: face(BASE_ABSENT) },
    });
    out.push({
      id: `slot-${slot}-absent-charset`,
      kind: 'slot',
      sz: SZ,
      text: SCRIPT_TEXT[slot],
      faces: { latin: face('Arial'), [slot]: face(BASE_ABSENT, { charset: charsetOf[slot] }) },
    });
  }
  out.push({
    id: 'slot-ea-present',
    kind: 'slot',
    sz: SZ,
    text: SCRIPT_TEXT.ea,
    faces: { latin: face('Arial'), ea: face('MS Gothic') },
    control: 'MS Gothic',
  });
  out.push({
    id: 'slot-cs-present',
    kind: 'slot',
    sz: SZ,
    text: SCRIPT_TEXT.cs,
    faces: { latin: face('Arial'), cs: face('Leelawadee UI') },
    control: 'Leelawadee UI',
  });
  return out;
}

/**
 * The same strings in the absent face and in every candidate substitute.
 *
 * Both readings come from one instrument, so a constant bias in it cancels and
 * only the equality is read.
 */
export function widthProbes(): Probe[] {
  const out: Probe[] = [];
  for (const typeface of WIDTH_FACES) {
    for (const sz of WIDTH_SIZES) {
      for (const [i, text] of WIDTH_TEXT.entries()) {
        out.push({
          id: `width-${typeface.replace(/\s+/g, '')}-${String(sz)}-${String(i)}`,
          kind: 'width',
          sz,
          text,
          faces: { latin: face(typeface) },
        });
      }
    }
  }
  return out;
}

/** `+mj-lt` and its siblings, against a theme that is present, empty or absent. */
export function themeProbes(): Probe[] {
  return [
    {
      id: 'theme-major-latin',
      kind: 'theme',
      sz: SZ,
      text: SCRIPT_TEXT.latin,
      faces: { latin: face('+mj-lt') },
    },
    {
      id: 'theme-minor-latin',
      kind: 'theme',
      sz: SZ,
      text: SCRIPT_TEXT.latin,
      faces: { latin: face('+mn-lt') },
    },
    {
      id: 'theme-minor-ea',
      kind: 'theme',
      sz: SZ,
      text: SCRIPT_TEXT.ea,
      faces: { latin: face('+mn-lt'), ea: face('+mn-ea') },
    },
    {
      id: 'theme-minor-cs',
      kind: 'theme',
      sz: SZ,
      text: SCRIPT_TEXT.cs,
      faces: { latin: face('+mn-lt'), cs: face('+mn-cs') },
    },
    // The controls: the theme's Latin face put in the slot by hand, which is
    // what "an empty collection falls back to the Latin one" predicts.
    {
      id: 'theme-ea-as-latin',
      kind: 'theme',
      sz: SZ,
      text: SCRIPT_TEXT.ea,
      faces: { latin: face('+mn-lt'), ea: face('+mn-lt') },
    },
    {
      id: 'theme-cs-as-latin',
      kind: 'theme',
      sz: SZ,
      text: SCRIPT_TEXT.cs,
      faces: { latin: face('+mn-lt'), cs: face('+mn-lt') },
    },
    {
      id: 'theme-minor-cs-arabic',
      kind: 'theme',
      sz: SZ,
      text: SCRIPT_TEXT.cs2,
      faces: { latin: face('+mn-lt'), cs: face('+mn-cs') },
    },
    {
      id: 'theme-cs-as-latin-arabic',
      kind: 'theme',
      sz: SZ,
      text: SCRIPT_TEXT.cs2,
      faces: { latin: face('+mn-lt'), cs: face('+mn-lt') },
    },
  ];
}

export function aliasProbes(): Probe[] {
  return ALIAS_FACES.map((typeface, i) => ({
    id: `alias-${String(i)}`,
    kind: 'alias' as const,
    sz: SZ,
    text: SCRIPT_TEXT.latin,
    faces: { latin: face(typeface) },
  }));
}

/** One package per theme, because a package has exactly one theme. */
export interface ProbePackage {
  readonly key: string;
  readonly themeFonts: ThemeFonts;
  readonly probes: readonly Probe[];
}

export function allPackages(): ProbePackage[] {
  return [
    {
      key: 'faces',
      themeFonts: DEFAULT_THEME_FONTS,
      probes: [...absentProbes(), ...installedProbes(), ...aliasProbes()],
    },
    { key: 'hints', themeFonts: DEFAULT_THEME_FONTS, probes: hintProbes() },
    { key: 'slots', themeFonts: DEFAULT_THEME_FONTS, probes: slotProbes() },
    { key: 'widths', themeFonts: DEFAULT_THEME_FONTS, probes: widthProbes() },
    {
      key: 'theme-present',
      themeFonts: {
        majorLatin: 'Georgia',
        minorLatin: 'Verdana',
        minorEa: 'MS Gothic',
        minorCs: 'Leelawadee UI',
      },
      probes: themeProbes(),
    },
    {
      key: 'theme-empty',
      themeFonts: { majorLatin: 'Georgia', minorLatin: 'Verdana', minorEa: '', minorCs: '' },
      probes: themeProbes(),
    },
    {
      key: 'theme-empty-alt',
      themeFonts: { majorLatin: 'Cambria', minorLatin: 'Courier New', minorEa: '', minorCs: '' },
      probes: themeProbes(),
    },
    {
      key: 'theme-absent',
      themeFonts: {
        majorLatin: BASE_ABSENT,
        minorLatin: BASE_ABSENT,
        minorEa: BASE_ABSENT,
        minorCs: BASE_ABSENT,
      },
      probes: themeProbes(),
    },
  ];
}
