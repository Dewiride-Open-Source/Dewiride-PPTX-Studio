/**
 * Experiment C2 - the colour questions 0.7-C did not ask.
 *
 * 0.7-C settled the thing that was actively contested: which space `tint` and
 * `shade` blend in, and where every transform lives. It did that with 214
 * swatches built on `a:srgbClr` and `a:schemeClr`, because a disagreement had to
 * be attributable to the transform rather than to the base.
 *
 * That left the bases themselves unmeasured, and sub-phase 2.6 cannot ship
 * without them:
 *
 *   - **`a:scrgbClr`.** Three percentages. Are they sRGB percentages, or linear
 *     light as the name "scRGB" suggests? The two answers differ by 60/255 at
 *     50% - the largest single ambiguity left anywhere in colour - and both
 *     readings have implementations behind them.
 *   - **`a:hslClr`.** A hue in sixtieths of a degree and two percentages. Is the
 *     conversion the same bi-hexcone the `lumMod` family measurably uses?
 *   - **`a:prstClr`.** 147 names, and the plan says flatly that they "are not
 *     CSS colour names". That is a claim with a table behind it, and the table
 *     is not in this repository. Asking PowerPoint produces one that is ground
 *     truth and carries no third party's copyright.
 *   - **`a:sysClr`.** `@lastClr` is preferred - 0.7-C confirmed that much
 *     through the theme's `dk1` - but what a renderer should do when there is no
 *     `lastClr` is unmeasured, and the fallback table does not exist.
 *   - **Alpha.** `a:alpha`, `a:alphaMod` and `a:alphaOff` were in the swatch
 *     type and in none of the swatches.
 *   - **The clamp boundary.** 0.7-C proved saturation is not clamped *between*
 *     consecutive HSL transforms. It did not ask what happens when a transform
 *     in another space interrupts the run, and `apply.ts` had to guess.
 *   - **A non-identity `clrMap`.** Every swatch in 0.7-C went through an
 *     identity map, which is precisely the configuration in which the
 *     `dk1`-bypasses-the-map bug is invisible.
 *
 * Each block below states the question it settles and what a wrong answer would
 * look like. Blocks that risk a package PowerPoint refuses outright - an
 * enumeration value that may not exist, an attribute grammar that may not be
 * accepted in a Transitional document - go in their own deck, so one refusal
 * costs one block rather than the whole experiment.
 */

import { CLR_SCHEME } from '../../../lib/pptx.ts';
import type { Swatch, Transform } from '../transforms/probes.ts';

/* -------------------------------------------------------------------------- */
/* the bases C2 adds                                                          */
/* -------------------------------------------------------------------------- */

export type Base2 =
  | { readonly kind: 'srgb'; readonly value: string }
  | { readonly kind: 'scheme'; readonly value: string }
  | { readonly kind: 'scrgb'; readonly r: number; readonly g: number; readonly b: number }
  | { readonly kind: 'hsl'; readonly hue: number; readonly sat: number; readonly lum: number }
  | { readonly kind: 'sys'; readonly value: string; readonly lastClr?: string }
  | { readonly kind: 'prst'; readonly value: string }
  /** A raw element, for grammar probes that no typed base can express. */
  | { readonly kind: 'raw'; readonly xml: string; readonly value: string };

export interface Swatch2 {
  readonly id: string;
  readonly deck: string;
  readonly group: string;
  readonly base: Base2;
  readonly transforms: readonly Transform[];
}

/* -------------------------------------------------------------------------- */
/* the enumerations                                                           */
/* -------------------------------------------------------------------------- */

/**
 * `ST_PresetColorVal`, the spelling with `Gray`.
 *
 * If PowerPoint refuses the deck built from this list, one of these names is not
 * in the enumeration and the refusal is the measurement - which is why they are
 * alone in a deck.
 */
export const PRESET_COLOR_NAMES: readonly string[] = [
  'aliceBlue',
  'antiqueWhite',
  'aqua',
  'aquamarine',
  'azure',
  'beige',
  'bisque',
  'black',
  'blanchedAlmond',
  'blue',
  'blueViolet',
  'brown',
  'burlyWood',
  'cadetBlue',
  'chartreuse',
  'chocolate',
  'coral',
  'cornflowerBlue',
  'cornsilk',
  'crimson',
  'cyan',
  'dkBlue',
  'dkCyan',
  'dkGoldenrod',
  'dkGray',
  'dkGreen',
  'dkKhaki',
  'dkMagenta',
  'dkOliveGreen',
  'dkOrange',
  'dkOrchid',
  'dkRed',
  'dkSalmon',
  'dkSeaGreen',
  'dkSlateBlue',
  'dkSlateGray',
  'dkTurquoise',
  'dkViolet',
  'deepPink',
  'deepSkyBlue',
  'dimGray',
  'dodgerBlue',
  'firebrick',
  'floralWhite',
  'forestGreen',
  'fuchsia',
  'gainsboro',
  'ghostWhite',
  'gold',
  'goldenrod',
  'gray',
  'green',
  'greenYellow',
  'honeydew',
  'hotPink',
  'indianRed',
  'indigo',
  'ivory',
  'khaki',
  'lavender',
  'lavenderBlush',
  'lawnGreen',
  'lemonChiffon',
  'ltBlue',
  'ltCoral',
  'ltCyan',
  'ltGoldenrodYellow',
  'ltGray',
  'ltGreen',
  'ltPink',
  'ltSalmon',
  'ltSeaGreen',
  'ltSkyBlue',
  'ltSlateGray',
  'ltSteelBlue',
  'ltYellow',
  'lime',
  'limeGreen',
  'linen',
  'magenta',
  'maroon',
  'medAquamarine',
  'medBlue',
  'medOrchid',
  'medPurple',
  'medSeaGreen',
  'medSlateBlue',
  'medSpringGreen',
  'medTurquoise',
  'medVioletRed',
  'midnightBlue',
  'mintCream',
  'mistyRose',
  'moccasin',
  'navajoWhite',
  'navy',
  'oldLace',
  'olive',
  'oliveDrab',
  'orange',
  'orangeRed',
  'orchid',
  'paleGoldenrod',
  'paleGreen',
  'paleTurquoise',
  'paleVioletRed',
  'papayaWhip',
  'peachPuff',
  'peru',
  'pink',
  'plum',
  'powderBlue',
  'purple',
  'red',
  'rosyBrown',
  'royalBlue',
  'saddleBrown',
  'salmon',
  'sandyBrown',
  'seaGreen',
  'seaShell',
  'sienna',
  'silver',
  'skyBlue',
  'slateBlue',
  'slateGray',
  'snow',
  'springGreen',
  'steelBlue',
  'tan',
  'teal',
  'thistle',
  'tomato',
  'turquoise',
  'violet',
  'wheat',
  'white',
  'whiteSmoke',
  'yellow',
  'yellowGreen',
];

/**
 * The `Grey` spellings, if they exist.
 *
 * ECMA is reported to carry both, as separate enumeration values naming the same
 * colour. "Reported" is doing real work in that sentence: nobody on this project
 * has read the schema, so these are in their own deck and a refusal answers the
 * question as cleanly as a pixel would.
 */
export const PRESET_COLOR_GREY_NAMES: readonly string[] = [
  'dkGrey',
  'dimGrey',
  'dkSlateGrey',
  'grey',
  'ltGrey',
  'ltSlateGrey',
  'slateGrey',
];

/** `ST_SystemColorVal`. Two of them start with a digit, which is legal in a value. */
export const SYSTEM_COLOR_NAMES: readonly string[] = [
  'scrollBar',
  'background',
  'activeCaption',
  'inactiveCaption',
  'menu',
  'window',
  'windowFrame',
  'menuText',
  'windowText',
  'captionText',
  'activeBorder',
  'inactiveBorder',
  'appWorkspace',
  'highlight',
  'highlightText',
  'btnFace',
  'btnShadow',
  'grayText',
  'btnText',
  'inactiveCaptionText',
  'btnHighlight',
  '3dDkShadow',
  '3dLight',
  'infoText',
  'infoBk',
  'hotLight',
  'gradientActiveCaption',
  'gradientInactiveCaption',
  'menuHighlight',
  'menuBar',
];

/* -------------------------------------------------------------------------- */
/* the probes                                                                 */
/* -------------------------------------------------------------------------- */

const pct = (n: number): number => Math.round(n * 1000);

export function swatches2(): Swatch2[] {
  const out: Swatch2[] = [];
  // One deck per group, and that is a decision the first run forced. Six decks
  // went to PowerPoint and one came back repaired, with the only sentence it
  // ever says - 'The file or directory is corrupted and unreadable' - and no
  // hint which of sixty-seven swatches caused it. A repaired deck is not a
  // measurement: PowerPoint may have rewritten anything in it. Grouping one
  // question per package means a refusal names its own cause and costs only the
  // block that asked for it.
  const add = (group: string, base: Base2, transforms: Transform[] = []): void => {
    out.push({
      id: `t${String(out.length).padStart(3, '0')}`,
      deck: group,
      group,
      base,
      transforms,
    });
  };

  /* -- `main` --------------------------------------------------------- */

  // 1 - scrgbClr. The discriminator is 50%: sRGB percentages give 808080,
  //     linear light gives BCBCBC. 0% and 100% are the control, where every
  //     reading agrees and a disagreement means the readback is broken.
  for (const v of [0, 10, 25, 50, 75, 90, 100]) {
    add('scrgb-grey', { kind: 'scrgb', r: pct(v), g: pct(v), b: pct(v) });
  }
  add('scrgb', { kind: 'scrgb', r: pct(100), g: 0, b: 0 });
  add('scrgb', { kind: 'scrgb', r: 0, g: pct(100), b: 0 });
  add('scrgb', { kind: 'scrgb', r: 0, g: 0, b: pct(100) });
  // 4472C4 written as percentages, so the answer can be compared with a swatch
  // 0.7-C already measured under a different base.
  add('scrgb', { kind: 'scrgb', r: 26667, g: 44706, b: 76863 });
  add('scrgb', { kind: 'scrgb', r: pct(20), g: pct(60), b: pct(80) });
  // Out of range in both directions. `ST_Percentage` permits it; the clamp is
  // the question.
  add('scrgb-range', { kind: 'scrgb', r: pct(150), g: pct(50), b: 0 });
  add('scrgb-range', { kind: 'scrgb', r: -pct(50), g: pct(50), b: pct(100) });

  // 2 - hslClr. Hue in sixtieths of a degree; saturation and lightness in
  //     hundred-thousandths. If the conversion is the same bi-hexcone the
  //     lum/sat transforms use, every one of these is predictable from 0.7-C.
  for (const deg of [0, 60, 120, 180, 240, 300]) {
    add('hsl-hue', { kind: 'hsl', hue: deg * 60000, sat: pct(100), lum: pct(50) });
  }
  for (const s of [0, 25, 50, 75, 100]) {
    add('hsl-sat', { kind: 'hsl', hue: 218 * 60000, sat: pct(s), lum: pct(50) });
  }
  for (const l of [0, 25, 50, 75, 100]) {
    add('hsl-lum', { kind: 'hsl', hue: 218 * 60000, sat: pct(52), lum: pct(l) });
  }
  // A hue past a full turn, and one written as a negative.
  add('hsl-range', { kind: 'hsl', hue: 400 * 60000, sat: pct(100), lum: pct(50) });
  add('hsl-range', { kind: 'hsl', hue: -60 * 60000, sat: pct(100), lum: pct(50) });

  // 3 - sysClr with a lastClr that disagrees with any plausible system value.
  //     If the painted colour is FF00FF the file wins; if it is anything else
  //     the machine does, and every deck authored on another Windows theme is a
  //     different colour here.
  add('sys-lastclr', { kind: 'sys', value: 'windowText', lastClr: 'FF00FF' });
  add('sys-lastclr', { kind: 'sys', value: 'window', lastClr: '00FF00' });
  add('sys-lastclr', { kind: 'sys', value: 'btnFace', lastClr: '123456' });

  // 4 - alpha. Read back through Fill.Transparency, not through a pixel: a
  //     translucent swatch's pixel is a composite with whatever is behind it,
  //     and what is behind it is another swatch.
  for (const v of [0, 25, 50, 75, 100]) {
    add('alpha', { kind: 'srgb', value: '4472C4' }, [{ name: 'alpha', val: pct(v) }]);
  }
  add('alpha-chain', { kind: 'srgb', value: '4472C4' }, [
    { name: 'alpha', val: pct(50) },
    { name: 'alphaMod', val: pct(50) },
  ]);
  add('alpha-chain', { kind: 'srgb', value: '4472C4' }, [
    { name: 'alpha', val: pct(50) },
    { name: 'alphaOff', val: pct(25) },
  ]);
  add('alpha-chain', { kind: 'srgb', value: '4472C4' }, [
    { name: 'alpha', val: pct(50) },
    { name: 'alphaOff', val: -pct(25) },
  ]);
  // Does an alpha transform disturb the colour, or the colour transforms the
  // alpha? Both directions, so neither can hide.
  add('alpha-mixed', { kind: 'srgb', value: '4472C4' }, [
    { name: 'alpha', val: pct(50) },
    { name: 'lumMod', val: pct(60) },
    { name: 'lumOff', val: pct(40) },
  ]);
  add('alpha-mixed', { kind: 'srgb', value: '4472C4' }, [
    { name: 'lumMod', val: pct(60) },
    { name: 'alpha', val: pct(50) },
    { name: 'lumOff', val: pct(40) },
  ]);

  // 5 - the clamp boundary. `satMod 300%` overflows saturation on this base;
  //     `satMod 33.333%` brings it back. If nothing clamps in between, rows two
  //     and three land on the base colour. If the gamma round trip forces a
  //     clamp, row three does not.
  add('clamp', { kind: 'srgb', value: '4472C4' }, [{ name: 'satMod', val: pct(300) }]);
  add('clamp', { kind: 'srgb', value: '4472C4' }, [
    { name: 'satMod', val: pct(300) },
    { name: 'satMod', val: 33333 },
  ]);
  add('clamp', { kind: 'srgb', value: '4472C4' }, [
    { name: 'satMod', val: pct(300) },
    { name: 'gamma' },
    { name: 'invGamma' },
    { name: 'satMod', val: 33333 },
  ]);
  add('clamp', { kind: 'srgb', value: '4472C4' }, [
    { name: 'satMod', val: pct(300) },
    { name: 'alpha', val: pct(100) },
    { name: 'satMod', val: 33333 },
  ]);
  // The same question for lightness, which overflows the other way.
  add('clamp', { kind: 'srgb', value: '4472C4' }, [
    { name: 'lumOff', val: pct(60) },
    { name: 'lumOff', val: -pct(60) },
  ]);
  add('clamp', { kind: 'srgb', value: '4472C4' }, [
    { name: 'lumOff', val: pct(60) },
    { name: 'gray' },
  ]);

  // 6 - the three absolute channel setters. 0.7-C measured `redMod` and
  //     `redOff` and inferred these by symmetry; symmetry is not evidence.
  for (const name of ['red', 'green', 'blue'] as const) {
    add('channel-set', { kind: 'srgb', value: '4472C4' }, [{ name, val: pct(50) }]);
  }

  // 7 - edges of the percentage range.
  add('edge', { kind: 'srgb', value: '4472C4' }, [{ name: 'tint', val: 0 }]);
  add('edge', { kind: 'srgb', value: '4472C4' }, [{ name: 'tint', val: pct(100) }]);
  add('edge', { kind: 'srgb', value: '4472C4' }, [{ name: 'shade', val: 0 }]);
  add('edge', { kind: 'srgb', value: '4472C4' }, [{ name: 'shade', val: pct(100) }]);
  add('edge', { kind: 'srgb', value: '4472C4' }, [{ name: 'lumOff', val: -pct(20) }]);
  add('edge', { kind: 'srgb', value: '4472C4' }, [{ name: 'satOff', val: -pct(50) }]);
  add('edge', { kind: 'srgb', value: '4472C4' }, [{ name: 'hueOff', val: -60 * 60000 }]);
  // Grey has no hue. Every hue transform on it should be a no-op, and a model
  // that leaves saturation out of the round trip fails here rather than subtly.
  add('edge', { kind: 'srgb', value: '808080' }, [{ name: 'hueOff', val: 90 * 60000 }]);
  add('edge', { kind: 'srgb', value: '808080' }, [{ name: 'satMod', val: pct(200) }]);
  add('edge', { kind: 'srgb', value: '000000' }, [{ name: 'lumOff', val: pct(50) }]);

  // 8 - the identity control for this deck, so a systematic error in the
  //     readback shows up here rather than being attributed to a base.
  for (const value of ['000000', 'FFFFFF', '4472C4']) {
    add('identity', { kind: 'srgb', value });
  }

  /* -- `prst` --------------------------------------------------------- */

  for (const name of PRESET_COLOR_NAMES) add('prst', { kind: 'prst', value: name });

  /* -- `grey` --------------------------------------------------------- */

  for (const name of PRESET_COLOR_GREY_NAMES) add('prst-grey', { kind: 'prst', value: name });

  /* -- `sys` ---------------------------------------------------------- */

  // No `lastClr` at all, which is the case a fallback table exists for. Its own
  // deck because a `sysClr` with no `lastClr` is legal but is not something
  // PowerPoint writes, and "PowerPoint refuses it" would itself be an answer.
  for (const name of SYSTEM_COLOR_NAMES) add('sys', { kind: 'sys', value: name });

  /* -- `map` ---------------------------------------------------------- */

  // A master whose clrMap is deliberately crossed. `dk1`/`lt1`/`dk2`/`lt2` name
  // theme slots and must ignore it; `bg1`/`tx1`/`bg2`/`tx2` must follow it. On
  // the identity map 0.7-C used, both readings agree on all eight.
  for (const name of [
    'dk1',
    'lt1',
    'dk2',
    'lt2',
    'bg1',
    'tx1',
    'bg2',
    'tx2',
    'accent1',
    'accent2',
  ]) {
    add('clrmap', { kind: 'scheme', value: name });
  }

  /* -- `grammar` ------------------------------------------------------ */

  // `ST_Percentage` is a union: a Transitional document writes `40000`, and the
  // Strict form is `40%`. Whether PowerPoint accepts the Strict spelling in a
  // Transitional part decides whether our parser must handle both. A refusal is
  // as good an answer as a pixel, which is why this is alone in a deck.
  add('percent-literal', {
    kind: 'raw',
    xml: '<a:srgbClr val="4472C4"><a:lumMod val="60%"/><a:lumOff val="40%"/></a:srgbClr>',
    value: 'lumMod 60% lumOff 40%',
  });
  add('percent-literal', {
    kind: 'raw',
    xml: '<a:srgbClr val="4472C4"><a:tint val="50%"/></a:srgbClr>',
    value: 'tint 50%',
  });
  // A decimal in the Strict spelling, which the integer form cannot express at
  // this precision.
  add('percent-literal', {
    kind: 'raw',
    xml: '<a:srgbClr val="4472C4"><a:lumMod val="60.5%"/></a:srgbClr>',
    value: 'lumMod 60.5%',
  });
  // The control: the same colours in the Transitional spelling, in the same
  // deck. If the deck opens and these three are right while the Strict ones are
  // not, the failure is the grammar and not the package.
  add('percent-control', { kind: 'srgb', value: '4472C4' }, [
    { name: 'lumMod', val: pct(60) },
    { name: 'lumOff', val: pct(40) },
  ]);
  add('percent-control', { kind: 'srgb', value: '4472C4' }, [{ name: 'tint', val: pct(50) }]);

  return out;
}

/* -------------------------------------------------------------------------- */
/* rendering a swatch to markup                                               */
/* -------------------------------------------------------------------------- */

/** The crossed `clrMap` used by the `map` deck. Deliberately not the identity. */
export const CROSSED_CLR_MAP = {
  bg1: 'dk1',
  tx1: 'lt1',
  bg2: 'dk2',
  tx2: 'lt2',
  accent1: 'accent2',
  accent2: 'accent1',
  accent3: 'accent3',
  accent4: 'accent4',
  accent5: 'accent5',
  accent6: 'accent6',
  hlink: 'hlink',
  folHlink: 'folHlink',
} as const;

export function colorXml2(swatch: Swatch2): string {
  const base = swatch.base;
  if (base.kind === 'raw') return base.xml;

  const inner = swatch.transforms
    .map((t) => ('val' in t ? `<a:${t.name} val="${String(t.val)}"/>` : `<a:${t.name}/>`))
    .join('');
  let open: string;
  let tag: string;
  switch (base.kind) {
    case 'srgb':
      tag = 'srgbClr';
      open = `<a:srgbClr val="${base.value}"`;
      break;
    case 'scheme':
      tag = 'schemeClr';
      open = `<a:schemeClr val="${base.value}"`;
      break;
    case 'prst':
      tag = 'prstClr';
      open = `<a:prstClr val="${base.value}"`;
      break;
    case 'sys':
      tag = 'sysClr';
      open =
        `<a:sysClr val="${base.value}"` +
        (base.lastClr === undefined ? '' : ` lastClr="${base.lastClr}"`);
      break;
    case 'scrgb':
      tag = 'scrgbClr';
      open = `<a:scrgbClr r="${String(base.r)}" g="${String(base.g)}" b="${String(base.b)}"`;
      break;
    case 'hsl':
      tag = 'hslClr';
      open = `<a:hslClr hue="${String(base.hue)}" sat="${String(base.sat)}" lum="${String(base.lum)}"`;
      break;
  }
  return inner === '' ? `${open}/>` : `${open}>${inner}</a:${tag}>`;
}

export function describe2(swatch: Swatch2): string {
  const base = swatch.base;
  let head: string;
  switch (base.kind) {
    case 'srgb':
      head = `#${base.value}`;
      break;
    case 'scheme':
      head = `scheme:${base.value}`;
      break;
    case 'prst':
      head = `prst:${base.value}`;
      break;
    case 'sys':
      head = `sys:${base.value}${base.lastClr === undefined ? '' : `@${base.lastClr}`}`;
      break;
    case 'scrgb':
      head = `scrgb(${String(base.r)},${String(base.g)},${String(base.b)})`;
      break;
    case 'hsl':
      head = `hsl(${String(base.hue)},${String(base.sat)},${String(base.lum)})`;
      break;
    case 'raw':
      head = `raw:${base.value}`;
      break;
  }
  if (swatch.transforms.length === 0) return head;
  const parts = swatch.transforms.map((t) =>
    'val' in t ? `${t.name}=${String(t.val)}` : `${t.name}`,
  );
  return `${head} ${parts.join(' ')}`;
}

/** Re-exported so the analyser can resolve `scheme` bases without a second copy. */
export const SCHEME = CLR_SCHEME;

/** Narrowing helper: the 0.7-C swatch type this file deliberately does not reuse. */
export type LegacySwatch = Swatch;
