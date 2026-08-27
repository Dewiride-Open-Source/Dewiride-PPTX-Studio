/**
 * Experiment C - the colour-transform swatch set.
 *
 * ECMA-376 defines `a:tint` as "10% of the input colour combined with 90%
 * white" and stops there. It does not say which colour space the combining
 * happens in, and the credible implementations disagree:
 *
 *   - LibreOffice (`oox/source/drawingml/color.cxx`) uses a power law with
 *     gamma 2.3.
 *   - Apache POI (`org.apache.poi.sl.draw.DrawPaint`) uses the sRGB piecewise
 *     transfer function.
 *   - Doing the arithmetic straight on the 0-255 values is a third answer, and
 *     it is what a first implementation naturally writes.
 *
 * They differ by a few units out of 255 in the midtones - on *every* themed
 * fill in *every* deck. That is not a rounding argument: it is the difference
 * between a renderer that matches PowerPoint and one that is visibly off on the
 * accent colour of every corporate template. So we ask PowerPoint.
 *
 * Design notes:
 *
 *   - Most swatches use `a:srgbClr`, not `a:schemeClr`, so a disagreement is
 *     attributable to the transform and not to theme resolution. A separate
 *     block does exercise scheme resolution and the `clrMap`.
 *   - `808080` appears at fine percentage steps because mid-grey is where the
 *     candidate models are furthest apart.
 *   - `000000` with a tint and `FFFFFF` with a shade are included because they
 *     are the cases where a linearising model and a naive one *agree*, which
 *     makes them a control rather than a probe.
 *   - Transform order is exercised explicitly: `lumMod` then `lumOff` is not
 *     the same colour as `lumOff` then `lumMod`, and a renderer that applies
 *     transforms in a fixed order rather than document order gets one of them
 *     wrong.
 */

export type Transform =
  | { readonly name: 'tint' | 'shade' | 'alpha' | 'alphaOff' | 'alphaMod'; readonly val: number }
  | { readonly name: 'hue' | 'hueOff' | 'hueMod'; readonly val: number }
  | { readonly name: 'sat' | 'satOff' | 'satMod'; readonly val: number }
  | { readonly name: 'lum' | 'lumOff' | 'lumMod'; readonly val: number }
  | {
      readonly name:
        | 'red'
        | 'redOff'
        | 'redMod'
        | 'green'
        | 'greenOff'
        | 'greenMod'
        | 'blue'
        | 'blueOff'
        | 'blueMod';
      readonly val: number;
    }
  | { readonly name: 'comp' | 'inv' | 'gray' | 'gamma' | 'invGamma' };

export interface Swatch {
  /** Stable identifier; also the shape name, which is how readback joins. */
  readonly id: string;
  /** What group of questions this swatch belongs to, for the report. */
  readonly group: string;
  readonly base:
    | { readonly kind: 'srgb'; readonly value: string }
    | { readonly kind: 'scheme'; readonly value: string };
  readonly transforms: readonly Transform[];
}

const pct = (n: number): number => Math.round(n * 1000);

function srgb(value: string): Swatch['base'] {
  return { kind: 'srgb', value };
}
function scheme(value: string): Swatch['base'] {
  return { kind: 'scheme', value };
}

/** The bases the tint/shade sweep runs on. */
const SWEEP_BASES = ['FF0000', '4472C4', '808080', '70AD47'] as const;
const SWEEP_PERCENTS = [10, 20, 25, 30, 40, 50, 60, 75, 80, 90] as const;
/** Mid-grey at fine steps: the maximum-divergence probe. */
const FINE_PERCENTS = [5, 15, 35, 45, 55, 65, 85, 95] as const;

export function swatches(): Swatch[] {
  const out: Swatch[] = [];
  const add = (group: string, base: Swatch['base'], transforms: Transform[]): void => {
    out.push({ id: `s${String(out.length).padStart(3, '0')}`, group, base, transforms });
  };

  // 1 - the main event: tint and shade across four bases.
  for (const name of ['tint', 'shade'] as const) {
    for (const base of SWEEP_BASES) {
      for (const p of SWEEP_PERCENTS) {
        add(`${name}-sweep`, srgb(base), [{ name, val: pct(p) }]);
      }
    }
  }

  // 2 - mid-grey at fine steps, where the models diverge most.
  for (const name of ['tint', 'shade'] as const) {
    for (const p of FINE_PERCENTS) {
      add(`${name}-fine`, srgb('808080'), [{ name, val: pct(p) }]);
    }
  }

  // 3 - controls. Every candidate model agrees on these, so a disagreement here
  //     means the readback itself is wrong, not the model.
  for (const p of [20, 50, 80]) {
    add('tint-of-black', srgb('000000'), [{ name: 'tint', val: pct(p) }]);
    add('shade-of-white', srgb('FFFFFF'), [{ name: 'shade', val: pct(p) }]);
  }

  // 4 - lumMod and lumOff alone.
  for (const base of ['FF0000', '4472C4', '808080'] as const) {
    for (const v of [20, 50, 60, 75, 120, 200]) {
      add('lumMod', srgb(base), [{ name: 'lumMod', val: pct(v) }]);
    }
  }
  for (const base of ['FF0000', '4472C4', '808080'] as const) {
    for (const v of [10, 25, 50, 75]) {
      add('lumOff', srgb(base), [{ name: 'lumOff', val: pct(v) }]);
    }
  }

  // 5 - the pairs PowerPoint's own "Lighter / Darker" gallery emits. If a
  //     renderer gets only one row of this product right, make it this one.
  const gallery: [string, Transform[]][] = [
    [
      'lighter-80',
      [
        { name: 'lumMod', val: pct(20) },
        { name: 'lumOff', val: pct(80) },
      ],
    ],
    [
      'lighter-60',
      [
        { name: 'lumMod', val: pct(40) },
        { name: 'lumOff', val: pct(60) },
      ],
    ],
    [
      'lighter-40',
      [
        { name: 'lumMod', val: pct(60) },
        { name: 'lumOff', val: pct(40) },
      ],
    ],
    ['darker-25', [{ name: 'lumMod', val: pct(75) }]],
    ['darker-50', [{ name: 'lumMod', val: pct(50) }]],
  ];
  for (const [, transforms] of gallery) {
    add('ui-gallery', srgb('4472C4'), transforms);
    add('ui-gallery-scheme', scheme('accent1'), transforms);
  }

  // 6 - order dependence. Same two transforms, both orders.
  for (const base of ['4472C4', '70AD47'] as const) {
    add('order', srgb(base), [
      { name: 'lumMod', val: pct(60) },
      { name: 'lumOff', val: pct(40) },
    ]);
    add('order', srgb(base), [
      { name: 'lumOff', val: pct(40) },
      { name: 'lumMod', val: pct(60) },
    ]);
    add('order', srgb(base), [
      { name: 'tint', val: pct(50) },
      { name: 'shade', val: pct(50) },
    ]);
    add('order', srgb(base), [
      { name: 'shade', val: pct(50) },
      { name: 'tint', val: pct(50) },
    ]);
  }

  // 7 - saturation.
  for (const base of ['4472C4', '808080'] as const) {
    for (const v of [20, 50, 150, 200, 300]) {
      add('satMod', srgb(base), [{ name: 'satMod', val: pct(v) }]);
    }
  }
  for (const v of [20, 60, 100]) add('sat', srgb('4472C4'), [{ name: 'sat', val: pct(v) }]);
  for (const v of [10, 30]) add('satOff', srgb('4472C4'), [{ name: 'satOff', val: pct(v) }]);

  // 8 - hue, in sixtieths of a degree.
  for (const deg of [60, 120, 180, 240]) {
    add('hueOff', srgb('4472C4'), [{ name: 'hueOff', val: deg * 60000 }]);
  }
  add('hue', srgb('4472C4'), [{ name: 'hue', val: 120 * 60000 }]);
  for (const v of [50, 150]) add('hueMod', srgb('4472C4'), [{ name: 'hueMod', val: pct(v) }]);

  // 9 - absolute luminance, and the flag-shaped transforms.
  for (const v of [25, 50, 75]) add('lum', srgb('4472C4'), [{ name: 'lum', val: pct(v) }]);
  for (const name of ['comp', 'inv', 'gray', 'gamma', 'invGamma'] as const) {
    add('flag', srgb('4472C4'), [{ name }]);
    add('flag', srgb('FF0000'), [{ name }]);
  }

  // 10 - per-channel.
  for (const name of ['redMod', 'greenMod', 'blueMod'] as const) {
    add('channelMod', srgb('4472C4'), [{ name, val: pct(50) }]);
  }
  for (const name of ['redOff', 'greenOff', 'blueOff'] as const) {
    add('channelOff', srgb('4472C4'), [{ name, val: pct(20) }]);
  }

  // 11 - theme resolution and the clrMap. `dk1`/`lt1`/`dk2`/`lt2` bypass the
  //      map; `bg1`/`tx1`/`bg2`/`tx2` go through it. Both spellings are here so
  //      the fixture can tell a resolver that confuses them.
  for (const name of [
    'dk1',
    'lt1',
    'dk2',
    'lt2',
    'tx1',
    'bg1',
    'tx2',
    'bg2',
    'accent1',
    'accent2',
    'accent3',
    'accent4',
    'accent5',
    'accent6',
    'hlink',
    'folHlink',
  ]) {
    add('scheme', scheme(name), []);
  }

  // 12 - plain bases, as the identity control on the whole pipeline.
  for (const base of ['000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', '4472C4', '808080']) {
    add('identity', srgb(base), []);
  }

  return out;
}

/** The `a:srgbClr` / `a:schemeClr` element for a swatch, transforms in order. */
export function colorXml(swatch: Swatch): string {
  const inner = swatch.transforms
    .map((t) => ('val' in t ? `<a:${t.name} val="${String(t.val)}"/>` : `<a:${t.name}/>`))
    .join('');
  const tag = swatch.base.kind === 'srgb' ? 'srgbClr' : 'schemeClr';
  return inner === ''
    ? `<a:${tag} val="${swatch.base.value}"/>`
    : `<a:${tag} val="${swatch.base.value}">${inner}</a:${tag}>`;
}

/** A one-line human-readable description, for the fixture file and the report. */
export function describe(swatch: Swatch): string {
  const base =
    swatch.base.kind === 'srgb' ? `#${swatch.base.value}` : `scheme:${swatch.base.value}`;
  if (swatch.transforms.length === 0) return base;
  const parts = swatch.transforms.map((t) =>
    'val' in t ? `${t.name}=${String(t.val)}` : `${t.name}`,
  );
  return `${base} ${parts.join(' ')}`;
}
