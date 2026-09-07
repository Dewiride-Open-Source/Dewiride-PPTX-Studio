/**
 * Experiment C6 - what `a:blipFill` means, as questions with candidate answers.
 *
 * The authoring step settled the units: `@tx` is EMU, `@sx` is a percentage of a
 * natural size, `a:alphaModFix/@amt` is opacity rather than transparency. What
 * it cannot settle is where the tile grid is anchored, whether `@flip` mirrors
 * every tile or alternate ones, and what curve each `a:blip` effect applies.
 * Those need pixels, and these are the probes that produce them.
 */

import { EMU_PER_INCH, SLIDE_HEIGHT, SLIDE_WIDTH } from '../../lib/pptx.ts';

/** One point, in EMU. Every rectangle here is a whole number of points. */
export const PT = 12700;

/** The probe image, in pixels. Small enough that a tile repeats several times. */
export const QUAD_PX = 32;

/** The ramp image: 256 columns of input, four bands of channel. */
export const RAMP_W = 256;
export const RAMP_H = 64;

/** The four quadrant colours, and the mark that makes orientation readable. */
export const QUADRANTS = [0xc0392b, 0x27ae60, 0x2980b9, 0xf1c40f] as const;
export const MARK = 0xffffff;

/**
 * 32 x 32, four coloured quadrants with a white L in the top-left one.
 *
 * The L is the whole point: four quadrants cannot tell a horizontal flip from a
 * vertical one plus a rotation, and `@flip`, `a:srcRect` and `@algn` are each
 * invisible on a symmetric image.
 */
export function quadPixel(x: number, y: number): number {
  if (x >= 4 && x < 7 && y >= 4 && y < 12) return MARK;
  if (x >= 4 && x < 12 && y >= 9 && y < 12) return MARK;
  return QUADRANTS[(y < 16 ? 0 : 2) + (x < 16 ? 0 : 1)] ?? 0;
}

/**
 * 256 x 64: a grey ramp over a red, a green and a blue ramp.
 *
 * Every effect gets 256 inputs per band rather than the four a quadrant image
 * would give, which is what makes a fitted curve refutable instead of merely
 * consistent.
 */
export function rampPixel(x: number, y: number): number {
  const level = x & 0xff;
  const band = Math.floor(y / 16);
  if (band === 0) return (level << 16) | (level << 8) | level;
  if (band === 1) return level << 16;
  if (band === 2) return level << 8;
  return level;
}

export interface Rect {
  readonly x: number;
  readonly y: number;
  readonly cx: number;
  readonly cy: number;
}

export type DeckId = 'geom' | 'dpi' | 'effects' | 'lum';

export interface BlipProbe {
  readonly id: string;
  readonly deck: DeckId;
  readonly question: string;
  /** The whole `a:blipFill` element, with `r:embed` already filled in. */
  readonly fill: string;
  readonly rect: Rect;
  /** Which image the fill draws, so the analysis knows what it is looking at. */
  readonly image: 'quad' | 'quad150' | 'ramp';
}

/** `rId2` is the first media relationship `buildPptx` writes, in media order. */
const QUAD = 'rId2';
const QUAD150 = 'rId3';
const RAMP = 'rId4';

function blip(embed: string, effects = ''): string {
  return effects === ''
    ? `<a:blip r:embed="${embed}"/>`
    : `<a:blip r:embed="${embed}">${effects}</a:blip>`;
}

/** A grid of cells that are a whole number of points on every edge. */
function grid(cols: number, cellW: number, cellH: number, gap: number, margin: number) {
  return (index: number): Rect => {
    const col = index % cols;
    const row = Math.floor(index / cols);
    return {
      x: (margin + col * (cellW + gap)) * PT,
      y: (margin + row * (cellH + gap)) * PT,
      cx: cellW * PT,
      cy: cellH * PT,
    };
  };
}

/* -------------------------------------------------------------------------- */
/* deck 1 - tile and crop geometry                                            */
/* -------------------------------------------------------------------------- */

const ALIGNMENTS = ['tl', 't', 'tr', 'l', 'ctr', 'r', 'bl', 'b', 'br'] as const;

function geomProbes(): BlipProbe[] {
  const at = grid(8, 100, 80, 10, 20);
  const out: BlipProbe[] = [];
  const add = (id: string, question: string, fill: string): void => {
    out.push({ id, deck: 'geom', question, fill, rect: at(out.length), image: 'quad' });
  };

  add(
    'stretch',
    'Does a bare a:stretch map the whole image onto the whole shape, aspect ignored?',
    `<a:blipFill>${blip(QUAD)}<a:stretch><a:fillRect/></a:stretch></a:blipFill>`,
  );
  add(
    'stretch-fillrect-in',
    'Is a:fillRect an inset of the destination, leaving the shape bare outside it?',
    `<a:blipFill>${blip(QUAD)}<a:stretch>` +
      '<a:fillRect l="10000" t="10000" r="10000" b="10000"/></a:stretch></a:blipFill>',
  );
  add(
    'stretch-fillrect-out',
    'Does a negative a:fillRect outset the image and clip it to the shape?',
    `<a:blipFill>${blip(QUAD)}<a:stretch>` +
      '<a:fillRect l="-20000" t="-20000" r="0" b="0"/></a:stretch></a:blipFill>',
  );
  add(
    'srcrect-crop',
    'Is a:srcRect an inset of the SOURCE, measured as a fraction of the image?',
    `<a:blipFill>${blip(QUAD)}<a:srcRect l="25000" t="25000" r="12500" b="12500"/>` +
      '<a:stretch><a:fillRect/></a:stretch></a:blipFill>',
  );
  add(
    'srcrect-and-fillrect',
    'When both are present, is the cropped source mapped onto the inset destination?',
    `<a:blipFill>${blip(QUAD)}<a:srcRect l="25000" t="25000"/>` +
      '<a:stretch><a:fillRect l="20000" t="20000"/></a:stretch></a:blipFill>',
  );

  for (const algn of ALIGNMENTS) {
    add(
      `tile-algn-${algn}`,
      `Where does @algn="${algn}" anchor the tile grid inside the shape?`,
      `<a:blipFill>${blip(QUAD)}<a:tile tx="0" ty="0" sx="100000" sy="100000" ` +
        `flip="none" algn="${algn}"/></a:blipFill>`,
    );
  }

  for (const flip of ['none', 'x', 'y', 'xy'] as const) {
    add(
      `tile-flip-${flip}`,
      `Does @flip="${flip}" mirror every tile, or alternate ones?`,
      `<a:blipFill>${blip(QUAD)}<a:tile tx="0" ty="0" sx="100000" sy="100000" ` +
        `flip="${flip}" algn="tl"/></a:blipFill>`,
    );
  }

  add(
    'tile-scale-half',
    'Is @sx a percentage of the natural tile size rather than of the shape?',
    `<a:blipFill>${blip(QUAD)}<a:tile tx="0" ty="0" sx="50000" sy="50000" ` +
      'flip="none" algn="tl"/></a:blipFill>',
  );
  add(
    'tile-scale-uneven',
    'Do @sx and @sy scale independently?',
    `<a:blipFill>${blip(QUAD)}<a:tile tx="0" ty="0" sx="50000" sy="150000" ` +
      'flip="none" algn="tl"/></a:blipFill>',
  );
  add(
    'tile-offset-pos',
    'Does a positive @tx move the grid right and down from the anchor?',
    `<a:blipFill>${blip(QUAD)}<a:tile tx="${9 * PT}" ty="${6 * PT}" sx="100000" sy="100000" ` +
      'flip="none" algn="tl"/></a:blipFill>',
  );
  add(
    'tile-offset-neg',
    'And a negative one left and up?',
    `<a:blipFill>${blip(QUAD)}<a:tile tx="${-9 * PT}" ty="${-6 * PT}" sx="100000" sy="100000" ` +
      'flip="none" algn="tl"/></a:blipFill>',
  );
  add(
    'tile-offset-centre',
    'Is @tx measured from the anchor even when the anchor is not a corner?',
    `<a:blipFill>${blip(QUAD)}<a:tile tx="${6 * PT}" ty="${6 * PT}" sx="100000" sy="100000" ` +
      'flip="none" algn="ctr"/></a:blipFill>',
  );
  // An asymmetric crop, because cropping to half leaves one solid quadrant and
  // a flat fill cannot say whether the crop happened before the tiling or after.
  add(
    'tile-srcrect',
    'Does a:srcRect crop the source BEFORE it is tiled?',
    `<a:blipFill>${blip(QUAD)}<a:srcRect l="25000" t="12500" r="12500"/>` +
      '<a:tile tx="0" ty="0" sx="100000" sy="100000" flip="none" algn="tl"/></a:blipFill>',
  );

  // `a:clrChange` on the quadrant image, where a replaced colour covers a whole
  // quadrant rather than the single ramp column that antialiasing contaminates.
  add(
    'clrchange-exact',
    'Does a:clrChange replace one exact colour and leave its neighbours alone?',
    `<a:blipFill>${blip(
      QUAD,
      '<a:clrChange><a:clrFrom><a:srgbClr val="F1C40F"/></a:clrFrom>' +
        '<a:clrTo><a:srgbClr val="FFFFFF"/></a:clrTo></a:clrChange>',
    )}<a:stretch><a:fillRect/></a:stretch></a:blipFill>`,
  );
  add(
    'clrchange-near',
    'A colour one level away: is the match exact, or is there a tolerance?',
    `<a:blipFill>${blip(
      QUAD,
      '<a:clrChange><a:clrFrom><a:srgbClr val="F1C40E"/></a:clrFrom>' +
        '<a:clrTo><a:srgbClr val="FFFFFF"/></a:clrTo></a:clrChange>',
    )}<a:stretch><a:fillRect/></a:stretch></a:blipFill>`,
  );
  add(
    'clrchange-alpha',
    'Does a clrTo carrying alpha=0 erase the colour rather than repaint it?',
    `<a:blipFill>${blip(
      QUAD,
      '<a:clrChange><a:clrFrom><a:srgbClr val="C0392B"/></a:clrFrom>' +
        '<a:clrTo><a:srgbClr val="C0392B"><a:alpha val="0"/></a:srgbClr></a:clrTo></a:clrChange>',
    )}<a:stretch><a:fillRect/></a:stretch></a:blipFill>`,
  );
  return out;
}

/* -------------------------------------------------------------------------- */
/* deck 2 - what sets a tile's natural size                                   */
/* -------------------------------------------------------------------------- */

/**
 * The one question the geometry deck cannot answer on its own.
 *
 * A tile at `sx="100000"` occupies some size in EMU, and the three candidates -
 * a fixed 96 dpi, the image's own declared resolution, and `a:blipFill/@dpi` -
 * agree on an image that declares 96 and a fill that declares nothing.
 */
function dpiProbes(): BlipProbe[] {
  const at = grid(8, 100, 80, 10, 20);
  const out: BlipProbe[] = [];
  const add = (id: string, question: string, fill: string, image: BlipProbe['image']): void => {
    out.push({ id, deck: 'dpi', question, fill, rect: at(out.length), image });
  };

  const tile = '<a:tile tx="0" ty="0" sx="100000" sy="100000" flip="none" algn="tl"/>';
  add(
    'dpi-image96',
    'A 96 dpi image with no @dpi: what is the tile period?',
    `<a:blipFill>${blip(QUAD)}${tile}</a:blipFill>`,
    'quad',
  );
  add(
    'dpi-image150',
    'Same pixels declaring 150 dpi: does the period shrink?',
    `<a:blipFill>${blip(QUAD150)}${tile}</a:blipFill>`,
    'quad150',
  );
  add(
    'dpi-attr-0',
    'dpi="0" is what PowerPoint writes: does it mean "use the image"?',
    `<a:blipFill dpi="0">${blip(QUAD150)}${tile}</a:blipFill>`,
    'quad150',
  );
  add(
    'dpi-attr-96',
    'Does an explicit dpi="96" override a 150 dpi image?',
    `<a:blipFill dpi="96">${blip(QUAD150)}${tile}</a:blipFill>`,
    'quad150',
  );
  add(
    'dpi-attr-300',
    'And does dpi="300" shrink the tile further still?',
    `<a:blipFill dpi="300">${blip(QUAD)}${tile}</a:blipFill>`,
    'quad',
  );

  // A resample to @dpi lands on a whole pixel count, and these three values put
  // that count at .827, .533 and .947 of a pixel - which is where rounding up,
  // rounding down and rounding to nearest stop agreeing.
  for (const dpi of [32, 40, 56]) {
    add(
      `dpi-attr-${String(dpi)}`,
      `At dpi="${String(dpi)}" the resampled pixel count is fractional: which way does it go?`,
      `<a:blipFill dpi="${String(dpi)}">${blip(QUAD150)}${tile}</a:blipFill>`,
      'quad150',
    );
  }
  return out;
}

/* -------------------------------------------------------------------------- */
/* deck 4 - brightness and contrast together                                  */
/* -------------------------------------------------------------------------- */

/**
 * `a:lum` with both attributes, which is what every Watermark and every
 * brightness slider writes.
 *
 * Each attribute alone is a straight line, and the two lines do not compose into
 * what the pair produces, so the composition is its own measurement.
 */
function lumProbes(): BlipProbe[] {
  const at = grid(3, RAMP_W, RAMP_H, 16, 24);
  const out: BlipProbe[] = [];
  for (const bright of [-40000, 40000, 70000]) {
    for (const contrast of [-70000, -30000, 30000, 60000]) {
      const effects = `<a:lum bright="${String(bright)}" contrast="${String(contrast)}"/>`;
      out.push({
        id: `lum-b${String(bright / 1000)}-c${String(contrast / 1000)}`,
        deck: 'lum',
        question: 'In which order, and about which pivot, do bright and contrast compose?',
        fill: `<a:blipFill>${blip(RAMP, effects)}<a:stretch><a:fillRect/></a:stretch></a:blipFill>`,
        rect: at(out.length),
        image: 'ramp',
      });
    }
  }
  return out;
}

/* -------------------------------------------------------------------------- */
/* deck 3 - the a:blip colour effects                                         */
/* -------------------------------------------------------------------------- */

/**
 * One shape per effect, each stretching the ramp so that a source column lands
 * on a whole number of device pixels.
 */
function effectProbes(): BlipProbe[] {
  const at = grid(3, RAMP_W, RAMP_H, 16, 24);
  const out: BlipProbe[] = [];
  const add = (id: string, question: string, effects: string): void => {
    out.push({
      id,
      deck: 'effects',
      question,
      fill: `<a:blipFill>${blip(RAMP, effects)}<a:stretch><a:fillRect/></a:stretch></a:blipFill>`,
      rect: at(out.length),
      image: 'ramp',
    });
  };

  add('plain', 'The control: does an unfiltered stretch reproduce the source exactly?', '');
  add('grayscl', 'Which luminance weights does a:grayscl use?', '<a:grayscl/>');
  add(
    'bilevel-25',
    'Is a:biLevel/@thresh compared against luminance?',
    '<a:biLevel thresh="25000"/>',
  );
  add(
    'bilevel-50',
    'And is the comparison inclusive at the threshold?',
    '<a:biLevel thresh="50000"/>',
  );
  add(
    'bilevel-75',
    'A third threshold, because two points fit any straight line.',
    '<a:biLevel thresh="75000"/>',
  );
  add(
    'lum-bright-40',
    'Is @bright an additive offset in sRGB or in linear light?',
    '<a:lum bright="40000"/>',
  );
  add(
    'lum-bright-neg40',
    'Does a negative @bright subtract by the same rule?',
    '<a:lum bright="-40000"/>',
  );
  add('lum-contrast-60', 'What pivot does @contrast scale about?', '<a:lum contrast="60000"/>');
  add(
    'lum-contrast-neg60',
    'And does a negative @contrast compress toward the same pivot?',
    '<a:lum contrast="-60000"/>',
  );
  add(
    'lum-watermark',
    'The pair PowerPoint writes for Watermark, as one operation.',
    '<a:lum bright="70000" contrast="-70000"/>',
  );
  add(
    'duotone-black-amber',
    'Does a:duotone map luminance onto a two-colour ramp?',
    '<a:duotone><a:prstClr val="black"/><a:srgbClr val="FFC000"/></a:duotone>',
  );
  add(
    'duotone-blue-red',
    'Two chromatic ends, so the interpolation space is visible.',
    '<a:duotone><a:srgbClr val="1F4E79"/><a:srgbClr val="E74C3C"/></a:duotone>',
  );
  add(
    'alphamodfix-60',
    'Is @amt the resulting opacity rather than the transparency?',
    '<a:alphaModFix amt="60000"/>',
  );
  add(
    'alphamodfix-30',
    'A second amount, to separate a scale from a constant.',
    '<a:alphaModFix amt="30000"/>',
  );
  add(
    'grayscl-then-alpha',
    'Do two effects apply in document order?',
    '<a:grayscl/><a:alphaModFix amt="60000"/>',
  );
  add(
    'clrchange',
    'Does a:clrChange match a colour exactly, or within a tolerance?',
    '<a:clrChange><a:clrFrom><a:srgbClr val="808080"/></a:clrFrom>' +
      '<a:clrTo><a:srgbClr val="FFFFFF"/></a:clrTo></a:clrChange>',
  );
  return out;
}

export function blipProbes(): readonly BlipProbe[] {
  return [...geomProbes(), ...dpiProbes(), ...effectProbes(), ...lumProbes()];
}

/** The natural size of one unscaled tile under a given resolution, in EMU. */
export function naturalTileEmu(pixels: number, dpi: number): number {
  return (pixels * EMU_PER_INCH) / dpi;
}

/** Guard: nothing may be authored off the slide, where it cannot be exported. */
export function assertOnSlide(probes: readonly BlipProbe[]): void {
  for (const probe of probes) {
    const { x, y, cx, cy } = probe.rect;
    if (x < 0 || y < 0 || x + cx > SLIDE_WIDTH || y + cy > SLIDE_HEIGHT) {
      throw new Error(`probe ${probe.id} is not entirely on the slide`);
    }
  }
}
