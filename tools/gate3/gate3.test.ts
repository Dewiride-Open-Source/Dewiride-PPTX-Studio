/**
 * What Gate 3 checks, and what it refuses to let through.
 *
 * The gate's value is that each clause fails alone, so every function here is tried at its
 * near miss: a mask that would hide a real change, a rename that would touch the wrong id, a
 * request that is same-path on another origin, a verdict with one count off zero. ADR 0054.
 */

import { describe, expect, it } from 'vitest';

import { FidelityError } from '../fidelity/errors.ts';

import {
  cellAt,
  classifyRequest,
  DISPLAY_RATIO,
  GATE_SLIDES,
  gateHolds,
  integerClip,
  invarianceOf,
  maskRootSize,
  maskZoom,
  renameIdPrefix,
  requestVerdict,
  STAGE_ZOOMS,
  STRIP_ZOOM,
  widthAt,
  zoomKey,
  type GateFacts,
} from './checks.ts';

const ORIGIN = 'http://127.0.0.1:5201';
const DECK = '/corpus/decks/a46-hundred-slides.pptx';

const SVG =
  '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 12192000 6858000" width="960" height="540">' +
  '<defs><clipPath id="slide3-1"><path d="M0 0H10V10H0Z"/></clipPath>' +
  '<marker id="slide3-2" markerUnits="userSpaceOnUse" markerWidth="127000" markerHeight="127000" ' +
  'refX="127000" refY="63500" orient="auto"><path d="M0,0 L127000,63500 L0,127000 Z" fill="#000"/></marker></defs>' +
  '<g data-shape="2"><path d="M0 0L100 0" stroke-width="12700" stroke-dasharray="50800 25400" marker-end="url(#slide3-2)"/>' +
  '<rect width="200" height="100" fill="url(#slide3-1)"/></g></svg>';

describe('the zooms and their widths', () => {
  it('draws 100 % first, so every other zoom has something to be compared with', () => {
    expect(STAGE_ZOOMS[0]).toBe(1);
    expect(STAGE_ZOOMS).toContain(0.25);
    expect(STAGE_ZOOMS).toContain(4);
    expect(STRIP_ZOOM).toBe(0.125);
  });

  it('keys a zoomed digest by the width in pixels, which is what the oracle is exported at', () => {
    expect(widthAt(1)).toBe(960);
    expect(widthAt(0.125)).toBe(120);
    expect(widthAt(4)).toBe(3840);
    expect(zoomKey('a46-hundred-slides-07', 240)).toBe('a46-hundred-slides-07@240');
  });

  it('keeps every grid 120 cells wide, and refuses a width whose cell is not whole', () => {
    expect(cellAt(120)).toBe(1);
    expect(cellAt(240)).toBe(2);
    expect(cellAt(960)).toBe(8);
    expect(cellAt(3840)).toBe(32);
    expect(() => cellAt(1000)).toThrow(FidelityError);
    expect(() => cellAt(60)).toThrow(FidelityError);
  });
});

describe('the zoom mask', () => {
  it('drops the root size and blanks the two attributes the stroke rule owns', () => {
    const masked = maskZoom(SVG);
    expect(masked).not.toContain('width="960"');
    expect(masked).not.toContain('height="540"');
    expect(masked).toContain('stroke-width="*"');
    expect(masked).toContain('stroke-dasharray="*"');
    // The viewBox, and the `<rect>`'s own width, are not the root's size and survive.
    expect(masked).toContain('viewBox="0 0 12192000 6858000"');
    expect(masked).toContain('<rect width="200" height="100"');
  });

  it('is the same for two zooms of one slide, and different for two slides', () => {
    const at4 = SVG.replace('width="960" height="540"', 'width="3840" height="2160"')
      .replace('stroke-width="12700"', 'stroke-width="6350"')
      .replace('stroke-dasharray="50800 25400"', 'stroke-dasharray="25400 12700"');
    expect(maskZoom(at4)).toBe(maskZoom(SVG));
    const moved = SVG.replace('M0 0L100 0', 'M0 0L101 0');
    expect(maskZoom(moved)).not.toBe(maskZoom(SVG));
  });

  it('lets a line end grow with the pen, and nothing else about it change', () => {
    // The same head at a pen a quarter wider: every number scales, no command moves.
    const wider = SVG.replace(/127000/g, '158750').replace(/63500/g, '79375');
    expect(maskZoom(wider)).toBe(maskZoom(SVG));
    // A path outside a marker is not the pen's to change, and nor is the head's outline.
    expect(maskZoom(SVG.replace('M0 0H10V10H0Z', 'M0 0H11V10H0Z'))).not.toBe(maskZoom(SVG));
    expect(maskZoom(SVG.replace('L0,127000 Z', 'L0,127000 L31750,63500 Z'))).not.toBe(
      maskZoom(SVG),
    );
    // Nor is the end's orientation, its paint, or which end it is on.
    expect(maskZoom(SVG.replace('orient="auto"', 'orient="auto-start-reverse"'))).not.toBe(
      maskZoom(SVG),
    );
    expect(maskZoom(SVG.replace('fill="#000"', 'fill="#111"'))).not.toBe(maskZoom(SVG));
    expect(maskZoom(SVG.replace('marker-end=', 'marker-start='))).not.toBe(maskZoom(SVG));
  });

  it('lets a crisp path move half a device pixel with its pen, and nothing else move at all', () => {
    // An odd pen at 100 % is half of 1/2 pt on; at 200 % the pen is even and the shift is gone.
    const odd = SVG.replace(
      '<path d="M0 0L100 0"',
      '<path d="M0 0L100 0" shape-rendering="crispEdges" transform="translate(6350 6350)"',
    );
    const even = SVG.replace(
      '<path d="M0 0L100 0"',
      '<path d="M0 0L100 0" shape-rendering="crispEdges"',
    );
    expect(maskZoom(odd)).toBe(maskZoom(even));
    // Turned a quarter, the shift turns with it; the mask owns that too.
    const turned = odd.replace('translate(6350 6350)', 'translate(6350 -6350)');
    expect(maskZoom(turned)).toBe(maskZoom(even));
    // Whether the path is crisp is the geometry's, not the zoom's, and a translate on an
    // antialiased path or on a group is a shape moving.
    expect(maskZoom(even)).not.toBe(maskZoom(SVG));
    expect(
      maskZoom(
        SVG.replace(
          '<path d="M0 0L100 0"',
          '<path d="M0 0L100 0" transform="translate(6350 6350)"',
        ),
      ),
    ).not.toBe(maskZoom(SVG));
    expect(
      maskZoom(
        SVG.replace('<g data-shape="2">', '<g data-shape="2" transform="translate(6350 6350)">'),
      ),
    ).not.toBe(maskZoom(SVG));
  });

  it('refuses anything that is not an <svg> root', () => {
    expect(() => maskZoom('<div/>')).toThrow(FidelityError);
    expect(() => maskRootSize('<div/>')).toThrow(FidelityError);
  });
});

describe('the display ratio', () => {
  it('is two, the ratio whose 100 % has the device pixels of 200 %', () => {
    expect(DISPLAY_RATIO).toBe(2);
    expect(widthAt(1) * DISPLAY_RATIO).toBe(widthAt(2));
  });

  it('masks the root size alone, so a stroke rounded to the wrong pixels still shows', () => {
    const bare = maskRootSize(SVG);
    expect(bare).not.toContain('width="960"');
    expect(bare).toContain('stroke-width="12700"');
    expect(bare).toContain('markerWidth="127000"');
    expect(bare).toContain('<rect width="200" height="100"');
    const atTwo = SVG.replace('width="960" height="540"', 'width="1920" height="1080"');
    expect(maskRootSize(atTwo)).toBe(bare);
    const wrongPixels = SVG.replace('stroke-width="12700"', 'stroke-width="6350"');
    expect(maskRootSize(wrongPixels)).not.toBe(bare);
  });
});

describe('the id rename', () => {
  it('renames ids, references and url(#…) on the prefix boundary only', () => {
    const svg =
      '<clipPath id="thumb3-1"/><use href="#thumb3-2"/><rect fill="url(#thumb3-1)"/>' +
      '<clipPath id="thumb30-1"/><rect fill="url(#thumb30-1)"/>';
    const renamed = renameIdPrefix(svg, 'thumb3', 'slide3');
    expect(renamed).toContain('id="slide3-1"');
    expect(renamed).toContain('href="#slide3-2"');
    expect(renamed).toContain('url(#slide3-1)');
    expect(renamed).toContain('id="thumb30-1"');
    expect(renamed).toContain('url(#thumb30-1)');
  });

  it('brings a thumbnail to the stage prefix so the mask can compare them', () => {
    const thumb = SVG.replace(/slide3-/g, 'thumb3-').replace(
      'width="960" height="540"',
      'width="120" height="68"',
    );
    expect(maskZoom(renameIdPrefix(thumb, 'thumb3', 'slide3'))).toBe(maskZoom(SVG));
  });
});

describe('invariance', () => {
  it('names the zooms whose masked SVG is not the one at 100 %', () => {
    expect(
      invarianceOf([
        { zoom: 1, masked: 'a' },
        { zoom: 0.25, masked: 'a' },
        { zoom: 2, masked: 'b' },
        { zoom: 0.125, masked: 'a' },
      ]),
    ).toEqual([2]);
  });

  it('refuses a set with no 100 % entry rather than declaring it invariant', () => {
    expect(() => invarianceOf([{ zoom: 2, masked: 'a' }])).toThrow(FidelityError);
  });
});

describe('the requests', () => {
  it('knows the page, its bundle and the deck, and nothing else', () => {
    expect(classifyRequest(`${ORIGIN}/`, ORIGIN, DECK)).toBe('static');
    expect(classifyRequest(`${ORIGIN}/dist/main.js`, ORIGIN, DECK)).toBe('static');
    expect(classifyRequest(`${ORIGIN}/dist/worker.js`, ORIGIN, DECK)).toBe('static');
    expect(classifyRequest(`${ORIGIN}${DECK}`, ORIGIN, DECK)).toBe('deck');
    expect(classifyRequest(`${ORIGIN}/corpus/decks/a01-minimal.pptx`, ORIGIN, DECK)).toBe('other');
    expect(classifyRequest(`${ORIGIN}/packages/text/dist/index.js`, ORIGIN, DECK)).toBe('other');
  });

  it('calls the same path on another origin an offender', () => {
    expect(classifyRequest(`https://example.invalid${DECK}`, ORIGIN, DECK)).toBe('other');
    expect(classifyRequest(`https://example.invalid/dist/main.js`, ORIGIN, DECK)).toBe('other');
    expect(classifyRequest('not a url', ORIGIN, DECK)).toBe('other');
  });

  it('counts, names offenders, and names anything after going offline whatever it was for', () => {
    const verdict = requestVerdict(
      [
        { url: `${ORIGIN}/`, afterOffline: false },
        { url: `${ORIGIN}${DECK}`, afterOffline: false },
        { url: 'https://fonts.example.invalid/x.woff2', afterOffline: false },
        { url: `${ORIGIN}${DECK}`, afterOffline: true },
      ],
      ORIGIN,
      DECK,
    );
    expect(verdict.total).toBe(4);
    expect(verdict.static).toBe(1);
    expect(verdict.deck).toBe(2);
    expect(verdict.other).toEqual(['https://fonts.example.invalid/x.woff2']);
    expect(verdict.afterOffline).toEqual([`${ORIGIN}${DECK}`]);
  });
});

describe('the verdict', () => {
  const holds: GateFacts = {
    slides: GATE_SLIDES,
    notDrawn: 0,
    pageErrors: 0,
    breaks: 0,
    changed: 0,
    vanished: 0,
    offenders: 0,
    afterOffline: 0,
    ratio: 0,
  };

  it('holds with a hundred slides and every count at zero', () => {
    expect(gateHolds(holds)).toBe(true);
    expect(gateHolds({ ...holds, slides: 101 })).toBe(true);
  });

  it('fails on each clause alone', () => {
    expect(gateHolds({ ...holds, slides: 99 })).toBe(false);
    for (const clause of Object.keys(holds) as (keyof GateFacts)[]) {
      if (clause === 'slides') continue;
      expect(gateHolds({ ...holds, [clause]: 1 }), clause).toBe(false);
    }
  });
});

describe('the clip', () => {
  it('takes a box on whole pixels and refuses one that is not', () => {
    const box = { x: 1, y: 1, width: 3840, height: 2160 };
    expect(integerClip(box, 'stage')).toEqual(box);
    expect(() => integerClip({ ...box, y: 817.1875 }, 'stage')).toThrow(FidelityError);
    expect(() => integerClip({ ...box, width: 120.5 }, 'thumb')).toThrow(FidelityError);
  });
});
