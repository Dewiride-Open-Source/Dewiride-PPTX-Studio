/**
 * The debug overlay, driven by a real pointer in a real browser.
 *
 * The thing worth testing here is not that a drag fires a callback. It is that
 * the pointer arrives in the right coordinate system. A handle lives in the
 * shape's own space, a pointer arrives in the slide's, and the two differ by
 * exactly the transform the shape is drawn with - so the test drags the same
 * handle on an unturned shape and on a rotated, mirrored one, and asserts the
 * two produce the same adjust value. If the inverse transform is wrong the
 * first passes and the second does not, which is precisely the failure a user
 * would report as "the handles are broken on some shapes".
 */

import { dragHandle, getPreset, type Geometry } from '@pptx-studio/geometry';
import { parseSheet, parseTheme, type Sheet } from '@pptx-studio/model';
import {
  framePoint,
  layoutSheet,
  serializeSvg,
  shapeOverlay,
  type Placed,
} from '@pptx-studio/render-svg';
import { parseXmlString } from '@pptx-studio/xml';
import { afterEach, describe, expect, it } from 'vitest';

import { RenderDomError } from './errors.js';
import { mergeAdjust, mountOverlay, type HandleEdit } from './overlay.js';

const NS =
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"' +
  ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"' +
  ' xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"';

const CLR_MAP =
  '<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2"' +
  ' accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6"' +
  ' hlink="hlink" folHlink="folHlink"/>';

const THEME_XML =
  '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="T">' +
  '<a:themeElements><a:clrScheme name="T">' +
  '<a:dk1><a:srgbClr val="000000"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>' +
  '<a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>' +
  '<a:accent1><a:srgbClr val="4472C4"/></a:accent1><a:accent2><a:srgbClr val="ED7D31"/></a:accent2>' +
  '<a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4>' +
  '<a:accent5><a:srgbClr val="5B9BD5"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6>' +
  '<a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink>' +
  '</a:clrScheme><a:fontScheme name="T"><a:majorFont><a:latin typeface="Calibri"/></a:majorFont>' +
  '<a:minorFont><a:latin typeface="Calibri"/></a:minorFont></a:fontScheme>' +
  '<a:fmtScheme name="T"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill>' +
  '</a:fillStyleLst><a:lnStyleLst><a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill>' +
  '</a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>' +
  '<a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst>' +
  '</a:fmtScheme></a:themeElements></a:theme>';

const SVG_NS = 'http://www.w3.org/2000/svg';
const SIZE = { cx: 12192000, cy: 6858000 };
const DEGREE = 60000;
const RECT = { x: 2000000, y: 1200000, cx: 4000000, cy: 2400000 };

interface Turn {
  readonly rot?: number;
  readonly flipH?: boolean;
  readonly flipV?: boolean;
}

function place(prst: string, turn: Turn = {}): Placed {
  const attrs =
    (turn.rot === undefined ? '' : ` rot="${String(Math.round(turn.rot * DEGREE))}"`) +
    (turn.flipH === true ? ' flipH="1"' : '') +
    (turn.flipV === true ? ' flipV="1"' : '');
  const xml =
    `<p:sldMaster ${NS}><p:cSld><p:spTree>` +
    '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>' +
    '<p:sp><p:nvSpPr><p:cNvPr id="9" name="probe"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr>' +
    `<a:xfrm${attrs}><a:off x="${String(RECT.x)}" y="${String(RECT.y)}"/>` +
    `<a:ext cx="${String(RECT.cx)}" cy="${String(RECT.cy)}"/></a:xfrm>` +
    `<a:prstGeom prst="${prst}"><a:avLst/></a:prstGeom>` +
    '</p:spPr></p:sp>' +
    `</p:spTree></p:cSld>${CLR_MAP}</p:sldMaster>`;

  const sheet: Sheet = {
    ...parseSheet(parseXmlString(xml).root, '/ppt/slideMaster1.xml'),
    parent: null,
    theme: parseTheme(parseXmlString(THEME_XML).root, '/ppt/theme/theme1.xml'),
  };
  const placed = layoutSheet(sheet)[0];
  if (placed === undefined) throw new Error('nothing was placed');
  return placed;
}

/** A shape whose `a:ahXY` names no guide at all, which no preset does. */
function placeCustom(): Placed {
  const xml =
    `<p:sldMaster ${NS}><p:cSld><p:spTree>` +
    '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>' +
    '<p:sp><p:nvSpPr><p:cNvPr id="11" name="custom"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr>' +
    `<a:xfrm><a:off x="${String(RECT.x)}" y="${String(RECT.y)}"/>` +
    `<a:ext cx="${String(RECT.cx)}" cy="${String(RECT.cy)}"/></a:xfrm>` +
    '<a:custGeom><a:avLst/><a:gdLst/>' +
    '<a:ahLst><a:ahXY><a:pos x="hc" y="t"/></a:ahXY></a:ahLst>' +
    '<a:cxnLst/><a:rect l="l" t="t" r="r" b="b"/>' +
    '<a:pathLst><a:path w="100" h="100"><a:moveTo><a:pt x="0" y="0"/></a:moveTo>' +
    '<a:lnTo><a:pt x="100" y="100"/></a:lnTo><a:close/></a:path></a:pathLst></a:custGeom>' +
    '</p:spPr></p:sp>' +
    `</p:spTree></p:cSld>${CLR_MAP}</p:sldMaster>`;
  const sheet: Sheet = {
    ...parseSheet(parseXmlString(xml).root, '/ppt/slideMaster1.xml'),
    parent: null,
    theme: parseTheme(parseXmlString(THEME_XML).root, '/ppt/theme/theme1.xml'),
  };
  const placed = layoutSheet(sheet)[0];
  if (placed === undefined) throw new Error('nothing was placed');
  return placed;
}

/** A slide `<svg>` on the page, sized so one EMU maps to a predictable pixel. */
function host(): SVGSVGElement {
  const svg = document.createElementNS(SVG_NS, 'svg');
  svg.setAttribute('viewBox', `0 0 ${String(SIZE.cx)} ${String(SIZE.cy)}`);
  svg.setAttribute('width', '960');
  svg.setAttribute('height', '540');
  svg.setAttribute('preserveAspectRatio', 'xMidYMid meet');
  // Pinned to the corner so the CTM is stable whatever the runner's viewport.
  svg.style.position = 'fixed';
  svg.style.left = '0';
  svg.style.top = '0';
  document.body.appendChild(svg);
  return svg;
}

const mounted: (() => void)[] = [];
afterEach(() => {
  for (const clean of mounted.splice(0)) clean();
  document.body.replaceChildren();
});

/** Client coordinates for a point in slide space, through the browser's own matrix. */
function clientOf(root: SVGSVGElement, x: number, y: number): { x: number; y: number } {
  const ctm = root.getScreenCTM();
  if (ctm === null) throw new Error('the host svg is not being rendered');
  const point = root.createSVGPoint();
  point.x = x;
  point.y = y;
  const screen = point.matrixTransform(ctm);
  return { x: screen.x, y: screen.y };
}

function pointer(type: string, at: { x: number; y: number }, target: Element): void {
  target.dispatchEvent(
    new PointerEvent(type, {
      bubbles: true,
      cancelable: true,
      pointerId: 1,
      clientX: at.x,
      clientY: at.y,
    }),
  );
}

/* -------------------------------------------------------------------------- */

describe('mergeAdjust', () => {
  it('replaces a guide in place and keeps the order', () => {
    const before = [
      { name: 'adj1', fmla: ['val', '10000'] },
      { name: 'adj2', fmla: ['val', '20000'] },
      { name: 'adj3', fmla: ['val', '30000'] },
    ];
    const after = mergeAdjust(before, { adj2: 44444 });
    expect(after.map((g) => g.name)).toEqual(['adj1', 'adj2', 'adj3']);
    expect(after[1]?.fmla).toEqual(['val', '44444']);
    expect(after[0]).toBe(before[0]);
  });

  it('appends a guide the shape had left at the preset default', () => {
    const after = mergeAdjust([], { adj: 12345.6 });
    expect(after).toEqual([{ name: 'adj', fmla: ['val', '12346'] }]);
  });

  it('rounds on both paths, because an adjust value is written as an integer', () => {
    // `ST_AdjCoordinate` is an integer. A fractional `val` is a file PowerPoint
    // has to interpret, and the number in the panel would stop being the number
    // a save writes.
    const replaced = mergeAdjust([{ name: 'adj', fmla: ['val', '10000'] }], { adj: 33333.7 });
    expect(replaced[0]?.fmla).toEqual(['val', '33334']);
    const appended = mergeAdjust([], { adj: 33333.7 });
    expect(appended[0]?.fmla).toEqual(['val', '33334']);
  });
});

describe('the mounted overlay', () => {
  it('mounts the same tree render-svg would serialise', () => {
    const root = host();
    const placed = place('roundRect');
    const overlay = mountOverlay(root, placed);
    mounted.push(() => {
      overlay.unmount();
    });
    // `XMLSerializer` must declare the namespace on the outermost element it
    // is handed. The node tree deliberately carries none: an overlay is always
    // a child of the slide `<svg>` and inherits it there.
    const mountedMarkup = new XMLSerializer()
      .serializeToString(overlay.root)
      .replace(' xmlns="http://www.w3.org/2000/svg"', '');
    expect(mountedMarkup).toBe(serializeSvg(shapeOverlay(placed).node));
  });

  it('refuses a host that is not an element', () => {
    expect(() => mountOverlay(null as unknown as SVGSVGElement, place('rect'))).toThrow(
      RenderDomError,
    );
  });

  it('redraws in place when the shape is updated', () => {
    const root = host();
    const overlay = mountOverlay(root, place('roundRect'));
    mounted.push(() => {
      overlay.unmount();
    });
    const before = overlay.root;
    overlay.update(place('chevron'));
    expect(overlay.root).not.toBe(before);
    expect(root.querySelectorAll('[data-role="overlay"]')).toHaveLength(1);
  });

  it('takes its nodes away and stops listening when unmounted', () => {
    // Two overlays on one root, and the first is unmounted. Dragging the
    // second's handle must reach only the second: the listeners live on the
    // shared root, so a `unmount` that forgot to remove them would leave the
    // first overlay editing a shape that is no longer on the page.
    const root = host();
    const gone: HandleEdit[] = [];
    const live: HandleEdit[] = [];
    const first = mountOverlay(root, place('roundRect'), {
      onChange: (edit) => gone.push(edit),
    });
    first.unmount();
    expect(root.querySelector('[data-role="overlay"]')).toBeNull();

    const second = mountOverlay(root, place('roundRect'), {
      onChange: (edit) => live.push(edit),
    });
    mounted.push(() => {
      second.unmount();
    });
    const handle = root.querySelector('[data-role="handle"]');
    expect(handle).not.toBeNull();
    const at = clientOf(root, RECT.x + 600000, RECT.y);
    pointer('pointerdown', at, handle as Element);
    pointer('pointermove', clientOf(root, RECT.x + 700000, RECT.y), root);
    expect(live.length).toBeGreaterThan(0);
    expect(gone).toHaveLength(0);
  });

  it('will not drag a handle that controls no guide', () => {
    // No preset declares one - all 120 with an `a:ahLst` give every handle at
    // least one `gdRef` - but an `a:custGeom` in a real file may, and a handle
    // that cannot write anything must not swallow the gesture.
    const root = host();
    const edits: HandleEdit[] = [];
    const overlay = mountOverlay(root, placeCustom(), {
      onChange: (edit) => edits.push(edit),
    });
    mounted.push(() => {
      overlay.unmount();
    });
    const handle = root.querySelector('[data-role="handle"]');
    expect(handle).not.toBeNull();
    expect(handle?.getAttribute('data-axes')).toBe('');
    pointer('pointerdown', clientOf(root, RECT.x, RECT.y), handle as Element);
    pointer('pointermove', clientOf(root, RECT.x + 500000, RECT.y), root);
    expect(overlay.dragging()).toBe(false);
    expect(edits).toHaveLength(0);
  });
});

describe('dragging a handle', () => {
  /** Drag `roundRect`'s corner handle to a point given in the shape's own space. */
  function drag(turn: Turn, local: { x: number; y: number }): HandleEdit[] {
    const root = host();
    const placed = place('roundRect', turn);
    const edits: HandleEdit[] = [];
    const overlay = mountOverlay(root, placed, {
      onChange: (edit) => edits.push(edit),
      onCommit: (edit) => edits.push(edit),
    });
    mounted.push(() => {
      overlay.unmount();
    });

    const handle = root.querySelector('[data-role="handle"]');
    if (handle === null) throw new Error('no handle was drawn');
    // The pointer travels in slide space, which is where a user's pointer is.
    const onSlide = framePoint(placed.frame, local);
    pointer('pointerdown', clientOf(root, placed.frame.x, placed.frame.y), handle);
    pointer('pointermove', clientOf(root, onSlide.x, onSlide.y), root);
    pointer('pointerup', clientOf(root, onSlide.x, onSlide.y), root);
    return edits;
  }

  /**
   * A destination well inside the handle's range.
   *
   * `roundRect`'s handle rides the top edge at `x1 = ss * adj / 100000`, with
   * `ss` the smaller extent - 2400000 here - and `adj` capped at 50000. So the
   * handle can reach 1200000 and no further, and a target beyond that clamps.
   * That matters: a saturating target makes every orientation return the same
   * number and the comparison below would pass without testing anything.
   */
  const IN_RANGE = { x: 600000, y: 0 };
  const EXPECTED_ADJ = 25000;

  it('writes an adjust value, not a position', () => {
    const edits = drag({}, IN_RANGE);
    expect(edits.length).toBeGreaterThan(0);
    const last = edits[edits.length - 1] as HandleEdit;
    expect(Object.keys(last.adjust)).toEqual(['adj']);
    expect(last.avLst).toEqual([{ name: 'adj', fmla: ['val', String(last.adjust['adj'])] }]);
    expect(Number.isInteger(last.adjust['adj'])).toBe(true);
  });

  it('agrees with dragging in shape space, so the inverse transform is right', () => {
    // The same handle, the same destination in the shape's own coordinates,
    // reached through the slide on four different orientations. `dragHandle`
    // called directly is the reference: it never sees the frame at all.
    const local = IN_RANGE;
    const preset = getPreset('roundRect') as Geometry;
    const expected = dragHandle(
      preset,
      { w: RECT.cx, h: RECT.cy },
      preset.ahLst[0] as Geometry['ahLst'][number],
      local,
    );
    // Strictly inside the declared bounds, or the test is vacuous: a clamped
    // target returns the same number whatever the transform did to the pointer.
    expect(expected['adj']).toBeCloseTo(EXPECTED_ADJ, 0);
    expect(expected['adj']).toBeGreaterThan(0);
    expect(expected['adj']).toBeLessThan(50000);

    for (const turn of [
      {},
      { rot: 30 },
      { rot: 30, flipH: true },
      { rot: 137.5, flipH: true, flipV: false },
      { rot: 200, flipV: true },
    ]) {
      const edits = drag(turn, local);
      const last = edits[edits.length - 1];
      expect(last, JSON.stringify(turn)).toBeDefined();
      // One adjust unit is a thousandth of a percent of the shape; the pointer
      // has been through a float32 screen matrix and back, so a couple of units
      // is the honest bound and a wrong transform misses by thousands.
      expect(
        Math.abs((last?.adjust['adj'] ?? 0) - (expected['adj'] ?? 0)),
        JSON.stringify(turn),
      ).toBeLessThan(50);
    }
  });

  it('refutes using the forward transform where the inverse belongs', () => {
    // On a mirrored, turned shape the two differ. Feeding the handle a point
    // mapped the wrong way lands somewhere else entirely, which is what makes
    // the previous test worth running on more than an unturned rectangle.
    const turn = { rot: 30, flipH: true };
    const placed = place('roundRect', turn);
    const local = IN_RANGE;
    const onSlide = framePoint(placed.frame, local);
    // The wrong reading: treat the slide point as though it were already local.
    const preset = getPreset('roundRect') as Geometry;
    const handle = preset.ahLst[0] as Geometry['ahLst'][number];
    const right = dragHandle(preset, { w: RECT.cx, h: RECT.cy }, handle, local);
    const wrong = dragHandle(preset, { w: RECT.cx, h: RECT.cy }, handle, onSlide);
    expect(Math.abs((right['adj'] ?? 0) - (wrong['adj'] ?? 0))).toBeGreaterThan(1000);

    const edits = drag(turn, local);
    const last = edits[edits.length - 1] as HandleEdit;
    expect(Math.abs((last.adjust['adj'] ?? 0) - (right['adj'] ?? 0))).toBeLessThan(50);
  });

  it('ignores a pointer that did not land on a handle', () => {
    const root = host();
    const edits: HandleEdit[] = [];
    const overlay = mountOverlay(root, place('roundRect'), {
      onChange: (edit) => edits.push(edit),
    });
    mounted.push(() => {
      overlay.unmount();
    });
    const outline = root.querySelector('[data-role="outline"]');
    expect(outline).not.toBeNull();
    pointer('pointerdown', clientOf(root, RECT.x, RECT.y), outline as Element);
    pointer('pointermove', clientOf(root, RECT.x + 500000, RECT.y), root);
    expect(overlay.dragging()).toBe(false);
    expect(edits).toHaveLength(0);
  });
});
