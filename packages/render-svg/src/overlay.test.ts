/**
 * The debug overlay, and the transform arithmetic it stands on.
 *
 * Two things here are load-bearing and neither is checked by comparing this
 * package against itself.
 *
 * **`framePoint` is checked against the browser.** The renderer emits a
 * `transform` attribute and the overlay does the same transform as arithmetic,
 * and a handle sits where the second one says while the shape is drawn where
 * the first one does. So the test builds the real attribute, asks Chromium for
 * the matrix it produced, and compares. Two implementations that must agree,
 * one of them somebody else's - which is the only kind of agreement worth
 * asserting.
 *
 * **The overlay is checked against `@pptx-studio/geometry`.** Path count, text
 * rectangle, connection sites and handle positions are all re-derived from
 * `resolveGeometry`/`resolveHandles` rather than from a recorded expectation,
 * so a change in the evaluator moves the test and the code together and a
 * change in only one of them fails.
 */

import {
  getPreset,
  presetNames,
  resolveGeometry,
  resolveHandles,
  type Geometry,
  type Point,
} from '@pptx-studio/geometry';
import { parseSheet, parseTheme, type Sheet } from '@pptx-studio/model';
import { parseXmlString } from '@pptx-studio/xml';
import { describe, expect, it } from 'vitest';

import { layoutSheet, type Placed } from './layout.js';
import { serializeSvg, type SvgElement, type SvgNode } from './node.js';
import { DEFAULT_UNIT, LOCUS_SAMPLES, PATH_COLORS, shapeOverlay } from './overlay.js';
import {
  framePoint,
  frameTransform,
  inverseFramePoint,
  type Frame,
  type Vec,
} from './transform.js';

/* -------------------------------------------------------------------------- */
/* a slide with one shape on it                                               */
/* -------------------------------------------------------------------------- */

const NS =
  ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"' +
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
  '</a:clrScheme><a:fontScheme name="T"><a:majorFont><a:latin typeface="Calibri Light"/></a:majorFont>' +
  '<a:minorFont><a:latin typeface="Calibri"/></a:minorFont></a:fontScheme>' +
  '<a:fmtScheme name="T"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill>' +
  '</a:fillStyleLst><a:lnStyleLst><a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill>' +
  '</a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>' +
  '<a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst>' +
  '</a:fmtScheme></a:themeElements></a:theme>';

const DEGREE = 60000;
const SIZE = { cx: 3000000, cy: 2000000 };

interface ShapeSpec {
  readonly prst?: string;
  readonly custGeom?: string;
  /** `a:avLst` entries, as `name` to value. */
  readonly adj?: Readonly<Record<string, number>>;
  readonly cx?: number;
  readonly cy?: number;
  readonly rot?: number;
  readonly flipH?: boolean;
  readonly flipV?: boolean;
}

/**
 * One shape on one master, laid out.
 *
 * A master rather than a slide so nothing inherits and the test is looking at
 * exactly the geometry it wrote.
 */
function place(spec: ShapeSpec = {}): Placed {
  const avLst =
    spec.adj === undefined
      ? '<a:avLst/>'
      : `<a:avLst>${Object.entries(spec.adj)
          .map(([name, value]) => `<a:gd name="${name}" fmla="val ${String(value)}"/>`)
          .join('')}</a:avLst>`;
  const geom =
    spec.custGeom ?? `<a:prstGeom prst="${spec.prst ?? 'roundRect'}">${avLst}</a:prstGeom>`;
  const turn =
    (spec.rot === undefined ? '' : ` rot="${String(Math.round(spec.rot * DEGREE))}"`) +
    (spec.flipH === true ? ' flipH="1"' : '') +
    (spec.flipV === true ? ' flipV="1"' : '');
  const xml =
    `<p:sldMaster ${NS}><p:cSld><p:spTree>` +
    '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>' +
    '<p:sp><p:nvSpPr><p:cNvPr id="7" name="probe"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr>' +
    `<a:xfrm${turn}><a:off x="500000" y="400000"/>` +
    `<a:ext cx="${String(spec.cx ?? SIZE.cx)}" cy="${String(spec.cy ?? SIZE.cy)}"/></a:xfrm>` +
    geom +
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

function walk(node: SvgNode, out: SvgElement[] = []): SvgElement[] {
  if (node.kind !== 'element') return out;
  out.push(node);
  for (const child of node.children) walk(child, out);
  return out;
}

const byRole = (node: SvgNode, role: string): SvgElement[] =>
  walk(node).filter((el) => el.attrs['data-role'] === role);

/* -------------------------------------------------------------------------- */
/* the transform, against the browser's own matrix                            */
/* -------------------------------------------------------------------------- */

const SVG_NS = 'http://www.w3.org/2000/svg';

/** What Chromium makes of the attribute the renderer emits. */
function browserPoint(frame: Frame, point: Vec): Vec {
  const svg = document.createElementNS(SVG_NS, 'svg');
  const group = document.createElementNS(SVG_NS, 'g');
  const transform = frameTransform(frame);
  if (transform !== '') group.setAttribute('transform', transform);
  svg.appendChild(group);
  document.body.appendChild(svg);
  try {
    const matrix = group.getCTM() ?? svg.createSVGMatrix();
    const p = svg.createSVGPoint();
    p.x = point.x;
    p.y = point.y;
    const moved = p.matrixTransform(matrix);
    return { x: moved.x, y: moved.y };
  } finally {
    svg.remove();
  }
}

const FRAMES: readonly Frame[] = [
  { x: 500000, y: 400000, cx: 3000000, cy: 2000000, rot: 0, flipH: false, flipV: false },
  { x: 500000, y: 400000, cx: 3000000, cy: 2000000, rot: 30, flipH: false, flipV: false },
  { x: 500000, y: 400000, cx: 3000000, cy: 2000000, rot: 30, flipH: true, flipV: false },
  { x: 500000, y: 400000, cx: 3000000, cy: 2000000, rot: 30, flipH: false, flipV: true },
  { x: 500000, y: 400000, cx: 3000000, cy: 2000000, rot: 30, flipH: true, flipV: true },
  { x: 120000, y: 900000, cx: 1200000, cy: 2600000, rot: 200, flipH: true, flipV: false },
  { x: 0, y: 0, cx: 900000, cy: 900000, rot: 137.5, flipH: false, flipV: true },
];

const CORNERS: readonly Vec[] = [
  { x: 0, y: 0 },
  { x: 3000000, y: 0 },
  { x: 0, y: 2000000 },
  { x: 812345, y: 1298765 },
];

/**
 * How far the browser is allowed to be, in EMU.
 *
 * An `SVGMatrix` is single precision. At a slide's magnitude - a coordinate
 * near 1.2 million EMU - a float32 mantissa is worth about a tenth of an EMU,
 * and the observed disagreement is 0.019. One EMU is a hundred-thousandth of a
 * point and eleven orders of magnitude below anything visible, so this bound
 * proves the two transforms are the same transform without pretending the
 * browser can answer more precisely than it can.
 */
const FLOAT32_EMU = 1;

describe('the frame transform, as arithmetic', () => {
  it('agrees with the transform attribute the renderer emits', () => {
    for (const frame of FRAMES) {
      for (const corner of CORNERS) {
        const ours = framePoint(frame, corner);
        const theirs = browserPoint(frame, corner);
        expect(Math.abs(ours.x - theirs.x)).toBeLessThan(FLOAT32_EMU);
        expect(Math.abs(ours.y - theirs.y)).toBeLessThan(FLOAT32_EMU);
      }
    }
  });

  it('inverts exactly, so a pointer lands where a handle was drawn', () => {
    for (const frame of FRAMES) {
      for (const corner of CORNERS) {
        const back = inverseFramePoint(frame, framePoint(frame, corner));
        expect(back.x).toBeCloseTo(corner.x, 6);
        expect(back.y).toBeCloseTo(corner.y, 6);
      }
    }
  });

  it('refutes rotating before mirroring, on the frames where the two differ', () => {
    // The wrong model: rotate, then mirror. `F R(t) F = R(-t)`, so it differs
    // from the measured order by the sign of the angle - on exactly the frames
    // that mirror one axis and are turned.
    const wrong = (frame: Frame, point: Vec): Vec => {
      const radians = (frame.rot * Math.PI) / 180;
      const x0 = point.x - frame.cx / 2;
      const y0 = point.y - frame.cy / 2;
      let x = x0 * Math.cos(radians) - y0 * Math.sin(radians);
      let y = x0 * Math.sin(radians) + y0 * Math.cos(radians);
      if (frame.flipH) x = -x;
      if (frame.flipV) y = -y;
      return { x: x + frame.x + frame.cx / 2, y: y + frame.y + frame.cy / 2 };
    };

    const oneAxis = FRAMES.filter((f) => f.rot !== 0 && f.flipH !== f.flipV);
    expect(oneAxis.length).toBeGreaterThan(0);
    for (const frame of oneAxis) {
      const corner = CORNERS[1] as Vec;
      const theirs = browserPoint(frame, corner);
      expect(Math.abs(framePoint(frame, corner).x - theirs.x)).toBeLessThan(FLOAT32_EMU);
      // And the wrong order really is wrong, rather than a distinction the
      // test cannot see: it misses by more than a shape's width.
      expect(Math.abs(wrong(frame, corner).x - theirs.x)).toBeGreaterThan(1000);
    }
  });
});

/* -------------------------------------------------------------------------- */
/* what the overlay draws                                                     */
/* -------------------------------------------------------------------------- */

describe('the overlay', () => {
  it('draws one outline per resolved path, in its own colour', () => {
    // `smileyFace` is the reason paths are kept separate: a filled head and an
    // unfilled mouth.
    const placed = place({ prst: 'smileyFace' });
    const overlay = shapeOverlay(placed);
    const outlines = byRole(overlay.node, 'outline');
    const drawable = placed.geometry?.paths.filter((p) => p.finite && p.d !== '') ?? [];

    expect(drawable.length).toBeGreaterThan(1);
    expect(outlines).toHaveLength(drawable.length);
    outlines.forEach((el, index) => {
      expect(el.attrs['d']).toBe(drawable[index]?.d);
      expect(el.attrs['stroke']).toBe(PATH_COLORS[index % PATH_COLORS.length]);
      expect(el.attrs['fill']).toBe('none');
    });
    expect(new Set(outlines.map((el) => el.attrs['stroke'])).size).toBe(outlines.length);
  });

  it('marks an unstroked subpath dashed rather than dropping it', () => {
    const placed = place({ prst: 'actionButtonSound' });
    const outlines = byRole(shapeOverlay(placed).node, 'outline');
    const unstroked = placed.geometry?.paths.filter((p) => p.finite && p.d !== '' && !p.stroke);
    expect(unstroked?.length).toBeGreaterThan(0);
    expect(outlines.filter((el) => el.attrs['stroke-dasharray'] !== undefined)).toHaveLength(
      unstroked?.length ?? -1,
    );
  });

  it('draws the text rectangle where geometry says it is', () => {
    const placed = place({ prst: 'chevron' });
    const overlay = shapeOverlay(placed);
    const rect = placed.geometry?.textRect;
    expect(rect).toBeDefined();
    const [drawn] = byRole(overlay.node, 'text-rect');
    expect(drawn?.attrs['x']).toBe(rect?.l);
    expect(drawn?.attrs['width']).toBe((rect?.r ?? 0) - (rect?.l ?? 0));
    // A chevron's text box is inset from its own frame by the arrowhead, which
    // is the whole reason this is drawn rather than assumed.
    expect(rect?.l).toBeGreaterThan(0);
    expect(rect?.r).toBeLessThan(placed.frame.cx);
  });

  it('draws no text rectangle for the five presets that declare none', () => {
    const none = presetNames().filter((name) => getPreset(name)?.rect === null);
    expect(none).toEqual(['chartPlus', 'chartStar', 'chartX', 'line', 'lineInv']);
    for (const name of none) {
      const overlay = shapeOverlay(place({ prst: name }));
      expect(overlay.textRect).toBeNull();
      expect(byRole(overlay.node, 'text-rect')).toHaveLength(0);
    }
  });

  it('draws every connection site, keeping the angle in 60000ths', () => {
    const placed = place({ prst: 'rect' });
    const overlay = shapeOverlay(placed);
    const sites = placed.geometry?.connectionSites ?? [];
    expect(sites.length).toBe(4);
    const drawn = walk(overlay.node).filter((el) => el.attrs['data-site'] !== undefined);
    expect(drawn).toHaveLength(sites.length);
    drawn.forEach((el, index) => {
      expect(el.attrs['data-ang']).toBe(sites[index]?.ang);
    });
    expect(overlay.connectionSites).toEqual(sites);
  });

  it('reports a non-finite path instead of emitting a d that holds NaN', () => {
    // At zero width a handful of presets divide by it. A `d` with a NaN in it
    // is not merely wrong: the browser drops the element and says nothing.
    const placed = place({ prst: 'moon', cx: 0 });
    const overlay = shapeOverlay(placed);
    expect(overlay.brokenPaths.length).toBeGreaterThan(0);
    expect(byRole(overlay.node, 'outline')).toHaveLength(0);
    expect(serializeSvg(overlay.node)).not.toContain('NaN');
  });

  it('dumps every guide, separating the seeded built-ins from the shape', () => {
    const placed = place({ prst: 'roundRect', adj: { adj: 25000 } });
    const overlay = shapeOverlay(placed);
    const own = overlay.guides.find((g) => g.name === 'adj');
    expect(own?.value).toBe(25000);
    expect(own?.builtin).toBe(false);
    // `w` and `h` are two of the 44 seeded built-ins and hold the shape's size.
    expect(overlay.guides.find((g) => g.name === 'w')?.value).toBe(placed.frame.cx);
    expect(overlay.guides.find((g) => g.name === 'w')?.builtin).toBe(true);
    expect(overlay.guides.every((g) => g.finite)).toBe(true);
  });

  it('is empty for a shape with no geometry rather than throwing', () => {
    const placed = place();
    const overlay = shapeOverlay({ ...placed, geometry: null, geometrySource: null });
    expect(overlay.handles).toEqual([]);
    expect(overlay.guides).toEqual([]);
    expect(walk(overlay.node)).toHaveLength(1);
  });

  it('carries the shape transform, so the overlay turns with the shape', () => {
    const placed = place({ rot: 30, flipH: true });
    expect(shapeOverlay(placed).node.attrs['transform']).toBe(frameTransform(placed.frame));
  });
});

/* -------------------------------------------------------------------------- */
/* handles and their range                                                    */
/* -------------------------------------------------------------------------- */

describe('adjust handles', () => {
  it('draws one handle per a:ahLst entry, where geometry puts it', () => {
    const placed = place({ prst: 'blockArc' });
    const source = placed.geometrySource as { geometry: Geometry; adjust: [] };
    const expected = resolveHandles(source.geometry, {
      w: placed.frame.cx,
      h: placed.frame.cy,
    });
    const drawn = byRole(shapeOverlay(placed).node, 'handle');

    expect(expected).toHaveLength(2);
    expect(drawn).toHaveLength(2);
    drawn.forEach((el, index) => {
      const half = DEFAULT_UNIT * 4;
      expect(el.attrs['x']).toBeCloseTo((expected[index]?.pos.x ?? 0) - half, 6);
      expect(el.attrs['y']).toBeCloseTo((expected[index]?.pos.y ?? 0) - half, 6);
      expect(el.attrs['data-kind']).toBe('polar');
    });
    // The first handle drives only the angle; the second drives both.
    expect(drawn[0]?.attrs['data-axes']).toBe('adj1');
    expect(drawn[1]?.attrs['data-axes']).toBe('adj2 adj3');
  });

  it('draws the locus by sampling the real evaluator, not a chord', () => {
    const placed = place({ prst: 'roundRect' });
    const overlay = shapeOverlay(placed);
    const [handle] = overlay.handles;
    const locus = handle?.locus[0] ?? [];
    expect(locus).toHaveLength(LOCUS_SAMPLES);

    // `roundRect`'s handle rides the top edge: `pos` is `(x1, t)`, so every
    // sample shares a `y` and the ends are the shape's own min and max.
    const axis = handle?.resolved.axes[0];
    const source = placed.geometrySource as { geometry: Geometry };
    const at = (value: number): Point => {
      const [first] = resolveHandles(
        source.geometry,
        { w: placed.frame.cx, h: placed.frame.cy },
        { adjust: { [axis?.guide ?? '']: value } },
      );
      return first?.pos ?? { x: 0, y: 0 };
    };
    expect(locus[0]?.x).toBeCloseTo(at(axis?.min ?? 0).x, 6);
    expect(locus[locus.length - 1]?.x).toBeCloseTo(at(axis?.max ?? 0).x, 6);
    expect(new Set(locus.map((p) => p.y)).size).toBe(1);
  });

  it('draws a polar handle a curve, which a two-point range would miss', () => {
    const placed = place({ prst: 'blockArc' });
    const overlay = shapeOverlay(placed);
    const angleLocus = overlay.handles[0]?.locus[0] ?? [];
    expect(angleLocus.length).toBe(LOCUS_SAMPLES);

    // Every sample sits on the ellipse through the handle, so no three of them
    // are collinear - the distance from the straight line between the ends to
    // the middle sample is a real fraction of the shape.
    const first = angleLocus[0] as Point;
    const last = angleLocus[angleLocus.length - 1] as Point;
    const mid = angleLocus[Math.floor(angleLocus.length / 2)] as Point;
    const span = Math.hypot(last.x - first.x, last.y - first.y);
    const area = Math.abs(
      (last.x - first.x) * (first.y - mid.y) - (first.x - mid.x) * (last.y - first.y),
    );
    expect(area / Math.max(span, 1)).toBeGreaterThan(placed.frame.cx * 0.1);
  });

  it('flags an axis whose declared bounds had to be widened', () => {
    // 2.5 found 16 axis-and-size combinations where the shape's own value falls
    // outside the `minX`/`maxX` the definition declares. The overlay colours
    // those differently, and the flag is what it colours from.
    const placed = place({ prst: 'roundRect', adj: { adj: 90000 } });
    const [handle] = shapeOverlay(placed).handles;
    const axis = handle?.resolved.axes[0];
    expect(axis?.widened).toBe(true);
    expect(axis?.max).toBeGreaterThanOrEqual(90000);
  });
});

/* -------------------------------------------------------------------------- */
/* all 187, which is the sub-phase's own verification                         */
/* -------------------------------------------------------------------------- */

describe('the whole preset gallery', () => {
  const names = presetNames();

  it('has 187 presets and draws an overlay for every one', () => {
    expect(names).toHaveLength(187);
    for (const name of names) {
      const overlay = shapeOverlay(place({ prst: name }));
      const markup = serializeSvg(overlay.node);
      expect(markup, name).not.toContain('NaN');
      expect(markup, name).not.toContain('Infinity');
      expect(overlay.brokenPaths, name).toEqual([]);
    }
  });

  it('gives every handled preset a handle with at least one live axis', () => {
    const handled = names.filter((name) => (getPreset(name)?.ahLst.length ?? 0) > 0);
    expect(handled).toHaveLength(120);
    for (const name of handled) {
      const overlay = shapeOverlay(place({ prst: name }));
      expect(overlay.handles.length, name).toBe(getPreset(name)?.ahLst.length);
      for (const handle of overlay.handles) {
        expect(handle.resolved.finite, `${name} handle ${String(handle.index)}`).toBe(true);
        expect(handle.resolved.axes.length, name).toBeGreaterThan(0);
      }
    }
  });

  it('gives every live axis a locus that stays finite across its whole range', () => {
    const handled = names.filter((name) => (getPreset(name)?.ahLst.length ?? 0) > 0);
    let axes = 0;
    for (const name of handled) {
      for (const handle of shapeOverlay(place({ prst: name })).handles) {
        handle.locus.forEach((points, index) => {
          const axis = handle.resolved.axes[index];
          if (axis === undefined || axis.min === axis.max) return;
          axes += 1;
          expect(points.length, `${name} ${axis.guide}`).toBe(LOCUS_SAMPLES);
          for (const point of points) {
            expect(
              Number.isFinite(point.x) && Number.isFinite(point.y),
              `${name} ${axis.guide}`,
            ).toBe(true);
          }
        });
      }
    }
    expect(axes).toBeGreaterThan(120);
  });

  /**
   * A shape collapsed on one axis or both.
   *
   * Not a corrupt file: a user drags a handle onto the opposite edge and this
   * is what the evaluator is handed. The guides then divide by the size, and
   * across the 187 presets that leaves 14 with a non-finite text rectangle at
   * zero width, 7 with a non-finite connection site, and 39 with a handle that
   * cannot be placed. None of that may reach the markup, because a `NaN` in an
   * attribute makes a browser drop the element in silence.
   */
  const DEGENERATE = [
    { label: 'zero width', cx: 0, cy: 1000000 },
    { label: 'zero height', cx: 1000000, cy: 0 },
    { label: 'both zero', cx: 0, cy: 0 },
  ];

  it.each(DEGENERATE)('emits nothing non-finite at $label', ({ cx, cy }) => {
    let brokenSeen = 0;
    let rectDropped = 0;
    let siteDropped = 0;
    let handleDropped = 0;

    for (const name of names) {
      const placed = place({ prst: name, cx, cy });
      const overlay = shapeOverlay(placed);
      const markup = serializeSvg(overlay.node);
      expect(markup, name).not.toContain('NaN');
      expect(markup, name).not.toContain('Infinity');
      brokenSeen += overlay.brokenPaths.length;

      const rect = placed.geometry?.textRect;
      if (
        rect !== null &&
        rect !== undefined &&
        [rect.l, rect.t, rect.r, rect.b].some((v) => !Number.isFinite(v))
      ) {
        rectDropped += 1;
        expect(byRole(overlay.node, 'text-rect'), name).toHaveLength(0);
      }

      const bad = (placed.geometry?.connectionSites ?? []).filter(
        (site) => !Number.isFinite(site.pos.x) || !Number.isFinite(site.pos.y),
      ).length;
      if (bad > 0) {
        siteDropped += 1;
        const drawn = walk(overlay.node).filter((el) => el.attrs['data-site'] !== undefined);
        expect(drawn.length, name).toBe((placed.geometry?.connectionSites.length ?? 0) - bad);
      }

      const unplaceable = overlay.handles.filter((handle) => !handle.resolved.finite).length;
      if (unplaceable > 0) {
        handleDropped += 1;
        expect(byRole(overlay.node, 'handle').length, name).toBe(
          overlay.handles.length - unplaceable,
        );
      }
    }

    // The counts are asserted so the sweep cannot quietly stop exercising the
    // guards - a change that made every preset finite here would pass every
    // assertion above and test nothing at all.
    expect(brokenSeen).toBeGreaterThan(0);
    expect(rectDropped).toBeGreaterThan(0);
    expect(siteDropped).toBeGreaterThan(0);
    expect(handleDropped).toBeGreaterThan(0);
  });

  it('marks a widened bound differently from an ordinary one', () => {
    const ordinary = shapeOverlay(place({ prst: 'roundRect' }));
    const widened = shapeOverlay(place({ prst: 'roundRect', adj: { adj: 90000 } }));
    expect(ordinary.handles[0]?.resolved.axes[0]?.widened).toBe(false);
    expect(widened.handles[0]?.resolved.axes[0]?.widened).toBe(true);

    const strokeOf = (overlay: ReturnType<typeof shapeOverlay>): unknown =>
      byRole(overlay.node, 'handle-locus')[0]?.attrs['stroke'];
    expect(strokeOf(ordinary)).toBeDefined();
    expect(strokeOf(widened)).toBeDefined();
    expect(strokeOf(widened)).not.toBe(strokeOf(ordinary));
  });

  it('resolves a custGeom the same way, since a preset is only one written out', () => {
    // 2.4's conformance test, from the overlay's side: a `custGeom` carrying
    // `roundRect`'s own definition must produce the identical overlay.
    const preset = getPreset('roundRect') as Geometry;
    const size = { w: SIZE.cx, h: SIZE.cy };
    const asPreset = resolveGeometry(preset, size, {
      adjust: [{ name: 'adj', fmla: ['val', '20000'] }],
    });
    const custom = shapeOverlay(
      place({
        custGeom:
          '<a:custGeom><a:avLst><a:gd name="adj" fmla="val 20000"/></a:avLst>' +
          `<a:gdLst>${preset.gdLst
            .map((g) => `<a:gd name="${g.name}" fmla="${g.fmla.join(' ')}"/>`)
            .join('')}</a:gdLst>` +
          '<a:ahLst/><a:cxnLst/>' +
          '<a:pathLst><a:path><a:moveTo><a:pt x="l" y="t"/></a:moveTo>' +
          '<a:lnTo><a:pt x="x1" y="t"/></a:lnTo></a:path></a:pathLst></a:custGeom>',
      }),
    );
    // The shared guide is what matters: the same `adj` through the same
    // evaluator must give the same `x1` either way.
    expect(custom.guides.find((g) => g.name === 'x1')?.value).toBeCloseTo(
      asPreset.guides.get('x1') ?? Number.NaN,
      6,
    );
  });
});
