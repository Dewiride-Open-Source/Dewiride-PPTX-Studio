import { quadrantPng } from '../png.ts';
import type { ProbeSlide } from '../package.ts';
import { grid, group, line, prstGeom, scheme, shape, solidFill, srgb } from '../shapes.ts';
import type { ProbeDeck } from '../types.ts';

/**
 * Every way DrawingML can fill a shape.
 *
 * This deck closes three census keys that no generator in the repository had
 * ever emitted, and one of them was a trap worth naming. `blipFill` read zero
 * across every benchmark deck **even though every benchmark deck is full of
 * pictures**, because a `p:pic` holds `<p:blipFill>` in the PresentationML
 * namespace while the census rule matches `{drawingml/2006/main}blipFill`. They
 * are different elements with the same local name, and a picture is not a
 * picture *fill*. The `a:blipFill` the rule is looking for is a shape filled
 * with an image, which lives inside `a:spPr`, and until this deck nothing wrote
 * one. `groupFill` and `customGeom` were simply never written.
 *
 * The image is a 32 x 32 PNG with four coloured quadrants and an asymmetric
 * white mark. That mark is doing work: `a:tile/@flip` and `a:srcRect` are both
 * invisible on a symmetric picture, so a probe that used noise or a solid
 * colour would pass while rendering nothing.
 */

const IMAGE_RID = 'rId2';

// ------------------------------------------------------- slide 1: solid, gradient

function solidAndGradient(): ProbeSlide {
  const at = grid(6, 3);
  let id = 10;
  const stroke = line({ width: 6350, fill: solidFill(scheme('tx1', '<a:alpha val="40000"/>')) });
  const shapes: string[] = [];
  const add = (name: string, fill: string): void => {
    shapes.push(shape({ id: id++, name, ...at(shapes.length), fill, line: stroke }));
  };

  // --- the six colour models, all of which resolve to a solid fill ----------
  add('noFill', '<a:noFill/>');
  add('srgbClr', solidFill(srgb('4472C4')));
  // Transforms apply in **document order**: lumMod then lumOff is not the same
  // colour as lumOff then lumMod, and this is the pair every theme uses.
  add(
    'schemeClr lumMod lumOff',
    solidFill(scheme('accent2', '<a:lumMod val="60000"/><a:lumOff val="40000"/>')),
  );
  add('sysClr', solidFill('<a:sysClr val="windowText" lastClr="000000"/>'));
  // `a:prstClr` names are DrawingML's own and are **not** CSS colour names.
  add('prstClr', solidFill('<a:prstClr val="dkSlateBlue"/>'));
  add('scrgbClr', solidFill('<a:scrgbClr r="20000" g="60000" b="80000"/>'));
  add('hslClr', solidFill('<a:hslClr hue="14400000" sat="80000" lum="50000"/>'));
  add('srgbClr alpha', solidFill(srgb('C0392B', '<a:alpha val="45000"/>')));
  add('schemeClr tint', solidFill(scheme('accent4', '<a:tint val="40000"/>')));
  add(
    'schemeClr shade satMod',
    solidFill(scheme('accent6', '<a:shade val="60000"/><a:satMod val="140000"/>')),
  );

  // --- gradients ------------------------------------------------------------
  const stops = (...entries: readonly (readonly [number, string])[]): string =>
    '<a:gsLst>' +
    entries.map(([pos, color]) => `<a:gs pos="${String(pos)}">${color}</a:gs>`).join('') +
    '</a:gsLst>';
  const twoStop = stops(
    [0, scheme('accent1', '<a:tint val="70000"/>')],
    [100000, scheme('accent1', '<a:shade val="60000"/>')],
  );

  add(
    'gradFill lin scaled=0',
    `<a:gradFill flip="none">${twoStop}<a:lin ang="0" scaled="0"/></a:gradFill>`,
  );
  add(
    'gradFill lin scaled=1',
    `<a:gradFill rotWithShape="1">${twoStop}<a:lin ang="2700000" scaled="1"/></a:gradFill>`,
  );
  add(
    'gradFill three stops',
    '<a:gradFill>' +
      stops([0, srgb('C0392B')], [50000, srgb('F1C40F')], [100000, srgb('27AE60')]) +
      '<a:lin ang="5400000" scaled="0"/></a:gradFill>',
  );
  // `a:path` gradients reverse stop order relative to a linear one: the first
  // stop is the *focus*, not the edge.
  add(
    'gradFill path circle',
    `<a:gradFill>${twoStop}<a:path path="circle">` +
      '<a:fillToRect l="50000" t="50000" r="50000" b="50000"/></a:path></a:gradFill>',
  );
  add(
    'gradFill path rect',
    `<a:gradFill>${twoStop}<a:path path="rect">` +
      '<a:fillToRect l="100000" t="100000"/></a:path></a:gradFill>',
  );
  add(
    'gradFill path shape',
    `<a:gradFill>${twoStop}<a:path path="shape">` +
      '<a:fillToRect l="50000" t="50000" r="50000" b="50000"/></a:path></a:gradFill>',
  );
  // `a:tileRect` is the third child of `a:gradFill` and the one nothing writes.
  add(
    'gradFill tileRect flip',
    `<a:gradFill flip="xy" rotWithShape="0">${twoStop}` +
      '<a:lin ang="1800000" scaled="0"/><a:tileRect r="-50000" b="-50000"/></a:gradFill>',
  );

  return {
    title: 'a03 — solid and gradient fills',
    body: shapes.join(''),
  };
}

// ---------------------------------------------------- slide 2: picture, group

function pictureAndGroup(): ProbeSlide {
  // Three rows for nine cells. The group is the ninth, and a 4 x 2 grid put it
  // a slide-height below the bottom edge: legal markup, unreadable fixture.
  const at = grid(4, 3);
  let id = 40;
  const stroke = line({ width: 6350, fill: solidFill(scheme('tx1', '<a:alpha val="40000"/>')) });
  const shapes: string[] = [];
  const add = (name: string, fill: string): void => {
    shapes.push(shape({ id: id++, name, ...at(shapes.length), fill, line: stroke }));
  };

  const blip = (effects = ''): string =>
    effects === ''
      ? `<a:blip r:embed="${IMAGE_RID}"/>`
      : `<a:blip r:embed="${IMAGE_RID}">${effects}</a:blip>`;

  add('blipFill stretch', `<a:blipFill>${blip()}<a:stretch><a:fillRect/></a:stretch></a:blipFill>`);
  // `a:srcRect` crops the source; `a:fillRect` insets the destination. Two
  // different rectangles, routinely confused, and both are percentages.
  add(
    'blipFill srcRect crop',
    `<a:blipFill>${blip()}<a:srcRect l="25000" t="25000" r="12500" b="12500"/>` +
      '<a:stretch><a:fillRect/></a:stretch></a:blipFill>',
  );
  add(
    'blipFill stretch fillRect',
    `<a:blipFill>${blip()}<a:stretch><a:fillRect l="10000" t="10000" r="10000" b="10000"/></a:stretch></a:blipFill>`,
  );
  add(
    'blipFill tile flip',
    `<a:blipFill dpi="0" rotWithShape="1">${blip()}` +
      '<a:tile tx="0" ty="0" sx="40000" sy="40000" flip="xy" algn="tl"/></a:blipFill>',
  );
  add(
    'blip duotone',
    `<a:blipFill>${blip('<a:duotone><a:prstClr val="black"/><a:srgbClr val="FFC000"/></a:duotone>')}` +
      '<a:stretch><a:fillRect/></a:stretch></a:blipFill>',
  );
  add(
    'blip grayscl alphaModFix',
    `<a:blipFill>${blip('<a:grayscl/><a:alphaModFix amt="60000"/>')}` +
      '<a:stretch><a:fillRect/></a:stretch></a:blipFill>',
  );
  add(
    'blip clrChange',
    `<a:blipFill>${blip(
      '<a:clrChange useA="1"><a:clrFrom><a:srgbClr val="F1C40F"/></a:clrFrom>' +
        '<a:clrTo><a:srgbClr val="F1C40F"><a:alpha val="0"/></a:srgbClr></a:clrTo></a:clrChange>',
    )}<a:stretch><a:fillRect/></a:stretch></a:blipFill>`,
  );
  add(
    'blip biLevel lum',
    `<a:blipFill>${blip('<a:lum bright="20000" contrast="-40000"/><a:biLevel thresh="45000"/>')}` +
      '<a:stretch><a:fillRect/></a:stretch></a:blipFill>',
  );

  // --- the group, which is the only place a:grpFill means anything ----------
  const cell = at(8);
  const wide = { x: cell.x, y: cell.y, cx: 4400000, cy: 1600000 };
  const children = [0, 1, 2].map((n) =>
    shape({
      id: id++,
      name: 'grpFill child ' + String(n + 1),
      x: wide.x + n * 1466666,
      y: wide.y + 200000,
      cx: 1200000,
      cy: 1000000,
      geometry: prstGeom(['ellipse', 'roundRect', 'triangle'][n] ?? 'rect'),
      // Not `a:noFill`. `a:grpFill` inherits the *enclosing group's* fill, and
      // treating the two as the same is how a themed shape renders empty.
      fill: '<a:grpFill/>',
      line: stroke,
    }),
  );
  shapes.push(
    group({
      id: id++,
      name: 'Group with a gradient fill',
      ...wide,
      fill:
        '<a:gradFill rotWithShape="1"><a:gsLst>' +
        `<a:gs pos="0">${scheme('accent5', '<a:tint val="60000"/>')}</a:gs>` +
        `<a:gs pos="100000">${scheme('accent1', '<a:shade val="70000"/>')}</a:gs>` +
        '</a:gsLst><a:lin ang="0" scaled="0"/></a:gradFill>',
      children: children.join(''),
    }),
  );

  return {
    title: 'a03 — picture fills, and a:grpFill inside a group',
    body: shapes.join(''),
    rels: [
      {
        id: IMAGE_RID,
        type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
        target: '../media/image1.png',
      },
    ],
  };
}

// ------------------------------------------------------- slide 3: 54 patterns

/**
 * Every value of `ST_PresetPatternVal`, in the order the enumeration declares
 * them.
 *
 * All 54, deliberately, and not a representative handful. Sub-phase 2.7
 * generates the pattern tiles from Mono's libgdiplus, and a codegen with no
 * fixture to diff against is a codegen nobody can check. This slide is that
 * fixture: the names are here, in a file PowerPoint has opened.
 */
const PATTERNS = [
  'pct5',
  'pct10',
  'pct20',
  'pct25',
  'pct30',
  'pct40',
  'pct50',
  'pct60',
  'pct70',
  'pct75',
  'pct80',
  'pct90',
  'horz',
  'vert',
  'ltHorz',
  'ltVert',
  'dkHorz',
  'dkVert',
  'narHorz',
  'narVert',
  'dashHorz',
  'dashVert',
  'cross',
  'dnDiag',
  'upDiag',
  'ltDnDiag',
  'ltUpDiag',
  'dkDnDiag',
  'dkUpDiag',
  'wdDnDiag',
  'wdUpDiag',
  'dashDnDiag',
  'dashUpDiag',
  'diagCross',
  'smCheck',
  'lgCheck',
  'smGrid',
  'lgGrid',
  'dotGrid',
  'smConfetti',
  'lgConfetti',
  'horzBrick',
  'diagBrick',
  'solidDmnd',
  'openDmnd',
  'dotDmnd',
  'plaid',
  'sphere',
  'weave',
  'divot',
  'shingle',
  'wave',
  'trellis',
  'zigZag',
] as const;

function patterns(): ProbeSlide {
  const at = grid(9, 6, 22860);
  let id = 100;
  return {
    title: 'a03 — all 54 ST_PresetPatternVal values',
    body: PATTERNS.map((preset, index) =>
      shape({
        id: id++,
        name: preset,
        ...at(index),
        fill:
          `<a:pattFill prst="${preset}">` +
          `<a:fgClr>${scheme('tx1')}</a:fgClr>` +
          `<a:bgClr>${scheme('bg1')}</a:bgClr>` +
          '</a:pattFill>',
        line: line({ width: 3175, fill: solidFill(scheme('tx1', '<a:alpha val="30000"/>')) }),
      }),
    ).join(''),
  };
}

export const a03Fills: ProbeDeck = {
  id: 'a03-fills',
  title: 'PPTX Studio corpus: a03 fills',
  description:
    'Every DrawingML fill: the six colour models under a:solidFill, seven a:gradFill shapes ' +
    'including path gradients and a:tileRect, eight a:blipFill variants with crop, tile-flip and ' +
    'a:blip colour effects, a:grpFill inside a real group, and all 54 ST_PresetPatternVal values. ' +
    'First deck in the corpus to emit {a}blipFill, a:grpFill and a non-theme a:pattFill.',
  features: {
    // Chassis 3 + 3 slide titles + 17 fills + 8 picture fills + 3 group
    // children + 54 patterns. The `p:grpSp` is `group`, never `shape`.
    shape: 88,
    placeholder: 6,
    // Every shape but the six placeholders, which carry no geometry at all.
    presetGeom: 82,
    // The theme's two, plus seven on slide 1 and the group's own on slide 2.
    gradientFill: 10,
    blipFill: 8,
    groupFill: 3,
    group: 1,
    patternFill: 54,
  },
  build: () => ({
    title: 'PPTX Studio corpus: a03 fills',
    slides: [solidAndGradient(), pictureAndGroup(), patterns()],
    parts: [
      {
        name: 'ppt/media/image1.png',
        bytes: quadrantPng(),
        contentType: { kind: 'default', extension: 'png', type: 'image/png' },
      },
    ],
  }),
};
