/** The selected shape's frame, geometry and fill, as rows a person can read. */

import type { Placed } from '@pptx-studio/render-svg';

import { el } from '../element.js';

const EMU_PER_POINT = 12700;

export function describe(placed: Placed): HTMLElement {
  const rows = el('div', 'rows');
  const add = (key: string, value: string): void => {
    const row = el('div', 'row');
    row.append(el('span', 'k', key), el('span', 'v', value));
    rows.append(row);
  };
  const frame = placed.frame;
  const pt = (value: number): string => (value / EMU_PER_POINT).toFixed(1);

  add('name', placed.shape.name === '' ? `#${String(placed.shape.cNvPrId)}` : placed.shape.name);
  add('kind', placed.shape.kind);
  add('position', `${pt(frame.x)}, ${pt(frame.y)} pt`);
  add('size', `${pt(frame.cx)} × ${pt(frame.cy)} pt`);
  if (frame.rot !== 0) add('rotation', `${frame.rot.toFixed(2)}°`);
  if (frame.flipH || frame.flipV) {
    add('mirrored', [frame.flipH ? 'H' : '', frame.flipV ? 'V' : ''].filter(Boolean).join(' + '));
  }
  const geometry = placed.geometrySource?.geometry;
  add('geometry', geometry === undefined ? 'none' : (geometry.name ?? 'a:custGeom'));
  if (placed.geometry !== null) {
    const broken = placed.geometry.paths.filter((p) => !p.finite).length;
    add(
      'paths',
      String(placed.geometry.paths.length) + (broken === 0 ? '' : ` (${String(broken)} broken)`),
    );
  }
  add('fill', placed.fill === null ? 'none' : placed.fill.type);
  add('sheet', placed.sheet.kind);
  return rows;
}
