/**
 * The preset gallery: all 187 shapes, with their handles live.
 *
 * Sub-phase 2.11's own verification, and the debugging tool the rest of Phase 2
 * is developed against. Every shape here goes through the whole pipeline - an
 * `a:prstGeom` in a real `p:sldMaster`, parsed by `@pptx-studio/model`, laid out
 * by `render-svg`, mounted by `render-dom` - rather than calling the geometry
 * evaluator directly. A gallery that reached past the renderer would be a
 * geometry demo and would keep working while the renderer was broken.
 *
 * Dragging a handle produces an `a:avLst`, not a position. The shape is then
 * rebuilt from that XML, which is the same round trip an editor will do and the
 * reason the value in the panel is exactly what a saved file would hold.
 */

import { presetNames, type PresetGuide } from '@pptx-studio/geometry';
import { parseSheet, parseTheme, type Sheet } from '@pptx-studio/model';
import { mountOverlay, mountSlide, type MountedOverlay } from '@pptx-studio/render-dom';
import type { OverlayGuide, Placed } from '@pptx-studio/render-svg';
import { parseXmlString } from '@pptx-studio/xml';

const NS =
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"' +
  ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"' +
  ' xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"';

const CLR_MAP =
  '<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2"' +
  ' accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6"' +
  ' hlink="hlink" folHlink="folHlink"/>';

const THEME_XML =
  '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Gallery">' +
  '<a:themeElements><a:clrScheme name="Gallery">' +
  '<a:dk1><a:srgbClr val="16181D"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>' +
  '<a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>' +
  '<a:accent1><a:srgbClr val="4472C4"/></a:accent1><a:accent2><a:srgbClr val="ED7D31"/></a:accent2>' +
  '<a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4>' +
  '<a:accent5><a:srgbClr val="5B9BD5"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6>' +
  '<a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink>' +
  '</a:clrScheme><a:fontScheme name="Gallery">' +
  '<a:majorFont><a:latin typeface="Calibri Light"/></a:majorFont>' +
  '<a:minorFont><a:latin typeface="Calibri"/></a:minorFont></a:fontScheme>' +
  '<a:fmtScheme name="Gallery"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill>' +
  '</a:fillStyleLst><a:lnStyleLst><a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill>' +
  '</a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle>' +
  '</a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill>' +
  '</a:bgFillStyleLst></a:fmtScheme></a:themeElements></a:theme>';

const THEME = parseTheme(parseXmlString(THEME_XML).root, '/ppt/theme/theme1.xml');

/** The stage, in EMU. Four by three, so a tall preset is not cropped. */
const STAGE = { cx: 4064000, cy: 3048000 };
const INSET = 508000;

const FILL = '<a:solidFill><a:srgbClr val="D6E4F5"/></a:solidFill>';
const LINE = '<a:ln w="12700"><a:solidFill><a:srgbClr val="1B4F86"/></a:solidFill></a:ln>';

function guideXml(guides: readonly PresetGuide[]): string {
  if (guides.length === 0) return '<a:avLst/>';
  return `<a:avLst>${guides
    .map((g) => `<a:gd name="${g.name}" fmla="${g.fmla.join(' ')}"/>`)
    .join('')}</a:avLst>`;
}

/** One preset alone on one master, at the stage size. */
function sheetFor(prst: string, adjust: readonly PresetGuide[], size: number): Sheet {
  const cx = STAGE.cx - INSET * 2;
  const cy = Math.round((STAGE.cy - INSET * 2) * size);
  const xml =
    `<p:sldMaster ${NS}><p:cSld><p:spTree>` +
    '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>' +
    `<p:sp><p:nvSpPr><p:cNvPr id="2" name="${prst}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr>` +
    `<a:xfrm><a:off x="${String(INSET)}" y="${String(Math.round((STAGE.cy - cy) / 2))}"/>` +
    `<a:ext cx="${String(cx)}" cy="${String(cy)}"/></a:xfrm>` +
    `<a:prstGeom prst="${prst}">${guideXml(adjust)}</a:prstGeom>${FILL}${LINE}` +
    '</p:spPr></p:sp>' +
    `</p:spTree></p:cSld>${CLR_MAP}</p:sldMaster>`;
  return {
    ...parseSheet(parseXmlString(xml).root, '/ppt/slideMasters/slideMaster1.xml'),
    parent: null,
    theme: THEME,
  };
}

function el<K extends keyof HTMLElementTagNameMap>(
  tag: K,
  className = '',
  text = '',
): HTMLElementTagNameMap[K] {
  const node = document.createElement(tag);
  if (className !== '') node.className = className;
  if (text !== '') node.textContent = text;
  return node;
}

function need(id: string): HTMLElement {
  const node = document.getElementById(id);
  if (node === null) throw new Error(`the page has no #${id}`);
  return node;
}

/* -------------------------------------------------------------------------- */
/* state                                                                      */
/* -------------------------------------------------------------------------- */

const NAMES = presetNames();

interface Selection {
  readonly prst: string;
  adjust: readonly PresetGuide[];
  size: number;
}

let selection: Selection = { prst: 'roundRect', adjust: [], size: 0.7 };
let overlay: MountedOverlay | null = null;

const stage = need('stage');
const grid = need('grid');
const guidesPanel = need('guides');
const adjustPanel = need('adjust');
const xmlPanel = need('xml');
const title = need('shape-name');
const counter = need('counter');

/* -------------------------------------------------------------------------- */
/* the stage                                                                  */
/* -------------------------------------------------------------------------- */

const STAGE_PX = 640;

/**
 * Rebuild the stage from the current selection.
 *
 * Everything downstream of a drag comes through here: the handle writes an
 * `a:avLst`, the shape is rebuilt from that XML, and the handles are resolved
 * again at whatever the new geometry says. Only the stage is touched - the
 * thumbnail list is left alone, because rebuilding 187 of them inside a
 * `pointermove` is how a smooth drag becomes a slideshow.
 */
function draw(): void {
  const sheet = sheetFor(selection.prst, selection.adjust, selection.size);
  const mounted = mountSlide(stage, sheet, STAGE, { width: STAGE_PX, height: 480 });
  const placed = mounted.placed[0];
  if (placed === undefined) return;

  overlay?.unmount();
  overlay = mountOverlay(mounted.root, placed, {
    // EMU per pixel of the stage, so the chrome is the same weight here as it
    // will be in the editor at 100%.
    unit: STAGE.cx / STAGE_PX,
    onChange: (edit) => {
      selection = { ...selection, adjust: edit.avLst };
      draw();
    },
  });
  title.textContent = selection.prst;
  showPanels(placed);
}

/* -------------------------------------------------------------------------- */
/* the panels                                                                 */
/* -------------------------------------------------------------------------- */

function guideRows(guides: readonly OverlayGuide[]): HTMLElement {
  const list = el('div', 'rows');
  const shown = guides.filter((g) => !g.builtin).concat(guides.filter((g) => g.builtin));
  for (const guide of shown) {
    const row = el('div', `row${guide.builtin ? ' dim' : ''}${guide.finite ? '' : ' bad'}`);
    row.append(el('span', 'k', guide.name), el('span', 'v', format(guide.value)));
    list.append(row);
  }
  return list;
}

function format(value: number): string {
  if (!Number.isFinite(value)) return String(value);
  return Math.abs(value) >= 1000 ? Math.round(value).toLocaleString('en') : value.toFixed(2);
}

function showPanels(placed: Placed): void {
  const data = overlay?.overlay;
  if (data === undefined) return;

  guidesPanel.replaceChildren(guideRows(data.guides));

  adjustPanel.replaceChildren();
  if (data.handles.length === 0) {
    adjustPanel.append(el('p', 'note', 'This preset declares no adjust handles.'));
  }
  for (const handle of data.handles) {
    const box = el('div', 'handle');
    box.append(el('h4', '', `handle ${String(handle.index + 1)} · ${handle.resolved.kind}`));
    for (const axis of handle.resolved.axes) {
      const row = el('div', `row${axis.widened ? ' widened' : ''}`);
      row.append(
        el('span', 'k', axis.guide),
        el('span', 'v', format(axis.value)),
        el('span', 'range', `${format(axis.min)} … ${format(axis.max)}`),
      );
      box.append(row);
    }
    if (handle.resolved.axes.some((axis) => axis.widened)) {
      box.append(
        el('p', 'note', 'A bound was widened: the shape sits outside its declared range.'),
      );
    }
    adjustPanel.append(box);
  }

  const broken =
    data.brokenPaths.length === 0
      ? ''
      : `\n<!-- ${String(data.brokenPaths.length)} path(s) did not evaluate finitely -->`;
  const avLst =
    selection.adjust.length === 0
      ? '  <a:avLst/>'
      : [
          '  <a:avLst>',
          ...selection.adjust.map((g) => `    <a:gd name="${g.name}" fmla="${g.fmla.join(' ')}"/>`),
          '  </a:avLst>',
        ].join('\n');
  xmlPanel.textContent = `<a:prstGeom prst="${selection.prst}">\n${avLst}\n</a:prstGeom>${broken}`;

  const label = placed.geometry?.paths.length ?? 0;
  counter.textContent = `${String(label)} path${label === 1 ? '' : 's'} · ${String(
    data.connectionSites.length,
  )} connection site${data.connectionSites.length === 1 ? '' : 's'} · ${String(
    data.guides.length,
  )} guides`;
}

/* -------------------------------------------------------------------------- */
/* the list                                                                   */
/* -------------------------------------------------------------------------- */

function thumbnail(prst: string): HTMLElement {
  const cell = el('button', 'cell');
  cell.type = 'button';
  cell.dataset['prst'] = prst;
  const holder = el('div', 'thumb');
  mountSlide(holder, sheetFor(prst, [], 0.7), STAGE, { width: 104, height: 78 });
  cell.append(holder, el('span', 'name', prst));
  cell.addEventListener('click', () => {
    selection = { prst, adjust: [], size: selection.size };
    markSelected();
    draw();
  });
  return cell;
}

function markSelected(): void {
  for (const cell of grid.querySelectorAll<HTMLElement>('.cell')) {
    cell.classList.toggle('on', cell.dataset['prst'] === selection.prst);
  }
}

function buildGrid(filter: string): void {
  const needle = filter.trim().toLowerCase();
  const shown = needle === '' ? NAMES : NAMES.filter((n) => n.toLowerCase().includes(needle));
  grid.replaceChildren(...shown.map(thumbnail));
  markSelected();
  need('found').textContent = `${String(shown.length)} of ${String(NAMES.length)}`;
}

/* -------------------------------------------------------------------------- */

function wire(): void {
  const search = need('search') as HTMLInputElement;
  search.addEventListener('input', () => buildGrid(search.value));

  const size = need('size') as HTMLInputElement;
  size.addEventListener('input', () => {
    selection = { ...selection, size: Number(size.value) / 100 };
    draw();
  });

  need('reset').addEventListener('click', () => {
    selection = { ...selection, adjust: [] };
    draw();
  });
}

wire();
buildGrid('');
draw();
