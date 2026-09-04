/**
 * The live renderer, in a real browser.
 *
 * The suite runs in Chromium rather than jsdom - the repository's deliberate
 * choice - and here it earns its keep twice over: `createElementNS` versus
 * `createElement` is the difference between a shape and an invisible HTML
 * element spelled `path`, and jsdom will happily accept the wrong one.
 *
 * The load-bearing assertion is the last one: the markup this package mounts
 * and the markup `render-svg` serialises come from the same tree, so they must
 * agree character for character. That is what "two renderers over one layout
 * engine" has to mean if it is to mean anything.
 */

import { parseSheet, parseTheme, type Sheet } from '@pptx-studio/model';
import { renderSlide } from '@pptx-studio/render-svg';
import { parseXmlString } from '@pptx-studio/xml';
import { describe, expect, it } from 'vitest';

import { RenderDomError } from './errors.js';
import { mountSlide, renderSlideMarkup } from './mount.js';

const NS =
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"' +
  ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"' +
  ' xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"';

const SP_TREE_HEAD =
  '<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
  '<p:grpSpPr/>';

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

const SIZE = { cx: 12192000, cy: 6858000 };

function shape(id: number, x: number, y: number, cx: number, cy: number, hex: string): string {
  return (
    '<p:sp><p:nvSpPr>' +
    `<p:cNvPr id="${String(id)}" name="shape${String(id)}"/><p:cNvSpPr/><p:nvPr/>` +
    '</p:nvSpPr><p:spPr>' +
    `<a:xfrm><a:off x="${String(x)}" y="${String(y)}"/><a:ext cx="${String(cx)}" cy="${String(cy)}"/></a:xfrm>` +
    '<a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom>' +
    `<a:solidFill><a:srgbClr val="${hex}"/></a:solidFill>` +
    '</p:spPr></p:sp>'
  );
}

function sheetOf(shapes: readonly string[]): Sheet {
  const body = (kind: 'slide' | 'layout' | 'master', inner: readonly string[]): string => {
    const tag = kind === 'slide' ? 'p:sld' : kind === 'layout' ? 'p:sldLayout' : 'p:sldMaster';
    const attrs = kind === 'layout' ? ' type="obj"' : '';
    const tail = kind === 'master' ? CLR_MAP : '<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>';
    return (
      `<${tag} ${NS}${attrs}><p:cSld name="${kind}">${SP_TREE_HEAD}${inner.join('')}` +
      `</p:spTree></p:cSld>${tail}</${tag}>`
    );
  };
  const parsed = (kind: 'slide' | 'layout' | 'master', inner: readonly string[]) =>
    parseSheet(parseXmlString(body(kind, inner)).root, `/ppt/${kind}.xml`);
  const theme = parseTheme(parseXmlString(THEME_XML).root, '/ppt/theme/theme1.xml');
  const master: Sheet = { ...parsed('master', []), parent: null, theme };
  const layout: Sheet = { ...parsed('layout', []), parent: master, theme: null };
  return { ...parsed('slide', shapes), parent: layout, theme: null };
}

function host(): HTMLElement {
  const element = document.createElement('div');
  document.body.appendChild(element);
  return element;
}

describe('mountSlide', () => {
  it('builds SVG elements, not HTML elements that happen to be spelled path', () => {
    const mounted = mountSlide(host(), sheetOf([shape(2, 0, 0, 914400, 457200, 'FF0000')]), SIZE);
    expect(mounted.root.namespaceURI).toBe('http://www.w3.org/2000/svg');
    const path = mounted.root.querySelector('path');
    expect(path).not.toBeNull();
    expect(path?.namespaceURI).toBe('http://www.w3.org/2000/svg');
    // The one thing an HTML element spelled `path` cannot do.
    expect(typeof (path as SVGPathElement).getTotalLength).toBe('function');
    expect((path as SVGPathElement).getTotalLength()).toBeGreaterThan(0);
    mounted.unmount();
  });

  it('puts a shape where the layout says, in the browser own reckoning', () => {
    const mounted = mountSlide(
      host(),
      sheetOf([shape(2, 914400, 457200, 914400, 457200, '00FF00')]),
      SIZE,
    );
    expect(mounted.element(2)).not.toBeNull();
    // `getBBox` on the group returns its *own* user space, which is before its
    // transform - so the question goes to the root, where the units are EMU and
    // the answer includes every transform between.
    const box = mounted.root.getBBox();
    expect(box.x).toBeCloseTo(914400, 0);
    expect(box.y).toBeCloseTo(457200, 0);
    expect(box.width).toBeCloseTo(914400, 0);
    mounted.unmount();
  });

  it('finds a shape by its cNvPr id, and does not invent one', () => {
    const mounted = mountSlide(
      host(),
      sheetOf([shape(2, 0, 0, 100, 100, '000000'), shape(3, 200, 0, 100, 100, 'FFFFFF')]),
      SIZE,
    );
    expect(mounted.element(2)).not.toBeNull();
    expect(mounted.element(3)).not.toBeNull();
    expect(mounted.element(99)).toBeNull();
    mounted.unmount();
  });

  it('resizes without re-laying anything out', () => {
    const mounted = mountSlide(host(), sheetOf([shape(2, 0, 0, 100, 100, '000000')]), SIZE);
    const before = mounted.root.querySelector('path')?.getAttribute('d');
    mounted.resize(640, 360);
    expect(mounted.root.getAttribute('width')).toBe('640');
    expect(mounted.root.getAttribute('height')).toBe('360');
    // The viewBox does the work, so nothing about the geometry changed.
    expect(mounted.root.querySelector('path')?.getAttribute('d')).toBe(before);
    mounted.unmount();
  });

  it('empties the host and can be detached again', () => {
    const container = host();
    container.appendChild(document.createElement('span'));
    const mounted = mountSlide(container, sheetOf([shape(2, 0, 0, 100, 100, '000000')]), SIZE);
    expect(container.querySelector('span')).toBeNull();
    expect(container.children.length).toBe(1);
    mounted.unmount();
    expect(container.children.length).toBe(0);
  });

  it('refuses a host that is not an element', () => {
    expect(() => mountSlide(null as unknown as Element, sheetOf([]), SIZE)).toThrow(RenderDomError);
    try {
      mountSlide({} as Element, sheetOf([]), SIZE);
    } catch (error) {
      expect((error as RenderDomError).code).toBe('RENDER_DOM_NO_HOST');
    }
  });

  it('mounts the same tree render-svg serializes', () => {
    const sheet = sheetOf([
      shape(2, 0, 0, 914400, 457200, 'FF0000'),
      shape(3, 914400, 0, 457200, 457200, '0000FF'),
    ]);
    const mounted = mountSlide(host(), sheet, SIZE, { idPrefix: 'same' });
    const fromString = renderSlide(sheet, SIZE, { idPrefix: 'same' });

    // The browser's own serializer over the mounted nodes, against the string
    // emitter's output. Attribute order is insertion order in both, so this is
    // an exact comparison rather than a structural one.
    const serialized = new XMLSerializer().serializeToString(mounted.root);
    expect(serialized.replace(/ xmlns="[^"]*"/, ' xmlns="http://www.w3.org/2000/svg"')).toBe(
      fromString,
    );
    expect(renderSlideMarkup(sheet, SIZE, { idPrefix: 'same' })).toBe(fromString);
    mounted.unmount();
  });
});
