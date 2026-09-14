/**
 * The editor's gestures over every slide the site ships: what they write sits in
 * schema order, parses back into a sheet, and their inverses restore the part
 * byte for byte.
 */

import assert from 'node:assert/strict';
import { readdirSync, readFileSync } from 'node:fs';
import { join } from 'node:path';
import { describe, it } from 'node:test';

import { parseSheet } from '@pptx-studio/model';
import { PartStore } from '@pptx-studio/opc';
import {
  applyEdits,
  childElements,
  descendantElements,
  outOfOrderChildren,
  parseXml,
  serializeXml,
  type XDocument,
  type XElement,
} from '@pptx-studio/xml';

import {
  NotEditable,
  applyAll,
  planDelete,
  planMove,
  planRecolour,
  planRecolourText,
  planRetext,
} from './edits.ts';
import {
  NS_A,
  NS_P,
  attributeOf,
  child,
  childScale,
  fillElement,
  is,
  paragraphElements,
  segmentsOf,
  shapeProperties,
  textBody,
  transformElement,
} from './locate.ts';

const DECKS = join(import.meta.dirname, '../../../public/decks');
const FRAME = {
  x: 914400,
  y: 685800,
  cx: 7315200,
  cy: 1143000,
  rot: 0,
  flipH: false,
  flipV: false,
};

interface Slide {
  readonly deck: string;
  readonly partName: string;
  readonly source: Uint8Array;
}

function* slides(): Generator<Slide, void, undefined> {
  for (const file of readdirSync(DECKS).filter((name) => /\.pptx$|\.pptm$/.test(name))) {
    const store = PartStore.open(new Uint8Array(readFileSync(join(DECKS, file))));
    for (const partName of store.partNames.filter((name) =>
      /^\/ppt\/slides\/slide\d+\.xml$/.test(name),
    )) {
      yield { deck: file, partName, source: store.read(partName) };
    }
  }
}

const shapesOf = (tree: XDocument): XElement[] =>
  [...descendantElements(tree.root)].filter(
    (one) =>
      namespaceIsP(one) && ['sp', 'pic', 'grpSp', 'graphicFrame', 'cxnSp'].includes(one.local),
  );
const namespaceIsP = (element: XElement): boolean => is(element, NS_P, element.local);

/** Nothing this edit touched sits out of schema order. */
function inOrder(root: XElement, where: string): void {
  for (const element of [root, ...descendantElements(root)]) {
    assert.deepEqual(outOfOrderChildren(element), [], `${where}: ${element.qname}`);
  }
}

/** Apply, check, undo, and hold the part to its original bytes. */
function roundTrip(
  slide: Slide,
  tree: XDocument,
  plan: Parameters<typeof applyAll>[0],
  check: () => void,
): void {
  const inverses = applyAll(plan);
  check();
  parseSheet(tree.root, slide.partName);
  applyEdits(inverses);
  assert.deepEqual(serializeXml(tree), slide.source, `${slide.deck} ${slide.partName} restored`);
}

describe('retext', () => {
  it('writes every paragraph it is given, in order, and no others', () => {
    let seen = 0;
    for (const slide of slides()) {
      const tree = parseXml(slide.source);
      for (const shape of shapesOf(tree)) {
        const body = textBody(shape);
        if (body === undefined) continue;
        seen += 1;
        const wanted = [['Hello'], ['one', 'soft break'], ['']];
        roundTrip(slide, tree, planRetext(shape, wanted), () => {
          assert.deepEqual(
            paragraphElements(body).map(segmentsOf),
            wanted,
            `${slide.deck} ${slide.partName}`,
          );
          inOrder(body, `${slide.deck} ${slide.partName}`);
        });
      }
    }
    assert.ok(seen > 50, `${String(seen)} text bodies`);
  });

  it('leaves a paragraph whose text did not change untouched, and drops surplus ones', () => {
    for (const slide of slides()) {
      const tree = parseXml(slide.source);
      for (const shape of shapesOf(tree)) {
        const body = textBody(shape);
        if (body === undefined) continue;
        const before = paragraphElements(body);
        if (before.length < 2) continue;
        const first = before[0];
        assert.ok(first !== undefined);
        const wanted = [segmentsOf(first)];
        roundTrip(slide, tree, planRetext(shape, wanted), () => {
          assert.equal(paragraphElements(body).length, 1);
          assert.equal(first.dirty, false, 'the unchanged paragraph is still clean');
        });
        return;
      }
    }
    assert.fail('no slide with a two-paragraph shape');
  });

  it("keeps the first run's properties when the runs collapse", () => {
    for (const slide of slides()) {
      const tree = parseXml(slide.source);
      for (const shape of shapesOf(tree)) {
        const body = textBody(shape);
        if (body === undefined) continue;
        const paragraph = paragraphElements(body).find(
          (one) => childElements(one).filter((run) => is(run, NS_A, 'r')).length >= 2,
        );
        if (paragraph === undefined) continue;
        const firstRun = childElements(paragraph).find((run) => is(run, NS_A, 'r'));
        assert.ok(firstRun !== undefined);
        const rPr = child(firstRun, NS_A, 'rPr');
        roundTrip(slide, tree, planRetext(shape, [['typed over']]), () => {
          const runs = childElements(paragraph).filter((run) => is(run, NS_A, 'r'));
          assert.equal(runs.length, 1);
          assert.equal(runs[0], firstRun);
          assert.equal(child(firstRun, NS_A, 'rPr'), rPr);
        });
        return;
      }
    }
    assert.fail('no slide with a two-run paragraph');
  });

  it('refuses a shape without text', () => {
    for (const slide of slides()) {
      const tree = parseXml(slide.source);
      const picture = shapesOf(tree).find((one) => is(one, NS_P, 'pic'));
      if (picture === undefined) continue;
      assert.throws(() => applyAll(planRetext(picture, [['x']])), NotEditable);
      return;
    }
    assert.fail('no slide with a picture');
  });
});

describe('move', () => {
  it("shifts a declared transform by the delta, in the group's units inside a group", () => {
    let seen = 0;
    for (const slide of slides()) {
      const tree = parseXml(slide.source);
      for (const shape of shapesOf(tree)) {
        const xfrm = transformElement(shape);
        const off = xfrm === undefined ? undefined : child(xfrm, NS_A, 'off');
        if (off === undefined) continue;
        seen += 1;
        const x = Number(attributeOf(off, 'x'));
        const y = Number(attributeOf(off, 'y'));
        const { sx, sy } = childScale(shape);
        roundTrip(slide, tree, planMove(shape, FRAME, 12700, -25400), () => {
          assert.equal(Number(attributeOf(off, 'x')), Math.round(x + 12700 / sx));
          assert.equal(Number(attributeOf(off, 'y')), Math.round(y - 25400 / sy));
        });
      }
    }
    assert.ok(seen > 50, `${String(seen)} declared transforms`);
  });

  it('gives a shape that inherits its position one, at the frame it resolved to', () => {
    let seen = 0;
    for (const slide of slides()) {
      const tree = parseXml(slide.source);
      for (const shape of shapesOf(tree)) {
        if (transformElement(shape) !== undefined || shapeProperties(shape) === undefined) continue;
        seen += 1;
        roundTrip(slide, tree, planMove(shape, FRAME, 12700, 0), () => {
          const xfrm = transformElement(shape);
          assert.ok(xfrm !== undefined);
          const properties = shapeProperties(shape);
          assert.ok(properties !== undefined);
          assert.equal(childElements(properties)[0], xfrm, 'a:xfrm is the first child of p:spPr');
          assert.equal(attributeOf(child(xfrm, NS_A, 'off')!, 'x'), String(FRAME.x + 12700));
          assert.equal(attributeOf(child(xfrm, NS_A, 'ext')!, 'cx'), String(FRAME.cx));
          inOrder(properties, `${slide.deck} ${slide.partName}`);
        });
      }
    }
    assert.ok(seen > 0, `${String(seen)} inherited positions`);
  });
});

describe('recolour', () => {
  it('writes a solid fill over whatever fill the shape declared, or none', () => {
    const kinds = new Map<string, number>();
    for (const slide of slides()) {
      const tree = parseXml(slide.source);
      for (const shape of shapesOf(tree)) {
        if (!is(shape, NS_P, 'sp')) continue;
        const properties = shapeProperties(shape);
        assert.ok(properties !== undefined);
        const had = fillElement(properties)?.local ?? 'inherited';
        kinds.set(had, (kinds.get(had) ?? 0) + 1);
        roundTrip(slide, tree, planRecolour(shape, '#e8453c'), () => {
          const fill = fillElement(properties);
          assert.ok(fill !== undefined && is(fill, NS_A, 'solidFill'));
          const color = childElements(fill)[0];
          assert.ok(color !== undefined && is(color, NS_A, 'srgbClr'));
          assert.equal(attributeOf(color, 'val'), 'E8453C');
          assert.equal(
            childElements(properties).filter((one) => fillElement(one) === one).length,
            0,
          );
          inOrder(properties, `${slide.deck} ${slide.partName}`);
        });
      }
    }
    for (const kind of ['inherited', 'noFill', 'solidFill']) {
      assert.ok((kinds.get(kind) ?? 0) > 0, `${kind} covered`);
    }
  });

  it('keeps the transparency of the colour it replaces', () => {
    for (const slide of slides()) {
      const tree = parseXml(slide.source);
      for (const shape of shapesOf(tree)) {
        const properties = shapeProperties(shape);
        const fill = properties === undefined ? undefined : fillElement(properties);
        const color = fill === undefined ? undefined : childElements(fill)[0];
        const alpha = color === undefined ? undefined : child(color, NS_A, 'alpha');
        if (alpha === undefined || properties === undefined) continue;
        const value = attributeOf(alpha, 'val');
        roundTrip(slide, tree, planRecolour(shape, '#0b62b0'), () => {
          const now = fillElement(properties);
          const carried = child(childElements(now!)[0]!, NS_A, 'alpha');
          assert.equal(attributeOf(carried!, 'val'), value);
        });
        return;
      }
    }
    assert.fail('no slide with a transparent solid fill');
  });

  it("colours a connector's line and refuses a picture", () => {
    let connectors = 0;
    let pictures = 0;
    for (const slide of slides()) {
      const tree = parseXml(slide.source);
      for (const shape of shapesOf(tree)) {
        if (is(shape, NS_P, 'pic')) {
          pictures += 1;
          assert.throws(() => applyAll(planRecolour(shape, '#000000')), NotEditable);
        }
        if (!is(shape, NS_P, 'cxnSp')) continue;
        connectors += 1;
        roundTrip(slide, tree, planRecolour(shape, '#2fa869'), () => {
          const line = child(shapeProperties(shape)!, NS_A, 'ln');
          assert.ok(line !== undefined);
          const fill = fillElement(line);
          assert.ok(fill !== undefined && is(fill, NS_A, 'solidFill'));
          inOrder(shapeProperties(shape)!, `${slide.deck} ${slide.partName}`);
        });
      }
    }
    assert.ok(
      connectors > 0 && pictures > 0,
      `${String(connectors)} connectors, ${String(pictures)} pictures`,
    );
  });
});

describe('recolour text', () => {
  const solid = (props: XElement | undefined, where: string): void => {
    assert.ok(props !== undefined, `${where}: run properties present`);
    const fill = fillElement(props);
    assert.ok(fill !== undefined && is(fill, NS_A, 'solidFill'), `${where}: a solid fill`);
    const color = childElements(fill)[0];
    assert.ok(color !== undefined && is(color, NS_A, 'srgbClr'));
    assert.equal(attributeOf(color, 'val'), '2FA869');
  };

  it('writes the colour into every run, break, field and paragraph end', () => {
    const had = new Map<string, number>();
    for (const slide of slides()) {
      const tree = parseXml(slide.source);
      for (const shape of shapesOf(tree)) {
        const body = textBody(shape);
        if (body === undefined) continue;
        for (const paragraph of paragraphElements(body)) {
          for (const holder of childElements(paragraph)) {
            if (!['r', 'br', 'fld'].includes(holder.local)) continue;
            const rPr = child(holder, NS_A, 'rPr');
            const kind = rPr === undefined ? 'no rPr' : (fillElement(rPr)?.local ?? 'inherited');
            had.set(kind, (had.get(kind) ?? 0) + 1);
          }
        }
        const where = `${slide.deck} ${slide.partName}`;
        roundTrip(slide, tree, planRecolourText(shape, '#2fa869'), () => {
          for (const paragraph of paragraphElements(body)) {
            const content = childElements(paragraph).filter((one) =>
              ['r', 'br', 'fld'].includes(one.local),
            );
            for (const holder of content) solid(child(holder, NS_A, 'rPr'), where);
            const end = child(paragraph, NS_A, 'endParaRPr');
            if (content.length === 0 || end !== undefined) solid(end, where);
          }
          inOrder(body, where);
        });
      }
    }
    for (const kind of ['inherited', 'solidFill']) {
      assert.ok((had.get(kind) ?? 0) > 0, `${kind} covered`);
    }
  });

  it('writes the properties a run or an empty paragraph lacks, first among its children', () => {
    // Every run the site ships carries an a:rPr, so the case is made by stripping them.
    const content = (paragraph: XElement): XElement[] =>
      childElements(paragraph).filter((one) => ['r', 'br', 'fld'].includes(one.local));
    const holds = (kind: 'run' | 'empty') => (shape: XElement) => {
      const body = textBody(shape);
      return (
        body !== undefined &&
        paragraphElements(body).some((p) => (kind === 'run') === content(p).length > 0)
      );
    };
    const seen = new Set<string>();
    for (const slide of slides()) {
      for (const kind of ['run', 'empty'] as const) {
        if (seen.has(kind)) continue;
        const shaped = parseXml(slide.source);
        const index = shapesOf(shaped).findIndex(holds(kind));
        if (index < 0) continue;
        seen.add(kind);
        applyEdits(
          [...descendantElements(textBody(shapesOf(shaped)[index]!)!)]
            .filter((one) => is(one, NS_A, 'rPr') || is(one, NS_A, 'endParaRPr'))
            .map((node) => ({ kind: 'removeChild' as const, parent: node.parent!, node })),
        );
        const stripped = { ...slide, source: serializeXml(shaped) };
        const tree = parseXml(stripped.source);
        const target = shapesOf(tree)[index];
        assert.ok(target !== undefined);
        const body = textBody(target);
        assert.ok(body !== undefined);
        roundTrip(stripped, tree, planRecolourText(target, '#2fa869'), () => {
          for (const paragraph of paragraphElements(body)) {
            for (const holder of content(paragraph)) {
              assert.equal(childElements(holder)[0]?.local, 'rPr', 'a:rPr first in its run');
              solid(child(holder, NS_A, 'rPr'), slide.partName);
            }
            const end = child(paragraph, NS_A, 'endParaRPr');
            if (content(paragraph).length === 0) solid(end, slide.partName);
            else assert.equal(end, undefined, 'no end properties invented');
          }
          inOrder(body, slide.partName);
        });
      }
      if (seen.size === 2) return;
    }
    assert.fail(`covered: ${[...seen].join(', ')}`);
  });

  it("keeps a run's other properties and its place", () => {
    for (const slide of slides()) {
      const tree = parseXml(slide.source);
      for (const shape of shapesOf(tree)) {
        const body = textBody(shape);
        if (body === undefined) continue;
        const rPr = [...descendantElements(body)].find(
          (one) =>
            is(one, NS_A, 'rPr') && fillElement(one) === undefined && one.attributes.length > 1,
        );
        if (rPr === undefined) continue;
        const attributes = rPr.attributes.map((attr) => [attr.qname, attr.value]);
        const run = rPr.parent;
        assert.ok(run !== undefined);
        roundTrip(slide, tree, planRecolourText(shape, '#2fa869'), () => {
          assert.equal(childElements(run)[0], rPr, 'a:rPr is still the first child of its run');
          assert.deepEqual(
            rPr.attributes.map((attr) => [attr.qname, attr.value]),
            attributes,
          );
          assert.equal(childElements(rPr).length, 1, 'the fill is all that was added');
        });
        return;
      }
    }
    assert.fail('no slide with an attributed, inherited-colour run');
  });

  it('refuses a shape without text', () => {
    for (const slide of slides()) {
      const tree = parseXml(slide.source);
      const picture = shapesOf(tree).find((one) => is(one, NS_P, 'pic'));
      if (picture === undefined) continue;
      assert.throws(() => applyAll(planRecolourText(picture, '#000000')), NotEditable);
      return;
    }
    assert.fail('no slide with a picture');
  });
});

describe('delete', () => {
  it('takes any top-level shape out and puts it back where it was', () => {
    let seen = 0;
    for (const slide of slides()) {
      const tree = parseXml(slide.source);
      for (const shape of shapesOf(tree)) {
        const parent = shape.parent;
        assert.ok(parent !== undefined);
        const index = parent.children.indexOf(shape);
        seen += 1;
        roundTrip(slide, tree, planDelete(shape), () => {
          assert.equal(parent.children.includes(shape), false);
        });
        assert.equal(parent.children.indexOf(shape), index);
      }
    }
    assert.ok(seen > 100, `${String(seen)} shapes`);
  });
});

describe('applyAll', () => {
  it('rolls an earlier batch back when a later one is refused', () => {
    const slide = slides().next().value;
    assert.ok(slide !== undefined);
    const tree = parseXml(slide.source);
    const shape = shapesOf(tree).find((one) => textBody(one) !== undefined);
    assert.ok(shape !== undefined);
    function* plan(): ReturnType<typeof planDelete> {
      yield* planRecolour(shape!, '#ffffff');
      throw new NotEditable('the second batch');
    }
    assert.throws(() => applyAll(plan()), NotEditable);
    assert.deepEqual(serializeXml(tree), slide.source);
  });
});
