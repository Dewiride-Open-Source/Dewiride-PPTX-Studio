import { inheritanceChain, loadDocument, resolveXfrm } from '@pptx-studio/model';
import { PartStore } from '@pptx-studio/opc';

declare const bytes: Uint8Array;

//#region example
const doc = loadDocument(PartStore.open(bytes));
const slide = doc.slides[0];
if (slide === undefined) throw new Error('no slides');

for (const shape of slide.shapes) {
  const xfrm = resolveXfrm(shape, slide);
  console.log(shape.name, xfrm?.origin, xfrm?.value); // 'Title 1' 'masterPh' { x, y, cx, cy, rot, … }
  console.log(
    inheritanceChain(shape, slide).map((link) => `${link.origin} ${link.sheet.partName}`),
  );
}
//#endregion
