import { loadDocument } from '@pptx-studio/model';
import { PartStore } from '@pptx-studio/opc';
import { mediaFromStore, renderSlide } from '@pptx-studio/render-svg';

declare const bytes: Uint8Array;

//#region example
const store = PartStore.open(bytes);
const doc = loadDocument(store);
const slide = doc.slides[0];
if (slide === undefined) throw new Error('no slides');

const svg: string = renderSlide(slide, doc.slideSize, {
  width: 1280,
  media: mediaFromStore(store), // pictures as data: URIs
  text: { defaultTextStyle: doc.defaultTextStyle }, // or `text: false` for geometry only
});
//#endregion

export { svg };
