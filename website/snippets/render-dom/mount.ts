import { loadDocument } from '@pptx-studio/model';
import { PartStore } from '@pptx-studio/opc';
import { mountOverlay, mountSlide } from '@pptx-studio/render-dom';
import { createTextEngine, mediaFromStore } from '@pptx-studio/render-svg';

declare const bytes: Uint8Array;

//#region example
const store = PartStore.open(bytes);
const doc = loadDocument(store);
const slide = doc.slides[0];
const host = document.querySelector('#stage');
if (slide === undefined || host === null) throw new Error('nothing to mount');

const text = createTextEngine({ defaultTextStyle: doc.defaultTextStyle }); // one measurer, every mount
const mounted = mountSlide(host, slide, doc.slideSize, {
  width: 960, // one CSS pixel per point
  devicePixelRatio: window.devicePixelRatio, // strokes round to device pixels
  media: mediaFromStore(store),
  text,
});
const first = mounted.placed[0];
if (first !== undefined) mountOverlay(mounted.root, first, { unit: doc.slideSize.cx / 960 });

// Zoom: unmount and mount again at the new width. The line breaks do not move.
mounted.unmount();
//#endregion
