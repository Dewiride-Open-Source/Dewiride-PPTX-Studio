import { PartStore, REL_TYPE } from '@pptx-studio/opc';

declare const bytes: Uint8Array; // the contents of a .pptx

//#region example
const store = PartStore.open(bytes); // hostile-input limits are on by default

const root = store.rootRelationships();
const main = root.firstOfType(REL_TYPE.officeDocument);
if (main === undefined) throw new Error('no officeDocument relationship');
const presentation = root.resolve(main); // '/ppt/presentation.xml'

for (const rel of store.relationships(presentation).byType(REL_TYPE.slide)) {
  const slide = store.relationships(presentation).resolve(rel);
  console.log(slide, store.contentTypeOf(slide), store.read(slide).byteLength);
}
//#endregion
