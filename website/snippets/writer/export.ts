import { exportPackage, openPackage } from '@pptx-studio/writer';
import { applyEdit, descendantElements, parseXml, serializeXml } from '@pptx-studio/xml';

declare const bytes: Uint8Array;

//#region example
const pkg = openPackage(bytes); // { store, baseline, baselineBytes }
const part = '/ppt/slides/slide1.xml';
const doc = parseXml(pkg.store.read(part));
const title = [...descendantElements(doc.root)].find((el) => el.qname === 'a:t');
const node = title?.children[0];
if (node !== undefined && node.type === 'text') {
  applyEdit({ kind: 'setValue', node, value: 'Hello from the browser' });
}
pkg.store.replacePart(part, serializeXml(doc));

const result = exportPackage(pkg); // throws rather than return a file PowerPoint would repair
result.rewritten; // ['/ppt/slides/slide1.xml']
result.streamed; // every other part, copied still compressed
const blob = new Blob([result.bytes.slice()], {
  type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
});
//#endregion

export { blob };
