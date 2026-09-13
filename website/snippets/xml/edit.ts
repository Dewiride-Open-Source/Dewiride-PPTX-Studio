import type { PartStore } from '@pptx-studio/opc';
import { applyEdit, descendantElements, parseXml, serializeXml } from '@pptx-studio/xml';

declare const store: PartStore;

//#region example
const doc = parseXml(store.read('/ppt/slides/slide1.xml'));
const cNvPr = [...descendantElements(doc.root)].find((el) => el.qname === 'p:cNvPr');
if (cNvPr === undefined) throw new Error('no shape');

const undo = applyEdit({ kind: 'setAttribute', element: cNvPr, qname: 'name', value: 'Renamed' });
serializeXml(doc); // the original bytes, except the value of that one attribute
applyEdit(undo);
serializeXml(doc); // byte-identical to what was read
//#endregion
