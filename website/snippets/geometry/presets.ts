import { dragHandle, getPreset, resolveGeometry, resolveHandles } from '@pptx-studio/geometry';

//#region example
const roundRect = getPreset('roundRect');
if (roundRect === undefined) throw new Error('unknown preset');
const size = { w: 200, h: 100 };

resolveGeometry(roundRect, size).paths[0]?.d; // SVG path data, e.g. 'M16.667 0L183.333 0A16.667 16.667 …'
resolveHandles(roundRect, size)[0]?.axes[0]?.value; // the adjust value as the shape stands

const handle = roundRect.ahLst[0];
if (handle !== undefined) {
  dragHandle(roundRect, size, handle, { x: 25, y: 0 }); // { adj: … } - what a:avLst should hold
}
//#endregion
