import {
  applyTransforms,
  resolveColor,
  toCss,
  type ClrMap,
  type ClrScheme,
  type Color,
} from '@pptx-studio/paint';

declare const scheme: ClrScheme; // the theme's, via schemeOf(sheet) from @pptx-studio/model
declare const map: ClrMap; // the sheet's colour map, via colorMapOf(sheet)

//#region example
// <a:schemeClr val="accent1"><a:lumMod val="60000"/><a:lumOff val="40000"/></a:schemeClr>
const color: Color = {
  space: 'scheme',
  name: 'accent1',
  transforms: [
    { op: 'lumMod', val: 60000 },
    { op: 'lumOff', val: 40000 },
  ],
};
toCss(resolveColor(color, { scheme, map })); // PowerPoint's own "Accent 1, Lighter 40%"

// Transforms apply in document order; reversed, they give a plausible wrong colour.
applyTransforms({ r: 68, g: 114, b: 196, a: 1 }, color.transforms);
//#endregion
