import { createCanvasMeasurer, formatAutonumber, lineAdvance, wrapText } from '@pptx-studio/text';

//#region example
const measure = createCanvasMeasurer(); // needs OffscreenCanvas: a tab or a Worker
const text = 'Measured against PowerPoint, not derived from the standard.';
const font = { family: 'Arial', sz: 1800 }; // sz in hundredths of a point

const lines = wrapText({
  text,
  widthPt: 200,
  hyphenWidthPt: measure.measure('-', font).width,
  measure: (start, end) => measure.measure(text.slice(start, end), font).width,
});
lines.length; // how many lines the package broke it into

lineAdvance(1800, { kind: 'percent', value: 150000 }); // 32.4 - the line box, not the font size
formatAutonumber('romanLcParenBoth', 4); // '(iv)'
//#endregion
