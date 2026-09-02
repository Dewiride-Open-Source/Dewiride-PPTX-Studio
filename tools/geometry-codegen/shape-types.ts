import type { PresetBucketName } from '../../packages/geometry/src/types.ts';

/**
 * `ST_ShapeType`, written down here rather than read out of the input.
 *
 * ## Why this list exists at all
 *
 * The sub-phase's stated verification is "all 187 `ST_ShapeType` names
 * present". Checked against the file the names were taken from, that sentence
 * is a tautology: the transcoder would report 187 whatever the file contained,
 * including a file that had lost a shape between POI releases.
 *
 * So the expected names are written out independently, from the enumeration in
 * ECMA-376 Part 1 (the `ST_ShapeType` simple type in `dml-main.xsd`), and the
 * generator compares the two sets. Agreement is then a fact about two sources
 * rather than about one.
 *
 * **A disagreement is a finding, not a chore.** The temptation, when this list
 * and the file differ, is to paste the file's names over this one - which
 * converts a real check back into the tautology it was written to avoid. If
 * they ever diverge, the right response is to work out which source is wrong
 * and say so in the ADR.
 *
 * Sorted, because the generator diffs the two sets and a sorted list makes the
 * diff readable in review.
 */
export const ST_SHAPE_TYPE: readonly string[] = [
  'accentBorderCallout1',
  'accentBorderCallout2',
  'accentBorderCallout3',
  'accentCallout1',
  'accentCallout2',
  'accentCallout3',
  'actionButtonBackPrevious',
  'actionButtonBeginning',
  'actionButtonBlank',
  'actionButtonDocument',
  'actionButtonEnd',
  'actionButtonForwardNext',
  'actionButtonHelp',
  'actionButtonHome',
  'actionButtonInformation',
  'actionButtonMovie',
  'actionButtonReturn',
  'actionButtonSound',
  'arc',
  'bentArrow',
  'bentConnector2',
  'bentConnector3',
  'bentConnector4',
  'bentConnector5',
  'bentUpArrow',
  'bevel',
  'blockArc',
  'borderCallout1',
  'borderCallout2',
  'borderCallout3',
  'bracePair',
  'bracketPair',
  'callout1',
  'callout2',
  'callout3',
  'can',
  'chartPlus',
  'chartStar',
  'chartX',
  'chevron',
  'chord',
  'circularArrow',
  'cloud',
  'cloudCallout',
  'corner',
  'cornerTabs',
  'cube',
  'curvedConnector2',
  'curvedConnector3',
  'curvedConnector4',
  'curvedConnector5',
  'curvedDownArrow',
  'curvedLeftArrow',
  'curvedRightArrow',
  'curvedUpArrow',
  'decagon',
  'diagStripe',
  'diamond',
  'dodecagon',
  'donut',
  'doubleWave',
  'downArrow',
  'downArrowCallout',
  'ellipse',
  'ellipseRibbon',
  'ellipseRibbon2',
  'flowChartAlternateProcess',
  'flowChartCollate',
  'flowChartConnector',
  'flowChartDecision',
  'flowChartDelay',
  'flowChartDisplay',
  'flowChartDocument',
  'flowChartExtract',
  'flowChartInputOutput',
  'flowChartInternalStorage',
  'flowChartMagneticDisk',
  'flowChartMagneticDrum',
  'flowChartMagneticTape',
  'flowChartManualInput',
  'flowChartManualOperation',
  'flowChartMerge',
  'flowChartMultidocument',
  'flowChartOfflineStorage',
  'flowChartOffpageConnector',
  'flowChartOnlineStorage',
  'flowChartOr',
  'flowChartPredefinedProcess',
  'flowChartPreparation',
  'flowChartProcess',
  'flowChartPunchedCard',
  'flowChartPunchedTape',
  'flowChartSort',
  'flowChartSummingJunction',
  'flowChartTerminator',
  'foldedCorner',
  'frame',
  'funnel',
  'gear6',
  'gear9',
  'halfFrame',
  'heart',
  'heptagon',
  'hexagon',
  'homePlate',
  'horizontalScroll',
  'irregularSeal1',
  'irregularSeal2',
  'leftArrow',
  'leftArrowCallout',
  'leftBrace',
  'leftBracket',
  'leftCircularArrow',
  'leftRightArrow',
  'leftRightArrowCallout',
  'leftRightCircularArrow',
  'leftRightRibbon',
  'leftRightUpArrow',
  'leftUpArrow',
  'lightningBolt',
  'line',
  'lineInv',
  'mathDivide',
  'mathEqual',
  'mathMinus',
  'mathMultiply',
  'mathNotEqual',
  'mathPlus',
  'moon',
  'noSmoking',
  'nonIsoscelesTrapezoid',
  'notchedRightArrow',
  'octagon',
  'parallelogram',
  'pentagon',
  'pie',
  'pieWedge',
  'plaque',
  'plaqueTabs',
  'plus',
  'quadArrow',
  'quadArrowCallout',
  'rect',
  'ribbon',
  'ribbon2',
  'rightArrow',
  'rightArrowCallout',
  'rightBrace',
  'rightBracket',
  'round1Rect',
  'round2DiagRect',
  'round2SameRect',
  'roundRect',
  'rtTriangle',
  'smileyFace',
  'snip1Rect',
  'snip2DiagRect',
  'snip2SameRect',
  'snipRoundRect',
  'squareTabs',
  'star10',
  'star12',
  'star16',
  'star24',
  'star32',
  'star4',
  'star5',
  'star6',
  'star7',
  'star8',
  'straightConnector1',
  'stripedRightArrow',
  'sun',
  'swooshArrow',
  'teardrop',
  'trapezoid',
  'triangle',
  'upArrow',
  'upArrowCallout',
  'upDownArrow',
  'upDownArrowCallout',
  'uturnArrow',
  'verticalScroll',
  'wave',
  'wedgeEllipseCallout',
  'wedgeRectCallout',
  'wedgeRoundRectCallout',
];

/**
 * PowerPoint's own Basic Shapes gallery, transcribed by hand.
 *
 * The other five buckets have a mechanical signature in the name - a
 * `flowChart` prefix, an `Arrow` inside it, a `star` and a number. "Basic" has
 * none, so it is a list, and being a list it is the one bucket rule that is a
 * judgement rather than a derivation.
 */
const BASIC = new Set([
  'arc',
  'bevel',
  'blockArc',
  'bracePair',
  'bracketPair',
  'can',
  'chord',
  'cloud',
  'corner',
  'cube',
  'decagon',
  'diagStripe',
  'diamond',
  'dodecagon',
  'donut',
  'ellipse',
  'foldedCorner',
  'frame',
  'halfFrame',
  'heart',
  'heptagon',
  'hexagon',
  'leftBrace',
  'leftBracket',
  'lightningBolt',
  'moon',
  'noSmoking',
  'octagon',
  'parallelogram',
  'pentagon',
  'pie',
  'plaque',
  'plus',
  'rect',
  'rightBrace',
  'rightBracket',
  'roundRect',
  'rtTriangle',
  'smileyFace',
  'sun',
  'teardrop',
  'trapezoid',
  'triangle',
]);

/**
 * Which bucket a preset is generated into. First rule that matches wins, and
 * the order is what puts `rightArrowCallout` with the arrows: PowerPoint files
 * the six `*ArrowCallout` shapes under Block Arrows rather than under Callouts,
 * and the name alone would say the opposite.
 *
 * These are **packaging** boundaries. Nothing about a shape's geometry depends
 * on which file it lands in, `presets/index.ts` resolves any name from any
 * bucket, and a future rebalance is not a change to what a preset draws - which
 * is why the README asks size-sensitive consumers importing one bucket directly
 * to pin a version.
 */
const RULES: readonly { readonly bucket: PresetBucketName; readonly test: RegExp }[] = [
  { bucket: 'flowchart', test: /^flowChart/ },
  { bucket: 'arrows', test: /Arrow|^chevron$|^homePlate$/ },
  { bucket: 'stars', test: /^star\d+$|^irregularSeal[12]$|ibbon|Scroll$|^wave$|^doubleWave$/ },
  { bucket: 'callouts', test: /allout/ },
];

export function bucketOf(name: string): PresetBucketName {
  for (const rule of RULES) {
    if (rule.test.test(name)) return rule.bucket;
  }
  return BASIC.has(name) ? 'basic' : 'misc';
}

export const BUCKET_NAMES: readonly PresetBucketName[] = [
  'basic',
  'arrows',
  'callouts',
  'flowchart',
  'misc',
  'stars',
];
