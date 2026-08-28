import { placeholderXml, TITLE_BOX, type ProbeLayout } from '../package.ts';
import { grid, shape } from '../shapes.ts';
import { textLine, txBody } from '../text.ts';
import type { ProbeDeck } from '../types.ts';

/**
 * Every `p:transition` variant ECMA-376 defines, on the three sheets that may
 * carry one.
 *
 * ## Twenty-one, not twenty-two
 *
 * `sub-phase 1.1`'s roster said twenty-two variants. `CT_SlideTransition`'s
 * choice group holds twenty-one child elements and the table says so, both in
 * `packages/xml/src/schema-order.gen.ts` - generated from the Transitional XSDs
 * - and in ISO/IEC 29500-1's own list: blinds, checker, circle, comb, cover,
 * cut, diamond, dissolve, fade, newsflash, plus, pull, push, random, randomBar,
 * split, strips, wedge, wheel, wipe, zoom. The other two children of the
 * element are `p:sndAc` and `p:extLst`, which are not effects. So the roster
 * was one over, and the correction lives here as well as there.
 *
 * ## Why a three-slide deck can hold twenty-one of them
 *
 * A transition is one per sheet, so twenty-one effects would need twenty-one
 * slides - except that `p:transition` is legal on `p:sld`, `p:sldLayout` **and**
 * `p:sldMaster`, and a slide with none of its own inherits its layout's. So the
 * master carries `cut`, eighteen otherwise-empty layouts carry one each, and
 * the three slides carry the four cases that are about something other than
 * which effect it is.
 *
 * That is also the probe: a renderer that reads transitions off slides alone
 * finds four of the twenty-three here.
 *
 * ## What PowerPoint actually writes, on slide 1
 *
 * Measured on 2026-08-27: setting a slide's transition to Fade in
 * PowerPoint 16.0.20326 produces neither `<p:transition><p:fade/></p:transition>`
 * nor an extension, but **both**, wrapped:
 *
 * ```xml
 * <mc:AlternateContent xmlns:mc="…">
 *   <mc:Choice xmlns:p14="…/powerpoint/2010/main" Requires="p14">
 *     <p:transition spd="slow" p14:dur="2000"><p:fade/></p:transition>
 *   </mc:Choice>
 *   <mc:Fallback>
 *     <p:transition spd="slow"><p:fade/></p:transition>
 *   </mc:Fallback>
 * </mc:AlternateContent>
 * ```
 *
 * The 2010 attribute is `p14:dur`, a duration in milliseconds that supersedes
 * the three-valued `@spd`. Both branches are real markup, so a census reading
 * pre-MCE tokens counts **two** transitions for one effect - the same policy
 * `a26-ole` will need for `oleObject`, and `features` says so rather than
 * pretending the wrapper is free.
 *
 * ## The attributes, per element
 *
 * Four different direction enumerations, which is the part that gets written
 * from memory and gets written wrong:
 *
 * | elements                        | attribute | values                       |
 * | ------------------------------- | --------- | ---------------------------- |
 * | blinds, checker, comb, randomBar | `dir`     | `horz`, `vert`               |
 * | cover, pull                     | `dir`     | `l u r d ld lu rd ru`        |
 * | push, wipe                      | `dir`     | `l u r d`                    |
 * | strips                          | `dir`     | `ld lu rd ru` only           |
 * | split, zoom                     | `dir`     | `in`, `out`                  |
 * | split                           | `orient`  | `horz`, `vert`               |
 * | cut, fade                       | `thruBlk` | boolean                      |
 * | wheel                           | `spokes`  | unsignedInt                  |
 *
 * and circle, diamond, dissolve, newsflash, plus, random and wedge take none.
 */

/** The eighteen that ride on layouts, each with the attributes it accepts. */
const LAYOUT_TRANSITIONS: readonly { readonly name: string; readonly xml: string }[] = [
  { name: 'blinds horz', xml: '<p:blinds dir="horz"/>' },
  { name: 'checker vert', xml: '<p:checker dir="vert"/>' },
  { name: 'circle', xml: '<p:circle/>' },
  { name: 'comb vert', xml: '<p:comb dir="vert"/>' },
  { name: 'cover ru', xml: '<p:cover dir="ru"/>' },
  { name: 'diamond', xml: '<p:diamond/>' },
  { name: 'dissolve', xml: '<p:dissolve/>' },
  { name: 'newsflash', xml: '<p:newsflash/>' },
  { name: 'plus', xml: '<p:plus/>' },
  { name: 'pull ld', xml: '<p:pull dir="ld"/>' },
  { name: 'push u', xml: '<p:push dir="u"/>' },
  { name: 'random', xml: '<p:random/>' },
  { name: 'split horz out', xml: '<p:split orient="horz" dir="out"/>' },
  { name: 'strips lu', xml: '<p:strips dir="lu"/>' },
  { name: 'wedge', xml: '<p:wedge/>' },
  { name: 'wheel 8 spokes', xml: '<p:wheel spokes="8"/>' },
  { name: 'wipe d', xml: '<p:wipe dir="d"/>' },
  { name: 'zoom in', xml: '<p:zoom dir="in"/>' },
];

const LAYOUTS: readonly ProbeLayout[] = [
  {
    type: 'titleOnly',
    name: 'Title Only',
    hasTitle: true,
    shapes: placeholderXml({ id: 2, name: 'Title 1', type: 'title', ...TITLE_BOX }),
  },
  ...LAYOUT_TRANSITIONS.map((entry, index) => ({
    type: 'blank',
    // The layout name is the effect, so the picker in PowerPoint reads as a
    // list of what this deck covers.
    name: 'Transition — ' + entry.name,
    transition:
      `<p:transition spd="${index % 3 === 0 ? 'slow' : index % 3 === 1 ? 'med' : 'fast'}">` +
      entry.xml +
      '</p:transition>',
  })),
];

const NS_MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
const NS_P14 = 'http://schemas.microsoft.com/office/powerpoint/2010/main';

/** Slide 1: the wrapped pair, exactly as measured. */
const MEASURED_FADE =
  `<mc:AlternateContent xmlns:mc="${NS_MC}">` +
  `<mc:Choice xmlns:p14="${NS_P14}" Requires="p14">` +
  '<p:transition spd="slow" p14:dur="2000"><p:fade/></p:transition>' +
  '</mc:Choice>' +
  '<mc:Fallback>' +
  '<p:transition spd="slow"><p:fade/></p:transition>' +
  '</mc:Fallback>' +
  '</mc:AlternateContent>';

/**
 * Slide 2: the timing attributes, and `p:sndAc`.
 *
 * `p:sndAc` is `CT_TransitionStartSoundAction`'s home and is the twenty-second
 * child of `p:transition` - a sound rather than an effect. `p:endSnd` is
 * `CT_Empty` and needs no media part, where `p:stSnd` would need a `p:snd` with
 * an `r:embed` pointing at a `.wav`. This deck has no audio, so it says the
 * cheap half and `a24-media` says the other.
 */
const TIMED_RANDOM_BAR =
  '<p:transition spd="slow" advClick="0" advTm="3000">' +
  '<p:randomBar dir="horz"/>' +
  '<p:sndAc><p:endSnd/></p:sndAc>' +
  '</p:transition>';

/**
 * Slide 3: every child optional, so this is a legal `p:transition` with no
 * effect in it at all. It still carries `@advTm`, which means the slide
 * auto-advances with no visible transition - a real authoring outcome, and the
 * case a reader that dereferences `firstElementChild` crashes on.
 */
const CHILDLESS = '<p:transition spd="fast" advTm="1500"/>';

function caption(id: number, lines: readonly string[]): string {
  return shape({
    id,
    name: 'Caption ' + String(id),
    ...grid(1, 3)(0),
    textBody: txBody({ paras: lines.map((line) => textLine(line, { sz: 1600 })).join('') }),
  });
}

export const a16Transitions: ProbeDeck = {
  id: 'a16-transitions',
  title: 'PPTX Studio corpus: a16 transitions',
  description:
    'All twenty-one p:transition effects ECMA-376 defines - not the twenty-two the plan claimed - ' +
    'spread across a slide master, eighteen slide layouts and three slides, because a transition ' +
    'is legal on all three sheets and a slide with none inherits its layout&apos;s. Slide 1 carries ' +
    'the mc:AlternateContent pair PowerPoint 16.0.20326 was measured writing for a plain fade, ' +
    'which is why the transition count is twenty-three for twenty-one effects. Preserved across ' +
    'an export and never played.',
  features: {
    shape: 9,
    placeholder: 6,
    presetGeom: 3,
    gradientFill: 2,
    transition: 23,
    alternateContent: 1,
  },
  build: () => ({
    title: 'PPTX Studio corpus: a16 transitions',
    masters: [
      {
        transition: '<p:transition spd="med"><p:cut thruBlk="1"/></p:transition>',
        layouts: LAYOUTS,
      },
    ],
    slides: [
      {
        title: 'a16 — the fade PowerPoint writes',
        layout: 0,
        body: caption(10, [
          'mc:Choice Requires="p14" carries p14:dur="2000"; mc:Fallback carries @spd alone.',
          'Both branches hold a p:transition, so the census counts two.',
        ]),
        tail: MEASURED_FADE,
      },
      {
        title: 'a16 — advance timing and a sound action',
        layout: 0,
        body: caption(10, [
          'advClick="0" advTm="3000": no click advances it, three seconds does.',
          'p:sndAc/p:endSnd stops the previous slide&apos;s sound and needs no media part.',
        ]),
        tail: TIMED_RANDOM_BAR,
      },
      {
        title: 'a16 — a transition with no effect in it',
        layout: 0,
        body: caption(10, [
          'Every child of p:transition is optional.',
          'This one auto-advances after 1.5s with nothing to draw.',
        ]),
        tail: CHILDLESS,
      },
    ],
  }),
};
