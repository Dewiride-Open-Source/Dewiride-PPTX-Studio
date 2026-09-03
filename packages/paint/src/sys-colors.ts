/**
 * `ST_SystemColorVal`, and what Windows 11's default light theme gives for each.
 *
 * ## `@lastClr` does not win, and that is the opposite of what we assumed
 *
 * The plan says "`a:sysClr` prefers `@lastClr`", and sub-phase 0.7-C recorded it
 * as confirmed. It is not. 2.6 wrote three swatches with a `@lastClr` no system
 * theme would ever produce - `<a:sysClr val="windowText" lastClr="FF00FF"/>` -
 * and PowerPoint painted all three the *machine's* colour and ignored the
 * attribute. 0.7-C only looked like a confirmation because the theme's
 * `windowText` had a `lastClr` of `000000`, which is also what this machine
 * says, so the two answers were the same number.
 *
 * `@lastClr` is what its name says: a cache of what the writer's machine
 * resolved, for a reader that has no machine to ask.
 *
 * ## So we use it anyway, deliberately
 *
 * A browser has no Windows system palette, and `CanvasText`/`Canvas` are the
 * viewer's theme rather than the author's. Preferring `@lastClr` reproduces what
 * the deck looked like where it was written, which is the closer answer to what
 * the user meant; the table below is the fallback for the case where there is no
 * `@lastClr` at all - which PowerPoint accepts, and which nothing PowerPoint
 * writes ever contains.
 *
 * The values are one machine's, on one theme, on one day: Windows 11 Enterprise
 * 10.0.26200 in its default light theme. `highlight` is `0078D7`, which is the
 * Windows 10/11 default accent and not a constant of the format. They are a
 * better guess than black and they are not ground truth about anything portable.
 */
export const SYSTEM_COLORS: Readonly<Record<string, string>> = {
  scrollBar: 'C8C8C8',
  background: '000000',
  activeCaption: '99B4D1',
  inactiveCaption: 'BFCDDB',
  menu: 'F0F0F0',
  window: 'FFFFFF',
  windowFrame: '646464',
  menuText: '000000',
  windowText: '000000',
  captionText: '000000',
  activeBorder: 'B4B4B4',
  inactiveBorder: 'F4F7FC',
  appWorkspace: 'ABABAB',
  highlight: '0078D7',
  highlightText: 'FFFFFF',
  btnFace: 'F0F0F0',
  btnShadow: 'A0A0A0',
  grayText: '6D6D6D',
  btnText: '000000',
  inactiveCaptionText: '000000',
  btnHighlight: 'FFFFFF',
  '3dDkShadow': '696969',
  '3dLight': 'E3E3E3',
  infoText: '000000',
  infoBk: 'FFFFE1',
  hotLight: '0066CC',
  gradientActiveCaption: 'B9D1EA',
  gradientInactiveCaption: 'D7E4F2',
  menuHighlight: '3399FF',
  menuBar: 'F0F0F0',
};

/** Whether a token is one of the thirty. */
export function isSystemColorName(name: string): boolean {
  return Object.prototype.hasOwnProperty.call(SYSTEM_COLORS, name);
}
