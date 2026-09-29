/**
 * Which style PowerPoint draws a table with. C8 measured one rule (ADR 0063): the built-in its
 * GUID names, in any case, and otherwise none; `ppt/tableStyles.xml` never changes the drawing.
 */

import { firstChild, parseXmlString } from '@pptx-studio/xml';

import { BUILTIN_TABLE_STYLES } from '../builtin/table-styles.js';
import { ModelError } from '../errors.js';
import { parseTableStyle } from '../parse/table.js';
import type { Table, TableStyle } from '../table.js';

const NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
/** The part name a built-in's errors carry: it comes from no package. */
const BUILTIN_PART = 'builtin table styles';
const BY_ID = new Map(BUILTIN_TABLE_STYLES.map((style) => [style.id, style]));

/** One of PowerPoint's 74 built-ins by GUID, in any case, parsed afresh; `undefined` for any other. */
export function builtinTableStyle(id: string): TableStyle | undefined {
  const style = BY_ID.get(id.toUpperCase());
  if (style === undefined) return undefined;
  const list = parseXmlString(`<a:tblStyleLst xmlns:a="${NS_A}">${style.xml}</a:tblStyleLst>`);
  const element = firstChild(list.root, 'a:tblStyle');
  if (element === undefined) {
    throw new ModelError('MODEL_TABLE_STYLE', `${style.name} is no a:tblStyle`, BUILTIN_PART);
  }
  return parseTableStyle(element, BUILTIN_PART);
}

/**
 * The built-in style a table is drawn with, or `null` for PowerPoint's default: a 1-pt black
 * grid and no fill, whatever the theme and whatever the package defines.
 */
export function tableStyleOf(table: Table): TableStyle | null {
  const ref = table.props?.style;
  return ref === undefined ? null : (builtinTableStyle(ref.id) ?? null);
}
