/**
 * A tiny SVG node type, and the string emitter over it.
 *
 * The plan calls for two renderers - a string one for thumbnails, PNG and PDF
 * export, and a live DOM one for the editor - "over the same layout engine", so
 * that they cannot disagree about what is on the screen. A shared *layout* is
 * not quite enough for that: emitting a `<linearGradient>` twice, once as text
 * and once as `createElementNS`, is two chances to get the same gradient wrong
 * in two different ways.
 *
 * So the shared surface goes one level further down. Everything is built as
 * this node tree, `serializeSvg` turns it into a string, and `render-dom` turns
 * the identical tree into elements. Neither knows anything the other does not.
 *
 * It is deliberately not a DOM: no namespaces, no parent links, no mutation.
 * The only thing it has to be is exactly serialisable and exactly constructible.
 */

export type AttributeValue = string | number | null | undefined;

export interface SvgElement {
  readonly kind: 'element';
  readonly tag: string;
  readonly attrs: Readonly<Record<string, AttributeValue>>;
  readonly children: readonly SvgNode[];
}

export interface SvgText {
  readonly kind: 'text';
  readonly text: string;
}

export type SvgNode = SvgElement | SvgText;

export function element(
  tag: string,
  attrs: Readonly<Record<string, AttributeValue>> = {},
  children: readonly SvgNode[] = [],
): SvgElement {
  return { kind: 'element', tag, attrs, children };
}

export function text(value: string): SvgText {
  return { kind: 'text', text: value };
}

/**
 * A number, short enough to read and long enough to be right.
 *
 * User units are EMU throughout - 12700 to the point - so three decimals is
 * about a nanometre and nothing anywhere needs more. Trailing zeros go, and so
 * does the negative zero that `toFixed` produces for a value just below the
 * axis, because `-0` in a `d` string is noise in every diff of a fixture.
 */
export function num(value: number, precision = 3): string {
  if (!Number.isFinite(value)) return '0';
  const factor = 10 ** precision;
  const rounded = Math.round(value * factor) / factor;
  return Object.is(rounded, -0) ? '0' : String(rounded);
}

const ESCAPES: Readonly<Record<string, string>> = {
  '&': '&amp;',
  '<': '&lt;',
  '>': '&gt;',
  '"': '&quot;',
  "'": '&apos;',
};

/**
 * Escape everything, including `>` and both quotes.
 *
 * More than XML strictly requires in text content, and that is on purpose: the
 * output of this package is handed to `DOMParser`, written into an `<img>` data
 * URI, and pasted into HTML, and the shortest rule that is safe in all three is
 * to escape the five.
 */
export function escapeXml(value: string): string {
  return value.replace(/[&<>"']/g, (char) => ESCAPES[char] ?? char);
}

function serializeInto(node: SvgNode, out: string[]): void {
  if (node.kind === 'text') {
    out.push(escapeXml(node.text));
    return;
  }
  out.push('<', node.tag);
  for (const [name, value] of Object.entries(node.attrs)) {
    if (value === null || value === undefined) continue;
    out.push(' ', name, '="', escapeXml(typeof value === 'number' ? num(value) : value), '"');
  }
  if (node.children.length === 0) {
    out.push('/>');
    return;
  }
  out.push('>');
  for (const child of node.children) serializeInto(child, out);
  out.push('</', node.tag, '>');
}

export function serializeSvg(node: SvgNode): string {
  const out: string[] = [];
  serializeInto(node, out);
  return out.join('');
}
