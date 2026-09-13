/**
 * JSON for a fixture: nested where a record does not fit on a line, one line per record where it does.
 *
 * Prettier keeps an object on one line when the source has no newline after its brace and the
 * line fits its width, so a record written the way prettier prints it stays that way and a
 * fixture of many records stays under the corpus cap.
 */

const PRINT_WIDTH = 100;

export function fixtureJson(value: unknown): string {
  return `${render(value, 0)}\n`;
}

/** The value as prettier prints it on one line: a space after every comma and colon. */
function oneLine(value: unknown): string {
  if (typeof value !== 'object' || value === null) return JSON.stringify(value ?? null);
  if (Array.isArray(value)) {
    return value.length === 0 ? '[]' : `[${value.map(oneLine).join(', ')}]`;
  }
  const entries = Object.entries(value as Record<string, unknown>).filter(
    ([, item]) => item !== undefined,
  );
  if (entries.length === 0) return '{}';
  const fields = entries.map(([key, item]) => `${JSON.stringify(key)}: ${oneLine(item)}`);
  return `{ ${fields.join(', ')} }`;
}

/** `prefix` is what precedes the value on its line: an object entry's key, colon and space. */
function render(value: unknown, depth: number, prefix = 0): string {
  if (typeof value !== 'object' || value === null) return JSON.stringify(value ?? null);
  const line = oneLine(value);
  // The line, its indent, its key and a trailing comma all count against the width.
  if (2 * depth + prefix + line.length + 1 <= PRINT_WIDTH) return line;
  const pad = '  '.repeat(depth + 1);
  const close = '  '.repeat(depth);
  if (Array.isArray(value)) {
    return `[\n${value.map((item) => pad + render(item, depth + 1)).join(',\n')}\n${close}]`;
  }
  const entries = Object.entries(value as Record<string, unknown>).filter(
    ([, item]) => item !== undefined,
  );
  return `{\n${entries
    .map(([key, item]) => {
      const head = `${JSON.stringify(key)}: `;
      return `${pad}${head}${render(item, depth + 1, head.length)}`;
    })
    .join(',\n')}\n${close}}`;
}
