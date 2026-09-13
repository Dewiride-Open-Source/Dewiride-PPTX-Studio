/**
 * The API reference, read from what was installed.
 *
 * ```
 * node prerender/reference.mjs
 * ```
 *
 * For each package, `dist/index.d.ts` is parsed with the TypeScript compiler:
 * every name in its final `export { ... }` becomes one entry, grouped by the
 * `//#region src/<file>.d.ts` it sits in, with its declaration text and the
 * first sentence of its JSDoc. Written to `src/reference/generated/<name>.json`,
 * so the reference can never describe a version other than the one the site runs.
 */

import { mkdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { createRequire } from 'node:module';
import { dirname, join } from 'node:path';

import ts from 'typescript';

const PACKAGES = [
  'opc',
  'xml',
  'census',
  'geometry',
  'paint',
  'text',
  'model',
  'render-svg',
  'render-dom',
  'validate',
  'writer',
  'cli',
];
const OUT = join(import.meta.dirname, '..', 'src', 'reference', 'generated');
const require = createRequire(import.meta.url);

/** @param {ts.Node} node @param {ts.SourceFile} file */
function docOf(node, file) {
  const ranges = ts.getLeadingCommentRanges(file.text, node.getFullStart()) ?? [];
  const last = ranges.at(-1);
  if (last === undefined) return '';
  const raw = file.text.slice(last.pos, last.end);
  if (!raw.startsWith('/**')) return '';
  const text = raw
    .replace(/^\/\*\*/, '')
    .replace(/\*\/$/, '')
    .split('\n')
    .map((line) => line.replace(/^\s*\* ?/, ''))
    .join('\n')
    .trim();
  const body = text.split(/\n\s*\n/)[0] ?? '';
  return body.replace(/\s+/g, ' ').trim();
}

/** @param {ts.Node} node */
function kindOf(node) {
  if (ts.isFunctionDeclaration(node)) return 'function';
  if (ts.isClassDeclaration(node)) return 'class';
  if (ts.isInterfaceDeclaration(node)) return 'interface';
  if (ts.isTypeAliasDeclaration(node)) return 'type';
  if (ts.isEnumDeclaration(node)) return 'enum';
  return 'const';
}

/** @param {ts.Node} node @param {ts.SourceFile} file */
function textOf(node, file) {
  return node
    .getText(file)
    .replace(/^declare\s+/, '')
    .replace(/^export\s+/, '');
}

/** @param {ts.Node} node */
function namesOf(node) {
  if (ts.isVariableStatement(node)) {
    return node.declarationList.declarations
      .map((declaration) => (ts.isIdentifier(declaration.name) ? declaration.name.text : null))
      .filter((name) => name !== null);
  }
  if (
    ts.isFunctionDeclaration(node) ||
    ts.isClassDeclaration(node) ||
    ts.isInterfaceDeclaration(node) ||
    ts.isTypeAliasDeclaration(node) ||
    ts.isEnumDeclaration(node)
  ) {
    return node.name === undefined ? [] : [node.name.text];
  }
  return [];
}

/** Every string literal in a type: the members of an `X_ERROR_CODES` tuple or an `XErrorCode` union. */
function literalsIn(node) {
  const found = [];
  const walk = (at) => {
    if (ts.isStringLiteral(at) || ts.isNoSubstitutionTemplateLiteral(at)) found.push(at.text);
    ts.forEachChild(at, walk);
  };
  walk(node);
  return found;
}

/**
 * Every declaration in one `.d.ts`, by name, with the region it sits in; plus its
 * `export { local as alias }` map and its `import { alias as name } from './x.js'` list.
 * @param {string} path
 */
function scan(path) {
  const source = readFileSync(path, 'utf8');
  const file = ts.createSourceFile(path, source, ts.ScriptTarget.Latest, true, ts.ScriptKind.TS);

  /** @type {{ start: number; end: number; file: string }[]} */
  const regions = [];
  const lines = source.split('\n');
  let offset = 0;
  let open = null;
  for (const line of lines) {
    const begin = /^\/\/#region (.+)$/.exec(line);
    if (begin?.[1] !== undefined) open = { start: offset, file: begin[1].trim() };
    if (line.startsWith('//#endregion') && open !== null) {
      regions.push({ start: open.start, end: offset, file: open.file });
      open = null;
    }
    offset += line.length + 1;
  }
  const regionOf = (position) =>
    regions.find((one) => position >= one.start && position < one.end)?.file ?? 'index.d.ts';

  /** @type {Map<string, { name: string; kind: string; signature: string; doc: string; region: string; position: number }>} */
  const declared = new Map();
  /** @type {{ name: string; isType: boolean; local: string }[]} */
  const exported = [];
  /** @type {{ name: string; alias: string; from: string }[]} */
  const imported = [];
  const codes = new Map();

  for (const statement of file.statements) {
    if (
      ts.isImportDeclaration(statement) &&
      statement.importClause?.namedBindings &&
      ts.isNamedImports(statement.importClause.namedBindings)
    ) {
      const from = statement.moduleSpecifier.text;
      for (const element of statement.importClause.namedBindings.elements) {
        const name = element.name.text;
        if (from.startsWith('@pptx-studio/')) {
          // Re-exported from a sibling package: documented there, named here.
          declared.set(name, {
            name,
            kind: 're-export',
            signature: `re-exported from ${from}`,
            doc: '',
            region: 'index.d.ts',
            position: statement.getStart(file),
          });
        } else if (from.startsWith('.')) {
          imported.push({
            name,
            alias: element.propertyName?.text ?? name,
            from: join(dirname(path), from.replace(/.js$/, '.d.ts')),
          });
        }
      }
      continue;
    }
    if (
      ts.isExportDeclaration(statement) &&
      statement.exportClause &&
      ts.isNamedExports(statement.exportClause)
    ) {
      for (const element of statement.exportClause.elements) {
        exported.push({
          name: element.name.text,
          isType: element.isTypeOnly,
          local: element.propertyName?.text ?? element.name.text,
        });
      }
      continue;
    }
    for (const declaredName of namesOf(statement)) {
      const position = statement.getStart(file);
      declared.set(declaredName, {
        name: declaredName,
        kind: kindOf(statement),
        signature: textOf(statement, file),
        doc: docOf(statement, file),
        region: regionOf(position),
        position,
      });
      if (/_ERROR_CODES$/.test(declaredName) || /ErrorCode$/.test(declaredName)) {
        const literals = literalsIn(statement);
        if (literals.length > 0) codes.set(declaredName, literals);
      }
    }
  }

  return { declared, exported, imported, codes };
}

/** @param {string} name */
function extract(name) {
  const manifest = require(`@pptx-studio/${name}/package.json`);
  const entry = join(
    dirname(require.resolve(`@pptx-studio/${name}/package.json`)),
    'dist',
    'index.d.ts',
  );
  const { declared, exported, imported, codes } = scan(entry);

  // A name re-exported from a sibling chunk is declared there, under the chunk's own alias.
  const siblings = new Map();
  for (const one of imported) {
    if (declared.has(one.name)) continue;
    const sibling = siblings.get(one.from) ?? scan(one.from);
    siblings.set(one.from, sibling);
    const local = sibling.exported.find((each) => each.name === one.alias)?.local ?? one.alias;
    const found = sibling.declared.get(local);
    if (found !== undefined) declared.set(one.name, { ...found, name: one.name });
    for (const [codeName, list] of sibling.codes) codes.set(codeName, list);
  }

  const missing = exported.filter((one) => !declared.has(one.name)).map((one) => one.name);
  if (missing.length > 0) {
    throw new Error(
      `@pptx-studio/${name}: exported but not declared in index.d.ts: ${missing.join(', ')}`,
    );
  }

  const groups = new Map();
  for (const one of exported) {
    const found = declared.get(one.name);
    const list = groups.get(found.region) ?? [];
    list.push({
      name: one.name,
      kind: found.kind,
      isType: one.isType,
      signature: found.signature,
      doc: found.doc,
      position: found.position,
    });
    groups.set(found.region, list);
  }

  const errorClass = exported.find((one) => /Error$/.test(one.name) && !one.isType)?.name ?? null;
  const codeUnion = [...codes.entries()].sort(([a], [b]) => (a < b ? -1 : 1));
  const errorCodes = codeUnion.reduce(
    (longest, [, list]) => (list.length > longest.length ? list : longest),
    [],
  );

  return {
    name,
    version: manifest.version,
    description: manifest.description,
    exported: exported.length,
    groups: [...groups.entries()]
      .map(([region, symbols]) => ({
        file: region.replace(/\.d\.ts$/, '.ts'),
        symbols: symbols
          .sort((a, b) => a.position - b.position)
          .map(({ position: _position, ...rest }) => rest),
      }))
      .sort((a, b) => ((a.symbols[0]?.name ?? '') < (b.symbols[0]?.name ?? '') ? 0 : 0)),
    errors: errorClass === null ? null : { className: errorClass, codes: errorCodes },
  };
}

rmSync(OUT, { recursive: true, force: true });
mkdirSync(OUT, { recursive: true });
for (const name of PACKAGES) {
  const reference = extract(name);
  writeFileSync(join(OUT, `${name}.json`), `${JSON.stringify(reference, null, 2)}\n`);
  console.log(
    `${name}@${reference.version}: ${String(reference.exported)} export(s) in ${String(reference.groups.length)} file(s)` +
      (reference.errors === null
        ? ''
        : `, ${String(reference.errors.codes.length)} error code(s) on ${reference.errors.className}`),
  );
}
