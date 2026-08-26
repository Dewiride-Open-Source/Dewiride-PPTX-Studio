import { RuleTester } from 'eslint';
import tsParser from '@typescript-eslint/parser';
import { describe, it } from 'vitest';
import rule from './no-top-level-await.ts';

// ESLint's RuleTester drives whatever test framework it is handed. Without this
// it falls back to an internal runner and the results never reach Vitest.
RuleTester.describe = describe;
RuleTester.it = it;
RuleTester.itOnly = it.only;

const ruleTester = new RuleTester({
  languageOptions: {
    parser: tsParser,
    ecmaVersion: 2024,
    sourceType: 'module',
  },
});

ruleTester.run('no-top-level-await', rule, {
  valid: [
    // The whole point: the same await, one scope deeper, is fine.
    'async function load() { await fetch("/x"); }',
    'const load = async () => { await fetch("/x"); };',
    'export async function* stream() { for await (const c of src) yield c; }',
    'class A { async m() { await x; } }',
    // Not an await at all.
    'const x = 1; export { x };',
    'function f() { return { await: 1 }; }',
    // Nested deeply inside a function is still inside a function.
    'async function a() { if (true) { for (const x of []) { await x; } } }',
  ],
  invalid: [
    {
      code: 'const config = await loadConfig();',
      errors: [{ messageId: 'topLevelAwait', data: { construct: 'await' } }],
    },
    {
      code: 'await import("./thing.js");',
      errors: [{ messageId: 'topLevelAwait' }],
    },
    {
      code: 'for await (const chunk of stream) { use(chunk); }',
      errors: [{ messageId: 'topLevelAwait', data: { construct: 'for await' } }],
    },
    {
      // Inside a block, but still no enclosing function.
      code: 'if (flag) { await ready; }',
      errors: [{ messageId: 'topLevelAwait' }],
    },
    {
      // A function declared earlier in the module must not make a later
      // top-level await look enclosed.
      code: 'const wrapped = () => 1; await wrapped;',
      errors: [{ messageId: 'topLevelAwait' }],
    },
    {
      // Explicit resource management at module scope is the same async-module
      // problem wearing a different syntax.
      code: 'await using handle = open();',
      errors: [{ messageId: 'topLevelAwait', data: { construct: 'await using' } }],
    },
  ],
});
