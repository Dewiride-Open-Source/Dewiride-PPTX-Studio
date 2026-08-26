import type { Rule } from 'eslint';

/**
 * Ban top-level `await` (and `for await` / `await using`) in library source.
 *
 * Why this is a hard rule and not a style preference: a module containing
 * top-level await is an *asynchronous* module. Node's ESM-from-CJS bridge
 * (`require(esm)`) refuses those with `ERR_REQUIRE_ASYNC_MODULE`, so a single
 * top-level `await` anywhere in a package's import graph makes that package
 * un-requireable from CommonJS - for every consumer, forever, with an error
 * message that points at Node rather than at us.
 *
 * There is no autofix: hoisting the await into a lazily-invoked async function
 * is a design decision, not a mechanical edit.
 */

/** Node types that introduce a new `await` scope. */
const FUNCTION_TYPES = new Set([
  'FunctionDeclaration',
  'FunctionExpression',
  'ArrowFunctionExpression',
]);

interface Linked {
  readonly type: string;
  readonly parent?: Linked | null;
}

/** True when `node` has no enclosing function. */
function isTopLevel(node: Linked): boolean {
  for (let parent = node.parent; parent; parent = parent.parent) {
    if (FUNCTION_TYPES.has(parent.type)) return false;
  }
  return true;
}

const rule: Rule.RuleModule = {
  meta: {
    type: 'problem',
    docs: {
      description:
        'Disallow top-level await, which makes a module asynchronous and un-requireable from CommonJS',
    },
    schema: [],
    messages: {
      topLevelAwait:
        'Top-level {{ construct }} makes this an async module; require() of any package importing it throws ERR_REQUIRE_ASYNC_MODULE. Move it inside an async function and call it lazily.',
    },
  },
  create(context) {
    const report = (node: Rule.Node, construct: string): void => {
      context.report({ node, messageId: 'topLevelAwait', data: { construct } });
    };
    return {
      AwaitExpression(node) {
        if (isTopLevel(node)) report(node, 'await');
      },
      ForOfStatement(node) {
        if (node.await && isTopLevel(node)) report(node, 'for await');
      },
      VariableDeclaration(node) {
        // `await using x = ...` is the same async-module problem in different
        // syntax. ESTree types predate it, hence the widened comparison.
        if ((node.kind as string) === 'await using' && isTopLevel(node)) {
          report(node, 'await using');
        }
      },
    };
  },
};

export default rule;
