#!/usr/bin/env node
import { main } from './main.js';

/**
 * The bin.
 *
 * Deliberately three lines. Everything testable lives in `main`, which takes
 * its streams as an argument and returns an exit code instead of calling
 * `process.exit` - so the whole command line is exercised in-process by the
 * test suite, with no child processes and no captured stdout.
 */
process.exitCode = main(process.argv.slice(2));
