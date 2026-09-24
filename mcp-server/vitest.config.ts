import { defineConfig } from 'vitest/config';

/**
 * Test layers — each is a separate CI step, so a failure names its layer.
 *
 *   unit        src/**\/*.test.ts          pure logic + in-process document API
 *               tests/unit/**
 *   module      tests/module/**             document API judged on saved/reopened files
 *   regression  tests/regression/**         one file per reported incident, with the report quoted
 *   e2e         tests/e2e/**                real MCP server over stdio (HWPX_MCP_SERVER picks the build)
 *
 * The legacy src/*.test.ts suite mixes unit and module-level checks; it runs
 * under "unit" until each file is moved. New tests go in tests/<layer>/.
 */
export default defineConfig({
  test: {
    include: ['src/**/*.test.ts', 'tests/**/*.test.ts'],
    testTimeout: 60_000,
    hookTimeout: 180_000,
  },
});
