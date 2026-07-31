// Tiny zero-dependency test runner. Each `*.test.mjs` file in this
// directory is loaded in alphabetical order; failing files cause the
// runner to exit non-zero. The runner does not swallow errors — any
// exception inside a test file bubbles up and aborts the run, which is
// the behaviour we want for CI.

import { spawnSync } from "node:child_process";
import { readdirSync } from "node:fs";
import { resolve as resolvePath, join, dirname } from "node:path";
import { fileURLToPath } from "node:url";

const here = dirname(fileURLToPath(import.meta.url));
const only = process.argv.find((a) => a.startsWith("--only="))?.slice("--only=".length);

const files = readdirSync(here)
  .filter((f) => f.endsWith(".test.mjs") && f !== "run.mjs")
  .filter((f) => !only || f === only || f.startsWith(only))
  .sort();

if (files.length === 0) {
  process.stderr.write("run.mjs: no test files found\n");
  process.exit(2);
}

let failed = 0;
for (const f of files) {
  const testPath = join(here, f);
  process.stdout.write(`\n=== ${f} ===\n`);
  const result = spawnSync(process.execPath, [testPath], {
    stdio: "inherit",
  });
  if (result.status !== 0) failed++;
}

if (failed > 0) {
  process.stderr.write(`\n${failed} file(s) failed\n`);
  process.exit(1);
}
process.stdout.write(`\nall ${files.length} file(s) passed\n`);
