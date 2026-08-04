import { readFileSync, writeFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

// Copies each file in examples/ into the README, between anchors:
//   <!-- example: read-node.ts -->
//   ```ts
//   ...file contents...
//   ```
//   <!-- /example -->
// The examples are type-checked by tsc, so this keeps the README from drifting
// as the API changes. Run with --check to fail instead of writing.

const root = join(dirname(fileURLToPath(import.meta.url)), "..");
const readmePath = join(root, "README.md");
const examplesDir = join(root, "examples");
const check = process.argv.includes("--check");

// The head and tail are captured verbatim, including any blank lines a markdown
// formatter puts around the anchors, so only the code between them is replaced.
const region = /(<!-- example: ([\w.-]+) -->\s*```ts\n)[\s\S]*?(\n```\s*<!-- \/example -->)/g;

const before = readFileSync(readmePath, "utf8");
const after = before.replace(region, (_match, head, name, tail) => {
  const code = readFileSync(join(examplesDir, name), "utf8").replace(/\n+$/, "");
  return `${head}${code}${tail}`;
});

if (after === before) {
  console.log("examples: README is up to date");
} else if (check) {
  console.error("examples: README is out of date. Run `npm run examples` to update it.");
  process.exit(1);
} else {
  writeFileSync(readmePath, after);
  console.log("examples: README updated");
}
