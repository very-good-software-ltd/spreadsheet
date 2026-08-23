import { readdirSync, readFileSync, writeFileSync } from "node:fs";
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

const anchor = /<!-- example: ([\w.-]+) -->/g;

// Only the two anchors are matched. Everything between them is thrown away and
// written again, fences included, so nothing in the block is depended on to be
// correct before the run. A fence that is missing, duplicated or run into the last
// line of code is replaced rather than having to be matched.
//
// The body cannot contain an anchor of either kind, so a block missing its closing
// anchor fails to match rather than swallowing the prose up to the next block.
const region = /<!-- example: ([\w.-]+) -->\n(?:(?!<!-- example:|<!-- \/example -->)[\s\S])*?<!-- \/example -->/g;

const before = readFileSync(readmePath, "utf8");

// A block the pattern cannot match would otherwise be skipped in silence, and
// the README would keep whatever stale code is in it.
const declared = [...before.matchAll(anchor)].map(([, name]) => name);
const matched = [...before.matchAll(region)].map((match) => match[1]);
// Only the first mismatch is reported. Past that the two lists no longer line up,
// so every later name would be named whether its own block is fine or not.
const malformed = declared.find((name, index) => matched[index] !== name);

if (malformed !== undefined) {
  console.error(`examples: malformed block in README.md for ${malformed}.`);
  console.error("It needs an opening <!-- example: name --> anchor and a closing <!-- /example --> one.");
  process.exit(1);
}

// An example nobody put an anchor in for is type-checked and then never seen, so
// adding one and forgetting to place it would go unnoticed.
const unplaced = readdirSync(examplesDir).filter((name) => !declared.includes(name));

if (unplaced.length > 0) {
  console.error(`examples: no anchor in README.md for ${unplaced.join(", ")}.`);
  console.error("Add <!-- example: name --> and <!-- /example --> where it should appear.");
  process.exit(1);
}

const after = before.replace(region, (_match, name) => {
  const code = readFileSync(join(examplesDir, name), "utf8").replace(/\n+$/, "");
  return `<!-- example: ${name} -->\n\n\`\`\`ts\n${code}\n\`\`\`\n\n<!-- /example -->`;
});

if (after === before) {
  console.log(`examples: README is up to date (${declared.length} examples)`);
} else if (check) {
  console.error("examples: README is out of date. Run `npm run examples` to update it.");
  process.exit(1);
} else {
  writeFileSync(readmePath, after);
  console.log(`examples: README updated (${declared.length} examples)`);
}
