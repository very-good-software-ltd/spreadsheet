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

const anchor = /<!-- example: ([\w.-]+) -->/g;

// The head and tail are captured verbatim, including any blank lines a markdown
// formatter puts around the anchors, so only the code between them is replaced.
//
// Two things here are load-bearing, both to stop one block's replacement running
// into the next and deleting the prose in between.
//
// The newline before the closing fence is optional, so an empty placeholder still
// matches. The head has already eaten the newline after the opening fence, so
// requiring another one means an empty block cannot match at its own position.
//
// And the body cannot contain a closing anchor, so a block missing its fence
// fails to match rather than swallowing everything up to the next block's anchor.
// A body that reaches too far is then caught below instead of being written out.
const region =
  /(<!-- example: ([\w.-]+) -->\s*```ts\n)((?:(?!<!-- \/example -->)[\s\S])*?)(\n?```\s*<!-- \/example -->)/g;

const before = readFileSync(readmePath, "utf8");

// A block the pattern cannot match would otherwise be skipped in silence, and
// the README would keep whatever stale code is in it.
const declared = [...before.matchAll(anchor)].map(([, name]) => name);
const matched = [...before.matchAll(region)].map((match) => match[2]);
// Only the first mismatch is reported. Past that the two lists no longer line up,
// so every later name would be named whether its own block is fine or not.
const malformed = declared.find((name, index) => matched[index] !== name);

if (malformed !== undefined) {
  console.error(`examples: malformed block in README.md for ${malformed}.`);
  console.error("It needs an opening ```ts fence, a closing ``` fence, then <!-- /example -->.");
  process.exit(1);
}

const after = before.replace(region, (_match, head, name, _body, tail) => {
  const code = readFileSync(join(examplesDir, name), "utf8").replace(/\n+$/, "");
  return `${head}${code}${tail}`;
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
