const CODE = `import { readFile } from "node:fs/promises";
import { Workbook } from "very-good-spreadsheet";

const workbook = await Workbook.open(await readFile("data.xlsx"));

for await (const row of workbook.worksheet("Sheet1").rows()) {
  for (const cell of row.cells) {
    console.log(cell.ref, cell.type, cell.value);
  }
}`;

export function Usage() {
  return (
    <section className="border-t border-gray-100 px-6 py-16 dark:border-gray-900">
      <div className="mx-auto max-w-3xl">
        <h2 className="text-2xl font-bold">Usage</h2>
        <p className="mt-2 text-sm text-gray-600 dark:text-gray-400">
          Open a workbook from its bytes, then stream a sheet's rows. In the browser, read the bytes with{" "}
          <code className="font-mono">file.arrayBuffer()</code> instead.
        </p>
        <pre className="mt-6 overflow-x-auto rounded-lg border border-gray-200 bg-gray-50 p-4 text-sm dark:border-gray-800 dark:bg-gray-900">
          <code>{CODE}</code>
        </pre>
      </div>
    </section>
  );
}
