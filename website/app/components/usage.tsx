import { SyntaxHighlight } from "~/lib/syntax-highlight";
import readNode from "../../../examples/read-node.ts?raw";

export function Usage() {
  return (
    <section className="border-t border-gray-100 px-6 py-16 dark:border-gray-900">
      <div className="mx-auto max-w-3xl">
        <h2 className="text-2xl font-bold">Usage</h2>
        <p className="mt-2 text-sm text-gray-600 dark:text-gray-400">
          Open a workbook, then stream a sheet's rows. In the browser, pass the chosen{" "}
          <code className="font-mono">File</code> straight to <code className="font-mono">Workbook.open</code>, which
          reads it in ranges instead of loading it all.
        </p>
        <pre className="mt-6 rounded-lg border font-mono text-sm border-gray-200 bg-gray-50 dark:border-gray-700 dark:bg-gray-900">
          <SyntaxHighlight>{readNode.trimEnd()}</SyntaxHighlight>
        </pre>
      </div>
    </section>
  );
}
