import { BENCHMARK } from "~/site";

const MAX_MEMORY = Math.max(...BENCHMARK.map((row) => row.peakMemoryMb));

export function Benchmarks() {
  return (
    <section className="border-t border-gray-100 px-6 py-16 dark:border-gray-900">
      <div className="mx-auto max-w-3xl">
        <h2 className="text-2xl font-bold">Benchmark</h2>
        <p className="mt-2 text-sm text-gray-600 dark:text-gray-400">
          Reading every cell of a 28 MB file whose single sheet is 170 MB uncompressed, 4.4 million cells, each library
          in its own Node process. The bar is peak memory, the highest the process reached.
        </p>

        <div className="mt-8 space-y-4">
          {BENCHMARK.map((row) => (
            <div key={`${row.library}-${row.mode}`}>
              <div className="flex items-baseline justify-between text-sm">
                <span className={row.us ? "font-semibold text-emerald-700 dark:text-emerald-400" : ""}>
                  {row.library} <span className="font-normal text-gray-400">({row.mode})</span>
                </span>
                <span className="tabular-nums text-gray-600 dark:text-gray-400">
                  {row.peakMemoryMb.toLocaleString()} MB / {row.timeSeconds}s
                </span>
              </div>
              <div className="mt-1 h-2 rounded bg-gray-100 dark:bg-gray-800">
                <div
                  className={`h-2 rounded ${row.us ? "bg-emerald-600" : "bg-gray-400 dark:bg-gray-600"}`}
                  style={{ width: `${(row.peakMemoryMb / MAX_MEMORY) * 100}%` }}
                />
              </div>
            </div>
          ))}
        </div>

        <p className="mt-6 text-xs text-gray-500">
          exceljs and SheetJS build a full model of the file. exceljs also has a streaming reader, shown here, but it is
          Node only. These numbers are from one file on one machine. Run{" "}
          <code className="font-mono">npm run benchmark</code> to reproduce.
        </p>
      </div>
    </section>
  );
}
