import { BENCHMARK, ODS_BENCHMARK } from "~/site";

const MAX_MEMORY = Math.max(...BENCHMARK.map((row) => row.peakMemoryMb));

export function Benchmarks() {
  return (
    <section className="border-t border-gray-100 px-6 py-16 dark:border-gray-900">
      <div className="mx-auto max-w-3xl">
        <h2 className="text-2xl font-bold">Benchmark</h2>
        <h3 className="mt-6 text-lg font-semibold">Reading xlsx</h3>
        <p className="mt-2 text-sm text-gray-600 dark:text-gray-400">
          Reading every cell of a 28 MB xlsx file whose single sheet is 170 MB uncompressed, 4.4 million cells, each
          library in its own Node process. The bar is peak memory, the highest the process reached.
        </p>

        <div className="mt-8 space-y-4">
          {BENCHMARK.map((row) => (
            <div key={`${row.library}-${row.mode}`}>
              <div className="flex items-baseline justify-between text-sm">
                <span className={row.us ? "font-semibold text-emerald-700 dark:text-emerald-400" : ""}>
                  {row.library} <span className="font-normal text-gray-400">({row.mode})</span>
                </span>
                <span className="tabular-nums text-right text-gray-600 dark:text-gray-400">
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

        <div className="mt-12">
          <h3 className="text-lg font-semibold">Reading .ods</h3>
          <p className="mt-2 text-sm text-gray-600 dark:text-gray-400">
            .ods is a different story. OpenDocument keeps the whole spreadsheet in one compressed stream you can't seek
            into, so we re-read from the start for each sheet. That makes us slower on .ods than on xlsx, and slower
            than readers that pull the entire file into memory. But the memory stays low. It barely moves as the file
            grows, 201 MB on a 14 MB file and 249 MB on a 48 MB one. So on a big .ods we finish where the
            load-everything readers can't. And .ods can get faster. Most of it is re-reading the stream for each sheet,
            which we can skip, so there is room to close the gap once a real workload needs it.
          </p>

          <div className="mt-6 overflow-x-auto">
            <table className="w-full text-sm">
              <thead>
                <tr className="text-left text-gray-500">
                  <th className="py-1 font-medium">file</th>
                  <th className="py-1 font-medium">ours, streaming</th>
                  <th className="py-1 font-medium">SheetJS, load</th>
                </tr>
              </thead>
              <tbody className="tabular-nums">
                {ODS_BENCHMARK.map((row) => (
                  <tr key={row.file} className="border-t border-gray-100 dark:border-gray-800">
                    <td className="py-2">
                      {row.file} <span className="text-gray-400">({row.cells})</span>
                    </td>
                    <td className="py-2 font-semibold text-emerald-700 dark:text-emerald-400">{row.ours}</td>
                    <td className="py-2 text-gray-600 dark:text-gray-400">{row.sheetjs}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>

          <p className="mt-4 text-xs text-gray-500">
            SheetJS crashes on the 48 MB file, it hits Node's maximum string size. exceljs doesn't read .ods at all, so
            neither can read the whole file.
          </p>
        </div>
      </div>
    </section>
  );
}
