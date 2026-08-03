import { useVirtualizer } from "@tanstack/react-virtual";
import { useRef, useState } from "react";
import type { Row } from "very-good-spreadsheet";
import { Workbook } from "very-good-spreadsheet";
import { columnLetter, formatBytes, formatCell } from "~/lib/format";

const DISPLAY_CAP = 100_000;
const ROW_HEIGHT = 28;
const COLUMN_WIDTH = 128;
const ROW_NUMBER_WIDTH = 64;
const FLUSH_EVERY = 5000;

type Status = "idle" | "reading" | "done" | "error";

export function Demo() {
  const workbookRef = useRef<Workbook | null>(null);
  const rowsRef = useRef<Row[]>([]);
  const scrollRef = useRef<HTMLDivElement>(null);
  const listRef = useRef<HTMLDivElement>(null);

  const [status, setStatus] = useState<Status>("idle");
  const [fileName, setFileName] = useState("");
  const [fileSize, setFileSize] = useState(0);
  const [sheetNames, setSheetNames] = useState<readonly string[]>([]);
  const [sheet, setSheet] = useState("");
  const [displayedCount, setDisplayedCount] = useState(0);
  const [streamedCount, setStreamedCount] = useState(0);
  const [columnCount, setColumnCount] = useState(0);
  const [elapsedMs, setElapsedMs] = useState(0);
  const [dragging, setDragging] = useState(false);
  const [error, setError] = useState("");

  const virtualizer = useVirtualizer({
    count: displayedCount,
    getScrollElement: () => scrollRef.current,
    estimateSize: () => ROW_HEIGHT,
    overscan: 20,
    scrollMargin: listRef.current?.offsetTop ?? 0,
  });

  async function streamSheet(name: string) {
    const workbook = workbookRef.current;
    if (workbook === null) {
      return;
    }
    rowsRef.current = [];
    setDisplayedCount(0);
    setStreamedCount(0);
    setColumnCount(0);
    setStatus("reading");
    setError("");

    const start = performance.now();
    let streamed = 0;
    let maxColumn = 0;
    try {
      for await (const row of workbook.worksheet(name).rows()) {
        streamed += 1;
        for (const cell of row.cells) {
          if (cell.columnIndex > maxColumn) {
            maxColumn = cell.columnIndex;
          }
        }
        if (rowsRef.current.length < DISPLAY_CAP) {
          rowsRef.current.push(row);
        }
        if (streamed % FLUSH_EVERY === 0) {
          setStreamedCount(streamed);
          setColumnCount(maxColumn + 1);
          setDisplayedCount(rowsRef.current.length);
          setElapsedMs(performance.now() - start);
          await new Promise((resolve) => setTimeout(resolve, 0));
        }
      }
      setStreamedCount(streamed);
      setColumnCount(maxColumn + 1);
      setDisplayedCount(rowsRef.current.length);
      setElapsedMs(performance.now() - start);
      setStatus("done");
    } catch (caught) {
      setError(caught instanceof Error ? caught.message : String(caught));
      setStatus("error");
    }
  }

  async function handleFile(file: File | undefined) {
    if (file === undefined) {
      return;
    }
    setFileName(file.name);
    setFileSize(file.size);
    setStatus("reading");
    setError("");
    try {
      // A File is a seekable Blob, so this reads it in ranges off disk rather
      // than loading the whole file into memory first.
      const workbook = await Workbook.open(file);
      workbookRef.current = workbook;
      setSheetNames(workbook.worksheetNames);
      const first = workbook.worksheetNames[0] ?? "";
      setSheet(first);
      await streamSheet(first);
    } catch (caught) {
      setError(caught instanceof Error ? caught.message : String(caught));
      setStatus("error");
    }
  }

  const gridWidth = ROW_NUMBER_WIDTH + columnCount * COLUMN_WIDTH;
  const capped = displayedCount < streamedCount;

  return (
    // biome-ignore lint/a11y/noStaticElementInteractions: drag and drop is an enhancement, the file input is the accessible path.
    <div
      className="flex flex-col gap-3"
      onDragOver={(event) => {
        event.preventDefault();
        setDragging(true);
      }}
      onDragLeave={() => setDragging(false)}
      onDrop={(event) => {
        event.preventDefault();
        setDragging(false);
        void handleFile(event.dataTransfer.files[0]);
      }}
    >
      <div className="flex flex-wrap items-center gap-4">
        <label className="cursor-pointer rounded-md bg-emerald-600 px-3 py-1.5 text-sm font-medium text-white hover:bg-emerald-700">
          Choose .xlsx
          <input
            type="file"
            accept=".xlsx"
            className="hidden"
            onChange={(event) => handleFile(event.target.files?.[0])}
          />
        </label>
        {sheetNames.length > 1 && (
          <select
            className="rounded-md border border-gray-300 px-2 py-1 text-sm dark:border-gray-600 dark:bg-gray-800"
            value={sheet}
            onChange={(event) => {
              setSheet(event.target.value);
              void streamSheet(event.target.value);
            }}
          >
            {sheetNames.map((name) => (
              <option key={name} value={name}>
                {name}
              </option>
            ))}
          </select>
        )}
        {status !== "idle" && status !== "error" && (
          <div className="flex flex-wrap items-center gap-x-4 gap-y-1 text-sm text-gray-600 dark:text-gray-400">
            <span className="font-medium text-gray-800 dark:text-gray-200">{fileName}</span>
            <span>{formatBytes(fileSize)}</span>
            <span>
              {streamedCount.toLocaleString()} rows × {columnCount} cols
            </span>
            <span>{(elapsedMs / 1000).toFixed(1)}s</span>
            {status === "reading" && <span className="text-emerald-600 dark:text-emerald-400">reading…</span>}
            {capped && (
              <span className="text-amber-700 dark:text-amber-400">
                showing first {displayedCount.toLocaleString()}
              </span>
            )}
          </div>
        )}
        {status === "error" && <span className="text-sm text-red-600">Failed: {error}</span>}
      </div>

      <div
        ref={scrollRef}
        className={`relative h-[70vh] overflow-auto rounded-lg border font-mono text-xs ${
          dragging
            ? "border-emerald-400 bg-emerald-50/50 dark:bg-emerald-950/30"
            : "border-gray-200 dark:border-gray-700"
        }`}
      >
        {displayedCount === 0 ? (
          <div className="flex h-full items-center justify-center p-8 text-center font-sans text-sm text-gray-400">
            {status === "reading" ? "Reading…" : "Drop an .xlsx here or choose a file above."}
          </div>
        ) : (
          <div style={{ width: gridWidth }}>
            <div className="sticky top-0 z-10 flex bg-gray-100 font-sans font-semibold dark:bg-gray-800">
              <div
                className="shrink-0 border-b border-gray-200 px-2 py-1 dark:border-gray-700"
                style={{ width: ROW_NUMBER_WIDTH }}
              />
              {Array.from({ length: columnCount }, (_, column) => (
                <div
                  // biome-ignore lint/suspicious/noArrayIndexKey: no real candidate for a better unique index.
                  key={column}
                  className="shrink-0 border-b border-l border-gray-200 px-2 py-1 dark:border-gray-700"
                  style={{ width: COLUMN_WIDTH }}
                >
                  {columnLetter(column)}
                </div>
              ))}
            </div>

            <div ref={listRef} style={{ height: virtualizer.getTotalSize(), position: "relative" }}>
              {virtualizer.getVirtualItems().map((item) => {
                const row = rowsRef.current[item.index];
                if (row === undefined) {
                  return null;
                }
                return (
                  <div
                    key={item.key}
                    className="absolute left-0 flex"
                    style={{
                      height: ROW_HEIGHT,
                      transform: `translateY(${item.start - virtualizer.options.scrollMargin}px)`,
                    }}
                  >
                    <div
                      className="shrink-0 border-b border-gray-100 bg-gray-50 px-2 py-1 text-right text-gray-400 dark:border-gray-800 dark:bg-gray-900"
                      style={{ width: ROW_NUMBER_WIDTH }}
                    >
                      {row.number}
                    </div>
                    {Array.from({ length: columnCount }, (_, column) => {
                      const cell = row.cell(column);
                      if (cell === undefined) {
                        return (
                          <div
                            // biome-ignore lint/suspicious/noArrayIndexKey: no real candidate for a better unique index.
                            key={column}
                            className="shrink-0 border-b border-l border-gray-100 px-2 py-1 dark:border-gray-800"
                            style={{ width: COLUMN_WIDTH }}
                          />
                        );
                      }
                      const formatted = formatCell(cell);
                      return (
                        <div
                          // biome-ignore lint/suspicious/noArrayIndexKey: no real candidate for a better unique index.
                          key={column}
                          className={`shrink-0 truncate border-b border-l border-gray-100 px-2 py-1 dark:border-gray-800 ${formatted.align} ${formatted.tone}`}
                          style={{ width: COLUMN_WIDTH }}
                          title={formatted.text}
                        >
                          {formatted.text}
                        </div>
                      );
                    })}
                  </div>
                );
              })}
            </div>
          </div>
        )}
      </div>
    </div>
  );
}
