import { Readable } from "node:stream";
import ExcelJS from "exceljs";
import { describe, expect, it } from "vitest";
import * as SheetJS from "xlsx";
import { BytesByteRange } from "../../src/io/byte-range";
import { Workbook } from "../../src/workbook";
import { openZip } from "../../src/zip/open-zip";
import type { StoredZipEntry, ZipArchive } from "../../src/zip/zip-archive";

// The input is written by exceljs rather than by our own fixture helper, because
// what matters here is surviving a file we did not produce. It carries parts we
// have no reader for at all, a theme and document properties, which is the point:
// a part nobody understands has to come out the other side untouched.

// The parts we rewrite on purpose. Everything else has to be byte-identical.
const REWRITTEN = new Set(["xl/worksheets/sheet1.xml", "xl/styles.xml", "xl/workbook.xml"]);

async function template(): Promise<Uint8Array> {
  const workbook = new ExcelJS.Workbook();
  const report = workbook.addWorksheet("Report");

  report.getColumn(1).width = 30;
  report.getCell("A1").value = "Corporate header";
  report.getCell("A1").font = { bold: true, size: 14 };
  report.getCell("A1").fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFDDEEFF" } };
  report.mergeCells("A1:C1");
  report.getCell("A2").value = "Qty";
  report.getCell("B2").value = 10;
  report.getCell("C2").value = { formula: "B2*2", result: 20 };
  report.getCell("A4").value = "Total";
  workbook.addWorksheet("Notes").getCell("A1").value = "untouched sheet";

  return new Uint8Array(await workbook.xlsx.writeBuffer());
}

async function bytesOf(stream: ReadableStream<Uint8Array>): Promise<Uint8Array> {
  const reader = stream.getReader();
  const chunks: Uint8Array[] = [];
  let total = 0;

  for (;;) {
    const { done, value } = await reader.read();
    if (done) {
      break;
    }
    chunks.push(value);
    total += value.length;
  }

  const all = new Uint8Array(total);
  let offset = 0;
  for (const chunk of chunks) {
    all.set(chunk, offset);
    offset += chunk.length;
  }

  return all;
}

async function filled(source: Uint8Array): Promise<Uint8Array> {
  const editor = (await Workbook.open(source)).edit();

  editor
    .worksheet("Report")
    .set("B2", 25)
    .appendRows([
      ["extra", 1],
      ["rows", 2],
    ]);

  return bytesOf(editor.save());
}

// Read through a stream rather than through exceljs's load(), whose declared
// parameter type is unsatisfiable: exceljs declares a global `interface Buffer
// extends ArrayBuffer`, which merges with Node's own Buffer and demands members
// no value has. Nothing in the library is affected, only how a test calls it.
async function readWithExcelJs(bytes: Uint8Array): Promise<ExcelJS.Workbook> {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.read(Readable.from([bytes]));
  return workbook;
}

function open(bytes: Uint8Array): Promise<ZipArchive> {
  return openZip(new BytesByteRange(bytes));
}

function fingerprint(entry: StoredZipEntry): unknown {
  return {
    method: entry.method,
    crc32: entry.crc32,
    compressedSize: entry.compressedSize,
    uncompressedSize: entry.uncompressedSize,
  };
}

describe("writing keeps the parts it does not touch", () => {
  it("copies every other entry across byte for byte, without recompressing it", async () => {
    const source = await template();
    const before = await open(source);
    const after = await open(await filled(source));

    const checked: string[] = [];

    for (const entry of before.entries()) {
      if (REWRITTEN.has(entry.path)) {
        continue;
      }

      expect(after.has(entry.path), `${entry.path} is missing from the written file`).toBe(true);
      expect(fingerprint(after.storedEntry(entry.path)), entry.path).toEqual(
        fingerprint(before.storedEntry(entry.path)),
      );
      expect(await after.read(entry.path), entry.path).toEqual(await before.read(entry.path));
      checked.push(entry.path);
    }

    // Guards the assertion above: if the loop ever stopped finding parts, the
    // test would pass by checking nothing.
    expect(checked).toContain("xl/theme/theme1.xml");
    expect(checked).toContain("docProps/core.xml");
    expect(checked).toContain("xl/worksheets/sheet2.xml");
    expect(checked).toContain("xl/sharedStrings.xml");
  });

  it("adds no entry and drops none", async () => {
    const source = await template();
    const before = await open(source);
    const after = await open(await filled(source));

    expect(
      after
        .entries()
        .map((entry) => entry.path)
        .sort(),
    ).toEqual(
      before
        .entries()
        .map((entry) => entry.path)
        .sort(),
    );
  });
});

describe("what other libraries make of the output", () => {
  it("exceljs reads it, and the template's formatting is still there", async () => {
    const written = await filled(await template());

    const workbook = await readWithExcelJs(written);
    const report = workbook.getWorksheet("Report");

    expect(report?.getCell("B2").value).toBe(25);
    expect(report?.getCell("A1").value).toBe("Corporate header");
    expect(report?.getCell("A1").font?.bold).toBe(true);
    expect(report?.getColumn(1).width).toBe(30);
    expect(report?.getCell("A4").value).toBe("Total");
    expect(workbook.getWorksheet("Notes")?.getCell("A1").value).toBe("untouched sheet");
  });

  it("exceljs still sees the merged range", async () => {
    const written = await filled(await template());

    const workbook = await readWithExcelJs(written);

    expect(workbook.getWorksheet("Report")?.getCell("B1").isMerged).toBe(true);
  });

  it("exceljs sees the formula, with no cached result to go stale", async () => {
    const written = await filled(await template());

    const workbook = await readWithExcelJs(written);
    const cell = workbook.getWorksheet("Report")?.getCell("C2");

    expect(cell?.formula).toBe("B2*2");
  });

  it("SheetJS reads it", async () => {
    const written = await filled(await template());

    const workbook = SheetJS.read(written, { type: "array" });
    const report = workbook.Sheets["Report"];

    expect(workbook.SheetNames).toEqual(["Report", "Notes"]);
    expect(report?.["B2"]?.v).toBe(25);
    expect(report?.["A1"]?.v).toBe("Corporate header");
  });

  // The formula cell comes back as its text, since the writer left the formula in
  // place and dropped its cached result for the application to recalculate.
  it("our own reader reads it, appended rows included", async () => {
    const written = await filled(await template());
    const workbook = await Workbook.open(written);
    const rows: unknown[][] = [];

    for await (const row of workbook.worksheet("Report").rows()) {
      rows.push([row.number, ...row.cells.map((cell) => cell.value)]);
    }

    expect(rows).toEqual([
      [1, "Corporate header"],
      [2, "Qty", 25, "B2*2"],
      [4, "Total"],
      [5, "extra", 1],
      [6, "rows", 2],
    ]);
  });
});

describe("writing a date into a template", () => {
  it("renders as a date and keeps the cell's other formatting", async () => {
    const source = await template();
    const editor = (await Workbook.open(source)).edit();

    editor.worksheet("Report").set("A1", new Date("2020-06-01T00:00:00.000Z"));
    const written = await bytesOf(editor.save());

    const workbook = await readWithExcelJs(written);
    const cell = workbook.getWorksheet("Report")?.getCell("A1");

    expect(cell?.value).toEqual(new Date("2020-06-01T00:00:00.000Z"));
    expect(cell?.font?.bold).toBe(true);
    expect(cell?.numFmt).toBe("yyyy-mm-dd");
  });

  it("reads back as a date through our own reader", async () => {
    const source = await template();
    const editor = (await Workbook.open(source)).edit();

    editor.worksheet("Report").set("A2", new Date("2020-06-01T00:00:00.000Z"));
    const workbook = await Workbook.open(await bytesOf(editor.save()));

    for await (const row of workbook.worksheet("Report").rows()) {
      if (row.number === 2) {
        expect(row.cell(0)).toMatchObject({ type: "date", value: new Date("2020-06-01T00:00:00.000Z") });
      }
    }
  });
});
