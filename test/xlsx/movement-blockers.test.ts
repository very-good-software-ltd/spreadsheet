import ExcelJS from "exceljs";
import { describe, expect, it } from "vitest";
import { Workbook } from "../../src/workbook";

// A one pixel PNG, so a drawing part exists without the test carrying an image.
const PIXEL = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==";

async function bytesOf(stream: ReadableStream<Uint8Array>): Promise<Uint8Array> {
  const chunks: Uint8Array[] = [];
  for await (const chunk of stream as unknown as AsyncIterable<Uint8Array>) {
    chunks.push(chunk);
  }
  return new Uint8Array(chunks.flatMap((chunk) => [...chunk]));
}

// A sheet with a region named over rows 3 to 5, so writing more rows moves row 5
// and everything under it.
async function templateWith(decorate: (workbook: ExcelJS.Workbook, sheet: ExcelJS.Worksheet) => void) {
  const source = new ExcelJS.Workbook();
  const sheet = source.addWorksheet("Report");
  for (const row of [3, 4, 5]) {
    sheet.getCell(`B${row}`).value = row;
  }
  source.definedNames.add("Report!B3:B5", "Data");
  decorate(source, sheet);

  return new Uint8Array(await source.xlsx.writeBuffer());
}

async function writeThreeRowsInto(bytes: Uint8Array) {
  const editor = (await Workbook.open(bytes)).edit();
  editor.worksheet("Report").writeRegion("Data", [[1], [2], [3], [4]]);

  return bytesOf(editor.save());
}

describe("what stops a worksheet's rows moving", () => {
  it("moves the rows when nothing is in the way", async () => {
    const bytes = await templateWith(() => {});

    await expect(writeThreeRowsInto(bytes)).resolves.toBeInstanceOf(Uint8Array);
  });

  it("refuses when an image is anchored where the rows would move", async () => {
    const bytes = await templateWith((workbook, sheet) => {
      const image = workbook.addImage({ base64: PIXEL, extension: "png" });
      sheet.addImage(image, "D6:E9");
    });

    await expect(writeThreeRowsInto(bytes)).rejects.toThrow(
      "that sheet has a drawing at or below row 5, which anchors charts and images to row numbers we do not rewrite",
    );
  });

  it("allows an image anchored above the rows that move", async () => {
    const bytes = await templateWith((workbook, sheet) => {
      const image = workbook.addImage({ base64: PIXEL, extension: "png" });
      sheet.addImage(image, "D1:E2");
    });

    await expect(writeThreeRowsInto(bytes)).resolves.toBeInstanceOf(Uint8Array);
  });

  it("refuses when a comment sits where the rows would move", async () => {
    const bytes = await templateWith((_workbook, sheet) => {
      sheet.getCell("B7").note = "check this";
    });

    await expect(writeThreeRowsInto(bytes)).rejects.toThrow(
      "that sheet has a cell comment at or below row 5, which is positioned by a drawing of its own",
    );
  });

  it("names the region and the row the move starts at", async () => {
    const bytes = await templateWith((_workbook, sheet) => {
      sheet.getCell("B7").note = "check this";
    });

    await expect(writeThreeRowsInto(bytes)).rejects.toThrow(
      'Cannot write into "Data": it moves the rows of worksheet "Report" from row 5',
    );
  });
});
