import ExcelJS from "exceljs";
import { strFromU8, unzipSync } from "fflate";
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

function anchorRowsIn(bytes: Uint8Array): (string | undefined)[] {
  const drawing = strFromU8(unzipSync(bytes)["xl/drawings/drawing1.xml"] ?? new Uint8Array());

  return [...drawing.matchAll(/<xdr:row>(\d+)<\/xdr:row>/g)].map((match) => match[1]);
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

  it("moves an image anchored where the rows move, rather than refusing", async () => {
    const bytes = await templateWith((workbook, sheet) => {
      const image = workbook.addImage({ base64: PIXEL, extension: "png" });
      sheet.addImage(image, "D6:E9");
    });

    // A drawing counts rows from zero and its lower corner is exclusive, so the
    // template anchors D6:E9 at 5 and 9. One row appears above it, so both move.
    expect(anchorRowsIn(await writeThreeRowsInto(bytes))).toEqual(["6", "10"]);
  });

  it("leaves an image anchored above the rows that move where it is", async () => {
    const bytes = await templateWith((workbook, sheet) => {
      const image = workbook.addImage({ base64: PIXEL, extension: "png" });
      sheet.addImage(image, "D1:E2");
    });

    expect(anchorRowsIn(await writeThreeRowsInto(bytes))).toEqual(["0", "2"]);
  });

  it("refuses when a comment sits where the rows would move", async () => {
    const bytes = await templateWith((_workbook, sheet) => {
      sheet.getCell("B7").note = "check this";
    });

    await expect(writeThreeRowsInto(bytes)).rejects.toThrow(
      "that sheet has a cell comment at or below row 3, which is positioned by a drawing of its own",
    );
  });

  it("names the region and the row the move starts at", async () => {
    const bytes = await templateWith((_workbook, sheet) => {
      sheet.getCell("B7").note = "check this";
    });

    await expect(writeThreeRowsInto(bytes)).rejects.toThrow(
      'Cannot write into "Data": it moves the rows of worksheet "Report" from row 3',
    );
  });
});
