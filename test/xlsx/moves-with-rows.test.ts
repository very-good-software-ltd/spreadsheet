import ExcelJS from "exceljs";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import { describe, expect, it } from "vitest";
import * as SheetJS from "xlsx";
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

const PIVOT_CACHE_PART = "xl/pivotCache/pivotCacheDefinition1.xml";
const PIVOT_CACHE_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition";

// No fixture library writes a pivot table, and only the cached source range is
// what stops a move, so the smallest thing that carries one is spliced in.
function withPivotOver(bytes: Uint8Array, sheet: string, extent: string): Uint8Array {
  const files = unzipSync(bytes);
  const relationships = strFromU8(files["xl/_rels/workbook.xml.rels"] ?? new Uint8Array());

  files[PIVOT_CACHE_PART] = strToU8(
    `<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><cacheSource type="worksheet"><worksheetSource sheet="${sheet}" ref="${extent}"/></cacheSource></pivotCacheDefinition>`,
  );
  files["xl/_rels/workbook.xml.rels"] = strToU8(
    relationships.replace(
      "</Relationships>",
      `<Relationship Id="rIdPivot" Type="${PIVOT_CACHE_TYPE}" Target="pivotCache/pivotCacheDefinition1.xml"/></Relationships>`,
    ),
  );

  return zipSync(files);
}

function anchorRowsIn(bytes: Uint8Array): (string | undefined)[] {
  const drawing = strFromU8(unzipSync(bytes)["xl/drawings/drawing1.xml"] ?? new Uint8Array());

  return [...drawing.matchAll(/<xdr:row>(\d+)<\/xdr:row>/g)].map((match) => match[1]);
}

function commentRefsIn(bytes: Uint8Array): (string | undefined)[] {
  const comments = strFromU8(unzipSync(bytes)["xl/comments1.xml"] ?? new Uint8Array());

  return [...comments.matchAll(/<comment ref="([^"]+)"/g)].map((match) => match[1]);
}

// The cell a comment's box hangs from, read back as a sheet row, where the VML
// counts from zero.
function commentCellRowsIn(bytes: Uint8Array): number[] {
  const vml = strFromU8(unzipSync(bytes)["xl/drawings/vmlDrawing1.vml"] ?? new Uint8Array());

  return [...vml.matchAll(/<x:Row>(\d+)<\/x:Row>/g)].map((match) => Number(match[1]) + 1);
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

  it("moves a comment below the rows that move, rather than refusing", async () => {
    const bytes = await templateWith((_workbook, sheet) => {
      sheet.getCell("B7").note = "check this";
    });

    expect(commentRefsIn(await writeThreeRowsInto(bytes))).toEqual(["B8"]);
  });

  it("moves the box the comment is drawn in with it", async () => {
    const bytes = await templateWith((_workbook, sheet) => {
      sheet.getCell("B7").note = "check this";
    });

    expect(commentCellRowsIn(await writeThreeRowsInto(bytes))).toEqual([8]);
  });

  // exceljs reads no comment back, from its own file or from ours, so that the
  // whole comment survived the move is proved by a third library instead.
  it("SheetJS reads the moved comment where it ended up", async () => {
    const bytes = await templateWith((_workbook, sheet) => {
      sheet.getCell("B7").note = "check this";
    });

    const report = SheetJS.read(await writeThreeRowsInto(bytes), { type: "array" }).Sheets["Report"];

    expect(report?.["B8"]?.c?.[0]?.t).toBe("check this");
    expect(report?.["B7"]?.c).toBeUndefined();
  });

  it("leaves a comment above the rows that move where it is", async () => {
    const bytes = await templateWith((_workbook, sheet) => {
      sheet.getCell("B1").note = "check this";
    });

    expect(commentRefsIn(await writeThreeRowsInto(bytes))).toEqual(["B1"]);
  });

  it("refuses when a pivot table reads from where the rows would move", async () => {
    const bytes = withPivotOver(await templateWith(() => {}), "Report", "B3:B5");

    await expect(writeThreeRowsInto(bytes)).rejects.toThrow(
      "that sheet has a pivot table reading from row 3 or below, whose cached source range we do not rewrite",
    );
  });

  it("leaves a pivot table reading from above the rows alone", async () => {
    const bytes = withPivotOver(await templateWith(() => {}), "Report", "B1:B2");

    await expect(writeThreeRowsInto(bytes)).resolves.toBeInstanceOf(Uint8Array);
  });

  it("names the region and the row the move starts at", async () => {
    const bytes = withPivotOver(await templateWith(() => {}), "Report", "B3:B5");

    await expect(writeThreeRowsInto(bytes)).rejects.toThrow(
      'Cannot write into "Data": it moves the rows of worksheet "Report" from row 3',
    );
  });
});
