import { describe, expect, it } from "vitest";
import { BytesByteRange } from "../../src/io/byte-range";
import { readWorkbook } from "../../src/xlsx/read-workbook";
import { createXmlReader } from "../../src/xml/create-xml-reader";
import { openZip } from "../../src/zip/open-zip";
import { xlsx } from "../support/xlsx-fixture";

async function workbookInfo(bytes: Uint8Array) {
  return readWorkbook(await openZip(new BytesByteRange(bytes)), createXmlReader());
}

const ONE_SHEET = [{ name: "Sheet1", rows: [[1]] }];

describe("readWorkbook defined names", () => {
  it("reads a workbook-wide name and where it points", async () => {
    const info = await workbookInfo(xlsx(ONE_SHEET, { definedNames: [{ name: "Data", target: "Sheet1!$B$4:$D$20" }] }));

    expect(info.definedNames).toEqual([
      {
        name: "Data",
        scope: undefined,
        target: {
          kind: "region",
          sheet: "Sheet1",
          firstRow: 4,
          lastRow: 20,
          firstColumnIndex: 1,
          lastColumnIndex: 3,
        },
      },
    ]);
  });

  it("resolves the sheet a scoped name belongs to", async () => {
    const info = await workbookInfo(
      xlsx([{ name: "First", rows: [[1]] }, ...ONE_SHEET], {
        definedNames: [{ name: "Data", target: "Sheet1!$A$1", scope: 1 }],
      }),
    );

    expect(info.definedNames[0]?.scope).toBe("Sheet1");
  });

  it("keeps a scoped name and a workbook-wide name of the same spelling apart", async () => {
    const info = await workbookInfo(
      xlsx(ONE_SHEET, {
        definedNames: [
          { name: "Data", target: "Sheet1!$A$1" },
          { name: "Data", target: "Sheet1!$C$3", scope: 0 },
        ],
      }),
    );

    expect(info.definedNames.map((defined) => defined.scope)).toEqual([undefined, "Sheet1"]);
  });

  it("keeps a name it cannot use, along with the reason", async () => {
    const info = await workbookInfo(xlsx(ONE_SHEET, { definedNames: [{ name: "Rate", target: "0.175" }] }));

    expect(info.definedNames[0]?.target).toEqual({ kind: "unusable", reason: "not a range" });
  });

  // Excel keeps its own bookkeeping in the same element, spelled with a reserved
  // prefix. A print area is an ordinary rectangle that no caller should write into.
  it("drops Excel's own reserved names", async () => {
    const info = await workbookInfo(
      xlsx(ONE_SHEET, {
        definedNames: [
          { name: "_xlnm.Print_Area", target: "Sheet1!$A$1:$D$9", scope: 0 },
          { name: "_xlnm._FilterDatabase", target: "Sheet1!$A$1:$D$9", scope: 0 },
          { name: "Data", target: "Sheet1!$A$1" },
        ],
      }),
    );

    expect(info.definedNames.map((defined) => defined.name)).toEqual(["Data"]);
  });

  it("has no names when the workbook defines none", async () => {
    const info = await workbookInfo(xlsx(ONE_SHEET));

    expect(info.definedNames).toEqual([]);
  });
});
