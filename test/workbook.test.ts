import { describe, expect, it } from "vitest";
import { Workbook } from "../src/xlsx/workbook";
import { xlsxWithSheets } from "./support/xlsx-fixture";

describe("Workbook", () => {
  it("lists worksheet names in document order", async () => {
    const workbook = await Workbook.open(xlsxWithSheets(["Summary", "Data"]));

    expect(workbook.worksheetNames).toEqual(["Summary", "Data"]);
  });
});
