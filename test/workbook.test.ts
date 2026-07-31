import { describe, expect, it } from "vitest";
import { Workbook } from "../src/xlsx/workbook";
import { xlsx } from "./support/xlsx-fixture";

describe("Workbook", () => {
  it("lists worksheet names in document order", async () => {
    const workbook = await Workbook.open(
      xlsx([
        { name: "Summary", rows: [] },
        { name: "Data", rows: [] },
      ]),
    );

    expect(workbook.worksheetNames).toEqual(["Summary", "Data"]);
  });

  it("reports worksheet visibility", async () => {
    const workbook = await Workbook.open(
      xlsx([
        { name: "Visible", rows: [] },
        { name: "Hidden", rows: [], hidden: true },
      ]),
    );

    expect(workbook.worksheets).toEqual([
      { name: "Visible", hidden: false },
      { name: "Hidden", hidden: true },
    ]);
  });
});
