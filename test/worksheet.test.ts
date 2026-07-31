import { describe, expect, it } from "vitest";
import type { Row } from "../src/xlsx/cell";
import { Workbook } from "../src/xlsx/workbook";
import { xlsx } from "./support/xlsx-fixture";

describe("Worksheet rows", () => {
  it("reads number and shared-string cells", async () => {
    const workbook = await Workbook.open(
      xlsx([
        {
          name: "Data",
          rows: [
            [1, "Hello"],
            ["World", 2.5],
          ],
        },
      ]),
    );

    const rows: Row[] = [];
    for await (const row of workbook.worksheet("Data").rows()) {
      rows.push(row);
    }

    expect(rows).toEqual([
      {
        number: 1,
        cells: [
          { ref: "A1", type: "number", value: 1 },
          { ref: "B1", type: "string", value: "Hello" },
        ],
      },
      {
        number: 2,
        cells: [
          { ref: "A2", type: "string", value: "World" },
          { ref: "B2", type: "number", value: 2.5 },
        ],
      },
    ]);
  });
});
