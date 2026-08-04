import { readFile } from "node:fs/promises";
import { Workbook } from "@very-good-software/spreadsheet";

const workbook = await Workbook.open(await readFile("data.xlsx"));

for (const name of workbook.worksheetNames) {
  console.log(name);
}

for await (const row of workbook.worksheet("Sheet1").rows()) {
  for (const cell of row.cells) {
    console.log(cell.ref, cell.type, cell.value);
  }
}
