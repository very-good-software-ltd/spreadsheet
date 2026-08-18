import { createWriteStream, openAsBlob } from "node:fs";
import { Writable } from "node:stream";
import { Workbook } from "@very-good-software/spreadsheet";

// The template's author selected the data rows in Excel and named them "Lines".
// They also put a total underneath, styled the region, and never told anyone which
// row anything is on, because none of that is this code's business.
const workbook = await Workbook.open(await openAsBlob("invoice-template.xlsx"));
const editor = workbook.edit();

const lines = await invoiceLines();

// However many lines there are, the region ends up that tall. The total below it
// moves with the rows and keeps covering all of them.
editor.worksheet("Invoice").writeRegion("Lines", lines);

await editor.save().pipeTo(Writable.toWeb(createWriteStream("invoice.xlsx")));

async function invoiceLines(): Promise<(string | number)[][]> {
  return [
    ["Consulting", 12, 950],
    ["Expenses", 1, 240],
    ["Support", 3, 180],
  ];
}
