import { createWriteStream, openAsBlob } from "node:fs";
import { Writable } from "node:stream";
import { Workbook } from "@very-good-software/spreadsheet";

// The template's author selected the data region in Excel and named it "Lines".
// Writing by that name rather than by cell reference means they can insert a row
// above it, or move the region to another sheet, without this code changing.
const workbook = await Workbook.open(await openAsBlob("invoice-template.xlsx"));
const editor = workbook.edit();

editor.worksheet("Invoice").writeRegion("Lines", [
  ["Consulting", 12, 950],
  ["Expenses", 1, 240],
]);

// Rows of the region left over from the last run are cleared, so the totals row
// underneath it can never add up stale numbers.
await editor.save().pipeTo(Writable.toWeb(createWriteStream("invoice.xlsx")));
