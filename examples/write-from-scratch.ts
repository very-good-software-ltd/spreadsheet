import { createWriteStream } from "node:fs";
import { Writable } from "node:stream";
import { formula, Workbook } from "@very-good-software/spreadsheet";

// Nothing to open, so start from an empty workbook. Everything after the first
// line is the same as filling a template.
const editor = (await Workbook.create()).edit();
const sheet = editor.worksheet(0);

sheet.appendRows([
  ["Region", "Units"],
  ["North", 120],
  ["South", 340],
]);
sheet.set("A5", "Total");
sheet.set("B5", formula("SUM(B2:B3)"));

await editor.save().pipeTo(Writable.toWeb(createWriteStream("report.xlsx")));
