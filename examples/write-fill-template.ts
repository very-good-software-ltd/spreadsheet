import { createWriteStream, openAsBlob } from "node:fs";
import { Writable } from "node:stream";
import { date, Workbook } from "@very-good-software/spreadsheet";

// Open a template, fill in the parts that change, and write a new file. The
// template's own charts, formatting and formulas are copied across untouched.
const workbook = await Workbook.open(await openAsBlob("template.xlsx"));
const editor = workbook.edit();

const invoice = editor.worksheet("Invoice");
invoice.set("C3", "Acme Ltd");
invoice.set("C4", date("2026-03-01"));

invoice.writeRows(8, [
  ["Consulting", 12, 950],
  ["Expenses", 1, 240],
]);

await editor.save().pipeTo(Writable.toWeb(createWriteStream("invoice.xlsx")));
