import { Workbook } from "very-good-spreadsheet";

// A File from an <input type="file"> is a seekable Blob, so this reads it in
// ranges off disk instead of loading the whole file into memory first.
export async function readSpreadsheet(file: File): Promise<void> {
  const workbook = await Workbook.open(file);

  for await (const row of workbook.firstWorksheet().rows()) {
    console.log(row.number, row.cells);
  }
}
