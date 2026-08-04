import type { Row, Workbook } from "@very-good-software/spreadsheet";

// There is no built-in "read everything" on purpose.
// Streaming keeps memory low by default, and holding the whole workbook
// is a plain loop when you want it.
export async function readAll(workbook: Workbook): Promise<Map<string, Row[]>> {
  const sheets = new Map<string, Row[]>();
  for (const name of workbook.worksheetNames) {
    const rows: Row[] = [];
    for await (const row of workbook.worksheet(name).rows()) {
      rows.push(row);
    }
    sheets.set(name, rows);
  }
  return sheets;
}

// Array.fromAsync(sheet.rows()) does the same for one sheet in a single line,
// where the runtime and your TypeScript lib provide it.
