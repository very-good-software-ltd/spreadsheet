import type { Worksheet } from "@very-good-software/spreadsheet";

export async function readGrid(sheet: Worksheet): Promise<void> {
  for await (const row of sheet.rows()) {
    const columnA = row.cell(0);
    console.log(row.number, columnA?.value, row.cells.length);
  }
}
