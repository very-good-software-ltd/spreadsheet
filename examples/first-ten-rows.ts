import type { Row, Worksheet } from "@very-good-software/spreadsheet";

// There is no take or limit on purpose. rows() is a stream, so breaking the loop
// stops the reading, and only the first ten rows of a huge file are ever parsed.
export async function firstTenRows(sheet: Worksheet): Promise<Row[]> {
  const preview: Row[] = [];
  for await (const row of sheet.rows()) {
    preview.push(row);
    if (preview.length === 10) {
      break;
    }
  }
  return preview;
}
