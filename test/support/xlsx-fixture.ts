import { strToU8, zipSync } from "fflate";

export function xlsxWithSheets(names: readonly string[]): Uint8Array {
  const sheets = names
    .map((name, index) => `<sheet name="${name}" sheetId="${index + 1}" r:id="rId${index + 1}"/>`)
    .join("");

  const workbook = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>${sheets}</sheets>
</workbook>`;

  return zipSync({ "xl/workbook.xml": strToU8(workbook) });
}
