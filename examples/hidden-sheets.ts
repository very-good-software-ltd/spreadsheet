import type { Workbook } from "@very-good-software/spreadsheet";

export function visibleSheetNames(workbook: Workbook): string[] {
  return workbook.worksheets.filter((sheet) => !sheet.hidden).map((sheet) => sheet.name);
}
