import type { Editor } from "../editor";
import type { Row } from "../row";
import type { WorkbookData, WorksheetInfo } from "../workbook";
import { createXmlReader } from "../xml/create-xml-reader";
import type { XmlReader } from "../xml/xml-reader";
import type { ZipArchive } from "../zip/zip-archive";
import { XlsxEditor } from "./edit-xlsx";
import type { CellContext } from "./interpret-cell";
import { readSharedStrings } from "./read-shared-strings";
import { readSheetRows } from "./read-sheet";
import { readStyles } from "./read-styles";
import { readTables, type TableOnSheet } from "./read-tables";
import { readWorkbook, type WorksheetRef } from "./read-workbook";

export async function readXlsx(archive: ZipArchive): Promise<WorkbookData> {
  const xml = createXmlReader();
  const info = await readWorkbook(archive, xml);
  const { worksheets: refs, date1904 } = info;
  const sharedStrings = await readSharedStrings(archive, xml);
  const styles = await readStyles(archive, xml);
  const context: CellContext = { sharedStrings, styles, date1904 };
  const tables = await readAllTables(archive, xml, refs);
  const worksheets: readonly WorksheetInfo[] = refs.map((ref) => ({ name: ref.name, hidden: ref.hidden }));

  return {
    worksheets,
    edit(): Editor {
      return new XlsxEditor(archive, info, styles, tables);
    },
    openRows(index: number): AsyncIterable<Row> {
      const ref = refs[index];

      if (ref === undefined || !archive.has(ref.path)) {
        throw new Error(`Worksheet "${ref?.name ?? index}" is missing its data part in the archive`);
      }

      return readSheetRows(archive, xml, ref.path, context);
    },
  };
}

// Read at open rather than at edit, because addressing a table by name is checked
// as the call is made and the editor is handed out synchronously. A sheet with no
// relationships part costs only a lookup in the archive's directory, which is what
// most sheets are.
async function readAllTables(
  archive: ZipArchive,
  xml: XmlReader,
  refs: readonly WorksheetRef[],
): Promise<readonly TableOnSheet[]> {
  const tables: TableOnSheet[] = [];

  for (const ref of refs) {
    for (const table of await readTables(archive, xml, ref.path)) {
      tables.push({ ...table, sheet: ref.name });
    }
  }

  return tables;
}
