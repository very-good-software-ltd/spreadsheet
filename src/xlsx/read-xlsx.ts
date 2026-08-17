import type { Editor } from "../editor";
import type { Row } from "../row";
import type { WorkbookData, WorksheetInfo } from "../workbook";
import { createXmlReader } from "../xml/create-xml-reader";
import type { ZipArchive } from "../zip/zip-archive";
import { XlsxEditor } from "./edit-xlsx";
import type { CellContext } from "./interpret-cell";
import { readSharedStrings } from "./read-shared-strings";
import { readSheetRows } from "./read-sheet";
import { readStyles } from "./read-styles";
import { readWorkbook } from "./read-workbook";

export async function readXlsx(archive: ZipArchive): Promise<WorkbookData> {
  const xml = createXmlReader();
  const info = await readWorkbook(archive, xml);
  const { worksheets: refs, date1904 } = info;
  const sharedStrings = await readSharedStrings(archive, xml);
  const styles = await readStyles(archive, xml);
  const context: CellContext = { sharedStrings, styles, date1904 };
  const worksheets: readonly WorksheetInfo[] = refs.map((ref) => ({ name: ref.name, hidden: ref.hidden }));

  return {
    worksheets,
    edit(): Editor {
      return new XlsxEditor(archive, info, styles);
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
