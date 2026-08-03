import type { ByteRange } from "../io/byte-range";
import type { Row } from "../row";
import type { WorkbookData, WorksheetInfo } from "../workbook";
import { createXmlReader } from "../xml/create-xml-reader";
import { openZip } from "../zip/open-zip";
import type { ZipArchive } from "../zip/zip-archive";
import type { CellContext } from "./interpret-cell";
import { readSharedStrings } from "./read-shared-strings";
import { readSheetRows } from "./read-sheet";
import { readStyles } from "./read-styles";
import { readWorkbook } from "./read-workbook";

export async function readXlsx(source: ByteRange): Promise<WorkbookData> {
  let archive: ZipArchive;
  try {
    archive = await openZip(source);
  } catch (cause) {
    throw new Error("Not a valid xlsx file: could not read it as a zip", { cause });
  }

  const xml = createXmlReader();
  const { worksheets: refs, date1904 } = await readWorkbook(archive, xml);
  const sharedStrings = await readSharedStrings(archive, xml);
  const styles = await readStyles(archive, xml);
  const context: CellContext = { sharedStrings, styles, date1904 };
  const worksheets: readonly WorksheetInfo[] = refs.map((ref) => ({ name: ref.name, hidden: ref.hidden }));

  return {
    worksheets,
    openRows(index: number): AsyncIterable<Row> {
      const ref = refs[index];

      if (ref === undefined || !archive.has(ref.path)) {
        throw new Error(`Worksheet "${ref?.name ?? index}" is missing its data part in the archive`);
      }

      return readSheetRows(archive, xml, ref.path, context);
    },
  };
}
