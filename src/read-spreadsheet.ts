import type { ByteRange } from "./io/byte-range";
import type { WorkbookData } from "./workbook";
import { readXlsx } from "./xlsx/read-xlsx";
import { openZip } from "./zip/open-zip";
import type { ZipArchive } from "./zip/zip-archive";

export async function readSpreadsheet(source: ByteRange): Promise<WorkbookData> {
  let archive: ZipArchive;
  try {
    archive = await openZip(source);
  } catch (cause) {
    throw new Error("Not a valid xlsx file: could not read it as a zip", { cause });
  }

  return readXlsx(archive);
}
