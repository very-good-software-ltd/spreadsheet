import type { ByteRange } from "./io/byte-range";
import { readOds } from "./ods/read-ods";
import type { WorkbookData } from "./workbook";
import { readXlsx } from "./xlsx/read-xlsx";
import { openZip } from "./zip/open-zip";
import type { ZipArchive } from "./zip/zip-archive";

const XLSX_WORKBOOK_PART = "xl/workbook.xml";
const ODS_MIMETYPE_PART = "mimetype";
const ODS_MIMETYPE = "application/vnd.oasis.opendocument.spreadsheet";

export async function readSpreadsheet(source: ByteRange): Promise<WorkbookData> {
  let archive: ZipArchive;
  try {
    archive = await openZip(source);
  } catch (cause) {
    throw new Error("Not a valid spreadsheet: could not read it as a zip", { cause });
  }

  if (archive.has(XLSX_WORKBOOK_PART)) {
    return readXlsx(archive);
  }
  if (await isOds(archive)) {
    return readOds(archive);
  }
  throw new Error("Not a valid spreadsheet: not an xlsx or ods file");
}

async function isOds(archive: ZipArchive): Promise<boolean> {
  if (!archive.has(ODS_MIMETYPE_PART)) {
    return false;
  }
  return new TextDecoder().decode(await archive.read(ODS_MIMETYPE_PART)) === ODS_MIMETYPE;
}
