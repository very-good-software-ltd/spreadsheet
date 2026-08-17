import { XML_DECLARATION } from "../xml/write-xml";
import { MemoryZipArchive } from "../zip/memory-zip-archive";
import type { ZipArchive } from "../zip/zip-archive";

// The smallest set of parts a spreadsheet application will open, so that creating
// a file from scratch is the same operation as filling a template whose template
// happens to be empty.
//
// The styles part is here even though nothing is styled yet. Writing a date has
// to put a number format somewhere, and having the part already present means
// that path never has to create one, along with its content type override and its
// relationship.

export const MAIN_NAMESPACE = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const RELATIONSHIPS_NAMESPACE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PACKAGE_RELATIONSHIPS_NAMESPACE = "http://schemas.openxmlformats.org/package/2006/relationships";
const CONTENT_TYPES_NAMESPACE = "http://schemas.openxmlformats.org/package/2006/content-types";
const OFFICE_DOCUMENT_TYPE = `${RELATIONSHIPS_NAMESPACE}/officeDocument`;
const WORKSHEET_TYPE = `${RELATIONSHIPS_NAMESPACE}/worksheet`;
const STYLES_TYPE = `${RELATIONSHIPS_NAMESPACE}/styles`;
const SPREADSHEET_TYPES = "application/vnd.openxmlformats-officedocument.spreadsheetml";

export const FIRST_WORKSHEET_NAME = "Sheet1";

const CONTENT_TYPES = `${XML_DECLARATION}<Types xmlns="${CONTENT_TYPES_NAMESPACE}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="${SPREADSHEET_TYPES}.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="${SPREADSHEET_TYPES}.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="${SPREADSHEET_TYPES}.styles+xml"/></Types>`;

const PACKAGE_RELATIONSHIPS = `${XML_DECLARATION}<Relationships xmlns="${PACKAGE_RELATIONSHIPS_NAMESPACE}"><Relationship Id="rId1" Type="${OFFICE_DOCUMENT_TYPE}" Target="xl/workbook.xml"/></Relationships>`;

const WORKBOOK = `${XML_DECLARATION}<workbook xmlns="${MAIN_NAMESPACE}" xmlns:r="${RELATIONSHIPS_NAMESPACE}"><sheets><sheet name="${FIRST_WORKSHEET_NAME}" sheetId="1" r:id="rId1"/></sheets></workbook>`;

const WORKBOOK_RELATIONSHIPS = `${XML_DECLARATION}<Relationships xmlns="${PACKAGE_RELATIONSHIPS_NAMESPACE}"><Relationship Id="rId1" Type="${WORKSHEET_TYPE}" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="${STYLES_TYPE}" Target="styles.xml"/></Relationships>`;

const WORKSHEET = `${XML_DECLARATION}<worksheet xmlns="${MAIN_NAMESPACE}"><sheetData/></worksheet>`;

// One of everything the format requires a style to point at. The second fill is
// here because Excel writes it into every file it saves and some readers assume
// a fill at index 1 exists.
const STYLES = `${XML_DECLARATION}<styleSheet xmlns="${MAIN_NAMESPACE}"><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs></styleSheet>`;

export function blankXlsxArchive(): ZipArchive {
  return MemoryZipArchive.ofText({
    "[Content_Types].xml": CONTENT_TYPES,
    "_rels/.rels": PACKAGE_RELATIONSHIPS,
    "xl/workbook.xml": WORKBOOK,
    "xl/_rels/workbook.xml.rels": WORKBOOK_RELATIONSHIPS,
    "xl/worksheets/sheet1.xml": WORKSHEET,
    "xl/styles.xml": STYLES,
  });
}
