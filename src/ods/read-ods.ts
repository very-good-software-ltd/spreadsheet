import type { Cell, CellValue } from "../cell";
import { cellReference } from "../cell-reference";
import { readPart } from "../read-part";
import { Row } from "../row";
import type { WorkbookData, WorksheetInfo } from "../workbook";
import { createXmlReader } from "../xml/create-xml-reader";
import type { XmlReader } from "../xml/xml-reader";
import type { ZipArchive } from "../zip/zip-archive";
import { cellFrom, emptyCell, interpretOdsCell, isEmpty, type OdsCell } from "./interpret-cell";

// ODF stores the whole spreadsheet in one content.xml, so both listing sheets
// and reading a sheet stream that part. It run-length encodes repeats: an empty
// row is padded to a huge table:number-rows-repeated and a row's trailing empty
// cells to table:number-columns-repeated, so the reader expands non-empty runs
// and skips empty ones rather than materialising the padding.

const Element = {
  Spreadsheet: "office:spreadsheet",
  Table: "table:table",
  Row: "table:table-row",
  Cell: "table:table-cell",
  CoveredCell: "table:covered-table-cell",
  Text: "text:p",
} as const;

const Attribute = {
  Name: "table:name",
  Visibility: "table:visibility",
  RowsRepeated: "table:number-rows-repeated",
  ColumnsRepeated: "table:number-columns-repeated",
} as const;

const CONTENT_PART = "content.xml";
const VISIBLE = "visible";

export async function readOds(archive: ZipArchive): Promise<WorkbookData> {
  const xml = createXmlReader();
  const worksheets = await readOdsWorksheets(archive, xml);

  return {
    worksheets,
    openRows(index: number): AsyncIterable<Row> {
      return readOdsSheet(archive, xml, index);
    },
  };
}

async function readOdsWorksheets(archive: ZipArchive, xml: XmlReader): Promise<readonly WorksheetInfo[]> {
  const worksheets: WorksheetInfo[] = [];
  let inSpreadsheet = false;

  for await (const batch of readPart(archive, xml, CONTENT_PART)) {
    for (const event of batch) {
      if (event.type === "open" && event.name === Element.Spreadsheet) {
        inSpreadsheet = true;
      } else if (event.type === "close" && event.name === Element.Spreadsheet) {
        inSpreadsheet = false;
      } else if (inSpreadsheet && event.type === "open" && event.name === Element.Table) {
        const visibility = event.attributes[Attribute.Visibility];
        worksheets.push({
          name: event.attributes[Attribute.Name] ?? "",
          hidden: visibility !== undefined && visibility !== VISIBLE,
        });
      }
    }
  }

  return worksheets;
}

async function* readOdsSheet(archive: ZipArchive, xml: XmlReader, index: number): AsyncIterable<Row> {
  let inSpreadsheet = false;
  let tableIndex = -1;
  let active = false;
  let rowNumber = 0;

  let partials: { readonly columnIndex: number; readonly value: CellValue }[] = [];
  let columnIndex = 0;
  let rowsRepeated = 1;

  let cell: OdsCell = emptyCell();
  let columnsRepeated = 1;
  let capturingText = false;

  for await (const batch of readPart(archive, xml, CONTENT_PART)) {
    for (const event of batch) {
      switch (event.type) {
        case "open":
          if (event.name === Element.Spreadsheet) {
            inSpreadsheet = true;
          } else if (inSpreadsheet && event.name === Element.Table) {
            tableIndex += 1;
            active = tableIndex === index;
            rowNumber = 0;
          } else if (active && event.name === Element.Row) {
            partials = [];
            columnIndex = 0;
            rowsRepeated = repeatOf(event.attributes[Attribute.RowsRepeated]);
          } else if (active && (event.name === Element.Cell || event.name === Element.CoveredCell)) {
            cell = cellFrom(event.attributes, event.name === Element.CoveredCell);
            columnsRepeated = repeatOf(event.attributes[Attribute.ColumnsRepeated]);
            capturingText = false;
          } else if (active && event.name === Element.Text) {
            capturingText = true;
          }
          break;
        case "text":
          if (capturingText) {
            cell.text += event.text;
          }
          break;
        case "close":
          if (event.name === Element.Text) {
            capturingText = false;
          } else if (active && (event.name === Element.Cell || event.name === Element.CoveredCell)) {
            if (isEmpty(cell)) {
              columnIndex += columnsRepeated;
            } else {
              const value = interpretOdsCell(cell);
              for (let i = 0; i < columnsRepeated; i += 1) {
                partials.push({ columnIndex: columnIndex + i, value });
              }
              columnIndex += columnsRepeated;
            }
          } else if (active && event.name === Element.Row) {
            if (partials.length === 0) {
              rowNumber += rowsRepeated;
            } else {
              for (let r = 0; r < rowsRepeated; r += 1) {
                rowNumber += 1;
                yield new Row(rowNumber, materialise(partials, rowNumber));
              }
            }
          } else if (active && event.name === Element.Table) {
            return;
          } else if (event.name === Element.Spreadsheet) {
            inSpreadsheet = false;
          }
          break;
      }
    }
  }
}

function materialise(
  partials: readonly { readonly columnIndex: number; readonly value: CellValue }[],
  row: number,
): Cell[] {
  return partials.map((partial) => ({
    ref: cellReference(row, partial.columnIndex),
    columnIndex: partial.columnIndex,
    ...partial.value,
  }));
}

function repeatOf(attribute: string | undefined): number {
  return attribute === undefined ? 1 : Number(attribute);
}
