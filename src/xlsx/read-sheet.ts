import type { Cell } from "../cell";
import { readPart } from "../read-part";
import { Row } from "../row";
import type { XmlReader } from "../xml/xml-reader";
import type { ZipArchive } from "../zip/zip-archive";
import { type CellContext, interpretCellValue } from "./interpret-cell";

const Element = {
  Row: "row",
  Cell: "c",
  Value: "v",
  Text: "t",
  Formula: "f",
} as const;

const Attribute = {
  Reference: "r",
  Type: "t",
  Style: "s",
} as const;

export async function* readSheetRows(
  archive: ZipArchive,
  xml: XmlReader,
  path: string,
  context: CellContext,
): AsyncIterable<Row> {
  let rowNumber = 0;
  let cells: Cell[] = [];
  let cellRef = "";
  let cellTypeCode = "";
  let cellStyleIndex: number | undefined;
  let valueText: string | null = null;
  let formulaText: string | null = null;
  let capturing = false;
  let capturingFormula = false;

  for await (const batch of readPart(archive, xml, path)) {
    for (const event of batch) {
      switch (event.type) {
        case "open":
          if (event.name === Element.Row) {
            rowNumber = Number(event.attributes[Attribute.Reference] ?? "0");
            cells = [];
          } else if (event.name === Element.Cell) {
            cellRef = event.attributes[Attribute.Reference] ?? "";
            cellTypeCode = event.attributes[Attribute.Type] ?? "";
            cellStyleIndex = styleIndexOf(event.attributes[Attribute.Style]);
            valueText = null;
            formulaText = null;
          } else if (event.name === Element.Formula) {
            capturingFormula = true;
            formulaText ??= "";
          } else if (event.name === Element.Value || event.name === Element.Text) {
            capturing = true;
            valueText ??= "";
          }
          break;
        case "text":
          if (capturingFormula && formulaText !== null) {
            formulaText += event.text;
          } else if (capturing && valueText !== null) {
            valueText += event.text;
          }
          break;
        case "close":
          if (event.name === Element.Formula) {
            capturingFormula = false;
          } else if (event.name === Element.Value || event.name === Element.Text) {
            capturing = false;
          } else if (event.name === Element.Cell) {
            const cached =
              valueText === null ? null : interpretCellValue(cellRef, cellTypeCode, cellStyleIndex, valueText, context);
            // A formula cell keeps its text and Excel's cached result. The text
            // is taken as stored, without a leading "=", and a shared-formula
            // dependent carries none, so its value is empty. A formula with no
            // cached <v> keeps a null cachedValue rather than being dropped.
            if (formulaText !== null) {
              cells.push({
                ref: cellRef,
                columnIndex: columnIndexOf(cellRef),
                type: "formula",
                value: formulaText,
                cachedValue: cached,
              });
            } else if (cached !== null) {
              cells.push({ ref: cellRef, columnIndex: columnIndexOf(cellRef), ...cached });
            }
          } else if (event.name === Element.Row) {
            yield new Row(rowNumber, cells);
          }
          break;
      }
    }
  }
}

function styleIndexOf(attribute: string | undefined): number | undefined {
  return attribute === undefined ? undefined : Number(attribute);
}

// A cell reference like "C1" starts with the column letters. Convert them from
// base-26 (A=1) to a zero-based index, so "A" is 0 and "AA" is 26.
function columnIndexOf(ref: string): number {
  let index = 0;
  for (const char of ref) {
    const code = char.charCodeAt(0);
    if (code < 65 || code > 90) {
      break;
    }
    index = index * 26 + (code - 64);
  }
  return index - 1;
}
