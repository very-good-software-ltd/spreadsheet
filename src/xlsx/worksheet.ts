import type { XmlReader } from "../xml/xml-reader";
import type { ZipArchive } from "../zip/zip-archive";
import type { Cell, Row } from "./cell";
import { type CellContext, interpretCell } from "./interpret-cell";

const Element = {
  Row: "row",
  Cell: "c",
  Value: "v",
  Text: "t",
} as const;

const Attribute = {
  Reference: "r",
  Type: "t",
  Style: "s",
} as const;

export class Worksheet {
  constructor(
    private readonly archive: ZipArchive,
    private readonly xml: XmlReader,
    private readonly path: string,
    private readonly context: CellContext,
  ) {}

  async *rows(): AsyncIterable<Row> {
    let rowNumber = 0;
    let cells: Cell[] = [];
    let cellRef = "";
    let cellTypeCode = "";
    let cellStyleIndex: number | undefined;
    let valueText: string | null = null;
    let capturing = false;

    for await (const event of this.xml.read(this.archive.openStream(this.path))) {
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
          } else if (event.name === Element.Value || event.name === Element.Text) {
            capturing = true;
            valueText ??= "";
          }
          break;
        case "text":
          if (capturing && valueText !== null) {
            valueText += event.text;
          }
          break;
        case "close":
          if (event.name === Element.Value || event.name === Element.Text) {
            capturing = false;
          } else if (event.name === Element.Cell) {
            if (valueText !== null) {
              cells.push(interpretCell(cellRef, cellTypeCode, cellStyleIndex, valueText, this.context));
            }
          } else if (event.name === Element.Row) {
            yield { number: rowNumber, cells };
          }
          break;
      }
    }
  }
}

function styleIndexOf(attribute: string | undefined): number | undefined {
  return attribute === undefined ? undefined : Number(attribute);
}
