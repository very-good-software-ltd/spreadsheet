import type { XmlReader } from "../xml/xml-reader";
import type { ZipArchive } from "../zip/zip-archive";
import type { Cell, Row } from "./cell";

export class Worksheet {
  constructor(
    private readonly archive: ZipArchive,
    private readonly xml: XmlReader,
    private readonly path: string,
    private readonly sharedStrings: readonly string[],
  ) {}

  async *rows(): AsyncIterable<Row> {
    let rowNumber = 0;
    let cells: Cell[] = [];
    let cellRef = "";
    let cellType = "";
    let valueText: string | null = null;
    let inValue = false;

    for await (const event of this.xml.read(this.archive.openStream(this.path))) {
      switch (event.type) {
        case "open":
          if (event.name === "row") {
            rowNumber = Number(event.attributes["r"] ?? "0");
            cells = [];
          } else if (event.name === "c") {
            cellRef = event.attributes["r"] ?? "";
            cellType = event.attributes["t"] ?? "";
            valueText = null;
          } else if (event.name === "v") {
            inValue = true;
            valueText = "";
          }
          break;
        case "text":
          if (inValue && valueText !== null) {
            valueText += event.text;
          }
          break;
        case "close":
          if (event.name === "v") {
            inValue = false;
          } else if (event.name === "c") {
            const cell = this.toCell(cellRef, cellType, valueText);
            if (cell !== null) {
              cells.push(cell);
            }
          } else if (event.name === "row") {
            yield { number: rowNumber, cells };
          }
          break;
      }
    }
  }

  private toCell(ref: string, type: string, valueText: string | null): Cell | null {
    if (valueText === null) {
      return null;
    }
    if (type === "s") {
      const index = Number(valueText);
      return { ref, type: "string", value: this.sharedStrings[index] ?? "" };
    }
    if (type === "" || type === "n") {
      return { ref, type: "number", value: Number(valueText) };
    }
    return null;
  }
}
