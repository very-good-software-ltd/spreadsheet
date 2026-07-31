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
    let capturing = false;

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
          } else if (event.name === "v" || event.name === "t") {
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
          if (event.name === "v" || event.name === "t") {
            capturing = false;
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
    switch (type) {
      case "":
      case "n":
        return { ref, type: "number", value: Number(valueText) };
      case "s":
        return { ref, type: "string", value: this.sharedStrings[Number(valueText)] ?? "" };
      case "str":
      case "inlineStr":
        return { ref, type: "string", value: valueText };
      case "b":
        return { ref, type: "boolean", value: valueText !== "0" };
      case "e":
        return { ref, type: "error", value: valueText };
      default:
        throw new Error(`Unsupported cell type "${type}" at ${ref}`);
    }
  }
}
