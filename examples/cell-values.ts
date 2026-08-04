import type { Cell } from "very-good-spreadsheet";

// Matching on a cell's type gives you a type-safe way to access its value.
export function show(cell: Cell): void {
  switch (cell.type) {
    case "number":
      console.log("number", cell.value);
      break;
    case "string":
      console.log("string", cell.value);
      break;
    case "boolean":
      console.log("boolean", cell.value);
      break;
    case "date":
      console.log("date, in UTC", cell.value);
      break;
    case "formula":
      // value is the formula text, cachedValue is the result Excel stored.
      console.log("formula", cell.value, cell.cachedValue?.value);
      break;
    case "error":
      console.log("error text", cell.value);
      break;
  }
}
