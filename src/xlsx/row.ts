import type { Cell } from "./cell";

export class Row {
  constructor(
    readonly number: number,
    readonly cells: readonly Cell[],
  ) {}

  cell(column: number): Cell | undefined {
    return this.cells.find((candidate) => candidate.columnIndex === column);
  }
}
