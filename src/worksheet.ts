import type { Row } from "./row";

export class Worksheet {
  constructor(private readonly open: () => AsyncIterable<Row>) {}

  rows(): AsyncIterable<Row> {
    return this.open();
  }
}
