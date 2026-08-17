import { createWriteStream } from "node:fs";
import { Writable } from "node:stream";
import { Workbook } from "@very-good-software/spreadsheet";

// The rows are pulled as the output drains, so this holds one row at a time
// however many there are.
async function* everyOrder(): AsyncIterable<readonly [number, string, number]> {
  for (let id = 1; id <= 5_000_000; id += 1) {
    yield [id, `Order ${id}`, id * 1.5];
  }
}

const editor = (await Workbook.create()).edit();
editor.worksheet(0).appendRows(everyOrder());

await editor.save().pipeTo(Writable.toWeb(createWriteStream("orders.xlsx")));
