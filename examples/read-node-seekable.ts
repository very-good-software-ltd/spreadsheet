import { openAsBlob } from "node:fs";
import { Workbook } from "@very-good-software/spreadsheet";

// openAsBlob gives a seekable Blob backed by the file on disk, so a large file
// is read in ranges and never held whole.
const workbook = await Workbook.open(await openAsBlob("data.xlsx"));

console.log(workbook.worksheetNames);
