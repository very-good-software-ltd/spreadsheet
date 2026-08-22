import { readFileSync } from "node:fs";
import { strFromU8, unzipSync } from "fflate";
import { describe, expect, it } from "vitest";
import { Workbook } from "../../src/workbook";

// Written by real Excel, because no library we depend on can write a pivot table.
// Its `Data` sheet holds five rows under a header, `Movements` names the five, and
// the pivot on the `Pivot` sheet reads `Data!A1:C6` through a cache holding its own
// copy of the rows.
const TEMPLATE = new Uint8Array(readFileSync(new URL("../fixtures/pivot-template.xlsx", import.meta.url)));

const CACHE_PART = "xl/pivotCache/pivotCacheDefinition1.xml";
const PIVOT_PART = "xl/pivotTables/pivotTable1.xml";

async function bytesOf(stream: ReadableStream<Uint8Array>): Promise<Uint8Array> {
  const chunks: Uint8Array[] = [];
  for await (const chunk of stream as unknown as AsyncIterable<Uint8Array>) {
    chunks.push(chunk);
  }
  return new Uint8Array(chunks.flatMap((chunk) => [...chunk]));
}

async function filledWith(rows: readonly (readonly (string | number)[])[]): Promise<Uint8Array> {
  const editor = (await Workbook.open(TEMPLATE)).edit();
  editor.writeRegion("Movements", rows);

  return bytesOf(editor.save());
}

function partOf(bytes: Uint8Array, path: string): string {
  return strFromU8(unzipSync(bytes)[path] ?? new Uint8Array());
}

const sourceRefIn = (bytes: Uint8Array): string | undefined =>
  /<worksheetSource[^>]*\bref="([^"]+)"/.exec(partOf(bytes, CACHE_PART))?.[1];

const locationRefIn = (bytes: Uint8Array): string | undefined =>
  /<location[^>]*\bref="([^"]+)"/.exec(partOf(bytes, PIVOT_PART))?.[1];

const THREE_ROWS = [
  ["North", "January", 100],
  ["South", "January", 200],
  ["East", "January", 300],
] as const;

const EIGHT_ROWS = Array.from({ length: 8 }, (_, index) => ["North", "January", index * 10] as const);

describe("filling a region a pivot table reads from", () => {
  it("writes the file instead of refusing", async () => {
    await expect(filledWith(THREE_ROWS)).resolves.toBeInstanceOf(Uint8Array);
  });

  it("closes the cache's source range up when rows go away", async () => {
    expect(sourceRefIn(await filledWith(THREE_ROWS))).toBe("A1:C4");
  });

  it("stretches the cache's source range when rows arrive", async () => {
    expect(sourceRefIn(await filledWith(EIGHT_ROWS))).toBe("A1:C9");
  });

  it("asks the application to rebuild the cache on open", async () => {
    expect(partOf(await filledWith(THREE_ROWS), CACHE_PART)).toContain('refreshOnLoad="1"');
  });

  it("leaves the pivot where it is, since it is on another sheet", async () => {
    expect(locationRefIn(await filledWith(THREE_ROWS))).toBe("A3:B7");
  });

  it("keeps the cached records, which the rebuild replaces", async () => {
    expect(partOf(await filledWith(THREE_ROWS), "xl/pivotCache/pivotCacheRecords1.xml")).toContain(
      "<pivotCacheRecords",
    );
  });
});
