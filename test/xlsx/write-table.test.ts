import { describe, expect, it } from "vitest";
import { withTableExtent } from "../../src/xlsx/write-table";
import type { XmlEvent } from "../../src/xml/xml-reader";

async function* eventsOf(events: readonly XmlEvent[]): AsyncIterable<XmlEvent> {
  yield* events;
}

async function refsAfter(events: readonly XmlEvent[], lastRow: number): Promise<(string | undefined)[]> {
  const refs: (string | undefined)[] = [];
  for await (const event of withTableExtent(eventsOf(events), lastRow)) {
    if (event.type === "open") {
      refs.push(event.attributes["ref"]);
    }
  }
  return refs;
}

const TABLE: readonly XmlEvent[] = [
  { type: "open", name: "table", attributes: { ref: "B2:D6", displayName: "Sales" } },
  { type: "open", name: "autoFilter", attributes: { ref: "B2:D6" } },
  { type: "close", name: "autoFilter" },
  { type: "open", name: "tableColumns", attributes: { count: "3" } },
  { type: "close", name: "tableColumns" },
  { type: "close", name: "table" },
];

describe("withTableExtent", () => {
  it("moves the table's last row without moving its first row or its columns", async () => {
    expect(await refsAfter(TABLE, 20)).toEqual(["B2:D20", "B2:D20", undefined]);
  });

  it("pulls the last row back when the table ends up smaller", async () => {
    expect(await refsAfter(TABLE, 4)).toEqual(["B2:D4", "B2:D4", undefined]);
  });

  it("leaves everything else about the table as it was", async () => {
    const kept: XmlEvent[] = [];
    for await (const event of withTableExtent(eventsOf(TABLE), 20)) {
      kept.push(event);
    }

    expect(kept).toHaveLength(TABLE.length);
    expect(kept[0]).toMatchObject({ attributes: { displayName: "Sales" } });
    expect(kept[3]).toEqual({ type: "open", name: "tableColumns", attributes: { count: "3" } });
  });

  // A table one row tall is a header and nothing else, which is what a ref with the
  // same row at both ends spells.
  it("copes with a table whose extent is a single row", async () => {
    const single: readonly XmlEvent[] = [
      { type: "open", name: "table", attributes: { ref: "B2:D2" } },
      { type: "close", name: "table" },
    ];

    expect(await refsAfter(single, 7)).toEqual(["B2:D7"]);
  });
});
