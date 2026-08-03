import { describe, expect, it } from "vitest";
import { SaxesXmlReader } from "../src/xml/saxes-xml-reader";
import type { XmlEvent } from "../src/xml/xml-reader";

function streamOf(text: string): ReadableStream<Uint8Array> {
  const bytes = new TextEncoder().encode(text);
  return new ReadableStream<Uint8Array>({
    start(controller): void {
      controller.enqueue(bytes);
      controller.close();
    },
  });
}

async function collect(source: ReadableStream<Uint8Array>): Promise<XmlEvent[]> {
  const events: XmlEvent[] = [];
  for await (const batch of new SaxesXmlReader().read(source)) {
    events.push(...batch);
  }
  return events;
}

describe("SaxesXmlReader", () => {
  it("emits open, text and close events with attributes", async () => {
    const events = await collect(streamOf('<row r="1"><c>7</c></row>'));

    expect(events).toEqual([
      { type: "open", name: "row", attributes: { r: "1" } },
      { type: "open", name: "c", attributes: {} },
      { type: "text", text: "7" },
      { type: "close", name: "c" },
      { type: "close", name: "row" },
    ]);
  });
});
