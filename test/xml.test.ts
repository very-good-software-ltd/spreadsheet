import { describe, expect, it } from "vitest";
import { SaxesXmlReader } from "../src/xml/saxes-xml-reader";
import { writeXmlEvent, XML_DECLARATION } from "../src/xml/write-xml";
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

describe("writeXmlEvent", () => {
  it("writes an open tag with its attributes", () => {
    expect(writeXmlEvent({ type: "open", name: "c", attributes: { r: "A1", t: "n" } })).toBe('<c r="A1" t="n">');
  });

  it("writes an open tag with no attributes", () => {
    expect(writeXmlEvent({ type: "open", name: "row", attributes: {} })).toBe("<row>");
  });

  it("writes a close tag", () => {
    expect(writeXmlEvent({ type: "close", name: "row" })).toBe("</row>");
  });

  it("escapes the markup characters in text", () => {
    expect(writeXmlEvent({ type: "text", text: `a & b < c > d "e" 'f'` })).toBe("a &amp; b &lt; c &gt; d \"e\" 'f'");
  });

  it("escapes the quote as well in an attribute value", () => {
    expect(writeXmlEvent({ type: "open", name: "c", attributes: { r: `a"b&c<d` } })).toBe(
      '<c r="a&quot;b&amp;c&lt;d">',
    );
  });

  it("escapes whitespace in an attribute value that a parser would otherwise fold into a space", () => {
    expect(writeXmlEvent({ type: "open", name: "t", attributes: { v: "a\nb\rc\td" } })).toBe(
      '<t v="a&#10;b&#13;c&#9;d">',
    );
  });

  it("leaves that whitespace alone in text, where it survives as written", () => {
    expect(writeXmlEvent({ type: "text", text: "a\nb\tc" })).toBe("a\nb\tc");
  });

  it("round-trips through the reader", async () => {
    const written = [
      { type: "open", name: "row", attributes: { r: "1" } },
      { type: "open", name: "c", attributes: { r: "A1" } },
      { type: "text", text: "a & b" },
      { type: "close", name: "c" },
      { type: "close", name: "row" },
    ] satisfies XmlEvent[];

    expect(await collect(streamOf(written.map(writeXmlEvent).join("")))).toEqual(written);
  });

  it("declares the version and encoding Excel expects", () => {
    expect(XML_DECLARATION).toBe(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n`);
  });
});
