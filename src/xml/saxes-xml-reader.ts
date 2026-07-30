import { SaxesParser } from "saxes";
import type { XmlEvent, XmlReader } from "./xml-reader";

export class SaxesXmlReader implements XmlReader {
  async *read(source: ReadableStream<Uint8Array>): AsyncIterable<XmlEvent> {
    const parser = new SaxesParser();
    const pending: XmlEvent[] = [];

    parser.on("opentag", (tag) => {
      pending.push({
        type: "open",
        name: tag.name,
        attributes: { ...tag.attributes },
      });
    });
    parser.on("text", (text) => {
      pending.push({ type: "text", text });
    });
    parser.on("closetag", (tag) => {
      pending.push({ type: "close", name: tag.name });
    });

    const decoder = new TextDecoder();
    const reader = source.getReader();
    try {
      for (;;) {
        const { done, value } = await reader.read();
        if (done) {
          break;
        }
        parser.write(decoder.decode(value, { stream: true }));
        yield* pending.splice(0);
      }
      parser.close();
      yield* pending.splice(0);
    } finally {
      reader.releaseLock();
    }
  }
}
