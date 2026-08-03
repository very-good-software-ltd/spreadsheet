export interface XmlReader {
  // Yields the events from one input chunk at a time, so a caller crosses the
  // async boundary once per chunk and then loops the batch, not once per event.
  read(source: ReadableStream<Uint8Array>): AsyncIterable<readonly XmlEvent[]>;
}

export type XmlEvent = XmlOpen | XmlText | XmlClose;

export interface XmlOpen {
  readonly type: "open";
  readonly name: string;
  readonly attributes: Readonly<Record<string, string>>;
}

export interface XmlText {
  readonly type: "text";
  readonly text: string;
}

export interface XmlClose {
  readonly type: "close";
  readonly name: string;
}
