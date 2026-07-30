export interface XmlReader {
  read(source: ReadableStream<Uint8Array>): AsyncIterable<XmlEvent>;
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
