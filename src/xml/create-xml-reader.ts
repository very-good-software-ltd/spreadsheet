import { SaxesXmlReader } from "./saxes-xml-reader";
import type { XmlReader } from "./xml-reader";

export function createXmlReader(): XmlReader {
  return new SaxesXmlReader();
}
