import type { XmlEvent } from "./xml-reader";

export const XML_DECLARATION = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n`;

export function writeXmlEvent(event: XmlEvent): string {
  switch (event.type) {
    case "open": {
      const attributes = Object.entries(event.attributes)
        .map(([name, value]) => ` ${name}="${escapeAttribute(value)}"`)
        .join("");
      return `<${event.name}${attributes}>`;
    }
    case "text":
      return escapeText(event.text);
    case "close":
      return `</${event.name}>`;
  }
}

function escapeText(text: string): string {
  return text.replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll(">", "&gt;");
}

// A parser folds a literal tab, newline or carriage return in an attribute value
// into a space before the value reaches the caller, so writing one loses it.
// The numeric references survive that normalisation and read back as written.
function escapeAttribute(value: string): string {
  return escapeText(value)
    .replaceAll('"', "&quot;")
    .replaceAll("\t", "&#9;")
    .replaceAll("\n", "&#10;")
    .replaceAll("\r", "&#13;");
}
