import type { XmlEvent } from "./xml-reader";

export const XML_DECLARATION = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n`;

export function writeXmlEvent(event: XmlEvent): string {
  switch (event.type) {
    case "open": {
      // Built by appending rather than by mapping and joining, which would allocate
      // two arrays for every tag. A sheet has several tags per cell.
      let out = `<${event.name}`;
      for (const name of Object.keys(event.attributes)) {
        out += ` ${name}="${escapeAttribute(event.attributes[name] as string)}"`;
      }
      return `${out}>`;
    }
    case "text":
      return escapeText(event.text);
    case "close":
      return `</${event.name}>`;
  }
}

// Almost nothing written to a spreadsheet needs escaping: a number never does and
// most text does not either. Testing once and returning the string untouched
// avoids several full scans per cell, which at a few million cells is the
// difference between this being free and it dominating the write.
const NEEDS_TEXT_ESCAPE = /[&<>]/;
const NEEDS_ATTRIBUTE_ESCAPE = /["&<>\t\n\r]/;

function escapeText(text: string): string {
  if (!NEEDS_TEXT_ESCAPE.test(text)) {
    return text;
  }
  return text.replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll(">", "&gt;");
}

// A parser folds a literal tab, newline or carriage return in an attribute value
// into a space before the value reaches the caller, so writing one loses it.
// The numeric references survive that normalisation and read back as written.
function escapeAttribute(value: string): string {
  if (!NEEDS_ATTRIBUTE_ESCAPE.test(value)) {
    return value;
  }
  return escapeText(value)
    .replaceAll('"', "&quot;")
    .replaceAll("\t", "&#9;")
    .replaceAll("\n", "&#10;")
    .replaceAll("\r", "&#13;");
}
