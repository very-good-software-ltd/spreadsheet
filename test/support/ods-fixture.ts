import { strToU8, zipSync } from "fflate";

const NAMESPACES = [
  'xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"',
  'xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"',
  'xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"',
].join(" ");

// Wraps hand-authored table markup in a minimal ODF spreadsheet package. The
// markup mirrors what LibreOffice writes, so the repeat and padding quirks can
// be exercised precisely.
export function odsWith(tables: string): Uint8Array {
  const content = `<office:document-content ${NAMESPACES}><office:body><office:spreadsheet>${tables}</office:spreadsheet></office:body></office:document-content>`;
  return zipSync({
    mimetype: strToU8("application/vnd.oasis.opendocument.spreadsheet"),
    "content.xml": strToU8(content),
  });
}
