import { strToU8, zipSync } from "fflate";
import { describe, expect, it } from "vitest";
import { readStyles } from "../src/xlsx/read-styles";
import { SaxesXmlReader } from "../src/xml/saxes-xml-reader";
import { FflateZipArchive } from "../src/zip/fflate-zip-archive";

async function stylesFrom(styleSheet: string) {
  const bytes = zipSync({ "xl/styles.xml": strToU8(styleSheet) });
  return readStyles(new FflateZipArchive(bytes), new SaxesXmlReader());
}

describe("readStyles", () => {
  it("treats a built-in date format id as a date style", async () => {
    const styles = await stylesFrom(
      '<styleSheet><cellXfs><xf numFmtId="0"/><xf numFmtId="14"/></cellXfs></styleSheet>',
    );

    expect(styles.isDateStyle(0)).toBe(false);
    expect(styles.isDateStyle(1)).toBe(true);
  });

  it("treats a custom format with date tokens as a date style", async () => {
    const styles = await stylesFrom(
      '<styleSheet><numFmts><numFmt numFmtId="164" formatCode="yyyy-mm-dd"/><numFmt numFmtId="165" formatCode="0.00"/></numFmts><cellXfs><xf numFmtId="164"/><xf numFmtId="165"/></cellXfs></styleSheet>',
    );

    expect(styles.isDateStyle(0)).toBe(true);
    expect(styles.isDateStyle(1)).toBe(false);
  });

  it("ignores xf elements outside cellXfs", async () => {
    const styles = await stylesFrom(
      '<styleSheet><cellStyleXfs><xf numFmtId="14"/></cellStyleXfs><cellXfs><xf numFmtId="0"/></cellXfs></styleSheet>',
    );

    expect(styles.isDateStyle(0)).toBe(false);
  });
});
