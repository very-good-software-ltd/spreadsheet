import { describe, expect, it } from "vitest";
import { readStyles, type Styles } from "../../src/xlsx/read-styles";
import { DateStyleTable, writeStylesPart } from "../../src/xlsx/write-styles";
import { SaxesXmlReader } from "../../src/xml/saxes-xml-reader";
import { XML_DECLARATION } from "../../src/xml/write-xml";
import type { ZipArchive } from "../../src/zip/zip-archive";

const SHEET_OPEN = `<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">`;

const CELL_STYLE_XFS = `<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"></xf></cellStyleXfs>`;

function stylesXml(inner: string): string {
  return `${SHEET_OPEN}${inner}</styleSheet>`;
}

function archiveOf(styles: string | undefined): ZipArchive {
  return {
    entries: () => [],
    has: (path) => path === "xl/styles.xml" && styles !== undefined,
    read: (path) => Promise.resolve(new TextEncoder().encode(path === "xl/styles.xml" ? (styles ?? "") : "")),
    openStream: (): ReadableStream<Uint8Array> => {
      const bytes = new TextEncoder().encode(styles ?? "");
      return new ReadableStream<Uint8Array>({
        start(controller): void {
          controller.enqueue(bytes);
          controller.close();
        },
      });
    },
    storedEntry: () => {
      throw new Error("not used");
    },
  };
}

async function load(styles: string | undefined): Promise<Styles> {
  return readStyles(archiveOf(styles), new SaxesXmlReader());
}

async function rewrite(styles: string, table: DateStyleTable): Promise<string> {
  const events = new SaxesXmlReader().read(archiveOf(styles).openStream("xl/styles.xml"));
  let out = "";
  for await (const chunk of writeStylesPart(events, table)) {
    out += chunk;
  }
  return out.slice(XML_DECLARATION.length);
}

const PLAIN_AND_DATE_XFS = `<cellXfs count="3"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"></xf><xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"></xf><xf numFmtId="0" fontId="4" fillId="7" borderId="2" xfId="0" applyFont="1" applyFill="1" applyBorder="1"></xf></cellXfs>`;

describe("DateStyleTable", () => {
  it("reuses a style that already formats as a date, which is the template case", async () => {
    const table = new DateStyleTable(await load(stylesXml(CELL_STYLE_XFS + PLAIN_AND_DATE_XFS)));

    expect(table.forDate("1", false)).toBe("1");
    expect(table.changed).toBe(false);
  });

  it("adds a style when the cell has none", async () => {
    const table = new DateStyleTable(await load(stylesXml(CELL_STYLE_XFS + PLAIN_AND_DATE_XFS)));

    expect(table.forDate(undefined, false)).toBe("3");
    expect(table.changed).toBe(true);
  });

  it("adds a style when the cell's own format is not a date", async () => {
    const table = new DateStyleTable(await load(stylesXml(CELL_STYLE_XFS + PLAIN_AND_DATE_XFS)));

    expect(table.forDate("2", false)).toBe("3");
  });

  it("hands back the same style for a repeated request, so ten thousand dates add one", async () => {
    const table = new DateStyleTable(await load(stylesXml(CELL_STYLE_XFS + PLAIN_AND_DATE_XFS)));

    expect(table.forDate("2", false)).toBe("3");
    expect(table.forDate("2", false)).toBe("3");
    expect(table.forDate(undefined, false)).toBe("4");
    expect(table.forDate("2", false)).toBe("3");
  });

  it("separates a date from a date and time", async () => {
    const table = new DateStyleTable(await load(stylesXml(CELL_STYLE_XFS + PLAIN_AND_DATE_XFS)));

    expect(table.forDate("2", false)).toBe("3");
    expect(table.forDate("2", true)).toBe("4");
  });
});

describe("writeStylesPart", () => {
  it("leaves the part alone when no date needed a style", async () => {
    const source = stylesXml(CELL_STYLE_XFS + PLAIN_AND_DATE_XFS);
    const table = new DateStyleTable(await load(source));

    table.forDate("1", false);

    expect(await rewrite(source, table)).toBe(source);
  });

  it("clones the cell's format so its font, fill and border survive", async () => {
    const source = stylesXml(CELL_STYLE_XFS + PLAIN_AND_DATE_XFS);
    const table = new DateStyleTable(await load(source));

    table.forDate("2", false);
    const result = await rewrite(source, table);

    expect(result).toContain(
      `<xf numFmtId="164" fontId="4" fillId="7" borderId="2" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyNumberFormat="1"></xf>`,
    );
  });

  it("bumps the cell format count so it matches what follows", async () => {
    const source = stylesXml(CELL_STYLE_XFS + PLAIN_AND_DATE_XFS);
    const table = new DateStyleTable(await load(source));

    table.forDate("2", false);
    table.forDate(undefined, true);

    expect(await rewrite(source, table)).toContain(`<cellXfs count="5">`);
  });

  it("leaves the other xf list alone", async () => {
    const source = stylesXml(CELL_STYLE_XFS + PLAIN_AND_DATE_XFS);
    const table = new DateStyleTable(await load(source));

    table.forDate("2", false);

    expect(await rewrite(source, table)).toContain(CELL_STYLE_XFS);
  });

  it("adds a number format list when the part has none", async () => {
    const source = stylesXml(CELL_STYLE_XFS + PLAIN_AND_DATE_XFS);
    const table = new DateStyleTable(await load(source));

    table.forDate(undefined, false);
    const result = await rewrite(source, table);

    expect(result).toContain(`<numFmts count="1"><numFmt numFmtId="164" formatCode="yyyy-mm-dd"></numFmt></numFmts>`);
    expect(result.indexOf("<numFmts")).toBeLessThan(result.indexOf("<cellStyleXfs"));
  });

  it("appends to an existing number format list and takes the next free id", async () => {
    const source = stylesXml(
      `<numFmts count="2"><numFmt numFmtId="164" formatCode="0.000"></numFmt><numFmt numFmtId="170" formatCode="#,##0"></numFmt></numFmts>${CELL_STYLE_XFS}${PLAIN_AND_DATE_XFS}`,
    );
    const table = new DateStyleTable(await load(source));

    table.forDate(undefined, false);
    const result = await rewrite(source, table);

    expect(result).toContain(`<numFmts count="3">`);
    expect(result).toContain(`<numFmt numFmtId="171" formatCode="yyyy-mm-dd"></numFmt>`);
    expect(result).toContain(`<numFmt numFmtId="164" formatCode="0.000"></numFmt>`);
  });

  it("writes a format that shows the time when the value carries one", async () => {
    const source = stylesXml(CELL_STYLE_XFS + PLAIN_AND_DATE_XFS);
    const table = new DateStyleTable(await load(source));

    table.forDate(undefined, true);

    expect(await rewrite(source, table)).toContain(`formatCode="yyyy-mm-dd hh:mm:ss"`);
  });

  it("uses one number format for many cloned styles", async () => {
    const source = stylesXml(CELL_STYLE_XFS + PLAIN_AND_DATE_XFS);
    const table = new DateStyleTable(await load(source));

    table.forDate("0", false);
    table.forDate("2", false);
    const result = await rewrite(source, table);

    expect(result).toContain(`<numFmts count="1">`);
    expect(result.match(/numFmtId="164"/g)).toHaveLength(3);
  });
});
