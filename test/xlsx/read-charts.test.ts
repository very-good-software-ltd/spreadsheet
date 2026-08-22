import { strToU8, zipSync } from "fflate";
import { describe, expect, it } from "vitest";
import { BytesByteRange } from "../../src/io/byte-range";
import { readChartPaths } from "../../src/xlsx/read-charts";
import { createXmlReader } from "../../src/xml/create-xml-reader";
import { openZip } from "../../src/zip/open-zip";

const CHART = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml";
const DRAWING = "application/vnd.openxmlformats-officedocument.drawing+xml";

async function chartsIn(overrides: readonly (readonly [string, string])[]): Promise<readonly string[]> {
  const declared = overrides.map(([part, type]) => `<Override PartName="${part}" ContentType="${type}"/>`).join("");
  const files: Record<string, Uint8Array> = {
    "[Content_Types].xml": strToU8(
      `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/>${declared}</Types>`,
    ),
  };

  for (const [part] of overrides) {
    files[part.slice(1)] = strToU8("<x/>");
  }

  return readChartPaths(await openZip(new BytesByteRange(zipSync(files))), createXmlReader());
}

describe("readChartPaths", () => {
  it("finds every chart the package declares", async () => {
    const charts = await chartsIn([
      ["/xl/charts/chart1.xml", CHART],
      ["/xl/drawings/drawing1.xml", DRAWING],
      ["/xl/charts/chart2.xml", CHART],
    ]);

    expect(charts).toEqual(["xl/charts/chart1.xml", "xl/charts/chart2.xml"]);
  });

  // A chart on its own tab hangs off a chartsheet rather than off a worksheet, so
  // walking out from the sheet whose rows moved would never reach it.
  it("finds a chart no worksheet points at", async () => {
    const charts = await chartsIn([["/xl/charts/chart1.xml", CHART]]);

    expect(charts).toEqual(["xl/charts/chart1.xml"]);
  });

  it("finds none in a workbook with no charts", async () => {
    expect(await chartsIn([["/xl/drawings/drawing1.xml", DRAWING]])).toEqual([]);
  });

  it("leaves out a chart the package declares but does not hold", async () => {
    const files = {
      "[Content_Types].xml": strToU8(
        `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/xl/charts/chart1.xml" ContentType="${CHART}"/></Types>`,
      ),
    };

    expect(await readChartPaths(await openZip(new BytesByteRange(zipSync(files))), createXmlReader())).toEqual([]);
  });
});
