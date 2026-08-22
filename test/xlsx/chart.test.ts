import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import { describe, expect, it } from "vitest";
import { Workbook } from "../../src/workbook";
import { type SheetInput, xlsx } from "../support/xlsx-fixture";

const CHART_PART = "xl/charts/chart1.xml";
const CONTENT_TYPES_PART = "[Content_Types].xml";
const CHART_TYPE = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml";

const HEADER = ["Item", "Qty", "Amount"];
const FIVE_ROWS = [HEADER, ...[1, 2, 3, 4, 5].map((n) => [`Row ${n}`, n, 100])];

const SHEETS: readonly SheetInput[] = [
  { name: "Data", rows: FIVE_ROWS },
  { name: "Elsewhere", rows: FIVE_ROWS },
];

// A series as Excel writes one, with its name from a header cell above the region
// and its values from a column of it.
function chartReading(...references: readonly string[]): string {
  const series = references
    .map(
      (reference, index) =>
        `<c:ser><c:idx val="${index}"/><c:val><c:numRef><c:f>${reference}</c:f><c:numCache><c:ptCount val="5"/></c:numCache></c:numRef></c:val></c:ser>`,
    )
    .join("");

  return `<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:plotArea><c:barChart>${series}</c:barChart></c:plotArea></c:chart></c:chartSpace>`;
}

function withChart(chart: string, ...names: readonly { name: string; target: string }[]): Uint8Array {
  const files = unzipSync(xlsx(SHEETS, { definedNames: names }));

  files[CHART_PART] = strToU8(chart);
  // The fixture builder writes no content types, and a chart is only found through
  // one, so this is where the package says the part is a chart.
  files[CONTENT_TYPES_PART] = strToU8(
    `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Override PartName="/${CHART_PART}" ContentType="${CHART_TYPE}"/></Types>`,
  );

  return zipSync(files);
}

async function bytesOf(stream: ReadableStream<Uint8Array>): Promise<Uint8Array> {
  const chunks: Uint8Array[] = [];
  for await (const chunk of stream as unknown as AsyncIterable<Uint8Array>) {
    chunks.push(chunk);
  }
  return new Uint8Array(chunks.flatMap((chunk) => [...chunk]));
}

async function referencesAfterFilling(
  source: Uint8Array,
  fill: (editor: ReturnType<Awaited<ReturnType<typeof Workbook.open>>["edit"]>) => void,
): Promise<readonly string[]> {
  const workbook = await Workbook.open(source);
  const editor = workbook.edit();
  fill(editor);

  const part = strFromU8(unzipSync(await bytesOf(editor.save()))[CHART_PART] ?? new Uint8Array());

  return [...part.matchAll(/<c:f>([^<]*)<\/c:f>/g)].map((match) => match[1] ?? "");
}

// Eight rows into a region five rows deep, so three appear at the region's last row.
const EIGHT_ROWS = [1, 2, 3, 4, 5, 6, 7, 8].map((n) => [`New ${n}`, n, 100]);

const OVER_DATA = { name: "Movements", target: "Data!$A$2:$C$6" };
const OVER_ELSEWHERE = { name: "Others", target: "Elsewhere!$A$2:$C$6" };

describe("a chart in a filled workbook", () => {
  it("follows the rows the region pushed down", async () => {
    const references = await referencesAfterFilling(withChart(chartReading("Data!$C$2:$C$6"), OVER_DATA), (editor) =>
      editor.writeRegion("Movements", EIGHT_ROWS),
    );

    expect(references).toEqual(["Data!$C$2:$C$9"]);
  });

  it("leaves a series reading a sheet that did not move alone", async () => {
    const references = await referencesAfterFilling(
      withChart(chartReading("Elsewhere!$C$2:$C$6"), OVER_DATA),
      (editor) => editor.writeRegion("Movements", EIGHT_ROWS),
    );

    expect(references).toEqual(["Elsewhere!$C$2:$C$6"]);
  });

  // One chart can read two sheets, which no other part positioned by row can, so
  // it is the only one that has to be given more than one move.
  it("follows both sheets when two regions move", async () => {
    const references = await referencesAfterFilling(
      withChart(chartReading("Data!$C$2:$C$6", "Elsewhere!$C$2:$C$6"), OVER_DATA, OVER_ELSEWHERE),
      (editor) => {
        editor.writeRegion("Movements", EIGHT_ROWS);
        editor.writeRegion("Others", EIGHT_ROWS);
      },
    );

    expect(references).toEqual(["Data!$C$2:$C$9", "Elsewhere!$C$2:$C$9"]);
  });

  it("is copied through untouched when nothing was written", async () => {
    const source = withChart(chartReading("Data!$C$2:$C$6"), OVER_DATA);
    const references = await referencesAfterFilling(source, () => {});

    expect(references).toEqual(["Data!$C$2:$C$6"]);
  });
});
