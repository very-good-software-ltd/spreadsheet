import { describe, expect, it } from "vitest";
import {
  withoutContentTypeOverride,
  withoutRelationshipTo,
  withRecalculationOnLoad,
} from "../../src/xlsx/write-package";
import { SaxesXmlReader } from "../../src/xml/saxes-xml-reader";
import { XML_DECLARATION } from "../../src/xml/write-xml";

function streamOf(text: string): ReadableStream<Uint8Array> {
  const bytes = new TextEncoder().encode(text);
  return new ReadableStream<Uint8Array>({
    start(controller): void {
      controller.enqueue(bytes);
      controller.close();
    },
  });
}

async function run(xml: string, transform: (events: never) => AsyncIterable<string>): Promise<string> {
  const events = new SaxesXmlReader().read(streamOf(xml)) as never;
  let out = "";
  for await (const chunk of transform(events)) {
    out += chunk;
  }
  return out.slice(XML_DECLARATION.length);
}

const WORKBOOK_OPEN = `<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">`;

describe("withRecalculationOnLoad", () => {
  it("adds the flag to a calcPr the workbook already has", async () => {
    const xml = `${WORKBOOK_OPEN}<sheets></sheets><calcPr calcId="181029"></calcPr></workbook>`;

    expect(await run(xml, withRecalculationOnLoad)).toBe(
      `${WORKBOOK_OPEN}<sheets></sheets><calcPr calcId="181029" fullCalcOnLoad="1"></calcPr></workbook>`,
    );
  });

  it("replaces the flag when it is already there and says otherwise", async () => {
    const xml = `${WORKBOOK_OPEN}<calcPr fullCalcOnLoad="0"></calcPr></workbook>`;

    expect(await run(xml, withRecalculationOnLoad)).toContain(`fullCalcOnLoad="1"`);
  });

  it("adds a calcPr at the end when the workbook has nothing that must follow it", async () => {
    const xml = `${WORKBOOK_OPEN}<sheets></sheets></workbook>`;

    expect(await run(xml, withRecalculationOnLoad)).toBe(
      `${WORKBOOK_OPEN}<sheets></sheets><calcPr fullCalcOnLoad="1"></calcPr></workbook>`,
    );
  });

  it("puts a new calcPr before the elements the schema orders after it", async () => {
    const xml = `${WORKBOOK_OPEN}<sheets></sheets><pivotCaches></pivotCaches><extLst></extLst></workbook>`;

    expect(await run(xml, withRecalculationOnLoad)).toBe(
      `${WORKBOOK_OPEN}<sheets></sheets><calcPr fullCalcOnLoad="1"></calcPr><pivotCaches></pivotCaches><extLst></extLst></workbook>`,
    );
  });

  it("is not fooled by a nested element that shares a name with one of those", async () => {
    const xml = `${WORKBOOK_OPEN}<sheets><extLst></extLst></sheets></workbook>`;

    expect(await run(xml, withRecalculationOnLoad)).toBe(
      `${WORKBOOK_OPEN}<sheets><extLst></extLst></sheets><calcPr fullCalcOnLoad="1"></calcPr></workbook>`,
    );
  });

  it("adds only one", async () => {
    const xml = `${WORKBOOK_OPEN}<pivotCaches></pivotCaches><extLst></extLst></workbook>`;

    expect((await run(xml, withRecalculationOnLoad)).match(/<calcPr/g)).toHaveLength(1);
  });
});

describe("withoutContentTypeOverride", () => {
  const types = `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"></Default><Override PartName="/xl/calcChain.xml" ContentType="calcChain"></Override><Override PartName="/xl/styles.xml" ContentType="styles"></Override></Types>`;

  it("drops the override naming the part", async () => {
    const result = await run(types, (events) => withoutContentTypeOverride(events, "/xl/calcChain.xml"));

    expect(result).not.toContain("calcChain");
    expect(result).toContain(`<Override PartName="/xl/styles.xml" ContentType="styles"></Override>`);
    expect(result).toContain(`<Default Extension="xml"`);
  });

  it("leaves the part alone when the override is not there", async () => {
    expect(await run(types, (events) => withoutContentTypeOverride(events, "/xl/missing.xml"))).toBe(types);
  });
});

describe("withoutRelationshipTo", () => {
  const rels = `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Target="worksheets/sheet1.xml"></Relationship><Relationship Id="rId2" Target="calcChain.xml"></Relationship></Relationships>`;

  it("drops the relationship pointing at the target", async () => {
    const result = await run(rels, (events) => withoutRelationshipTo(events, "calcChain.xml"));

    expect(result).not.toContain("calcChain");
    expect(result).toContain(`Target="worksheets/sheet1.xml"`);
  });

  it("leaves the part alone when nothing points there", async () => {
    expect(await run(rels, (events) => withoutRelationshipTo(events, "nothing.xml"))).toBe(rels);
  });
});
