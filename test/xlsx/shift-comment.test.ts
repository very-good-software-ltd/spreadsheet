import { describe, expect, it } from "vitest";
import { shiftCommentRefs } from "../../src/xlsx/shift-comment";
import type { RowShift } from "../../src/xlsx/shift-formula";
import { SaxesXmlReader } from "../../src/xml/saxes-xml-reader";
import { writeXmlEvent } from "../../src/xml/write-xml";

function streamOf(text: string): ReadableStream<Uint8Array> {
  return new ReadableStream<Uint8Array>({
    start(controller): void {
      controller.enqueue(new TextEncoder().encode(text));
      controller.close();
    },
  });
}

async function shifted(comments: string, shift: RowShift): Promise<string> {
  const events = shiftCommentRefs(new SaxesXmlReader().read(streamOf(comments)), shift);
  let out = "";
  for await (const batch of events) {
    for (const event of batch) {
      out += writeXmlEvent(event);
    }
  }
  return out;
}

function commentsOn(...refs: readonly string[]): string {
  const list = refs
    .map((ref) => `<comment ref="${ref}" authorId="0"><text><r><t>note ${ref}</t></r></text></comment>`)
    .join("");

  return `<comments xmlns="m"><authors><author>Author</author></authors><commentList>${list}</commentList></comments>`;
}

const refsOf = (comments: string): string[] =>
  [...comments.matchAll(/<comment ref="([^"]+)"/g)].map((match) => match[1] ?? "");

// Two rows appear at row 11.
const INSERT: RowShift = { sheet: "Report", at: 11, by: 2 };

// Rows 10 and 11 go away.
const REMOVE: RowShift = { sheet: "Report", at: 10, by: -2 };

describe("shiftCommentRefs", () => {
  it("moves a comment below the rows down by as many", async () => {
    expect(refsOf(await shifted(commentsOn("B12"), INSERT))).toEqual(["B14"]);
  });

  it("leaves a comment above the rows alone", async () => {
    expect(refsOf(await shifted(commentsOn("B3"), INSERT))).toEqual(["B3"]);
  });

  it("moves a comment on the first row that moves", async () => {
    expect(refsOf(await shifted(commentsOn("B11"), INSERT))).toEqual(["B13"]);
  });

  it("keeps the column when the row moves", async () => {
    expect(refsOf(await shifted(commentsOn("AB12"), INSERT))).toEqual(["AB14"]);
  });

  it("pulls a comment below the removed rows up", async () => {
    expect(refsOf(await shifted(commentsOn("B12"), REMOVE))).toEqual(["B10"]);
  });

  it("drops a comment whose row went away", async () => {
    expect(refsOf(await shifted(commentsOn("B10"), REMOVE))).toEqual([]);
  });

  it("drops the text of a comment whose row went away with it", async () => {
    expect(await shifted(commentsOn("B10"), REMOVE)).not.toContain("note B10");
  });

  it("keeps the comments either side of one that went away", async () => {
    expect(refsOf(await shifted(commentsOn("B3", "B10", "B12"), REMOVE))).toEqual(["B3", "B10"]);
  });

  it("leaves the authors alone", async () => {
    expect(await shifted(commentsOn("B12"), INSERT)).toContain("<author>Author</author>");
  });
});
