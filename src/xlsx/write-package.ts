import { writeXmlEvent, XML_DECLARATION } from "../xml/write-xml";
import type { XmlEvent, XmlOpen } from "../xml/xml-reader";

const Element = {
  Workbook: "workbook",
  Sheets: "sheets",
  Sheet: "sheet",
  CalculationProperties: "calcPr",
  Types: "Types",
  Override: "Override",
  Relationships: "Relationships",
  Relationship: "Relationship",
} as const;

const Attribute = {
  FullCalculationOnLoad: "fullCalcOnLoad",
  Name: "name",
  SheetId: "sheetId",
  RelationshipId: "r:id",
  Id: "Id",
  Type: "Type",
  PartName: "PartName",
  ContentType: "ContentType",
  Target: "Target",
} as const;

const WORKSHEET_RELATIONSHIP_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";
const WORKSHEET_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";

// Elements the schema places after calcPr inside the workbook part. Inserting a
// calcPr before the first of these keeps the sequence valid, and a workbook with
// none of them takes it just before the closing tag.
const AFTER_CALCULATION_PROPERTIES: ReadonlySet<string> = new Set([
  "oleSize",
  "customWorkbookViews",
  "pivotCaches",
  "smartTagPr",
  "smartTagTypes",
  "webPublishing",
  "fileRecoveryPr",
  "webPublishObjects",
  "extLst",
]);

/** A worksheet being added to a workbook, and the names it takes in each part. */
export interface AddedWorksheet {
  readonly name: string;

  /** The part path, for example `xl/worksheets/sheet3.xml`. */
  readonly path: string;

  /** The relationship id the workbook part points at it by. */
  readonly relationshipId: string;

  /** The workbook-wide identifier, distinct from the position and the path. */
  readonly sheetId: number;
}

/** Serialises an event stream, declaration included, as the part's text. */
export async function* asPart(events: AsyncIterable<XmlEvent>): AsyncIterable<string> {
  yield XML_DECLARATION;

  for await (const event of events) {
    yield writeXmlEvent(event);
  }
}

/** Every event of a part, so a transform can be built on a plain event stream. */
export async function* flatten(batches: AsyncIterable<readonly XmlEvent[]>): AsyncIterable<XmlEvent> {
  for await (const batch of batches) {
    yield* batch;
  }
}

/**
 * Makes a spreadsheet application recalculate every formula when it opens the
 * file.
 *
 * Changing a value leaves the cached result of every formula that depends on it
 * stale, and those cached results are what gets displayed until something
 * recalculates. This is what makes that happen.
 */
export async function* withRecalculationOnLoad(events: AsyncIterable<XmlEvent>): AsyncIterable<XmlEvent> {
  let written = false;
  let depth = 0;

  for await (const event of events) {
    if (event.type === "open") {
      depth += 1;

      if (event.name === Element.CalculationProperties) {
        written = true;
        yield withAttribute(event, Attribute.FullCalculationOnLoad, "1");
        continue;
      }

      if (!written && depth === 2 && AFTER_CALCULATION_PROPERTIES.has(event.name)) {
        written = true;
        yield* calculationProperties();
      }

      yield event;
      continue;
    }

    if (event.type === "close") {
      if (!written && event.name === Element.Workbook) {
        written = true;
        yield* calculationProperties();
      }
      depth -= 1;
    }

    yield event;
  }
}

/** Adds each worksheet to the workbook part's list of sheets. */
export function withAddedSheets(
  events: AsyncIterable<XmlEvent>,
  added: readonly AddedWorksheet[],
): AsyncIterable<XmlEvent> {
  return insertBefore(events, Element.Sheets, Element.Workbook, function* () {
    for (const sheet of added) {
      yield { type: "open", name: Element.Sheet, attributes: sheetAttributes(sheet) };
      yield { type: "close", name: Element.Sheet };
    }
  });
}

/** Adds a relationship from the workbook to each worksheet's part. */
export function withAddedRelationships(
  events: AsyncIterable<XmlEvent>,
  added: readonly AddedWorksheet[],
): AsyncIterable<XmlEvent> {
  return insertBefore(events, Element.Relationships, Element.Relationships, function* () {
    for (const sheet of added) {
      const attributes = {
        [Attribute.Id]: sheet.relationshipId,
        [Attribute.Type]: WORKSHEET_RELATIONSHIP_TYPE,
        [Attribute.Target]: sheet.path.replace(/^xl\//, ""),
      };
      yield { type: "open", name: Element.Relationship, attributes };
      yield { type: "close", name: Element.Relationship };
    }
  });
}

/** Declares the content type of each added worksheet's part. */
export function withAddedContentTypes(
  events: AsyncIterable<XmlEvent>,
  added: readonly AddedWorksheet[],
): AsyncIterable<XmlEvent> {
  return insertBefore(events, Element.Types, Element.Types, function* () {
    for (const sheet of added) {
      const attributes = {
        [Attribute.PartName]: `/${sheet.path}`,
        [Attribute.ContentType]: WORKSHEET_CONTENT_TYPE,
      };
      yield { type: "open", name: Element.Override, attributes };
      yield { type: "close", name: Element.Override };
    }
  });
}

/** Drops the content type override naming `partName`. */
export function withoutContentTypeOverride(events: AsyncIterable<XmlEvent>, partName: string): AsyncIterable<XmlEvent> {
  return withoutElements(
    events,
    (event) => event.name === Element.Override && event.attributes[Attribute.PartName] === partName,
  );
}

/** Drops the relationship pointing at `target`. */
export function withoutRelationshipTo(events: AsyncIterable<XmlEvent>, target: string): AsyncIterable<XmlEvent> {
  return withoutElements(
    events,
    (event) => event.name === Element.Relationship && event.attributes[Attribute.Target] === target,
  );
}

function sheetAttributes(sheet: AddedWorksheet): Record<string, string> {
  return {
    [Attribute.Name]: sheet.name,
    [Attribute.SheetId]: String(sheet.sheetId),
    [Attribute.RelationshipId]: sheet.relationshipId,
  };
}

// Puts the events just inside the close of `container`. A part that has no such
// element, an empty workbook with no sheets element at all, falls back to just
// inside the close of `root`.
async function* insertBefore(
  events: AsyncIterable<XmlEvent>,
  container: string,
  root: string,
  extra: () => Iterable<XmlEvent>,
): AsyncIterable<XmlEvent> {
  let inserted = false;

  for await (const event of events) {
    if (event.type === "close" && !inserted && (event.name === container || event.name === root)) {
      inserted = true;
      yield* extra();
    }
    yield event;
  }
}

function* calculationProperties(): Iterable<XmlEvent> {
  const attributes = { [Attribute.FullCalculationOnLoad]: "1" };

  yield { type: "open", name: Element.CalculationProperties, attributes };
  yield { type: "close", name: Element.CalculationProperties };
}

function withAttribute(event: XmlOpen, name: string, value: string): XmlEvent {
  return { ...event, attributes: { ...event.attributes, [name]: value } };
}

async function* withoutElements(
  events: AsyncIterable<XmlEvent>,
  matches: (event: XmlOpen) => boolean,
): AsyncIterable<XmlEvent> {
  let skipping = 0;

  for await (const event of events) {
    if (skipping > 0) {
      if (event.type === "open") {
        skipping += 1;
      } else if (event.type === "close") {
        skipping -= 1;
      }
      continue;
    }

    if (event.type === "open" && matches(event)) {
      skipping = 1;
      continue;
    }

    yield event;
  }
}
