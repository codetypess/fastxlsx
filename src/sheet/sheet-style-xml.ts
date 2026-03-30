import type { LocatedRow } from "./sheet-index.js";
import { buildXmlContainer, findWorksheetChildInsertionIndex, replaceXmlTagSource } from "./sheet-xml.js";
import { findFirstXmlTag, findXmlTags, getTagAttr } from "../utils/xml-read.js";
import { getXmlAttr, parseAttributes, serializeAttributes } from "../utils/xml.js";

interface ColumnDefinition {
  min: number;
  max: number;
  attributes: Array<[string, string]>;
}

const COLS_FOLLOWING_TAGS = [
  "sheetData",
  "autoFilter",
  "sortState",
  "mergeCells",
  "phoneticPr",
  "conditionalFormatting",
  "dataValidations",
  "hyperlinks",
  "printOptions",
  "pageMargins",
  "pageSetup",
  "headerFooter",
  "rowBreaks",
  "colBreaks",
  "customProperties",
  "cellWatches",
  "ignoredErrors",
  "smartTags",
  "drawing",
  "legacyDrawing",
  "legacyDrawingHF",
  "picture",
  "oleObjects",
  "controls",
  "webPublishItems",
  "tableParts",
  "extLst",
];

export function parseRowStyleId(attributesSource: string | undefined): number | null {
  if (!attributesSource) {
    return null;
  }

  const styleId = getXmlAttr(attributesSource, "s");
  return styleId === undefined ? null : Number(styleId);
}

export function parseColumnStyleId(sheetXml: string, columnNumber: number): number | null {
  let styleId: number | null = null;

  for (const definition of parseColumnDefinitions(sheetXml)) {
    if (columnNumber < definition.min || columnNumber > definition.max) {
      continue;
    }

    const styleText = getXmlAttr(serializeAttributes(definition.attributes), "style");
    styleId = styleText === undefined ? null : Number(styleText);
  }

  return styleId;
}

export function buildStyledRowXml(sheetXml: string, row: LocatedRow, styleId: number | null): string {
  const serializedAttributes = serializeAttributes(
    buildRowAttributesWithStyle(row.rowNumber, styleId, row.attributesSource),
  );

  if (row.selfClosing) {
    return `<row ${serializedAttributes}/>`;
  }

  return `<row ${serializedAttributes}>${sheetXml.slice(row.innerStart, row.innerEnd)}</row>`;
}

export function buildEmptyStyledRowXml(rowNumber: number, styleId: number): string {
  return `<row ${serializeAttributes(buildRowAttributesWithStyle(rowNumber, styleId))}/>`;
}

export function updateColumnStyleIdInSheetXml(
  sheetXml: string,
  columnNumber: number,
  styleId: number | null,
): string {
  const existingDefinitions = parseColumnDefinitions(sheetXml);
  if (existingDefinitions.length === 0 && styleId === null) {
    return sheetXml;
  }

  const nextDefinitions: ColumnDefinition[] = [];
  let handled = false;

  for (const definition of existingDefinitions) {
    if (columnNumber < definition.min || columnNumber > definition.max) {
      nextDefinitions.push(definition);
      continue;
    }

    handled = true;
    if (definition.min < columnNumber) {
      nextDefinitions.push(buildColumnDefinition(definition.min, columnNumber - 1, definition.attributes));
    }

    const styledDefinition = buildColumnDefinitionWithStyle(columnNumber, columnNumber, definition.attributes, styleId);
    if (styledDefinition) {
      nextDefinitions.push(styledDefinition);
    }

    if (columnNumber < definition.max) {
      nextDefinitions.push(buildColumnDefinition(columnNumber + 1, definition.max, definition.attributes));
    }
  }

  if (!handled && styleId !== null) {
    nextDefinitions.push(buildColumnDefinitionWithStyle(columnNumber, columnNumber, [], styleId)!);
  }

  return replaceColumnDefinitions(sheetXml, normalizeColumnDefinitions(nextDefinitions));
}

export function transformColumnStyleDefinitions(
  sheetXml: string,
  targetColumnNumber: number,
  count: number,
  mode: "shift" | "delete",
): string {
  const existingDefinitions = parseColumnDefinitions(sheetXml);
  if (existingDefinitions.length === 0) {
    return sheetXml;
  }

  const nextDefinitions: ColumnDefinition[] = [];

  for (const definition of existingDefinitions) {
    if (mode === "shift") {
      if (definition.max < targetColumnNumber) {
        nextDefinitions.push(definition);
        continue;
      }

      if (definition.min >= targetColumnNumber) {
        nextDefinitions.push(buildColumnDefinition(definition.min + count, definition.max + count, definition.attributes));
        continue;
      }

      nextDefinitions.push(buildColumnDefinition(definition.min, targetColumnNumber - 1, definition.attributes));
      nextDefinitions.push(buildColumnDefinition(targetColumnNumber + count, definition.max + count, definition.attributes));
      continue;
    }

    const deleteEnd = targetColumnNumber + count - 1;
    if (definition.max < targetColumnNumber) {
      nextDefinitions.push(definition);
      continue;
    }

    if (definition.min > deleteEnd) {
      nextDefinitions.push(buildColumnDefinition(definition.min - count, definition.max - count, definition.attributes));
      continue;
    }

    if (definition.min < targetColumnNumber) {
      nextDefinitions.push(buildColumnDefinition(definition.min, targetColumnNumber - 1, definition.attributes));
    }

    if (definition.max > deleteEnd) {
      nextDefinitions.push(buildColumnDefinition(targetColumnNumber, definition.max - count, definition.attributes));
    }
  }

  return replaceColumnDefinitions(sheetXml, normalizeColumnDefinitions(nextDefinitions));
}

function buildRowAttributesWithStyle(
  rowNumber: number,
  styleId: number | null,
  existingAttributesSource = "",
): Array<[string, string]> {
  const attributes = parseAttributes(existingAttributesSource);
  const preserved = attributes.filter(
    ([name]) => name !== "r" && name !== "s" && name !== "customFormat",
  );
  const nextAttributes: Array<[string, string]> = [["r", String(rowNumber)]];

  if (styleId !== null) {
    nextAttributes.push(["s", String(styleId)], ["customFormat", "1"]);
  }

  nextAttributes.push(...preserved);
  return nextAttributes;
}

function parseColumnDefinitions(sheetXml: string): ColumnDefinition[] {
  const colsTag = findFirstXmlTag(sheetXml, "cols");
  if (!colsTag || colsTag.innerXml === null) {
    return [];
  }

  return findXmlTags(colsTag.innerXml, "col")
    .map((colTag) => {
      const attributes = parseAttributes(colTag.attributesSource);
      const min = Number(getTagAttr(colTag, "min") ?? "0");
      const max = Number(getTagAttr(colTag, "max") ?? "0");

      return {
        min,
        max,
        attributes,
      };
    })
    .filter(
      (definition) =>
        Number.isInteger(definition.min) &&
        Number.isInteger(definition.max) &&
        definition.min > 0 &&
        definition.max >= definition.min,
    );
}

function replaceColumnDefinitions(sheetXml: string, definitions: ColumnDefinition[]): string {
  const colsTag = findFirstXmlTag(sheetXml, "cols");
  const colsXml =
    definitions.length === 0
      ? ""
      : buildXmlContainer(
          "cols",
          colsTag?.attributesSource ?? "",
          definitions.map((definition) => serializeColumnDefinition(definition)).join(""),
        );

  if (colsTag) {
    return replaceXmlTagSource(sheetXml, colsTag, colsXml);
  }

  if (definitions.length === 0) {
    return sheetXml;
  }

  const insertionIndex = findWorksheetChildInsertionIndex(sheetXml, COLS_FOLLOWING_TAGS);
  return sheetXml.slice(0, insertionIndex) + colsXml + sheetXml.slice(insertionIndex);
}

function normalizeColumnDefinitions(definitions: ColumnDefinition[]): ColumnDefinition[] {
  const filtered = definitions
    .filter((definition) => definition.min <= definition.max)
    .sort((left, right) => left.min - right.min || left.max - right.max);
  const merged: ColumnDefinition[] = [];

  for (const definition of filtered) {
    const previous = merged.at(-1);
    if (
      previous &&
      previous.max + 1 === definition.min &&
      haveEquivalentColumnDefinitionAttributes(previous.attributes, definition.attributes)
    ) {
      previous.max = definition.max;
      continue;
    }

    merged.push({
      min: definition.min,
      max: definition.max,
      attributes: [...definition.attributes],
    });
  }

  return merged;
}

function buildColumnDefinition(
  min: number,
  max: number,
  existingAttributes: Array<[string, string]>,
): ColumnDefinition {
  const preserved = existingAttributes.filter(([name]) => name !== "min" && name !== "max");
  return {
    min,
    max,
    attributes: [["min", String(min)], ["max", String(max)], ...preserved],
  };
}

function buildColumnDefinitionWithStyle(
  min: number,
  max: number,
  existingAttributes: Array<[string, string]>,
  styleId: number | null,
): ColumnDefinition | null {
  const preserved = existingAttributes.filter(
    ([name]) => name !== "min" && name !== "max" && name !== "style",
  );

  if (styleId === null && preserved.length === 0) {
    return null;
  }

  const attributes: Array<[string, string]> = [["min", String(min)], ["max", String(max)]];
  if (styleId !== null) {
    attributes.push(["style", String(styleId)]);
  }
  attributes.push(...preserved);

  return { min, max, attributes };
}

function serializeColumnDefinition(definition: ColumnDefinition): string {
  const attributes = definition.attributes.map(([name, value]) => {
    if (name === "min") {
      return [name, String(definition.min)] as [string, string];
    }
    if (name === "max") {
      return [name, String(definition.max)] as [string, string];
    }
    return [name, value] as [string, string];
  });

  return `<col ${serializeAttributes(attributes)}/>`;
}

function haveEquivalentColumnDefinitionAttributes(
  left: Array<[string, string]>,
  right: Array<[string, string]>,
): boolean {
  const normalize = (attributes: Array<[string, string]>) =>
    serializeAttributes(attributes.filter(([name]) => name !== "min" && name !== "max"));

  return normalize(left) === normalize(right);
}
