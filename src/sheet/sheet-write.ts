import { XlsxError } from "../errors.js";
import type { CellValue, SetFormulaOptions } from "../types.js";
import type { LocatedRow, SheetIndex } from "./sheet-index.js";
import { splitCellAddress } from "./sheet-address.js";
import { escapeXmlText, parseAttributes, serializeAttributes } from "../utils/xml.js";

export function buildValueCellXml(
  address: string,
  value: CellValue,
  existingAttributesSource?: string,
): string {
  const attributes = parseAttributes(existingAttributesSource ?? "");
  const preserved = attributes.filter(([name]) => name !== "r" && name !== "t");
  const nextAttributes: Array<[string, string]> = [["r", address]];

  if (typeof value === "string") {
    nextAttributes.push(["t", "inlineStr"]);
  } else if (typeof value === "boolean") {
    nextAttributes.push(["t", "b"]);
  }

  nextAttributes.push(...preserved);

  const serializedAttributes = serializeAttributes(nextAttributes);

  if (value === null) {
    return `<c ${serializedAttributes}/>`;
  }

  if (typeof value === "string") {
    const needsSpace = value.trim() !== value;
    const space = needsSpace ? ' xml:space="preserve"' : "";
    return `<c ${serializedAttributes}><is><t${space}>${escapeXmlText(value)}</t></is></c>`;
  }

  if (typeof value === "boolean") {
    return `<c ${serializedAttributes}><v>${value ? "1" : "0"}</v></c>`;
  }

  return `<c ${serializedAttributes}><v>${String(value)}</v></c>`;
}

export function buildStyledCellXml(
  address: string,
  styleId: number | null,
  existingAttributesSource?: string,
  existingCellXml?: string,
): string {
  const serializedAttributes = serializeAttributes(
    buildCellAttributesWithStyle(address, styleId, existingAttributesSource),
  );

  if (!existingCellXml || existingCellXml.endsWith("/>")) {
    return `<c ${serializedAttributes}/>`;
  }

  const openTagEnd = existingCellXml.indexOf(">");
  if (openTagEnd === -1) {
    throw new XlsxError("Cell XML is missing opening tag");
  }

  return `<c ${serializedAttributes}>${existingCellXml.slice(openTagEnd + 1)}`;
}

export function buildFormulaCellXml(
  address: string,
  formula: string,
  cachedValue: CellValue,
  existingAttributesSource?: string,
): string {
  const attributes = parseAttributes(existingAttributesSource ?? "");
  const preserved = attributes.filter(([name]) => name !== "r" && name !== "t");
  const nextAttributes: Array<[string, string]> = [["r", address]];

  if (typeof cachedValue === "string") {
    nextAttributes.push(["t", "str"]);
  } else if (typeof cachedValue === "boolean") {
    nextAttributes.push(["t", "b"]);
  }

  nextAttributes.push(...preserved);

  const serializedAttributes = serializeAttributes(nextAttributes);
  const valueXml = buildFormulaValueXml(cachedValue);

  return `<c ${serializedAttributes}><f>${escapeXmlText(formula)}</f>${valueXml}</c>`;
}

export function insertCell(sheetIndex: SheetIndex, address: string, cellXml: string): string {
  const { rowNumber, columnNumber } = splitCellAddress(address);
  const row = sheetIndex.rows.get(rowNumber);

  if (row) {
    if (row.selfClosing) {
      const nextRowXml = `<row ${row.attributesSource}>${cellXml}</row>`;
      return sheetIndex.xml.slice(0, row.start) + nextRowXml + sheetIndex.xml.slice(row.end);
    }

    const insertionIndex = findCellInsertionIndex(row, columnNumber);
    return sheetIndex.xml.slice(0, insertionIndex) + cellXml + sheetIndex.xml.slice(insertionIndex);
  }

  const rowXml = `<row r="${rowNumber}">${cellXml}</row>`;
  const insertionIndex = findRowInsertionIndex(sheetIndex, rowNumber);

  return sheetIndex.xml.slice(0, insertionIndex) + rowXml + sheetIndex.xml.slice(insertionIndex);
}

export function findRowInsertionIndex(sheetIndex: SheetIndex, rowNumber: number): number {
  for (const candidateRow of sheetIndex.rowNumbers) {
    if (candidateRow > rowNumber) {
      const row = sheetIndex.rows.get(candidateRow);
      if (!row) {
        break;
      }

      return row.start;
    }
  }

  return sheetIndex.sheetDataInnerEnd;
}

export function resolveSetCellValue(
  addressOrRowNumber: string | number,
  columnOrValue: number | string | CellValue,
  value: CellValue | undefined,
): CellValue {
  if (typeof addressOrRowNumber === "string") {
    return columnOrValue as CellValue;
  }

  if (value === undefined) {
    throw new XlsxError(`Missing cell value for row ${addressOrRowNumber}`);
  }

  return value;
}

export function resolveSetFormulaArguments(
  addressOrRowNumber: string | number,
  columnOrFormula: number | string,
  formulaOrOptions: string | SetFormulaOptions | undefined,
  options: SetFormulaOptions,
): { formula: string; formulaOptions: SetFormulaOptions } {
  if (typeof addressOrRowNumber === "string") {
    if (typeof columnOrFormula !== "string") {
      throw new XlsxError(`Invalid formula: ${String(columnOrFormula)}`);
    }

    return {
      formula: columnOrFormula,
      formulaOptions: (formulaOrOptions as SetFormulaOptions | undefined) ?? {},
    };
  }

  if (typeof formulaOrOptions !== "string") {
    throw new XlsxError(`Missing formula for row ${addressOrRowNumber}`);
  }

  return {
    formula: formulaOrOptions,
    formulaOptions: options,
  };
}

function buildCellAttributesWithStyle(
  address: string,
  styleId: number | null,
  existingAttributesSource = "",
): Array<[string, string]> {
  const attributes = parseAttributes(existingAttributesSource);
  const preserved = attributes.filter(([name]) => name !== "r" && name !== "s");
  const nextAttributes: Array<[string, string]> = [["r", address]];

  if (styleId !== null) {
    nextAttributes.push(["s", String(styleId)]);
  }

  nextAttributes.push(...preserved);
  return nextAttributes;
}

function buildFormulaValueXml(value: CellValue): string {
  if (value === null) {
    return "";
  }

  if (typeof value === "string") {
    return `<v>${escapeXmlText(value)}</v>`;
  }

  if (typeof value === "boolean") {
    return `<v>${value ? "1" : "0"}</v>`;
  }

  return `<v>${String(value)}</v>`;
}

function findCellInsertionIndex(row: LocatedRow, columnNumber: number): number {
  for (const cell of row.cells) {
    if (cell.columnNumber > columnNumber) {
      return cell.start;
    }
  }

  return row.innerEnd;
}
