import { XlsxError } from "../errors.js";
import { parseStringItemText } from "../workbook/shared-strings.js";
import type { CellSnapshot, CellType, CellValue } from "../types.js";
import type { Workbook } from "../workbook.js";
import { decodeXmlText } from "../utils/xml.js";

export interface LocatedCell {
  address: string;
  start: number;
  end: number;
  attributesSource: string;
  innerStart: number;
  innerEnd: number;
  rawType: string | null;
  snapshot: CellSnapshot;
  styleId: number | null;
  rowNumber: number;
  columnNumber: number;
}

export interface LocatedRow {
  start: number;
  end: number;
  attributesSource: string;
  selfClosing: boolean;
  rowNumber: number;
  innerStart: number;
  innerEnd: number;
  cells: LocatedCell[];
  cellsByColumn: Array<LocatedCell | undefined>;
  maxColumnNumber: number;
}

interface UsedRangeBounds {
  minRow: number;
  maxRow: number;
  minColumn: number;
  maxColumn: number;
}

export interface SheetIndex {
  xml: string;
  cells: Map<string, LocatedCell> | null;
  rows: Map<number, LocatedRow>;
  rowNumbers: number[];
  usedBounds: UsedRangeBounds | null;
  sheetDataInnerStart: number;
  sheetDataInnerEnd: number;
}

export function parseCellSnapshot(cell: LocatedCell | undefined): CellSnapshot {
  if (!cell) {
    return {
      exists: false,
      formula: null,
      rawType: null,
      styleId: null,
      type: "missing",
      value: null,
    };
  }

  return cell.snapshot;
}

export function buildSheetIndex(workbook: Workbook, sheetXml: string): SheetIndex {
  const sheetDataStart = sheetXml.indexOf("<sheetData");
  if (sheetDataStart === -1) {
    throw new XlsxError("Worksheet is missing <sheetData>");
  }

  const sheetDataOpenTagEnd = sheetXml.indexOf(">", sheetDataStart);
  const sheetDataCloseTagStart = sheetXml.indexOf("</sheetData>", sheetDataOpenTagEnd + 1);
  if (sheetDataOpenTagEnd === -1 || sheetDataCloseTagStart === -1) {
    throw new XlsxError("Worksheet is missing <sheetData>");
  }

  const sheetDataInnerStart = sheetDataOpenTagEnd + 1;
  const sheetDataInnerEnd = sheetDataCloseTagStart;
  const rows = new Map<number, LocatedRow>();
  const rowNumbers: number[] = [];
  let minRow = Number.POSITIVE_INFINITY;
  let maxRow = 0;
  let minColumn = Number.POSITIVE_INFINITY;
  let maxColumn = 0;
  let hasUsedBounds = false;
  let previousRowNumber = 0;
  let rowsAreSorted = true;
  let cursor = sheetDataInnerStart;

  while (cursor < sheetDataInnerEnd) {
    const rowStart = sheetXml.indexOf("<row", cursor);
    if (rowStart === -1 || rowStart >= sheetDataInnerEnd) {
      break;
    }

    const rowOpenTagEnd = sheetXml.indexOf(">", rowStart + 4);
    if (rowOpenTagEnd === -1 || rowOpenTagEnd >= sheetDataInnerEnd) {
      break;
    }

    const rowTagSource = sheetXml.slice(rowStart + 4, rowOpenTagEnd);
    const rowMetadata = parseRowTagMetadata(rowTagSource);
    const selfClosing = rowMetadata.selfClosing;
    const attributesSource = rowMetadata.attributesSource;
    const rowNumberText = rowMetadata.rowNumberText;
    const rowEnd = selfClosing
      ? rowOpenTagEnd + 1
      : sheetXml.indexOf("</row>", rowOpenTagEnd + 1);

    if (!rowNumberText || rowEnd === -1) {
      cursor = rowOpenTagEnd + 1;
      continue;
    }

    const rowNumber = Number(rowNumberText);
    const innerStart = selfClosing ? rowEnd : rowOpenTagEnd + 1;
    const innerEnd = selfClosing ? rowEnd : rowEnd;
    const row: LocatedRow = {
      start: rowStart,
      end: selfClosing ? rowEnd : rowEnd + ROW_CLOSE_TAG.length,
      attributesSource,
      selfClosing,
      rowNumber,
      innerStart,
      innerEnd,
      cells: [],
      cellsByColumn: [],
      maxColumnNumber: 0,
    };

    if (!selfClosing) {
      let cellCursor = innerStart;
      let previousColumnNumber = 0;
      let cellsAreSorted = true;

      while (cellCursor < innerEnd) {
        const cellStart = sheetXml.indexOf("<c", cellCursor);
        if (cellStart === -1 || cellStart >= innerEnd) {
          break;
        }

        const cellOpenTagEnd = sheetXml.indexOf(">", cellStart + 2);
        if (cellOpenTagEnd === -1 || cellOpenTagEnd > innerEnd) {
          break;
        }

        const cellTagSource = sheetXml.slice(cellStart + 2, cellOpenTagEnd);
        const cellMetadata = parseCellTagMetadata(cellTagSource);
        const cellSelfClosing = cellMetadata.selfClosing;
        const cellAttributesSource = cellMetadata.attributesSource;
        const addressSource = cellMetadata.addressSource;
        const cellEnd = cellSelfClosing
          ? cellOpenTagEnd + 1
          : sheetXml.indexOf(CELL_CLOSE_TAG, cellOpenTagEnd + 1);

        if (!addressSource || cellEnd === -1) {
          cellCursor = cellOpenTagEnd + 1;
          continue;
        }

        const address = addressSource.toUpperCase();
        const columnNumber = columnLabelToNumberFromAddress(address);
        const cellInnerStart = cellSelfClosing ? cellEnd : cellOpenTagEnd + 1;
        const cellInnerEnd = cellSelfClosing ? cellEnd : cellEnd;
        const rawType = cellMetadata.rawType;
        const styleId = cellMetadata.styleIdText === undefined ? null : Number(cellMetadata.styleIdText);
        const cell: LocatedCell = {
          address,
          start: cellStart,
          end: cellSelfClosing ? cellEnd : cellEnd + CELL_CLOSE_TAG.length,
          attributesSource: cellAttributesSource,
          innerStart: cellInnerStart,
          innerEnd: cellInnerEnd,
          rawType,
          snapshot: buildCellSnapshot(workbook, rawType, styleId, sheetXml, cellInnerStart, cellInnerEnd),
          styleId,
          rowNumber,
          columnNumber,
        };

        row.cells.push(cell);
        row.cellsByColumn[columnNumber] = cell;
        if (columnNumber < previousColumnNumber) {
          cellsAreSorted = false;
        }

        previousColumnNumber = columnNumber;
        cellCursor = cell.end;
      }

      if (!cellsAreSorted) {
        row.cells.sort((left, right) => left.columnNumber - right.columnNumber);
      }
    }

    const rowAnalysis = analyzeRowCells(row.cells);
    row.maxColumnNumber = rowAnalysis.maxColumnNumber;
    if (row.maxColumnNumber > 0) {
      minColumn = Math.min(minColumn, row.cells[0]?.columnNumber ?? row.maxColumnNumber);
      maxColumn = Math.max(maxColumn, row.maxColumnNumber);
    }
    if (rowAnalysis.hasUsedCells) {
      hasUsedBounds = true;
      minRow = Math.min(minRow, rowNumber);
      maxRow = Math.max(maxRow, rowNumber);
    }

    rows.set(rowNumber, row);
    rowNumbers.push(rowNumber);
    if (rowNumber < previousRowNumber) {
      rowsAreSorted = false;
    }

    previousRowNumber = rowNumber;
    cursor = row.end;
  }

  if (!rowsAreSorted) {
    rowNumbers.sort((left, right) => left - right);
  }

  return {
    xml: sheetXml,
    cells: null,
    rows,
    rowNumbers,
    usedBounds: hasUsedBounds ? { minRow, maxRow, minColumn, maxColumn } : null,
    sheetDataInnerStart,
    sheetDataInnerEnd,
  };
}

export function getLocatedCell(index: SheetIndex, address: string): LocatedCell | undefined {
  if (!index.cells) {
    index.cells = new Map<string, LocatedCell>();

    for (const rowNumber of index.rowNumbers) {
      const row = index.rows.get(rowNumber);
      if (!row) {
        continue;
      }

      for (const cell of row.cells) {
        index.cells.set(cell.address, cell);
      }
    }
  }

  return index.cells.get(address);
}

function analyzeRowCells(cells: LocatedCell[]): { hasUsedCells: boolean; maxColumnNumber: number } {
  if (cells.length === 0) {
    return { hasUsedCells: false, maxColumnNumber: 0 };
  }

  let lastUsedIndex = -1;

  for (let index = cells.length - 1; index >= 0; index -= 1) {
    if (isUsedCell(cells[index]?.snapshot)) {
      lastUsedIndex = index;
      break;
    }
  }

  if (lastUsedIndex === -1) {
    let logicalMaxColumnNumber = cells[0]?.columnNumber ?? 0;

    for (let index = 1; index < cells.length; index += 1) {
      const columnNumber = cells[index]?.columnNumber ?? 0;
      if (columnNumber > logicalMaxColumnNumber + 1) {
        break;
      }

      logicalMaxColumnNumber = columnNumber;
    }

    return { hasUsedCells: false, maxColumnNumber: logicalMaxColumnNumber };
  }

  let logicalMaxColumnNumber = cells[lastUsedIndex]?.columnNumber ?? 0;

  for (let index = lastUsedIndex + 1; index < cells.length; index += 1) {
    const cell = cells[index];
    if (!cell || isUsedCell(cell.snapshot) || cell.columnNumber > logicalMaxColumnNumber + 1) {
      break;
    }

    logicalMaxColumnNumber = cell.columnNumber;
  }

  return { hasUsedCells: true, maxColumnNumber: logicalMaxColumnNumber };
}

function isUsedCell(snapshot: CellSnapshot | undefined): boolean {
  return snapshot !== undefined && (snapshot.formula !== null || snapshot.value !== null);
}

export function parseCellAddressFast(address: string): { rowNumber: number; columnNumber: number } {
  let columnNumber = 0;
  let rowNumber = 0;
  let index = 0;

  while (index < address.length) {
    let characterCode = address.charCodeAt(index);
    if (characterCode === 36) {
      index += 1;
      continue;
    }

    if (characterCode >= 97 && characterCode <= 122) {
      characterCode -= 32;
    }

    if (characterCode < 65 || characterCode > 90) {
      break;
    }

    columnNumber = columnNumber * 26 + (characterCode - 64);
    index += 1;
  }

  while (index < address.length) {
    const characterCode = address.charCodeAt(index);
    if (characterCode === 36) {
      index += 1;
      continue;
    }

    if (characterCode < 48 || characterCode > 57) {
      throw new XlsxError(`Invalid cell address: ${address}`);
    }

    rowNumber = rowNumber * 10 + (characterCode - 48);
    index += 1;
  }

  if (columnNumber === 0 || rowNumber === 0) {
    throw new XlsxError(`Invalid cell address: ${address}`);
  }

  return { rowNumber, columnNumber };
}

function parseCellValue(
  workbook: Workbook,
  rawType: string | null,
  valueSource: string | undefined,
  inlineText: string | null,
  hasSelfClosingValue: boolean,
): CellValue {
  if (rawType === "inlineStr") {
    return inlineText ?? "";
  }

  if (rawType === "str") {
    if (valueSource !== undefined) {
      return decodeXmlText(valueSource);
    }

    return hasSelfClosingValue ? "" : null;
  }

  if (rawType === "s") {
    const indexText = valueSource;
    if (!indexText) {
      return null;
    }

    return workbook.getSharedString(Number(indexText));
  }

  if (rawType === "b") {
    return valueSource === "1";
  }

  if (valueSource === undefined) {
    return null;
  }

  const numericValue = Number(valueSource);
  return Number.isFinite(numericValue) ? numericValue : decodeXmlText(valueSource);
}

function buildCellSnapshot(
  workbook: Workbook,
  rawType: string | null,
  styleId: number | null,
  xml: string,
  innerStart: number,
  innerEnd: number,
): CellSnapshot {
  const { formulaSource, hasSelfClosingValue, valueSource } = extractCellContentsFast(xml, innerStart, innerEnd);
  const inlineText = rawType === "inlineStr" ? parseStringItemText(xml.slice(innerStart, innerEnd)) : null;
  const formula = formulaSource === null ? null : decodeXmlText(formulaSource);
  const value = parseCellValue(workbook, rawType, valueSource, inlineText, hasSelfClosingValue);

  if (formula !== null) {
    return {
      exists: true,
      formula,
      rawType,
      styleId,
      type: "formula",
      value,
    };
  }

  const type: CellType =
    value === null
      ? "blank"
      : typeof value === "string"
        ? "string"
        : typeof value === "number"
          ? "number"
          : "boolean";

  return {
    exists: true,
    formula: null,
    rawType,
    styleId,
    type,
    value,
  };
}

function extractCellTypeAttr(attributesSource: string): string | null {
  return readXmlAttrFast(attributesSource, "t") ?? null;
}

function extractCellStyleAttr(attributesSource: string): string | undefined {
  return readXmlAttrFast(attributesSource, "s");
}

function parseRowTagMetadata(source: string): {
  attributesSource: string;
  rowNumberText: string | undefined;
  selfClosing: boolean;
} {
  const selfClosing = isSelfClosingTagSource(source);
  const attributesSource = cleanTagAttributesSource(source);
  return {
    attributesSource,
    rowNumberText: readXmlAttrFast(attributesSource, "r"),
    selfClosing,
  };
}

function parseCellTagMetadata(source: string): {
  addressSource: string | undefined;
  attributesSource: string;
  rawType: string | null;
  selfClosing: boolean;
  styleIdText: string | undefined;
} {
  const selfClosing = isSelfClosingTagSource(source);
  const attributesSource = cleanTagAttributesSource(source);
  let addressSource: string | undefined;
  let rawType: string | null = null;
  let styleIdText: string | undefined;
  let index = 0;

  while (index < attributesSource.length) {
    while (index < attributesSource.length && isXmlWhitespaceCode(attributesSource.charCodeAt(index))) {
      index += 1;
    }

    const nameStart = index;
    while (index < attributesSource.length) {
      const code = attributesSource.charCodeAt(index);
      if (code === 61 || isXmlWhitespaceCode(code)) {
        break;
      }
      index += 1;
    }

    if (index <= nameStart) {
      break;
    }

    const name = attributesSource.slice(nameStart, index);
    while (index < attributesSource.length && isXmlWhitespaceCode(attributesSource.charCodeAt(index))) {
      index += 1;
    }
    if (attributesSource.charCodeAt(index) !== 61) {
      while (index < attributesSource.length && !isXmlWhitespaceCode(attributesSource.charCodeAt(index))) {
        index += 1;
      }
      continue;
    }

    index += 1;
    while (index < attributesSource.length && isXmlWhitespaceCode(attributesSource.charCodeAt(index))) {
      index += 1;
    }

    const quote = attributesSource.charCodeAt(index);
    if (quote !== 34 && quote !== 39) {
      continue;
    }

    const valueStart = index + 1;
    const valueEnd = attributesSource.indexOf(String.fromCharCode(quote), valueStart);
    if (valueEnd === -1) {
      break;
    }

    const value = attributesSource.slice(valueStart, valueEnd);
    if (name === "r") {
      addressSource = value;
    } else if (name === "t") {
      rawType = value;
    } else if (name === "s") {
      styleIdText = value;
    }

    index = valueEnd + 1;
  }

  return {
    addressSource,
    attributesSource,
    rawType,
    selfClosing,
    styleIdText,
  };
}

function cleanTagAttributesSource(source: string): string {
  let end = source.length;

  while (end > 0 && isXmlWhitespaceCode(source.charCodeAt(end - 1))) {
    end -= 1;
  }

  if (source.charCodeAt(end - 1) === 47) {
    end -= 1;
  }

  while (end > 0 && isXmlWhitespaceCode(source.charCodeAt(end - 1))) {
    end -= 1;
  }

  let start = 0;
  while (start < end && isXmlWhitespaceCode(source.charCodeAt(start))) {
    start += 1;
  }

  return source.slice(start, end);
}

function isSelfClosingTagSource(source: string): boolean {
  let index = source.length - 1;

  while (index >= 0 && isXmlWhitespaceCode(source.charCodeAt(index))) {
    index -= 1;
  }

  return index >= 0 && source.charCodeAt(index) === 47;
}

function readXmlAttrFast(source: string, attributeName: string): string | undefined {
  const pattern = attributeName;
  let searchStart = 0;

  while (searchStart < source.length) {
    const attributeStart = source.indexOf(pattern, searchStart);
    if (attributeStart === -1) {
      return undefined;
    }

    const previousCode = attributeStart === 0 ? 32 : source.charCodeAt(attributeStart - 1);
    if (isXmlAttributeBoundaryCode(previousCode)) {
      let cursor = attributeStart + pattern.length;

      while (cursor < source.length && isXmlWhitespaceCode(source.charCodeAt(cursor))) {
        cursor += 1;
      }

      if (source.charCodeAt(cursor) !== 61) {
        searchStart = attributeStart + pattern.length;
        continue;
      }

      cursor += 1;
      while (cursor < source.length && isXmlWhitespaceCode(source.charCodeAt(cursor))) {
        cursor += 1;
      }

      const quote = source.charCodeAt(cursor);
      if (quote !== 34 && quote !== 39) {
        searchStart = attributeStart + pattern.length;
        continue;
      }

      const valueStart = cursor + 1;
      const valueEnd = source.indexOf(String.fromCharCode(quote), valueStart);
      return valueEnd === -1 ? undefined : source.slice(valueStart, valueEnd);
    }

    searchStart = attributeStart + pattern.length;
  }

  return undefined;
}

function extractCellContentsFast(xml: string, start: number, end: number): {
  formulaSource: string | null;
  hasSelfClosingValue: boolean;
  valueSource: string | undefined;
} {
  let formulaSource: string | null = null;
  let valueSource: string | undefined;
  let hasSelfClosingValue = false;
  let cursor = start;

  while (cursor < end) {
    const tagStart = xml.indexOf("<", cursor);
    if (tagStart === -1) {
      break;
    }
    if (tagStart >= end) {
      break;
    }

    const nextCode = xml.charCodeAt(tagStart + 1);
    if (nextCode === 47 || nextCode === 33 || nextCode === 63) {
      cursor = tagStart + 1;
      continue;
    }

    const tagOpenEnd = xml.indexOf(">", tagStart + 1);
    if (tagOpenEnd === -1) {
      break;
    }
    if (tagOpenEnd >= end) {
      break;
    }

    const nameStart = tagStart + 1;
    let nameEnd = nameStart;
    while (nameEnd < tagOpenEnd) {
      const code = xml.charCodeAt(nameEnd);
      if (code === 47 || code === 62 || isXmlWhitespaceCode(code)) {
        break;
      }
      nameEnd += 1;
    }

    const tagName = xml.slice(nameStart, nameEnd);
    if (tagName === "f" || tagName === "v") {
      const selfClosing = isSelfClosingTagSource(xml.slice(nameEnd, tagOpenEnd));
      if (tagName === "v" && selfClosing) {
        hasSelfClosingValue = true;
      } else if (!selfClosing) {
        const closeTag = `</${tagName}>`;
        const closeStart = xml.indexOf(closeTag, tagOpenEnd + 1);
        if (closeStart !== -1 && closeStart < end) {
          const text = xml.slice(tagOpenEnd + 1, closeStart);
          if (tagName === "f") {
            formulaSource = text;
          } else {
            valueSource = text;
          }
          cursor = closeStart + closeTag.length;
          continue;
        }
      }
    }

    cursor = tagOpenEnd + 1;
  }

  return { formulaSource, hasSelfClosingValue, valueSource };
}

function columnLabelToNumberFromAddress(address: string): number {
  let value = 0;
  let index = 0;

  while (index < address.length) {
    let characterCode = address.charCodeAt(index);
    if (characterCode === 36) {
      index += 1;
      continue;
    }

    if (characterCode >= 97 && characterCode <= 122) {
      characterCode -= 32;
    }

    if (characterCode < 65 || characterCode > 90) {
      break;
    }

    value = value * 26 + (characterCode - 64);
    index += 1;
  }

  if (value === 0) {
    throw new XlsxError(`Invalid cell address: ${address}`);
  }

  return value;
}

function isXmlWhitespaceCode(code: number): boolean {
  return code === 9 || code === 10 || code === 13 || code === 32;
}

function isXmlAttributeBoundaryCode(code: number): boolean {
  return code === 47 || isXmlWhitespaceCode(code);
}

const ROW_CLOSE_TAG = "</row>";
const CELL_CLOSE_TAG = "</c>";
