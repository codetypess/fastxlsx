import { XlsxError } from "../errors.js";
import { parseCellAddressFast } from "./sheet-index.js";

export function splitCellAddress(address: string): { rowNumber: number; columnNumber: number } {
  return parseCellAddressFast(address);
}

export function columnLabelToNumber(label: string): number {
  let value = 0;

  for (const character of label.toUpperCase()) {
    value = value * 26 + (character.charCodeAt(0) - 64);
  }

  return value;
}

export function normalizeColumnNumber(column: number | string): number {
  if (typeof column === "number") {
    assertColumnNumber(column);
    return column;
  }

  if (!/^[A-Z]+$/i.test(column)) {
    throw new XlsxError(`Invalid column label: ${column}`);
  }

  return columnLabelToNumber(column.toUpperCase());
}

export function assertCellAddress(address: string): void {
  if (!/^[A-Z]+[1-9]\d*$/i.test(address)) {
    throw new XlsxError(`Invalid cell address: ${address}`);
  }
}

export function normalizeCellAddress(address: string): string {
  assertCellAddress(address);
  return address.toUpperCase();
}

export function normalizeRangeRef(range: string): string {
  const { startRow, endRow, startColumn, endColumn } = parseRangeRef(range);
  return formatRangeRef(startRow, startColumn, endRow, endColumn);
}

export function normalizeSqref(rangeList: string): string {
  const ranges = rangeList
    .trim()
    .split(/\s+/)
    .filter((range) => range.length > 0)
    .map((range) => normalizeRangeRef(range));

  if (ranges.length === 0) {
    throw new XlsxError(`Invalid sqref: ${rangeList}`);
  }

  return ranges.join(" ");
}

export function parseRangeRef(range: string): {
  startRow: number;
  endRow: number;
  startColumn: number;
  endColumn: number;
} {
  const normalizedRange = range.toUpperCase();
  const [startAddress, endAddress = normalizedRange] = normalizedRange.split(":");

  if (!startAddress || !endAddress) {
    throw new XlsxError(`Invalid range reference: ${range}`);
  }

  const start = splitCellAddress(startAddress);
  const end = splitCellAddress(endAddress);

  return {
    startRow: Math.min(start.rowNumber, end.rowNumber),
    endRow: Math.max(start.rowNumber, end.rowNumber),
    startColumn: Math.min(start.columnNumber, end.columnNumber),
    endColumn: Math.max(start.columnNumber, end.columnNumber),
  };
}

export function makeCellAddress(rowNumber: number, columnNumber: number): string {
  return `${numberToColumnLabel(columnNumber)}${rowNumber}`;
}

export function formatRangeRef(
  startRow: number,
  startColumn: number,
  endRow: number,
  endColumn: number,
): string {
  const startAddress = makeCellAddress(startRow, startColumn);
  const endAddress = makeCellAddress(endRow, endColumn);
  return startAddress === endAddress ? startAddress : `${startAddress}:${endAddress}`;
}

export function numberToColumnLabel(columnNumber: number): string {
  assertColumnNumber(columnNumber);

  let remaining = columnNumber;
  let label = "";

  while (remaining > 0) {
    const offset = (remaining - 1) % 26;
    label = String.fromCharCode(65 + offset) + label;
    remaining = Math.floor((remaining - 1) / 26);
  }

  return label;
}

export function compareCellAddresses(left: string, right: string): number {
  const leftCell = splitCellAddress(left);
  const rightCell = splitCellAddress(right);
  return leftCell.rowNumber - rightCell.rowNumber || leftCell.columnNumber - rightCell.columnNumber;
}

export function compareRangeRefs(left: string, right: string): number {
  const leftRange = parseRangeRef(left);
  const rightRange = parseRangeRef(right);

  return (
    leftRange.startRow - rightRange.startRow ||
    leftRange.startColumn - rightRange.startColumn ||
    leftRange.endRow - rightRange.endRow ||
    leftRange.endColumn - rightRange.endColumn
  );
}

function assertColumnNumber(columnNumber: number): void {
  if (!Number.isInteger(columnNumber) || columnNumber < 1) {
    throw new XlsxError(`Invalid column number: ${columnNumber}`);
  }
}
