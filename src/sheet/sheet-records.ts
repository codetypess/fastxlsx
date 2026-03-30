import { XlsxError } from "../errors.js";
import type { CellValue } from "../types.js";
import { makeCellAddress } from "./sheet-address.js";

export function buildHeaderMap(headers: CellValue[]): Map<string, number> {
  const headerMap = new Map<string, number>();

  headers.forEach((value, index) => {
    if (typeof value === "string" && value.length > 0 && !headerMap.has(value)) {
      headerMap.set(value, index + 1);
    }
  });

  return headerMap;
}

export function writeRecordValues(
  rowNumber: number,
  record: Record<string, CellValue>,
  headerMap: Map<string, number>,
  replaceMissingKeys: boolean,
  setCell: (address: string, value: CellValue) => void,
): void {
  const keys = Object.keys(record);

  for (const key of keys) {
    if (!headerMap.has(key)) {
      throw new XlsxError(`Header not found: ${key}`);
    }
  }

  if (replaceMissingKeys) {
    for (const [header, columnNumber] of headerMap) {
      const nextValue = Object.hasOwn(record, header) ? record[header] ?? null : null;
      setCell(makeCellAddress(rowNumber, columnNumber), nextValue);
    }
    return;
  }

  for (const key of keys) {
    const columnNumber = headerMap.get(key);
    if (!columnNumber) {
      continue;
    }

    setCell(makeCellAddress(rowNumber, columnNumber), record[key] ?? null);
  }
}
