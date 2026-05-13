import type {
  AutoFilterColumn,
  AutoFilterCondition,
  AutoFilterDefinition,
  CellEntry,
  SheetComment,
  CellStyleAlignment,
  CellSnapshot,
  CellValue,
  DateGroupItem,
  DefinedName,
  FreezePane,
  SheetValueWindowCell,
  SheetValueWindowSnapshot,
  SheetWindowCell,
  SheetWindowReadOptions,
  SheetWindowSnapshot,
  WorkbookManifest,
  WorkbookSheetManifest,
} from "../types.js";
import { XlsxError } from "../errors.js";
import type { Sheet } from "../sheet.js";
import type { Workbook } from "../workbook.js";
import {
  buildCellSnapshot,
  parseCellAddressFast,
  parseCellTagMetadata,
  parseRowTagMetadata,
} from "../sheet/sheet-index.js";
import { parseMergedRanges } from "../sheet/sheet-merge.js";
import { formatRangeRef, numberToColumnLabel, parseRangeRef } from "../sheet/sheet-address.js";
import { calculateCommentBounds, filterCommentsInWindow } from "../sheet/sheet-comments.js";
import { parseSheetFreezePane } from "../sheet/sheet-view-metadata.js";
import { parseWorksheetAutoFilterDefinition } from "../sheet/sheet-auto-filter.js";
import {
  parseColumnDefinitions,
  parseRowHeight,
  parseRowHidden,
  parseRowStyleId,
} from "../sheet/sheet-style-xml.js";
import { assertColumnNumber, assertRowNumber } from "../sheet/sheet-validation.js";
import { translateFormulaReferences } from "../sheet/sheet-structure.js";
import { parseAttributes, serializeAttributes, getXmlAttr } from "../utils/xml.js";

interface ReadState {
  manifest?: WorkbookManifest;
  sheetCaches: Map<string, SheetReadCache>;
  valueSheetCaches: Map<string, ValueSheetReadCache>;
}

interface SharedFormulaAnchor {
  columnNumber: number;
  formula: string;
  rowNumber: number;
}

interface SheetReadCache {
  autoFilter: AutoFilterDefinition | null;
  columnDefinitions: ColumnWindowDefinition[];
  comments: SheetComment[];
  freezePane: FreezePane | null;
  mergedRanges: WindowRange[];
  rowInfos: SheetRowReadInfo[];
  sheetXml: string;
  sharedFormulaAnchors: Map<string, SharedFormulaAnchor>;
  usedBounds: UsedBounds | null;
}

interface ValueSheetReadCache {
  rowInfos: SheetRowReadInfo[];
  sheetXml: string;
  usedBounds: UsedBounds | null;
}

interface SheetRowReadInfo {
  attributesSource: string;
  innerEnd: number;
  innerStart: number;
  maxColumnNumber: number;
  minColumnNumber: number;
  rowNumber: number;
  used: boolean;
}

interface WindowRange {
  endColumn: number;
  endRow: number;
  ref: string;
  startColumn: number;
  startRow: number;
}

interface UsedBounds {
  maxColumn: number;
  maxRow: number;
  minColumn: number;
  minRow: number;
}

interface ScratchCellInfo {
  columnNumber: number;
  snapshot: CellSnapshot;
}

interface RowWindowMetadata {
  hiddenRows: number[];
  rowAlignments: Record<string, CellStyleAlignment>;
  rowHeights: Record<string, number>;
  rowStyleIds: Record<string, number>;
}

interface ColumnWindowDefinition {
  hidden: boolean;
  max: number;
  min: number;
  styleId: number | null;
  width: number | null;
}

interface ColumnWindowMetadata {
  columnAlignments: Record<string, CellStyleAlignment>;
  columnStyleIds: Record<string, number>;
  columnWidths: Record<string, number>;
  hiddenColumns: string[];
}

interface CellWindowReadResult {
  cellAlignments: Record<string, CellStyleAlignment>;
  cells: SheetWindowCell[];
}

interface ValueCellWindowReadResult {
  cells: SheetValueWindowCell[];
}

interface ValueScratchCellInfo {
  columnNumber: number;
  logical: boolean;
}

const readStates = new WeakMap<Workbook, ReadState>();

export function invalidateWorkbookReadCaches(workbook: Workbook): void {
  readStates.delete(workbook);
}

export function readWorkbookManifest(workbook: Workbook): WorkbookManifest {
  const state = getOrCreateReadState(workbook);
  if (!state.manifest) {
    const sheetNames = workbook.getSheetNames();
    const sheets: WorkbookSheetManifest[] = sheetNames.map((name) => ({
      name,
      visibility: workbook.getSheetVisibility(name),
    }));
    state.manifest = {
      activeSheetName: sheets.length === 0 ? null : workbook.getActiveSheet().name,
      definedNames: workbook.getDefinedNames().map(cloneDefinedName),
      sheetCount: sheets.length,
      sheets,
      visibleSheetCount: sheets.filter((sheet) => sheet.visibility === "visible").length,
    };
  }

  return {
    activeSheetName: state.manifest.activeSheetName,
    definedNames: state.manifest.definedNames.map(cloneDefinedName),
    sheetCount: state.manifest.sheetCount,
    sheets: state.manifest.sheets.map((sheet) => ({ ...sheet })),
    visibleSheetCount: state.manifest.visibleSheetCount,
  };
}

export function readWorkbookSheetWindow(
  workbook: Workbook,
  sheet: Sheet,
  options: SheetWindowReadOptions,
): SheetWindowSnapshot {
  assertSheetWindowReadOptions(options);

  const state = getOrCreateReadState(workbook);
  let cache = state.sheetCaches.get(sheet.path);
  if (!cache) {
    cache = buildSheetReadCache(workbook, sheet);
    state.sheetCaches.set(sheet.path, cache);
  }

  const requestedRange = formatRangeRef(options.startRow, options.startColumn, options.endRow, options.endColumn);
  const rowCount = cache.usedBounds?.maxRow ?? 0;
  const columnCount = cache.usedBounds?.maxColumn ?? 0;
  const sheetRange = cache.usedBounds
    ? formatRangeRef(
        cache.usedBounds.minRow,
        cache.usedBounds.minColumn,
        cache.usedBounds.maxRow,
        cache.usedBounds.maxColumn,
      )
    : null;

  if (rowCount === 0 || columnCount === 0) {
    return buildEmptyWindowSnapshot(sheet.name, requestedRange, sheetRange, rowCount, columnCount, cache);
  }

  if (options.startRow > rowCount || options.startColumn > columnCount) {
    return buildEmptyWindowSnapshot(sheet.name, requestedRange, sheetRange, rowCount, columnCount, cache);
  }

  const clampedStartRow = Math.max(1, Math.min(options.startRow, rowCount));
  const clampedEndRow = Math.max(1, Math.min(options.endRow, rowCount));
  const clampedStartColumn = Math.max(1, Math.min(options.startColumn, columnCount));
  const clampedEndColumn = Math.max(1, Math.min(options.endColumn, columnCount));
  if (clampedStartRow > clampedEndRow || clampedStartColumn > clampedEndColumn) {
    return buildEmptyWindowSnapshot(sheet.name, requestedRange, sheetRange, rowCount, columnCount, cache);
  }

  const alignmentCache = new Map<number, CellStyleAlignment | null>();
  const rowMetadata = collectRowWindowMetadata(
    workbook,
    cache.rowInfos,
    clampedStartRow,
    clampedEndRow,
    alignmentCache,
  );
  const columnMetadata = collectColumnWindowMetadata(
    workbook,
    cache.columnDefinitions,
    clampedStartColumn,
    clampedEndColumn,
    alignmentCache,
  );
  const cellResult = readWindowCells(
    workbook,
    cache,
    clampedStartRow,
    clampedEndRow,
    clampedStartColumn,
    clampedEndColumn,
    alignmentCache,
  );

  return {
    autoFilter: cloneAutoFilterDefinition(cache.autoFilter),
    cellAlignments: cellResult.cellAlignments,
    cells: cellResult.cells,
    clampedRange: formatRangeRef(clampedStartRow, clampedStartColumn, clampedEndRow, clampedEndColumn),
    comments: filterCommentsInWindow(cache.comments, clampedStartRow, clampedEndRow, clampedStartColumn, clampedEndColumn),
    columnAlignments: columnMetadata.columnAlignments,
    columnCount,
    columnStyleIds: columnMetadata.columnStyleIds,
    columnWidths: columnMetadata.columnWidths,
    freezePane: cloneFreezePane(cache.freezePane),
    hiddenColumns: columnMetadata.hiddenColumns,
    hiddenRows: rowMetadata.hiddenRows,
    mergedRanges: cache.mergedRanges
      .filter((range) => rangesIntersect(range, clampedStartRow, clampedEndRow, clampedStartColumn, clampedEndColumn))
      .map((range) => range.ref),
    requestedRange,
    rowAlignments: rowMetadata.rowAlignments,
    rowCount,
    rowHeights: rowMetadata.rowHeights,
    rowStyleIds: rowMetadata.rowStyleIds,
    sheetName: sheet.name,
    sheetRange,
  };
}

export function readWorkbookSheetValueWindow(
  workbook: Workbook,
  sheet: Sheet,
  options: SheetWindowReadOptions,
): SheetValueWindowSnapshot {
  assertSheetWindowReadOptions(options);

  const state = getOrCreateReadState(workbook);
  let cache = state.valueSheetCaches.get(sheet.path);
  if (!cache) {
    cache = buildSheetValueReadCache(workbook, sheet);
    state.valueSheetCaches.set(sheet.path, cache);
  }

  const requestedRange = formatRangeRef(options.startRow, options.startColumn, options.endRow, options.endColumn);
  const rowCount = cache.usedBounds?.maxRow ?? 0;
  const columnCount = cache.usedBounds?.maxColumn ?? 0;
  const sheetRange = cache.usedBounds
    ? formatRangeRef(
        cache.usedBounds.minRow,
        cache.usedBounds.minColumn,
        cache.usedBounds.maxRow,
        cache.usedBounds.maxColumn,
      )
    : null;

  if (rowCount === 0 || columnCount === 0) {
    return buildEmptyValueWindowSnapshot(sheet.name, requestedRange, sheetRange, rowCount, columnCount);
  }

  if (options.startRow > rowCount || options.startColumn > columnCount) {
    return buildEmptyValueWindowSnapshot(sheet.name, requestedRange, sheetRange, rowCount, columnCount);
  }

  const clampedStartRow = Math.max(1, Math.min(options.startRow, rowCount));
  const clampedEndRow = Math.max(1, Math.min(options.endRow, rowCount));
  const clampedStartColumn = Math.max(1, Math.min(options.startColumn, columnCount));
  const clampedEndColumn = Math.max(1, Math.min(options.endColumn, columnCount));
  if (clampedStartRow > clampedEndRow || clampedStartColumn > clampedEndColumn) {
    return buildEmptyValueWindowSnapshot(sheet.name, requestedRange, sheetRange, rowCount, columnCount);
  }

  const cellResult = readValueWindowCells(
    workbook,
    cache,
    clampedStartRow,
    clampedEndRow,
    clampedStartColumn,
    clampedEndColumn,
  );

  return {
    cells: cellResult.cells,
    clampedRange: formatRangeRef(clampedStartRow, clampedStartColumn, clampedEndRow, clampedEndColumn),
    columnCount,
    requestedRange,
    rowCount,
    sheetName: sheet.name,
    sheetRange,
  };
}

function getOrCreateReadState(workbook: Workbook): ReadState {
  let state = readStates.get(workbook);
  if (!state) {
    state = { sheetCaches: new Map(), valueSheetCaches: new Map() };
    readStates.set(workbook, state);
  }
  return state;
}

function buildSheetReadCache(workbook: Workbook, sheet: Sheet): SheetReadCache {
  const sheetXml = workbook.readEntryText(sheet.path);
  const comments = sheet.getComments();
  const sharedFormulaAnchors = new Map<string, SharedFormulaAnchor>();
  const rowInfos: SheetRowReadInfo[] = [];
  const { sheetDataInnerEnd, sheetDataInnerStart } = locateSheetData(sheetXml);
  let cursor = sheetDataInnerStart;
  let previousRowNumber = 0;
  let rowsAreSorted = true;

  while (cursor < sheetDataInnerEnd) {
    const rowStart = sheetXml.indexOf("<row", cursor);
    if (rowStart === -1 || rowStart >= sheetDataInnerEnd) {
      break;
    }

    const rowOpenTagEnd = sheetXml.indexOf(">", rowStart + 4);
    if (rowOpenTagEnd === -1 || rowOpenTagEnd >= sheetDataInnerEnd) {
      break;
    }

    const rowMetadata = parseRowTagMetadata(sheetXml.slice(rowStart + 4, rowOpenTagEnd));
    const rowEnd = rowMetadata.selfClosing
      ? rowOpenTagEnd + 1
      : sheetXml.indexOf("</row>", rowOpenTagEnd + 1);
    if (!rowMetadata.rowNumberText || rowEnd === -1) {
      cursor = rowOpenTagEnd + 1;
      continue;
    }

    const rowNumber = Number(rowMetadata.rowNumberText);
    const innerStart = rowMetadata.selfClosing ? rowEnd : rowOpenTagEnd + 1;
    const innerEnd = rowMetadata.selfClosing ? rowEnd : rowEnd;
    rowInfos.push(
      buildSheetRowReadInfo(
        workbook,
        sheetXml,
        innerStart,
        innerEnd,
        rowMetadata.attributesSource,
        rowMetadata.selfClosing,
        rowNumber,
        sharedFormulaAnchors,
      ),
    );

    if (rowNumber < previousRowNumber) {
      rowsAreSorted = false;
    }
    previousRowNumber = rowNumber;
    cursor = rowMetadata.selfClosing ? rowEnd : rowEnd + "</row>".length;
  }

  if (!rowsAreSorted) {
    rowInfos.sort((left, right) => left.rowNumber - right.rowNumber);
  }

  return {
    autoFilter: parseWorksheetAutoFilterDefinition(sheetXml),
    columnDefinitions: parseColumnDefinitions(sheetXml).map((definition) => ({
      hidden: parseColumnDefinitionHidden(definition.attributes),
      max: definition.max,
      min: definition.min,
      styleId: parseColumnDefinitionStyleId(definition.attributes),
      width: parseColumnDefinitionWidth(definition.attributes),
    })),
    comments,
    freezePane: parseSheetFreezePane(sheetXml),
    mergedRanges: parseMergedRanges(sheetXml).map((ref) => {
      const parsed = parseRangeRef(ref);
      return {
        endColumn: parsed.endColumn,
        endRow: parsed.endRow,
        ref,
        startColumn: parsed.startColumn,
        startRow: parsed.startRow,
      };
    }),
    rowInfos,
    sheetXml,
    sharedFormulaAnchors,
    usedBounds: mergeUsedBounds(calculateUsedBounds(rowInfos), calculateCommentBounds(comments)),
  };
}

function buildSheetValueReadCache(workbook: Workbook, sheet: Sheet): ValueSheetReadCache {
  const sheetXml = workbook.readEntryText(sheet.path);
  const rowInfos: SheetRowReadInfo[] = [];
  const { sheetDataInnerEnd, sheetDataInnerStart } = locateSheetData(sheetXml);
  let cursor = sheetDataInnerStart;
  let previousRowNumber = 0;
  let rowsAreSorted = true;

  while (cursor < sheetDataInnerEnd) {
    const rowStart = sheetXml.indexOf("<row", cursor);
    if (rowStart === -1 || rowStart >= sheetDataInnerEnd) {
      break;
    }

    const rowOpenTagEnd = sheetXml.indexOf(">", rowStart + 4);
    if (rowOpenTagEnd === -1 || rowOpenTagEnd >= sheetDataInnerEnd) {
      break;
    }

    const rowMetadata = parseRowTagMetadata(sheetXml.slice(rowStart + 4, rowOpenTagEnd));
    const rowEnd = rowMetadata.selfClosing
      ? rowOpenTagEnd + 1
      : sheetXml.indexOf("</row>", rowOpenTagEnd + 1);
    if (!rowMetadata.rowNumberText || rowEnd === -1) {
      cursor = rowOpenTagEnd + 1;
      continue;
    }

    const rowNumber = Number(rowMetadata.rowNumberText);
    const innerStart = rowMetadata.selfClosing ? rowEnd : rowOpenTagEnd + 1;
    const innerEnd = rowMetadata.selfClosing ? rowEnd : rowEnd;
    rowInfos.push(
      buildValueSheetRowReadInfo(
        sheetXml,
        innerStart,
        innerEnd,
        rowMetadata.selfClosing,
        rowNumber,
      ),
    );

    if (rowNumber < previousRowNumber) {
      rowsAreSorted = false;
    }
    previousRowNumber = rowNumber;
    cursor = rowMetadata.selfClosing ? rowEnd : rowEnd + "</row>".length;
  }

  if (!rowsAreSorted) {
    rowInfos.sort((left, right) => left.rowNumber - right.rowNumber);
  }

  return {
    rowInfos,
    sheetXml,
    usedBounds: calculateUsedBounds(rowInfos),
  };
}

function locateSheetData(sheetXml: string): { sheetDataInnerEnd: number; sheetDataInnerStart: number } {
  const sheetDataStart = sheetXml.indexOf("<sheetData");
  if (sheetDataStart === -1) {
    throw new XlsxError("Worksheet is missing <sheetData>");
  }

  const sheetDataOpenTagEnd = sheetXml.indexOf(">", sheetDataStart);
  const sheetDataCloseTagStart = sheetXml.indexOf("</sheetData>", sheetDataOpenTagEnd + 1);
  if (sheetDataOpenTagEnd === -1 || sheetDataCloseTagStart === -1) {
    throw new XlsxError("Worksheet is missing <sheetData>");
  }

  return {
    sheetDataInnerEnd: sheetDataCloseTagStart,
    sheetDataInnerStart: sheetDataOpenTagEnd + 1,
  };
}

function buildValueSheetRowReadInfo(
  sheetXml: string,
  innerStart: number,
  innerEnd: number,
  selfClosing: boolean,
  rowNumber: number,
): SheetRowReadInfo {
  if (selfClosing) {
    return {
      attributesSource: "",
      innerEnd,
      innerStart,
      maxColumnNumber: 0,
      minColumnNumber: 0,
      rowNumber,
      used: false,
    };
  }

  const scratchCells: ValueScratchCellInfo[] = [];
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

    const cellMetadata = parseCellTagMetadata(sheetXml.slice(cellStart + 2, cellOpenTagEnd));
    const cellEnd = cellMetadata.selfClosing
      ? cellOpenTagEnd + 1
      : sheetXml.indexOf("</c>", cellOpenTagEnd + 1);
    if (!cellMetadata.addressSource || cellEnd === -1) {
      cellCursor = cellOpenTagEnd + 1;
      continue;
    }

    const { columnNumber } = parseCellAddressFast(cellMetadata.addressSource.toUpperCase());
    scratchCells.push({
      columnNumber,
      logical: hasLogicalValueWindowCellContent(
        sheetXml,
        cellMetadata.selfClosing ? cellEnd : cellOpenTagEnd + 1,
        cellMetadata.selfClosing ? cellEnd : cellEnd,
        cellMetadata.rawType,
      ),
    });

    if (columnNumber < previousColumnNumber) {
      cellsAreSorted = false;
    }
    previousColumnNumber = columnNumber;
    cellCursor = cellMetadata.selfClosing ? cellEnd : cellEnd + "</c>".length;
  }

  if (!cellsAreSorted) {
    scratchCells.sort((left, right) => left.columnNumber - right.columnNumber);
  }

  const analysis = analyzeValueScratchCells(scratchCells);
  return {
    attributesSource: "",
    innerEnd,
    innerStart,
    maxColumnNumber: analysis.maxColumnNumber,
    minColumnNumber: scratchCells[0]?.columnNumber ?? 0,
    rowNumber,
    used: analysis.used,
  };
}

function buildSheetRowReadInfo(
  workbook: Workbook,
  sheetXml: string,
  innerStart: number,
  innerEnd: number,
  attributesSource: string,
  selfClosing: boolean,
  rowNumber: number,
  sharedFormulaAnchors: Map<string, SharedFormulaAnchor>,
): SheetRowReadInfo {
  if (selfClosing) {
    return {
      attributesSource,
      innerEnd,
      innerStart,
      maxColumnNumber: 0,
      minColumnNumber: 0,
      rowNumber,
      used: false,
    };
  }

  const scratchCells: ScratchCellInfo[] = [];
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

    const cellMetadata = parseCellTagMetadata(sheetXml.slice(cellStart + 2, cellOpenTagEnd));
    const cellEnd = cellMetadata.selfClosing
      ? cellOpenTagEnd + 1
      : sheetXml.indexOf("</c>", cellOpenTagEnd + 1);
    if (!cellMetadata.addressSource || cellEnd === -1) {
      cellCursor = cellOpenTagEnd + 1;
      continue;
    }

    const { columnNumber } = parseCellAddressFast(cellMetadata.addressSource.toUpperCase());
    const snapshotResult = buildCellSnapshot(
      workbook,
      cellMetadata.rawType,
      cellMetadata.styleIdText === undefined ? null : Number(cellMetadata.styleIdText),
      sheetXml,
      cellMetadata.selfClosing ? cellEnd : cellOpenTagEnd + 1,
      cellMetadata.selfClosing ? cellEnd : cellEnd,
    );

    scratchCells.push({
      columnNumber,
      snapshot: snapshotResult.snapshot,
    });
    registerSharedFormulaAnchor(
      snapshotResult.formulaAttributesSource,
      snapshotResult.snapshot,
      rowNumber,
      columnNumber,
      sharedFormulaAnchors,
    );

    if (columnNumber < previousColumnNumber) {
      cellsAreSorted = false;
    }
    previousColumnNumber = columnNumber;
    cellCursor = cellMetadata.selfClosing ? cellEnd : cellEnd + "</c>".length;
  }

  if (!cellsAreSorted) {
    scratchCells.sort((left, right) => left.columnNumber - right.columnNumber);
  }

  const analysis = analyzeScratchCells(scratchCells);
  return {
    attributesSource,
    innerEnd,
    innerStart,
    maxColumnNumber: analysis.maxColumnNumber,
    minColumnNumber: scratchCells[0]?.columnNumber ?? 0,
    rowNumber,
    used: analysis.used,
  };
}

function registerSharedFormulaAnchor(
  formulaAttributesSource: string,
  snapshot: CellSnapshot,
  rowNumber: number,
  columnNumber: number,
  sharedFormulaAnchors: Map<string, SharedFormulaAnchor>,
): void {
  if (formulaAttributesSource.length === 0) {
    return;
  }

  const attributes = new Map(parseAttributes(formulaAttributesSource));
  if (attributes.get("t") !== "shared") {
    return;
  }

  const sharedIndex = attributes.get("si");
  if (!sharedIndex || snapshot.formula === null || snapshot.formula.length === 0) {
    return;
  }

  sharedFormulaAnchors.set(sharedIndex, {
    columnNumber,
    formula: snapshot.formula,
    rowNumber,
  });
}

function analyzeScratchCells(cells: ScratchCellInfo[]): { maxColumnNumber: number; used: boolean } {
  if (cells.length === 0) {
    return { maxColumnNumber: 0, used: false };
  }

  let lastUsedIndex = -1;
  for (let index = cells.length - 1; index >= 0; index -= 1) {
    if (isLogicalCellEntry(cells[index]!.snapshot)) {
      lastUsedIndex = index;
      break;
    }
  }

  if (lastUsedIndex === -1) {
    let logicalMaxColumnNumber = cells[0]!.columnNumber;

    for (let index = 1; index < cells.length; index += 1) {
      const columnNumber = cells[index]!.columnNumber;
      if (columnNumber > logicalMaxColumnNumber + 1) {
        break;
      }

      logicalMaxColumnNumber = columnNumber;
    }

    return {
      maxColumnNumber: logicalMaxColumnNumber,
      used: false,
    };
  }

  let logicalMaxColumnNumber = cells[lastUsedIndex]!.columnNumber;
  for (let index = lastUsedIndex + 1; index < cells.length; index += 1) {
    const cell = cells[index]!;
    if (isLogicalCellEntry(cell.snapshot) || cell.columnNumber > logicalMaxColumnNumber + 1) {
      break;
    }

    logicalMaxColumnNumber = cell.columnNumber;
  }

  return {
    maxColumnNumber: logicalMaxColumnNumber,
    used: true,
  };
}

function calculateUsedBounds(rowInfos: SheetRowReadInfo[]): UsedBounds | null {
  let minRow = Number.POSITIVE_INFINITY;
  let maxRow = 0;
  let minColumn = Number.POSITIVE_INFINITY;
  let maxColumn = 0;
  let hasUsedBounds = false;

  for (const rowInfo of rowInfos) {
    if (rowInfo.maxColumnNumber > 0) {
      minColumn = Math.min(minColumn, rowInfo.minColumnNumber || rowInfo.maxColumnNumber);
      maxColumn = Math.max(maxColumn, rowInfo.maxColumnNumber);
    }

    if (rowInfo.used) {
      hasUsedBounds = true;
      minRow = Math.min(minRow, rowInfo.rowNumber);
      maxRow = Math.max(maxRow, rowInfo.rowNumber);
    }
  }

  return hasUsedBounds
    ? {
        maxColumn,
        maxRow,
        minColumn,
        minRow,
      }
    : null;
}

function mergeUsedBounds(left: UsedBounds | null, right: UsedBounds | null): UsedBounds | null {
  if (!left) {
    return right;
  }
  if (!right) {
    return left;
  }

  return {
    maxColumn: Math.max(left.maxColumn, right.maxColumn),
    maxRow: Math.max(left.maxRow, right.maxRow),
    minColumn: Math.min(left.minColumn, right.minColumn),
    minRow: Math.min(left.minRow, right.minRow),
  };
}

function collectRowWindowMetadata(
  workbook: Workbook,
  rowInfos: SheetRowReadInfo[],
  startRow: number,
  endRow: number,
  alignmentCache: Map<number, CellStyleAlignment | null>,
): RowWindowMetadata {
  const hiddenRows: number[] = [];
  const rowAlignments: Record<string, CellStyleAlignment> = {};
  const rowHeights: Record<string, number> = {};
  const rowStyleIds: Record<string, number> = {};

  for (const rowInfo of rowInfos) {
    if (rowInfo.rowNumber < startRow) {
      continue;
    }
    if (rowInfo.rowNumber > endRow) {
      break;
    }

    const styleId = parseRowStyleId(rowInfo.attributesSource);
    if (styleId !== null) {
      rowStyleIds[String(rowInfo.rowNumber)] = styleId;
      const alignment = resolveStyleAlignment(workbook, styleId, alignmentCache);
      if (alignment) {
        rowAlignments[String(rowInfo.rowNumber)] = alignment;
      }
    }

    const height = parseRowHeight(rowInfo.attributesSource);
    if (height !== null) {
      rowHeights[String(rowInfo.rowNumber)] = height;
    }

    if (parseRowHidden(rowInfo.attributesSource)) {
      hiddenRows.push(rowInfo.rowNumber);
    }
  }

  return { hiddenRows, rowAlignments, rowHeights, rowStyleIds };
}

function collectColumnWindowMetadata(
  workbook: Workbook,
  columnDefinitions: ColumnWindowDefinition[],
  startColumn: number,
  endColumn: number,
  alignmentCache: Map<number, CellStyleAlignment | null>,
): ColumnWindowMetadata {
  const columnAlignments: Record<string, CellStyleAlignment> = {};
  const columnStyleIds: Record<string, number> = {};
  const columnWidths: Record<string, number> = {};
  const hiddenColumns: string[] = [];

  for (let columnNumber = startColumn; columnNumber <= endColumn; columnNumber += 1) {
    let styleId: number | null = null;
    let width: number | null = null;
    let hidden = false;

    for (const definition of columnDefinitions) {
      if (columnNumber < definition.min || columnNumber > definition.max) {
        continue;
      }

      styleId = definition.styleId;
      width = definition.width;
      hidden = definition.hidden;
    }

    const columnLabel = numberToColumnLabel(columnNumber);
    if (styleId !== null) {
      columnStyleIds[columnLabel] = styleId;
      const alignment = resolveStyleAlignment(workbook, styleId, alignmentCache);
      if (alignment) {
        columnAlignments[columnLabel] = alignment;
      }
    }
    if (width !== null) {
      columnWidths[columnLabel] = width;
    }
    if (hidden) {
      hiddenColumns.push(columnLabel);
    }
  }

  return { columnAlignments, columnStyleIds, columnWidths, hiddenColumns };
}

function readWindowCells(
  workbook: Workbook,
  cache: SheetReadCache,
  startRow: number,
  endRow: number,
  startColumn: number,
  endColumn: number,
  alignmentCache: Map<number, CellStyleAlignment | null>,
): CellWindowReadResult {
  const cellAlignments: Record<string, CellStyleAlignment> = {};
  const cells: SheetWindowCell[] = [];

  for (const rowInfo of cache.rowInfos) {
    if (rowInfo.rowNumber < startRow) {
      continue;
    }
    if (rowInfo.rowNumber > endRow) {
      break;
    }

    let cellCursor = rowInfo.innerStart;
    while (cellCursor < rowInfo.innerEnd) {
      const cellStart = cache.sheetXml.indexOf("<c", cellCursor);
      if (cellStart === -1 || cellStart >= rowInfo.innerEnd) {
        break;
      }

      const cellOpenTagEnd = cache.sheetXml.indexOf(">", cellStart + 2);
      if (cellOpenTagEnd === -1 || cellOpenTagEnd > rowInfo.innerEnd) {
        break;
      }

      const cellMetadata = parseCellTagMetadata(cache.sheetXml.slice(cellStart + 2, cellOpenTagEnd));
      const cellEnd = cellMetadata.selfClosing
        ? cellOpenTagEnd + 1
        : cache.sheetXml.indexOf("</c>", cellOpenTagEnd + 1);
      if (!cellMetadata.addressSource || cellEnd === -1) {
        cellCursor = cellOpenTagEnd + 1;
        continue;
      }

      const address = cellMetadata.addressSource.toUpperCase();
      const { columnNumber } = parseCellAddressFast(address);
      cellCursor = cellMetadata.selfClosing ? cellEnd : cellEnd + "</c>".length;
      if (columnNumber < startColumn || columnNumber > endColumn) {
        continue;
      }

      const snapshotResult = buildCellSnapshot(
        workbook,
        cellMetadata.rawType,
        cellMetadata.styleIdText === undefined ? null : Number(cellMetadata.styleIdText),
        cache.sheetXml,
        cellMetadata.selfClosing ? cellEnd : cellOpenTagEnd + 1,
        cellMetadata.selfClosing ? cellEnd : cellEnd,
      );
      const snapshot = resolveSharedFormulaSnapshot(
        snapshotResult.formulaAttributesSource,
        snapshotResult.snapshot,
        rowInfo.rowNumber,
        columnNumber,
        cache.sharedFormulaAnchors,
      );
      if (!isLogicalCellEntry(snapshot)) {
        continue;
      }

      if (snapshot.styleId !== null) {
        const alignment = resolveStyleAlignment(workbook, snapshot.styleId, alignmentCache);
        if (alignment) {
          cellAlignments[address] = alignment;
        }
      }

      const cellEntry: CellEntry = {
        address,
        columnNumber,
        rowNumber: rowInfo.rowNumber,
        ...snapshot,
      };
      cells.push({
        ...cellEntry,
        displayValue: formatCellDisplayValue(snapshot),
      });
    }
  }

  return { cellAlignments, cells };
}

function readValueWindowCells(
  workbook: Workbook,
  cache: ValueSheetReadCache,
  startRow: number,
  endRow: number,
  startColumn: number,
  endColumn: number,
): ValueCellWindowReadResult {
  const cells: SheetValueWindowCell[] = [];

  for (const rowInfo of cache.rowInfos) {
    if (rowInfo.rowNumber < startRow) {
      continue;
    }
    if (rowInfo.rowNumber > endRow) {
      break;
    }

    let cellCursor = rowInfo.innerStart;
    while (cellCursor < rowInfo.innerEnd) {
      const cellStart = cache.sheetXml.indexOf("<c", cellCursor);
      if (cellStart === -1 || cellStart >= rowInfo.innerEnd) {
        break;
      }

      const cellOpenTagEnd = cache.sheetXml.indexOf(">", cellStart + 2);
      if (cellOpenTagEnd === -1 || cellOpenTagEnd > rowInfo.innerEnd) {
        break;
      }

      const cellMetadata = parseCellTagMetadata(cache.sheetXml.slice(cellStart + 2, cellOpenTagEnd));
      const cellEnd = cellMetadata.selfClosing
        ? cellOpenTagEnd + 1
        : cache.sheetXml.indexOf("</c>", cellOpenTagEnd + 1);
      if (!cellMetadata.addressSource || cellEnd === -1) {
        cellCursor = cellOpenTagEnd + 1;
        continue;
      }

      const address = cellMetadata.addressSource.toUpperCase();
      const { columnNumber } = parseCellAddressFast(address);
      cellCursor = cellMetadata.selfClosing ? cellEnd : cellEnd + "</c>".length;
      if (columnNumber < startColumn || columnNumber > endColumn) {
        continue;
      }

      const innerStart = cellMetadata.selfClosing ? cellEnd : cellOpenTagEnd + 1;
      const innerEnd = cellMetadata.selfClosing ? cellEnd : cellEnd;
      if (!hasLogicalValueWindowCellContent(cache.sheetXml, innerStart, innerEnd, cellMetadata.rawType)) {
        continue;
      }

      const snapshotResult = buildCellSnapshot(
        workbook,
        cellMetadata.rawType,
        null,
        cache.sheetXml,
        innerStart,
        innerEnd,
      );
      cells.push({
        address,
        columnNumber,
        rowNumber: rowInfo.rowNumber,
        value: snapshotResult.snapshot.value,
      });
    }
  }

  return { cells };
}

function resolveSharedFormulaSnapshot(
  formulaAttributesSource: string,
  snapshot: CellSnapshot,
  rowNumber: number,
  columnNumber: number,
  sharedFormulaAnchors: Map<string, SharedFormulaAnchor>,
): CellSnapshot {
  if (formulaAttributesSource.length === 0) {
    return snapshot;
  }

  const attributes = new Map(parseAttributes(formulaAttributesSource));
  if (attributes.get("t") !== "shared") {
    return snapshot;
  }

  const sharedIndex = attributes.get("si");
  if (!sharedIndex || attributes.has("ref") || (snapshot.formula !== null && snapshot.formula.length > 0)) {
    return snapshot;
  }

  const anchor = sharedFormulaAnchors.get(sharedIndex);
  if (!anchor) {
    return snapshot;
  }

  return {
    ...snapshot,
    formula: translateFormulaReferences(
      anchor.formula,
      columnNumber - anchor.columnNumber,
      rowNumber - anchor.rowNumber,
    ),
    type: "formula",
  };
}

function buildEmptyWindowSnapshot(
  sheetName: string,
  requestedRange: string,
  sheetRange: string | null,
  rowCount: number,
  columnCount: number,
  cache: SheetReadCache,
): SheetWindowSnapshot {
  return {
    autoFilter: cloneAutoFilterDefinition(cache.autoFilter),
    cellAlignments: {},
    cells: [],
    clampedRange: null,
    comments: [],
    columnAlignments: {},
    columnCount,
    columnStyleIds: {},
    columnWidths: {},
    freezePane: cloneFreezePane(cache.freezePane),
    hiddenColumns: [],
    hiddenRows: [],
    mergedRanges: [],
    requestedRange,
    rowAlignments: {},
    rowCount,
    rowHeights: {},
    rowStyleIds: {},
    sheetName,
    sheetRange,
  };
}

function buildEmptyValueWindowSnapshot(
  sheetName: string,
  requestedRange: string,
  sheetRange: string | null,
  rowCount: number,
  columnCount: number,
): SheetValueWindowSnapshot {
  return {
    cells: [],
    clampedRange: null,
    columnCount,
    requestedRange,
    rowCount,
    sheetName,
    sheetRange,
  };
}

function assertSheetWindowReadOptions(options: SheetWindowReadOptions): void {
  assertRowNumber(options.startRow);
  assertRowNumber(options.endRow);
  assertColumnNumber(options.startColumn);
  assertColumnNumber(options.endColumn);

  if (options.endRow < options.startRow) {
    throw new XlsxError(`Invalid row window: ${options.startRow}:${options.endRow}`);
  }
  if (options.endColumn < options.startColumn) {
    throw new XlsxError(`Invalid column window: ${options.startColumn}:${options.endColumn}`);
  }
}

function isLogicalCellEntry(cell: Pick<CellSnapshot, "formula" | "value">): boolean {
  return cell.formula !== null || cell.value !== null;
}

function hasLogicalValueWindowCellContent(
  xml: string,
  start: number,
  end: number,
  rawType: string | null,
): boolean {
  let cursor = start;

  while (cursor < end) {
    const tagStart = xml.indexOf("<", cursor);
    if (tagStart === -1 || tagStart >= end) {
      break;
    }

    const nextCode = xml.charCodeAt(tagStart + 1);
    if (nextCode === 47 || nextCode === 33 || nextCode === 63) {
      cursor = tagStart + 1;
      continue;
    }

    const tagOpenEnd = xml.indexOf(">", tagStart + 1);
    if (tagOpenEnd === -1 || tagOpenEnd >= end) {
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
    if (tagName === "f" || tagName === "v" || (rawType === "inlineStr" && tagName === "is")) {
      return true;
    }

    cursor = tagOpenEnd + 1;
  }

  return false;
}

function analyzeValueScratchCells(cells: ValueScratchCellInfo[]): { maxColumnNumber: number; used: boolean } {
  if (cells.length === 0) {
    return { maxColumnNumber: 0, used: false };
  }

  let lastUsedIndex = -1;
  for (let index = cells.length - 1; index >= 0; index -= 1) {
    if (cells[index]!.logical) {
      lastUsedIndex = index;
      break;
    }
  }

  if (lastUsedIndex === -1) {
    let logicalMaxColumnNumber = cells[0]!.columnNumber;

    for (let index = 1; index < cells.length; index += 1) {
      const columnNumber = cells[index]!.columnNumber;
      if (columnNumber > logicalMaxColumnNumber + 1) {
        break;
      }

      logicalMaxColumnNumber = columnNumber;
    }

    return {
      maxColumnNumber: logicalMaxColumnNumber,
      used: false,
    };
  }

  let logicalMaxColumnNumber = cells[lastUsedIndex]!.columnNumber;
  for (let index = lastUsedIndex + 1; index < cells.length; index += 1) {
    const cell = cells[index]!;
    if (cell.logical || cell.columnNumber > logicalMaxColumnNumber + 1) {
      break;
    }

    logicalMaxColumnNumber = cell.columnNumber;
  }

  return {
    maxColumnNumber: logicalMaxColumnNumber,
    used: true,
  };
}

function rangesIntersect(
  range: WindowRange,
  startRow: number,
  endRow: number,
  startColumn: number,
  endColumn: number,
): boolean {
  return !(
    range.endRow < startRow ||
    range.startRow > endRow ||
    range.endColumn < startColumn ||
    range.startColumn > endColumn
  );
}

function formatCellDisplayValue(cell: Pick<CellSnapshot, "error" | "value">): string | null {
  if (cell.error) {
    return cell.error.text;
  }

  if (cell.value === null) {
    return null;
  }

  if (typeof cell.value === "boolean") {
    return cell.value ? "TRUE" : "FALSE";
  }

  return String(cell.value);
}

function cloneDefinedName(definedName: DefinedName): DefinedName {
  return { ...definedName };
}

function resolveStyleAlignment(
  workbook: Workbook,
  styleId: number,
  alignmentCache: Map<number, CellStyleAlignment | null>,
): CellStyleAlignment | null {
  if (alignmentCache.has(styleId)) {
    return cloneCellAlignment(alignmentCache.get(styleId) ?? null);
  }

  const alignment = workbook.getStyle(styleId)?.alignment ?? null;
  alignmentCache.set(styleId, alignment ? cloneCellAlignment(alignment) : null);
  return cloneCellAlignment(alignment);
}

function cloneFreezePane(freezePane: FreezePane | null): FreezePane | null {
  return freezePane ? { ...freezePane } : null;
}

function cloneCellAlignment(alignment: CellStyleAlignment | null): CellStyleAlignment | null {
  return alignment ? { ...alignment } : null;
}

function cloneAutoFilterDefinition(definition: AutoFilterDefinition | null): AutoFilterDefinition | null {
  if (!definition) {
    return null;
  }

  return {
    columns: definition.columns.map(cloneAutoFilterColumn),
    range: definition.range,
    sortState:
      definition.sortState === null
        ? null
        : definition.sortState
          ? {
              conditions: definition.sortState.conditions.map((condition) => ({ ...condition })),
              range: definition.sortState.range,
            }
          : undefined,
  };
}

function cloneAutoFilterColumn(column: AutoFilterColumn): AutoFilterColumn {
  switch (column.kind) {
    case "values":
      return {
        columnNumber: column.columnNumber,
        includeBlank: column.includeBlank,
        kind: "values",
        values: [...column.values],
      };
    case "blank":
      return { ...column };
    case "custom":
      return {
        columnNumber: column.columnNumber,
        conditions: column.conditions.map(cloneAutoFilterCondition),
        join: column.join,
        kind: "custom",
      };
    case "dateGroup":
      return {
        columnNumber: column.columnNumber,
        items: column.items.map(cloneDateGroupItem),
        kind: "dateGroup",
      };
    case "color":
      return { ...column };
    case "dynamic":
      return { ...column };
    case "top10":
      return { ...column };
    case "icon":
      return { ...column };
  }
}

function cloneAutoFilterCondition(condition: AutoFilterCondition): AutoFilterCondition {
  return { ...condition };
}

function cloneDateGroupItem(item: DateGroupItem): DateGroupItem {
  return { ...item };
}

function parseColumnDefinitionStyleId(attributes: Array<[string, string]>): number | null {
  const styleText = attributes.find(([name]) => name === "style")?.[1];
  if (styleText === undefined) {
    return null;
  }

  const styleId = Number(styleText);
  return Number.isInteger(styleId) ? styleId : null;
}

function parseColumnDefinitionHidden(attributes: Array<[string, string]>): boolean {
  const hiddenText = getXmlAttr(serializeAttributes(attributes), "hidden");
  return hiddenText === "1" || hiddenText === "true";
}

function parseColumnDefinitionWidth(attributes: Array<[string, string]>): number | null {
  const widthText = getXmlAttr(serializeAttributes(attributes), "width");
  if (widthText === undefined) {
    return null;
  }

  const width = Number(widthText);
  return Number.isFinite(width) ? width : null;
}

function isXmlWhitespaceCode(code: number): boolean {
  return code === 9 || code === 10 || code === 13 || code === 32;
}
