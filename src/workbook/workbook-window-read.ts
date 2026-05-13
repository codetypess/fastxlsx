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
import { decodeXmlText, serializeAttributes, getXmlAttr } from "../utils/xml.js";
import {
    cleanTagAttributesSource,
    isSelfClosingTagSource,
    isXmlWhitespaceCode,
    parseCellColumnNumberFast,
    readXmlAttrFast,
    scanValueWindowCellFast,
} from "./workbook-window-read-xml.js";

interface ReadState {
    manifest?: WorkbookManifest;
    baseSheetCaches: Map<string, BaseSheetReadCache>;
    fullSheetCaches: Map<string, SheetReadCache>;
}

interface SharedFormulaAnchor {
    columnNumber: number;
    formula: string;
    rowNumber: number;
}

interface SheetReadCache {
    base: BaseSheetReadCache;
    autoFilter: AutoFilterDefinition | null;
    columnDefinitions: ColumnWindowDefinition[];
    comments: SheetComment[];
    freezePane: FreezePane | null;
    mergedRanges: WindowRange[];
    sharedFormulaAnchors: Map<string, SharedFormulaAnchor>;
    usedBounds: UsedBounds | null;
}

interface BaseSheetReadCache {
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

interface WindowViewport {
    clampedEndColumn: number;
    clampedEndRow: number;
    clampedRange: string;
    clampedStartColumn: number;
    clampedStartRow: number;
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
    options: SheetWindowReadOptions
): SheetWindowSnapshot {
    assertSheetWindowReadOptions(options);

    const cache = resolveFullSheetReadCache(workbook, sheet);
    const resolvedWindow = resolveWindowReadBounds(options, cache.usedBounds);

    if (!resolvedWindow.viewport) {
        return buildEmptyWindowSnapshot(
            sheet.name,
            resolvedWindow.requestedRange,
            resolvedWindow.sheetRange,
            resolvedWindow.rowCount,
            resolvedWindow.columnCount,
            cache
        );
    }
    const viewport = resolvedWindow.viewport;

    const alignmentCache = new Map<number, CellStyleAlignment | null>();
    const rowMetadata = collectRowWindowMetadata(
        workbook,
        cache.base.rowInfos,
        viewport.clampedStartRow,
        viewport.clampedEndRow,
        alignmentCache
    );
    const columnMetadata = collectColumnWindowMetadata(
        workbook,
        cache.columnDefinitions,
        viewport.clampedStartColumn,
        viewport.clampedEndColumn,
        alignmentCache
    );
    const cellResult = readWindowCells(
        workbook,
        cache,
        viewport.clampedStartRow,
        viewport.clampedEndRow,
        viewport.clampedStartColumn,
        viewport.clampedEndColumn,
        alignmentCache
    );

    return {
        autoFilter: cloneAutoFilterDefinition(cache.autoFilter),
        cellAlignments: cellResult.cellAlignments,
        cells: cellResult.cells,
        clampedRange: viewport.clampedRange,
        comments: filterCommentsInWindow(
            cache.comments,
            viewport.clampedStartRow,
            viewport.clampedEndRow,
            viewport.clampedStartColumn,
            viewport.clampedEndColumn
        ),
        columnAlignments: columnMetadata.columnAlignments,
        columnCount: resolvedWindow.columnCount,
        columnStyleIds: columnMetadata.columnStyleIds,
        columnWidths: columnMetadata.columnWidths,
        freezePane: cloneFreezePane(cache.freezePane),
        hiddenColumns: columnMetadata.hiddenColumns,
        hiddenRows: rowMetadata.hiddenRows,
        mergedRanges: cache.mergedRanges
            .filter((range) =>
                rangesIntersect(
                    range,
                    viewport.clampedStartRow,
                    viewport.clampedEndRow,
                    viewport.clampedStartColumn,
                    viewport.clampedEndColumn
                )
            )
            .map((range) => range.ref),
        requestedRange: resolvedWindow.requestedRange,
        rowAlignments: rowMetadata.rowAlignments,
        rowCount: resolvedWindow.rowCount,
        rowHeights: rowMetadata.rowHeights,
        rowStyleIds: rowMetadata.rowStyleIds,
        sheetName: sheet.name,
        sheetRange: resolvedWindow.sheetRange,
    };
}

export function readWorkbookSheetValueWindow(
    workbook: Workbook,
    sheet: Sheet,
    options: SheetWindowReadOptions
): SheetValueWindowSnapshot {
    assertSheetWindowReadOptions(options);

    const cache = resolveBaseSheetReadCache(workbook, sheet);
    const resolvedWindow = resolveWindowReadBounds(options, cache.usedBounds);

    if (!resolvedWindow.viewport) {
        return buildEmptyValueWindowSnapshot(
            sheet.name,
            resolvedWindow.requestedRange,
            resolvedWindow.sheetRange,
            resolvedWindow.rowCount,
            resolvedWindow.columnCount
        );
    }
    const viewport = resolvedWindow.viewport;

    const cellResult = readValueWindowCells(
        workbook,
        cache,
        viewport.clampedStartRow,
        viewport.clampedEndRow,
        viewport.clampedStartColumn,
        viewport.clampedEndColumn
    );

    return {
        cells: cellResult.cells,
        clampedRange: viewport.clampedRange,
        columnCount: resolvedWindow.columnCount,
        requestedRange: resolvedWindow.requestedRange,
        rowCount: resolvedWindow.rowCount,
        sheetName: sheet.name,
        sheetRange: resolvedWindow.sheetRange,
    };
}

function resolveWindowReadBounds(
    options: SheetWindowReadOptions,
    usedBounds: UsedBounds | null
): {
    columnCount: number;
    requestedRange: string;
    rowCount: number;
    sheetRange: string | null;
    viewport: WindowViewport | null;
} {
    const requestedRange = formatRangeRef(
        options.startRow,
        options.startColumn,
        options.endRow,
        options.endColumn
    );
    const rowCount = usedBounds?.maxRow ?? 0;
    const columnCount = usedBounds?.maxColumn ?? 0;
    const sheetRange = usedBounds
        ? formatRangeRef(
              usedBounds.minRow,
              usedBounds.minColumn,
              usedBounds.maxRow,
              usedBounds.maxColumn
          )
        : null;

    if (
        rowCount === 0 ||
        columnCount === 0 ||
        options.startRow > rowCount ||
        options.startColumn > columnCount
    ) {
        return { columnCount, requestedRange, rowCount, sheetRange, viewport: null };
    }

    const clampedStartRow = Math.max(1, Math.min(options.startRow, rowCount));
    const clampedEndRow = Math.max(1, Math.min(options.endRow, rowCount));
    const clampedStartColumn = Math.max(1, Math.min(options.startColumn, columnCount));
    const clampedEndColumn = Math.max(1, Math.min(options.endColumn, columnCount));
    if (clampedStartRow > clampedEndRow || clampedStartColumn > clampedEndColumn) {
        return { columnCount, requestedRange, rowCount, sheetRange, viewport: null };
    }

    return {
        columnCount,
        requestedRange,
        rowCount,
        sheetRange,
        viewport: {
            clampedEndColumn,
            clampedEndRow,
            clampedRange: formatRangeRef(
                clampedStartRow,
                clampedStartColumn,
                clampedEndRow,
                clampedEndColumn
            ),
            clampedStartColumn,
            clampedStartRow,
        },
    };
}

function getOrCreateReadState(workbook: Workbook): ReadState {
    let state = readStates.get(workbook);
    if (!state) {
        state = { baseSheetCaches: new Map(), fullSheetCaches: new Map() };
        readStates.set(workbook, state);
    }
    return state;
}

function resolveBaseSheetReadCache(workbook: Workbook, sheet: Sheet): BaseSheetReadCache {
    const state = getOrCreateReadState(workbook);
    let cache = state.baseSheetCaches.get(sheet.path);
    if (!cache) {
        cache = buildBaseSheetReadCache(workbook, sheet).base;
        state.baseSheetCaches.set(sheet.path, cache);
    }

    return cache;
}

function resolveFullSheetReadCache(workbook: Workbook, sheet: Sheet): SheetReadCache {
    const state = getOrCreateReadState(workbook);
    let cache = state.fullSheetCaches.get(sheet.path);
    if (cache) {
        return cache;
    }

    let baseCache = state.baseSheetCaches.get(sheet.path);
    let sharedFormulaAnchors: Map<string, SharedFormulaAnchor>;
    if (!baseCache) {
        const built = buildBaseSheetReadCache(workbook, sheet, {
            collectSharedFormulaAnchors: true,
        });
        baseCache = built.base;
        sharedFormulaAnchors = built.sharedFormulaAnchors;
        state.baseSheetCaches.set(sheet.path, baseCache);
    } else {
        sharedFormulaAnchors = collectSharedFormulaAnchors(baseCache);
    }

    cache = buildFullSheetReadCache(sheet, baseCache, sharedFormulaAnchors);
    state.fullSheetCaches.set(sheet.path, cache);
    return cache;
}

function buildFullSheetReadCache(
    sheet: Sheet,
    base: BaseSheetReadCache,
    sharedFormulaAnchors: Map<string, SharedFormulaAnchor>
): SheetReadCache {
    const sheetXml = base.sheetXml;
    const comments = sheet.getComments();

    return {
        base,
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
        sharedFormulaAnchors,
        usedBounds: mergeUsedBounds(base.usedBounds, calculateCommentBounds(comments)),
    };
}

function buildBaseSheetReadCache(
    workbook: Workbook,
    sheet: Sheet,
    options: { collectSharedFormulaAnchors?: boolean } = {}
): {
    base: BaseSheetReadCache;
    sharedFormulaAnchors: Map<string, SharedFormulaAnchor>;
} {
    const sheetXml = workbook.readEntryText(sheet.path);
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
            options.collectSharedFormulaAnchors
                ? buildSheetRowReadInfo(
                      workbook,
                      sheetXml,
                      innerStart,
                      innerEnd,
                      rowMetadata.attributesSource,
                      rowMetadata.selfClosing,
                      rowNumber,
                      sharedFormulaAnchors
                  )
                : buildBaseSheetRowReadInfo(
                      sheetXml,
                      innerStart,
                      innerEnd,
                      rowMetadata.attributesSource,
                      rowMetadata.selfClosing,
                      rowNumber
                  )
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
        base: {
            rowInfos,
            sheetXml,
            usedBounds: calculateUsedBounds(rowInfos),
        },
        sharedFormulaAnchors,
    };
}

function locateSheetData(sheetXml: string): {
    sheetDataInnerEnd: number;
    sheetDataInnerStart: number;
} {
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

function buildBaseSheetRowReadInfo(
    sheetXml: string,
    innerStart: number,
    innerEnd: number,
    attributesSource: string,
    selfClosing: boolean,
    rowNumber: number
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

    const scratchCells: ValueScratchCellInfo[] = [];
    let cellCursor = innerStart;
    let previousColumnNumber = 0;
    let cellsAreSorted = true;

    let scannedCell = scanValueWindowCellFast(sheetXml, innerEnd, cellCursor);
    while (scannedCell) {
        scratchCells.push({
            columnNumber: scannedCell.columnNumber,
            logical: scannedCell.logical,
        });

        if (scannedCell.columnNumber < previousColumnNumber) {
            cellsAreSorted = false;
        }
        previousColumnNumber = scannedCell.columnNumber;
        cellCursor = scannedCell.nextCursor;
        scannedCell = scanValueWindowCellFast(sheetXml, innerEnd, cellCursor);
    }

    if (!cellsAreSorted) {
        scratchCells.sort((left, right) => left.columnNumber - right.columnNumber);
    }

    const analysis = analyzeValueScratchCells(scratchCells);
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

function buildSheetRowReadInfo(
    workbook: Workbook,
    sheetXml: string,
    innerStart: number,
    innerEnd: number,
    attributesSource: string,
    selfClosing: boolean,
    rowNumber: number,
    sharedFormulaAnchors: Map<string, SharedFormulaAnchor>
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

        const columnNumber = parseCellColumnNumberFast(cellMetadata.addressSource);
        const snapshotResult = buildCellSnapshot(
            workbook,
            cellMetadata.rawType,
            cellMetadata.styleIdText === undefined ? null : Number(cellMetadata.styleIdText),
            sheetXml,
            cellMetadata.selfClosing ? cellEnd : cellOpenTagEnd + 1,
            cellMetadata.selfClosing ? cellEnd : cellEnd
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
            sharedFormulaAnchors
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
    sharedFormulaAnchors: Map<string, SharedFormulaAnchor>
): void {
    const sharedFormula = parseSharedFormulaMetadata(formulaAttributesSource);
    if (!sharedFormula || snapshot.formula === null || snapshot.formula.length === 0) {
        return;
    }

    sharedFormulaAnchors.set(sharedFormula.sharedIndex, {
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

function collectSharedFormulaAnchors(base: BaseSheetReadCache): Map<string, SharedFormulaAnchor> {
    const sharedFormulaAnchors = new Map<string, SharedFormulaAnchor>();

    for (const rowInfo of base.rowInfos) {
        let cellCursor = rowInfo.innerStart;
        while (cellCursor < rowInfo.innerEnd) {
            const cellStart = base.sheetXml.indexOf("<c", cellCursor);
            if (cellStart === -1 || cellStart >= rowInfo.innerEnd) {
                break;
            }

            const cellOpenTagEnd = base.sheetXml.indexOf(">", cellStart + 2);
            if (cellOpenTagEnd === -1 || cellOpenTagEnd > rowInfo.innerEnd) {
                break;
            }

            const cellMetadata = parseCellTagMetadata(
                base.sheetXml.slice(cellStart + 2, cellOpenTagEnd)
            );
            const cellEnd = cellMetadata.selfClosing
                ? cellOpenTagEnd + 1
                : base.sheetXml.indexOf("</c>", cellOpenTagEnd + 1);
            if (!cellMetadata.addressSource || cellEnd === -1) {
                cellCursor = cellOpenTagEnd + 1;
                continue;
            }

            const columnNumber = parseCellColumnNumberFast(cellMetadata.addressSource);
            const anchor = extractSharedFormulaAnchorFast(
                base.sheetXml,
                cellMetadata.selfClosing ? cellEnd : cellOpenTagEnd + 1,
                cellMetadata.selfClosing ? cellEnd : cellEnd
            );
            if (anchor) {
                sharedFormulaAnchors.set(anchor.sharedIndex, {
                    columnNumber,
                    formula: anchor.formula,
                    rowNumber: rowInfo.rowNumber,
                });
            }

            cellCursor = cellMetadata.selfClosing ? cellEnd : cellEnd + "</c>".length;
        }
    }

    return sharedFormulaAnchors;
}

function parseSharedFormulaMetadata(
    formulaAttributesSource: string
): { hasRef: boolean; sharedIndex: string } | null {
    if (formulaAttributesSource.length === 0) {
        return null;
    }

    if (readXmlAttrFast(formulaAttributesSource, "t") !== "shared") {
        return null;
    }

    const sharedIndex = readXmlAttrFast(formulaAttributesSource, "si");
    if (!sharedIndex) {
        return null;
    }

    return {
        hasRef: readXmlAttrFast(formulaAttributesSource, "ref") !== undefined,
        sharedIndex,
    };
}

function extractSharedFormulaAnchorFast(
    xml: string,
    innerStart: number,
    innerEnd: number
): { formula: string; sharedIndex: string } | null {
    const formulaStart = xml.indexOf("<f", innerStart);
    if (formulaStart === -1 || formulaStart >= innerEnd) {
        return null;
    }

    const nextCode = xml.charCodeAt(formulaStart + 2);
    if (nextCode !== 62 && nextCode !== 47 && !isXmlWhitespaceCode(nextCode)) {
        return null;
    }

    const formulaOpenTagEnd = xml.indexOf(">", formulaStart + 2);
    if (formulaOpenTagEnd === -1 || formulaOpenTagEnd >= innerEnd) {
        return null;
    }

    const formulaTagSource = xml.slice(formulaStart + 2, formulaOpenTagEnd);
    if (isSelfClosingTagSource(formulaTagSource)) {
        return null;
    }

    const sharedFormula = parseSharedFormulaMetadata(cleanTagAttributesSource(formulaTagSource));
    if (!sharedFormula) {
        return null;
    }

    const formulaCloseStart = xml.indexOf("</f>", formulaOpenTagEnd + 1);
    if (formulaCloseStart === -1 || formulaCloseStart >= innerEnd) {
        return null;
    }

    const formulaSource = xml.slice(formulaOpenTagEnd + 1, formulaCloseStart);
    if (formulaSource.length === 0) {
        return null;
    }

    return {
        formula: decodeXmlText(formulaSource),
        sharedIndex: sharedFormula.sharedIndex,
    };
}

function collectRowWindowMetadata(
    workbook: Workbook,
    rowInfos: SheetRowReadInfo[],
    startRow: number,
    endRow: number,
    alignmentCache: Map<number, CellStyleAlignment | null>
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
    alignmentCache: Map<number, CellStyleAlignment | null>
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
    alignmentCache: Map<number, CellStyleAlignment | null>
): CellWindowReadResult {
    const cellAlignments: Record<string, CellStyleAlignment> = {};
    const cells: SheetWindowCell[] = [];
    const base = cache.base;

    for (const rowInfo of base.rowInfos) {
        if (rowInfo.rowNumber < startRow) {
            continue;
        }
        if (rowInfo.rowNumber > endRow) {
            break;
        }

        let cellCursor = rowInfo.innerStart;
        while (cellCursor < rowInfo.innerEnd) {
            const cellStart = base.sheetXml.indexOf("<c", cellCursor);
            if (cellStart === -1 || cellStart >= rowInfo.innerEnd) {
                break;
            }

            const cellOpenTagEnd = base.sheetXml.indexOf(">", cellStart + 2);
            if (cellOpenTagEnd === -1 || cellOpenTagEnd > rowInfo.innerEnd) {
                break;
            }

            const cellMetadata = parseCellTagMetadata(
                base.sheetXml.slice(cellStart + 2, cellOpenTagEnd)
            );
            const cellEnd = cellMetadata.selfClosing
                ? cellOpenTagEnd + 1
                : base.sheetXml.indexOf("</c>", cellOpenTagEnd + 1);
            if (!cellMetadata.addressSource || cellEnd === -1) {
                cellCursor = cellOpenTagEnd + 1;
                continue;
            }

            const address = cellMetadata.addressSource.toUpperCase();
            const columnNumber = parseCellColumnNumberFast(cellMetadata.addressSource);
            cellCursor = cellMetadata.selfClosing ? cellEnd : cellEnd + "</c>".length;
            if (columnNumber < startColumn || columnNumber > endColumn) {
                continue;
            }

            const snapshotResult = buildCellSnapshot(
                workbook,
                cellMetadata.rawType,
                cellMetadata.styleIdText === undefined ? null : Number(cellMetadata.styleIdText),
                base.sheetXml,
                cellMetadata.selfClosing ? cellEnd : cellOpenTagEnd + 1,
                cellMetadata.selfClosing ? cellEnd : cellEnd
            );
            const snapshot = resolveSharedFormulaSnapshot(
                snapshotResult.formulaAttributesSource,
                snapshotResult.snapshot,
                rowInfo.rowNumber,
                columnNumber,
                cache.sharedFormulaAnchors
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
    cache: BaseSheetReadCache,
    startRow: number,
    endRow: number,
    startColumn: number,
    endColumn: number
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
        let scannedCell = scanValueWindowCellFast(cache.sheetXml, rowInfo.innerEnd, cellCursor);
        while (scannedCell) {
            cellCursor = scannedCell.nextCursor;
            if (scannedCell.columnNumber < startColumn || scannedCell.columnNumber > endColumn) {
                scannedCell = scanValueWindowCellFast(cache.sheetXml, rowInfo.innerEnd, cellCursor);
                continue;
            }

            if (!scannedCell.logical) {
                scannedCell = scanValueWindowCellFast(cache.sheetXml, rowInfo.innerEnd, cellCursor);
                continue;
            }

            const snapshotResult = buildCellSnapshot(
                workbook,
                scannedCell.rawType,
                null,
                cache.sheetXml,
                scannedCell.innerStart,
                scannedCell.innerEnd
            );
            cells.push({
                address: scannedCell.addressSource.toUpperCase(),
                columnNumber: scannedCell.columnNumber,
                rowNumber: rowInfo.rowNumber,
                value: snapshotResult.snapshot.value,
            });
            scannedCell = scanValueWindowCellFast(cache.sheetXml, rowInfo.innerEnd, cellCursor);
        }
    }

    return { cells };
}

function resolveSharedFormulaSnapshot(
    formulaAttributesSource: string,
    snapshot: CellSnapshot,
    rowNumber: number,
    columnNumber: number,
    sharedFormulaAnchors: Map<string, SharedFormulaAnchor>
): CellSnapshot {
    const sharedFormula = parseSharedFormulaMetadata(formulaAttributesSource);
    if (
        !sharedFormula ||
        sharedFormula.hasRef ||
        (snapshot.formula !== null && snapshot.formula.length > 0)
    ) {
        return snapshot;
    }

    const anchor = sharedFormulaAnchors.get(sharedFormula.sharedIndex);
    if (!anchor) {
        return snapshot;
    }

    return {
        ...snapshot,
        formula: translateFormulaReferences(
            anchor.formula,
            columnNumber - anchor.columnNumber,
            rowNumber - anchor.rowNumber
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
    cache: SheetReadCache
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
    columnCount: number
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

function analyzeValueScratchCells(cells: ValueScratchCellInfo[]): {
    maxColumnNumber: number;
    used: boolean;
} {
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
    endColumn: number
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
    alignmentCache: Map<number, CellStyleAlignment | null>
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

function cloneAutoFilterDefinition(
    definition: AutoFilterDefinition | null
): AutoFilterDefinition | null {
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
                        conditions: definition.sortState.conditions.map((condition) => ({
                            ...condition,
                        })),
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
