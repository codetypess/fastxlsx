import type { CellSnapshot } from "../types.js";
import type { Workbook } from "../workbook.js";
import type { SheetIndex } from "../sheet/sheet-index.js";
import { getXmlAttr } from "../utils/xml.js";

interface MutationCell {
    columnNumber: number;
    rowNumber: number;
    kind: "set" | "delete";
    snapshot: CellSnapshot;
}
export interface PreparedReadCell {
    columnNumber: number;
    snapshot: CellSnapshot;
}
export interface PreparedSheetReadModel {
    sheetXml: string;
    rows: Array<{ rowNumber: number; attributesSource: string; cells: PreparedReadCell[] }>;
}
const models = new WeakMap<Workbook, Map<string, PreparedSheetReadModel>>();

export function invalidatePreparedSheetReads(workbook: Workbook, path: string, all: boolean): void {
    if (all) {
        models.delete(workbook);
    } else {
        models.get(workbook)?.delete(path);
    }
}

export function getPreparedSheetReadModel(
    workbook: Workbook,
    path: string
): PreparedSheetReadModel | undefined {
    return models.get(workbook)?.get(path);
}

// Transfer semantic snapshots, not offset indexes, to the full-window reader.
// A later raw write invalidates this model before it can be observed again.
export function publishPreparedSheetReadModel(
    workbook: Workbook,
    path: string,
    sheetXml: string,
    index: SheetIndex,
    mutations: readonly MutationCell[] = []
): void {
    // Overwriting shared anchors can change followers that were expanded from
    // the previous anchor. Reparse this uncommon case from final XML.
    if (
        mutations.some((mutation) => {
            const cell = index.rows.get(mutation.rowNumber)?.cellsByColumn[mutation.columnNumber];
            return cell && getXmlAttr(cell.formulaAttributesSource, "t") === "shared";
        })
    ) {
        return;
    }
    const patches = new Map<number, Map<number, MutationCell>>();
    for (const mutation of mutations) {
        let row = patches.get(mutation.rowNumber);
        if (!row) {
            row = new Map();
            patches.set(mutation.rowNumber, row);
        }
        row.set(mutation.columnNumber, mutation);
    }
    const rowNumbers = [...new Set([...index.rowNumbers, ...patches.keys()])].sort((a, b) => a - b);
    const rows: PreparedSheetReadModel["rows"] = [];
    for (const rowNumber of rowNumbers) {
        const original = index.rows.get(rowNumber);
        const patch = patches.get(rowNumber);
        const cells: PreparedReadCell[] = [];
        for (const cell of original?.cells ?? []) {
            const mutation = patch?.get(cell.columnNumber);
            if (mutation?.kind !== "delete") {
                cells.push({
                    columnNumber: cell.columnNumber,
                    snapshot: mutation?.snapshot ?? cell.snapshot,
                });
            }
            patch?.delete(cell.columnNumber);
        }
        for (const mutation of patch?.values() ?? []) {
            if (mutation.kind === "set") {
                cells.push({ columnNumber: mutation.columnNumber, snapshot: mutation.snapshot });
            }
        }
        if (patch?.size) {
            cells.sort((a, b) => a.columnNumber - b.columnNumber);
        }
        if (original || cells.length) {
            rows.push({
                rowNumber,
                attributesSource: original?.attributesSource ?? 'r="' + rowNumber + '"',
                cells,
            });
        }
    }
    let sheets = models.get(workbook);
    if (!sheets) {
        sheets = new Map();
        models.set(workbook, sheets);
    }
    sheets.set(path, { sheetXml, rows });
}
