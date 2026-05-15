import { resolve } from "node:path";
import { fileURLToPath } from "node:url";

import { Workbook } from "../src/index.js";

type WindowReadMode = "window" | "value-window";

interface WindowScenario {
    endColumn: number;
    endRow: number;
    label: string;
    startColumn: number;
    startRow: number;
}

interface WindowScenarioResult {
    averageMs: number;
    clampedRange: string | null;
    label: string;
    requestedRange: string;
    runs: number[];
    visitedCells: number;
}

interface WindowSequenceResult {
    averageMs: number;
    mode: WindowReadMode;
    runs: number[];
    steps: WindowScenarioResult[];
}

interface WindowReadDepthBenchmarkSummary {
    columnCount: number;
    deepValueWindow: WindowScenarioResult[];
    deepWindow: WindowScenarioResult[];
    rowCount: number;
    sequentialValueWindow: WindowSequenceResult;
    sequentialWindow: WindowSequenceResult;
    sheet: string;
}

const DEEP_WINDOW_SCENARIOS: WindowScenario[] = [
    { endColumn: 20, endRow: 200, label: "rows-1-200", startColumn: 1, startRow: 1 },
    { endColumn: 20, endRow: 5200, label: "rows-5000-5200", startColumn: 1, startRow: 5000 },
    { endColumn: 20, endRow: 10200, label: "rows-10000-10200", startColumn: 1, startRow: 10000 },
    { endColumn: 20, endRow: 20200, label: "rows-20000-20200", startColumn: 1, startRow: 20000 },
];

const SEQUENTIAL_WINDOW_SCENARIOS: WindowScenario[] = [
    { endColumn: 20, endRow: 8060, label: "rows-8000-8060", startColumn: 1, startRow: 8000 },
    { endColumn: 20, endRow: 8100, label: "rows-8040-8100", startColumn: 1, startRow: 8040 },
    { endColumn: 20, endRow: 8140, label: "rows-8080-8140", startColumn: 1, startRow: 8080 },
];

export async function runWindowReadDepthBenchmark(options: {
    filePath?: string;
    iterations?: number;
    sheetName?: string;
} = {}): Promise<{
    file: string;
    iterations: number;
    synthetic: WindowReadDepthBenchmarkSummary;
    workbook: WindowReadDepthBenchmarkSummary;
}> {
    const filePath = options.filePath ?? resolve(process.cwd(), "res/monster.xlsx");
    const iterations = options.iterations ?? 20;
    const workbook = await Workbook.open(filePath);
    const sheet = options.sheetName ? workbook.getSheet(options.sheetName) : selectTargetSheet(workbook);

    return {
        file: filePath,
        iterations,
        synthetic: benchmarkSheet(buildSyntheticDeepBenchmarkWorkbook().getSheet("SyntheticDeepWindowBenchmark"), iterations),
        workbook: benchmarkSheet(sheet, iterations),
    };
}

function benchmarkSheet(
    sheet: ReturnType<Workbook["getSheets"]>[number],
    iterations: number
): WindowReadDepthBenchmarkSummary {
    const warmWindow = sheet.readWindow(DEEP_WINDOW_SCENARIOS[0]!);
    sheet.readValueWindow(DEEP_WINDOW_SCENARIOS[0]!);

    return {
        columnCount: warmWindow.columnCount,
        deepValueWindow: benchmarkScenarioSet(sheet, DEEP_WINDOW_SCENARIOS, iterations, "value-window"),
        deepWindow: benchmarkScenarioSet(sheet, DEEP_WINDOW_SCENARIOS, iterations, "window"),
        rowCount: warmWindow.rowCount,
        sequentialValueWindow: benchmarkScenarioSequence(
            sheet,
            SEQUENTIAL_WINDOW_SCENARIOS,
            iterations,
            "value-window"
        ),
        sequentialWindow: benchmarkScenarioSequence(
            sheet,
            SEQUENTIAL_WINDOW_SCENARIOS,
            iterations,
            "window"
        ),
        sheet: sheet.name,
    };
}

function benchmarkScenarioSet(
    sheet: ReturnType<Workbook["getSheets"]>[number],
    scenarios: WindowScenario[],
    iterations: number,
    mode: WindowReadMode
): WindowScenarioResult[] {
    return scenarios.map((scenario) => benchmarkSingleScenario(sheet, scenario, iterations, mode));
}

function benchmarkSingleScenario(
    sheet: ReturnType<Workbook["getSheets"]>[number],
    scenario: WindowScenario,
    iterations: number,
    mode: WindowReadMode
): WindowScenarioResult {
    const runs: number[] = [];
    let clampedRange: string | null = null;
    let requestedRange = "";
    let visitedCells = 0;

    for (let index = 0; index < iterations; index += 1) {
        const startedAt = performance.now();
        const window = readScenario(sheet, scenario, mode);
        runs.push(Number((performance.now() - startedAt).toFixed(3)));
        clampedRange = window.clampedRange;
        requestedRange = window.requestedRange;
        visitedCells = window.cells.length;
    }

    return {
        averageMs: Number((runs.reduce((sum, value) => sum + value, 0) / runs.length).toFixed(3)),
        clampedRange,
        label: scenario.label,
        requestedRange,
        runs,
        visitedCells,
    };
}

function benchmarkScenarioSequence(
    sheet: ReturnType<Workbook["getSheets"]>[number],
    scenarios: WindowScenario[],
    iterations: number,
    mode: WindowReadMode
): WindowSequenceResult {
    const runs: number[] = [];
    let steps: WindowScenarioResult[] = [];

    for (let index = 0; index < iterations; index += 1) {
        const startedAt = performance.now();
        const currentSteps: WindowScenarioResult[] = [];

        for (const scenario of scenarios) {
            const window = readScenario(sheet, scenario, mode);
            currentSteps.push({
                averageMs: 0,
                clampedRange: window.clampedRange,
                label: scenario.label,
                requestedRange: window.requestedRange,
                runs: [],
                visitedCells: window.cells.length,
            });
        }

        steps = currentSteps;
        runs.push(Number((performance.now() - startedAt).toFixed(3)));
    }

    return {
        averageMs: Number((runs.reduce((sum, value) => sum + value, 0) / runs.length).toFixed(3)),
        mode,
        runs,
        steps,
    };
}

function readScenario(
    sheet: ReturnType<Workbook["getSheets"]>[number],
    scenario: WindowScenario,
    mode: WindowReadMode
) {
    if (mode === "window") {
        return sheet.readWindow(scenario);
    }

    return sheet.readValueWindow(scenario);
}

function selectTargetSheet(workbook: Workbook) {
    const sheets = workbook.getSheets();
    const [firstSheet] = sheets;
    if (!firstSheet) {
        throw new Error("Workbook has no worksheets to benchmark");
    }

    let targetSheet = firstSheet;
    let maxPhysicalCellCount = firstSheet.getPhysicalCellEntries().length;

    for (let index = 1; index < sheets.length; index += 1) {
        const sheet = sheets[index]!;
        const physicalCellCount = sheet.getPhysicalCellEntries().length;
        if (physicalCellCount > maxPhysicalCellCount) {
            targetSheet = sheet;
            maxPhysicalCellCount = physicalCellCount;
        }
    }

    return targetSheet;
}

function buildSyntheticDeepBenchmarkWorkbook(): Workbook {
    const workbook = Workbook.create("SyntheticDeepWindowBenchmark");
    const sheet = workbook.getSheet("SyntheticDeepWindowBenchmark");

    sheet.batch((currentSheet) => {
        for (let rowNumber = 1; rowNumber <= 22000; rowNumber += 1) {
            currentSheet.setCell(rowNumber, 1, `R${rowNumber}`);
            currentSheet.setCell(rowNumber, 2, rowNumber);
            currentSheet.setFormula(rowNumber, 3, `B${rowNumber}*2`, {
                cachedValue: rowNumber * 2,
            });
            currentSheet.setCell(rowNumber, 4, rowNumber % 2 === 0);
        }
    });

    return workbook;
}

async function main(): Promise<void> {
    const [filePathArg, iterationsArg, sheetNameArg] = process.argv.slice(2);
    const result = await runWindowReadDepthBenchmark({
        filePath: filePathArg ? resolve(process.cwd(), filePathArg) : undefined,
        iterations: iterationsArg ? Number(iterationsArg) : undefined,
        sheetName: sheetNameArg,
    });

    console.log(JSON.stringify(result, null, 2));
}

if (process.argv[1] && fileURLToPath(import.meta.url) === resolve(process.argv[1])) {
    await main();
}
