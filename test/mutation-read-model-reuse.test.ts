import test from "node:test";
import assert from "node:assert/strict";
import { Workbook } from "../src/index.ts";
import { getPreparedSheetReadModel } from "../src/workbook/workbook-mutation-read.ts";

const full = { startRow: 1, endRow: 1048576, startColumn: 1, endColumn: 16384 };
function verify(workbook: Workbook) {
    const entries = workbook.toEntries();
    const fresh = Workbook.fromEntries(entries);
    for (const name of workbook.getSheetNames()) {
        for (const bounds of [
            full,
            { startRow: 2, endRow: 5, startColumn: 2, endColumn: 4 },
            { startRow: 20, endRow: 30, startColumn: 10, endColumn: 12 },
        ]) {
            assert.deepEqual(
                workbook.readSheetWindow(name, bounds),
                fresh.readSheetWindow(name, bounds),
                name
            );
            assert.deepEqual(
                workbook.readSheetValueWindow(name, bounds),
                fresh.readSheetValueWindow(name, bounds),
                name
            );
        }
    }
    assert.deepEqual(workbook.toEntries(), entries, "reading must not change archive parts");
}
function fixture() {
    const workbook = Workbook.create("Data");
    const sheet = workbook.getSheet("Data");
    workbook.addSheet("References").setFormula("A1", "Data!A3", { cachedValue: "three" });
    workbook.batch(() => {
        sheet.setRange("A1", [
            ["one", 1],
            ["two", 2],
            ["three", 3],
        ]);
        sheet.setFormula("C3", "A3", { cachedValue: "three" });
        sheet.setComment("A3", "note", { author: "tester" });
        sheet.setRowHeight(3, 24);
        sheet.setColumnWidth(2, 18);
        sheet.freezePane(1, 1);
        sheet.setAlignment("L25", { horizontal: "right" });
    });
    return Workbook.fromEntries(workbook.toEntries());
}

test("finalized values and sparse overlays share read data without reparsing cells", () => {
    const workbook = fixture();
    const sheet = workbook.getSheet("Data");
    workbook.batch(() => {
        sheet.cell("A1").setValue("updated");
        sheet.cell("D9").setValue(true);
        sheet.cell("E9").setValue(null);
        sheet.cell("C3").setFormula("SUM(B1:B3)", { cachedValue: 6 });
        sheet.deleteCell("B2");
    });
    const prepared = getPreparedSheetReadModel(workbook, sheet.path);
    assert.ok(prepared);
    assert.ok(prepared.rows.every((row) => !Object.hasOwn(row, "start")));
    verify(workbook);
    const read = workbook.readEntryText.bind(workbook);
    workbook.readEntryText = (path) => {
        assert.notEqual(path, sheet.path, "warm prepared windows must not read worksheet XML");
        return read(path);
    };
    assert.equal(
        workbook.readSheetWindow("Data", full).cells.find((cell) => cell.address === "A1")?.value,
        "updated"
    );
});

for (const operation of ["insertRow", "deleteRow", "insertColumn", "deleteColumn"] as const) {
    test(`${operation} reads match reopened archives, including comments and references`, () => {
        const workbook = fixture();
        const sheet = workbook.getSheet("Data");
        workbook.batch(() => {
            sheet[operation](2, 1);
            sheet.cell("F2").setValue("last overlay");
        });
        assert.ok(getPreparedSheetReadModel(workbook, sheet.path));
        verify(workbook);
        workbook.batch(() => {
            sheet.cell("F2").setValue("again");
            sheet.deleteCell("A1");
        });
        verify(workbook);
    });
}

test("nested batches and intermediate reads never reuse stale prepared state", () => {
    const workbook = fixture();
    const sheet = workbook.getSheet("Data");
    workbook.batch(() => sheet.cell("A1").setValue("first"));
    workbook.readSheetWindow("Data", full);
    workbook.batch(() => {
        sheet.cell("A1").setValue("middle");
        assert.equal(workbook.readSheetWindow("Data", full).cells[0].value, "middle");
        workbook.batch(() => sheet.insertRow(2));
        sheet.cell("A2").setValue("last");
        const version = workbook.getSheetReadVersion("Data");
        assert.equal(
            workbook.readSheetWindow("Data", full).cells.find((cell) => cell.address === "A2")
                ?.value,
            "last"
        );
        sheet.cell("A2").setValue("after version");
        assert.notEqual(workbook.getSheetReadVersion("Data"), version);
    });
    verify(workbook);
});

test("read tokens isolate ordinary sheet writes and invalidate shared metadata conservatively", () => {
    const workbook = fixture();
    const data = workbook.getSheetReadVersion("Data");
    const refs = workbook.getSheetReadVersion("References");
    workbook.batch(() => workbook.getSheet("Data").cell("A1").setValue("changed"));
    assert.notEqual(workbook.getSheetReadVersion("Data"), data);
    assert.equal(workbook.getSheetReadVersion("References"), refs);
    workbook.batch(() => workbook.getSheet("Data").insertRow(2));
    assert.notEqual(workbook.getSheetReadVersion("References"), refs);
    const before = workbook.getSheetReadVersion("References");
    workbook.getSheet("Data").setComment("A1", "new comment");
    assert.notEqual(workbook.getSheetReadVersion("References"), before);
    assert.throws(() => workbook.getSheetReadVersion("missing"));
    verify(workbook);
});

test("raw worksheet replacement invalidates prepared values", () => {
    const workbook = fixture();
    const sheet = workbook.getSheet("Data");
    workbook.batch(() => sheet.cell("A1").setValue("cached"));
    const version = workbook.getSheetReadVersion("Data");
    assert.ok(getPreparedSheetReadModel(workbook, sheet.path));
    workbook.writeEntryText(
        sheet.path,
        workbook.readEntryText(sheet.path).replace("cached", "replaced")
    );
    assert.equal(getPreparedSheetReadModel(workbook, sheet.path), undefined);
    assert.notEqual(workbook.getSheetReadVersion("Data"), version);
    verify(workbook);
    assert.equal(workbook.readSheetWindow("Data", full).cells[0].value, "replaced");
});

test("shared-string replacements invalidate prepared semantics and tokens", () => {
    const workbook = Workbook.create("Data");
    const path = workbook.getSheet("Data").path;
    workbook.writeEntryText(
        "xl/_rels/workbook.xml.rels",
        workbook
            .readEntryText("xl/_rels/workbook.xml.rels")
            .replace(
                "</Relationships>",
                '<Relationship Id="rIdShared" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/></Relationships>'
            )
    );
    workbook.writeEntryText(
        "xl/sharedStrings.xml",
        '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><si><t>original</t></si></sst>'
    );
    workbook.writeEntryText(
        path,
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" t="s"><v>0</v></c></row></sheetData></worksheet>'
    );
    const sheet = workbook.getSheet("Data");
    workbook.batch(() => sheet.cell("B1").setValue("new"));
    assert.ok(getPreparedSheetReadModel(workbook, path));
    const before = workbook.getSheetReadVersion("Data");
    workbook.writeEntryText(
        "xl/sharedStrings.xml",
        workbook.readEntryText("xl/sharedStrings.xml").replace("original", "new shared")
    );
    assert.equal(getPreparedSheetReadModel(workbook, path), undefined);
    assert.notEqual(workbook.getSheetReadVersion("Data"), before);
    assert.equal(workbook.readSheetWindow("Data", full).cells[0].value, "new shared");
    verify(workbook);
});

test("shared-anchor overwrite uses final XML semantics", () => {
    const workbook = Workbook.create("Data");
    const path = workbook.getSheet("Data").path;
    workbook.writeEntryText(
        path,
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>' +
            '<row r="1"><c r="A1"><f t="shared" ref="A1:A2" si="0">B1</f><v>1</v></c></row>' +
            '<row r="2"><c r="A2"><f t="shared" si="0"/><v>2</v></c></row></sheetData></worksheet>'
    );
    const fresh = Workbook.fromEntries(workbook.toEntries());
    fresh.batch(() => fresh.getSheet("Data").cell("A1").setFormula("C1", { cachedValue: 3 }));
    assert.equal(getPreparedSheetReadModel(fresh, path), undefined);
    verify(fresh);
});

test("blank styles outside logical dimensions survive prepared and cold windows", () => {
    const workbook = fixture();
    workbook.batch(() => workbook.getSheet("Data").cell("A1").setValue("trigger"));
    const window = workbook.readSheetWindow("Data", {
        startRow: 20,
        endRow: 30,
        startColumn: 10,
        endColumn: 12,
    });
    assert.equal(window.rowCount, 3);
    assert.equal(window.cellAlignments.L25?.horizontal, "right");
    assert.deepEqual(window.cells, []);
    assert.equal(window.clampedRange, null);
    verify(workbook);
});

test("returned windows are detached from cached cells, errors, comments and alignments", () => {
    const workbook = fixture();
    workbook.batch(() => workbook.getSheet("Data").cell("A1").setValue("stable"));
    const before = workbook.readSheetWindow("Data", full);
    const mutable = workbook.readSheetWindow("Data", full);
    mutable.cells[0].value = "tampered";
    mutable.cellAlignments.L25.horizontal = "left";
    mutable.comments[0].text = "tampered";
    assert.deepEqual(workbook.readSheetWindow("Data", full), before);
    verify(workbook);
});

function sharedFixture() {
    const workbook = Workbook.create("Data");
    workbook.writeEntryText(
        workbook.getSheet("Data").path,
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>' +
            '<row r="1"><c r="A1"><f t="shared" ref="A1:A3" si="0">B1</f><v>1</v></c></row>' +
            '<row r="2"><c r="A2"><f t="shared" si="0"/><v>2</v></c></row>' +
            '<row r="3"><c r="A3"><f t="shared" si="0"/></c></row>' +
            "</sheetData></worksheet>"
    );
    return Workbook.fromEntries(workbook.toEntries());
}

for (const operation of ["insertRow", "deleteRow", "insertColumn", "deleteColumn"] as const) {
    test(`${operation} reuses shared followers after unrelated pending edits`, () => {
        const workbook = sharedFixture();
        const sheet = workbook.getSheet("Data");
        workbook.batch(() => {
            sheet[operation](2, 1);
            sheet.cell("D5").setValue("independent");
        });
        assert.ok(getPreparedSheetReadModel(workbook, sheet.path));
        verify(workbook);
        const previous = workbook.readSheetWindow("Data", full);
        workbook.batch(() => sheet.cell("D6").setValue("another edit"));
        verify(workbook);
        assert.equal(
            previous.cells.find((cell) => cell.address === "D6"),
            undefined
        );
    });
}

test("shared follower overwrites and anchor deletion conservatively use final XML", () => {
    for (const address of ["A1", "A2"]) {
        const workbook = sharedFixture();
        const sheet = workbook.getSheet("Data");
        workbook.batch(() => sheet.deleteCell(address));
        assert.equal(getPreparedSheetReadModel(workbook, sheet.path), undefined);
        verify(workbook);
    }
});

test("raw part removal invalidates prepared cells and read versions", () => {
    const workbook = fixture();
    const sheet = workbook.getSheet("Data");
    workbook.batch(() => sheet.cell("A1").setValue("cached"));
    const before = workbook.getSheetReadVersion("Data");
    workbook.removeEntry(sheet.path);
    assert.equal(getPreparedSheetReadModel(workbook, sheet.path), undefined);
    assert.notEqual(workbook.getSheetReadVersion("Data"), before);
    assert.throws(() => workbook.readSheetWindow("Data", full), /Entry not found/);
});

test("returned error details are detached from the prepared model", () => {
    const workbook = Workbook.create("Data");
    workbook.writeEntryText(
        workbook.getSheet("Data").path,
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>' +
            '<row r="1"><c r="A1" t="e"><v>#DIV/0!</v></c></row></sheetData></worksheet>'
    );
    const fresh = Workbook.fromEntries(workbook.toEntries());
    fresh.batch(() => fresh.getSheet("Data").cell("B1").setValue("overlay"));
    const before = fresh.readSheetWindow("Data", full);
    const mutable = fresh.readSheetWindow("Data", full);
    assert.ok(mutable.cells[0].error);
    Object.assign(mutable.cells[0].error, { code: "tampered" });
    assert.deepEqual(fresh.readSheetWindow("Data", full), before);
    verify(fresh);
});
