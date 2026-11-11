import test from "node:test";
import assert from "node:assert/strict";
import { Workbook } from "../src/index.ts";

function fromRows(rows: string): Workbook {
    const workbook = Workbook.create("Data");
    workbook.writeEntryText(
        workbook.getSheet("Data").path,
        `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>${rows}</sheetData></worksheet>`
    );
    return Workbook.fromEntries(workbook.toEntries());
}

test("cached shared values and used bounds do not expand follower formulas", () => {
    const workbook = fromRows(`
        <row r="1"><c r="A1"><f t="shared" ref="A1:B2" si="0">C1+$D$4</f><v>5</v></c></row>
        <row r="2"><c r="B2"><f t="shared" si="0"/></c></row>`);
    const sheet = workbook.getSheet("Data");
    const snapshot = sheet.readCellSnapshot("B2");
    assert.equal(snapshot.type, "formula");
    assert.equal(snapshot.value, null);
    assert.equal(sheet.rowCount, 2);
    assert.equal(sheet.columnCount, 2);
    assert.equal(sheet.cell("A1").value, 5);
    assert.equal(sheet.cell("B2").value, null);
    assert.equal(typeof Object.getOwnPropertyDescriptor(snapshot, "formula")?.get, "function");
    assert.equal(snapshot.formula, "D2+$D$4");
    assert.equal(Object.getOwnPropertyDescriptor(snapshot, "formula")?.get, undefined);
    assert.equal(JSON.parse(JSON.stringify(snapshot)).formula, "D2+$D$4");
    assert.equal({ ...snapshot }.formula, "D2+$D$4");
});

test("late anchors, mixed references, and missing anchors retain their formula semantics", () => {
    const workbook = fromRows(`
        <row r="1"><c r="A1"><f t="shared" si="0"/><v>7</v></c><c r="C1"><f t="shared" si="missing"/><v>9</v></c></row>
        <row r="2"><c r="B2"><f t="shared" ref="A1:B2" si="0">C2+$D$4+E$5+$F6+&quot;C2&quot;</f><v>8</v></c></row>`);
    const sheet = workbook.getSheet("Data");
    const snapshot = sheet.readCellSnapshot("A1");
    assert.equal(sheet.cell("A1").value, 7);
    assert.equal({ ...snapshot }.formula, 'B1+$D$4+D$5+$F5+"C2"');
    assert.equal(sheet.cell("C1").formula, null);
    assert.equal(sheet.cell("C1").value, 9);
    const reopened = Workbook.fromUint8Array(workbook.toUint8Array()).getSheet("Data");
    assert.equal(reopened.cell("A1").formula, snapshot.formula);
});

test("shared formula snapshots support freezing and sealing before materialization", () => {
    const workbook = fromRows(`
        <row r="1"><c r="A1"><f t="shared" ref="A1:B3" si="0">C1</f></c></row>
        <row r="2"><c r="B2"><f t="shared" si="0"/></c></row>
        <row r="3"><c r="B3"><f t="shared" si="0"/></c></row>`);
    const sheet = workbook.getSheet("Data");
    const frozen = Object.freeze(sheet.readCellSnapshot("B2"));
    assert.equal(frozen.formula, "D2");
    assert.equal(JSON.parse(JSON.stringify(frozen)).formula, "D2");
    assert.throws(() => Object.assign(frozen, { formula: "changed" }), TypeError);
    const sealed = Object.seal(sheet.readCellSnapshot("B3"));
    sealed.formula = "changed";
    assert.equal(sealed.formula, "changed");
});

test("shared formulas follow structural changes while retained snapshots keep the old anchor", () => {
    const workbook = fromRows(`
        <row r="1"><c r="A1"><f t="shared" ref="A1:B2" si="0">C1+$D$4</f><v>5</v></c></row>
        <row r="2"><c r="B2"><f t="shared" si="0"/><v>6</v></c></row>`);
    const sheet = workbook.getSheet("Data");
    const retained = sheet.readCellSnapshot("B2");
    workbook.batch(() => {
        sheet.insertRow(1, 2);
        sheet.insertColumn(1);
        sheet.cell("Z20").setValue("grow dimension");
    });
    assert.equal(retained.formula, "D2+$D$4");
    assert.equal(sheet.cell("C4").formula, "E4+$E$6");
    assert.equal(sheet.cell("C4").value, 6);
    workbook.batch(() => {
        sheet.deleteRow(1, 2);
        sheet.deleteColumn(1);
        sheet.deleteCell("Y18");
    });
    const reopened = Workbook.fromUint8Array(workbook.toUint8Array()).getSheet("Data");
    assert.equal(reopened.cell("B2").formula, "D2+$D$4");
    assert.equal(reopened.rowCount, 2);
    assert.equal(reopened.columnCount, 2);
});

test("sparse batch dimensions match intermediate flushes through clearing and repopulation", () => {
    const actual = Workbook.create("Data");
    const expected = Workbook.create("Data");
    const phases = [
        (sheet: ReturnType<Workbook["getSheet"]>) => {
            for (const address of ["A30", "B30", "D30", "E30"]) sheet.cell(address).setValue(null);
            sheet.cell("C30").setValue("bridge");
            sheet.cell("Z40").setFormula("C30", { cachedValue: "formula" });
        },
        (sheet: ReturnType<Workbook["getSheet"]>) => {
            sheet.deleteCell("Z40");
            sheet.cell("C30").setValue(null);
            sheet.cell("A1").setValue("read inside batch");
            assert.equal(sheet.cell("A1").value, "read inside batch");
            sheet.getCellEntries();
            sheet.cell("B2").setValue("after read");
        },
        (sheet: ReturnType<Workbook["getSheet"]>) => {
            for (const cell of sheet.getCellEntries()) sheet.deleteCell(cell.address);
        },
        (sheet: ReturnType<Workbook["getSheet"]>) => sheet.cell("AA50").setValue("repopulate"),
    ];
    for (const phase of phases) {
        actual.batch(() => phase(actual.getSheet("Data")));
        expected.batch(() => {
            phase(expected.getSheet("Data"));
            // Force the older intermediate-read path before finalization.
            expected.getSheet("Data").getCellEntries();
        });
        assert.deepEqual(actual.toEntries(), expected.toEntries());
        assert.deepEqual(actual.getSheet("Data").getCellEntries(), expected.getSheet("Data").getCellEntries());
    }
});

test("large structural rewrites preserve row gaps and untouched parts", { timeout: 15000 }, () => {
    const rows = Array.from({ length: 12000 }, (_, index) =>
        `<row r="${index + 1}"><c r="A${index + 1}"><v>${index + 1}</v></c></row>`
    );
    const workbook = fromRows(rows.join("\n<!-- keep gap -->\n"));
    const sheet = workbook.getSheet("Data");
    workbook.writeEntryText("custom/preserved.xml", "<custom>untouched</custom>");
    workbook.batch(() => {
        sheet.insertRow(607, 20);
        sheet.cell("B607").setValue("inserted");
    });
    const reopened = Workbook.fromUint8Array(workbook.toUint8Array());
    const result = reopened.getSheet("Data");
    assert.equal(result.cell("A606").value, 606);
    assert.equal(result.cell("A607").value, null);
    assert.equal(result.cell("B607").value, "inserted");
    assert.equal(result.cell("A627").value, 607);
    assert.equal(result.cell("A12020").value, 12000);
    assert.equal(reopened.readEntryText("custom/preserved.xml"), "<custom>untouched</custom>");
    assert.equal(reopened.readEntryText(result.path).match(/<!-- keep gap -->/g)?.length, 11999);
});
