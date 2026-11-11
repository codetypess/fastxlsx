import test from "node:test";
import assert from "node:assert/strict";
import { mkdtemp, readFile, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { unzipSync } from "fflate";
import { Workbook } from "../src/index.ts";
import { Zip } from "../src/zip.ts";

for (const hasComments of [false, true]) {
    test(`cold comments and full windows do not build a cell index (comments=${hasComments})`, () => {
        const source = Workbook.create("Data");
        source.getSheet("Data").cell("A1").setValue("original");
        if (hasComments) source.getSheet("Data").setComment("C7", "Keep", { author: "tester" });
        const workbook = Workbook.fromUint8Array(source.toUint8Array());
        const sheet = workbook.getSheet("Data");
        const expected = source.getSheet("Data").getComments();
        (sheet as any).getSheetIndex = () => assert.fail("Comment reads must not index all cells");
        assert.deepEqual(sheet.getComments(), expected);
        const window = workbook.readSheetWindow("Data", {
            startRow: 1,
            endRow: 10,
            startColumn: 1,
            endColumn: 5,
        });
        assert.deepEqual(window.comments, expected);
        assert.equal(window.cells.find((cell) => cell.address === "A1")?.value, "original");
        assert.equal(window.rowCount, hasComments ? 7 : 1);
    });
}

test("comment reads flush pending cell writes and preserve shifted comments through nested batches", () => {
    const workbook = Workbook.create("Data");
    const sheet = workbook.getSheet("Data");
    sheet.setComment("C7", "Keep", { author: "tester" });
    workbook.batch(() => {
        sheet.cell("A1").setValue("first");
        assert.equal(sheet.getComments()[0]?.address, "C7");
        assert.match(workbook.readEntryText(sheet.path), /first/);
        workbook.batch(() => {
            sheet.insertRow(3, 2);
            sheet.cell("A2").setFormula("A1", { cachedValue: "second" });
            assert.equal(sheet.getComments()[0]?.address, "C9");
            assert.match(workbook.readEntryText(sheet.path), /second/);
        });
    });
    const reopened = Workbook.fromUint8Array(workbook.toUint8Array());
    assert.equal(reopened.getSheet("Data").cell("A2").formula, "A1");
    assert.equal(reopened.getSheet("Data").cell("A2").value, "second");
    assert.deepEqual(reopened.getSheet("Data").getComments(), sheet.getComments());
});

test("native ZIP preserves Unicode names, binary and empty entries, order and last duplicate values", () => {
    const binary = Uint8Array.from({ length: 65_537 }, (_, index) => index % 256);
    const entries = [
        { path: "汉字/🙂.bin", data: binary },
        { path: "empty", data: new Uint8Array() },
        { path: "folder/", data: new Uint8Array() },
        { path: "duplicate", data: new Uint8Array([1]) },
        { path: "duplicate", data: new Uint8Array([2, 3]) },
    ];
    function* source() {
        yield* entries;
    }
    const output = new Zip().writeArchiveData(source());
    assert.equal(output.constructor, Uint8Array);
    const extracted = unzipSync(output);
    assert.deepEqual(Object.keys(extracted), ["汉字/🙂.bin", "empty", "folder/", "duplicate"]);
    assert.deepEqual(extracted["汉字/🙂.bin"], binary);
    assert.deepEqual(extracted.empty, new Uint8Array());
    assert.deepEqual(extracted["folder/"], new Uint8Array());
    assert.deepEqual(extracted.duplicate, new Uint8Array([2, 3]));
    assert.deepEqual(unzipSync(new Zip().writeArchiveData([])), {});
    assert.throws(() => new Zip().writeArchiveData([{ path: "a".repeat(65_536), data: binary }]));
});

test("native ZIP disk writes and workbook exports preserve all package entries", async () => {
    const directory = await mkdtemp(join(tmpdir(), "fastxlsx-native-zip-"));
    try {
        const workbook = Workbook.create("数据");
        workbook.getSheet("数据").cell("A1").setValue("你好🙂");
        workbook.getSheet("数据").setComment("B3", "Comment", { author: "tester" });
        const expected = workbook.toEntries().sort((left, right) => left.path.localeCompare(right.path));
        const file = join(directory, "nested", "book.xlsx");
        await workbook.save(file);
        const output = await readFile(file);
        const extracted = unzipSync(output);
        for (const entry of expected)
            assert.deepEqual(extracted[entry.path], entry.data, entry.path);
        assert.deepEqual((await Workbook.open(file)).toEntries(), expected);
        assert.deepEqual(Workbook.fromUint8Array(workbook.toUint8Array()).toEntries(), expected);
        await assert.rejects(new Zip().writeArchive(join(file, "blocked.xlsx"), expected));
        assert.deepEqual(await readFile(file), output);
    } finally {
        await rm(directory, { recursive: true, force: true });
    }
});
