import test from "node:test";
import assert from "node:assert/strict";
import { mkdtemp, readFile, readdir, rm, stat } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join, resolve } from "node:path";

import { Workbook, validateRoundtripFile } from "../src/index.ts";

test("deterministic XML formatting fuzz keeps roundtrip and write invariants stable", async () => {
  const fixtureDir = resolve("test/fixtures/lossless-source");
  const baseEntries = await loadFixtureEntries(fixtureDir);
  const tempRoot = await mkdtemp(join(tmpdir(), "fastxlsx-xml-fuzz-"));
  const pathsToMutate = [
    "_rels/.rels",
    "[Content_Types].xml",
    "docProps/app.xml",
    "docProps/core.xml",
    "xl/_rels/workbook.xml.rels",
    "xl/styles.xml",
    "xl/workbook.xml",
    "xl/worksheets/sheet1.xml",
  ];

  try {
    for (const seed of [1, 2, 3, 4, 5, 6]) {
      const mutatedEntries = mutateEntriesFormatting(baseEntries, seed, pathsToMutate);
      const inputPath = join(tempRoot, `seed-${seed}.xlsx`);
      const outputPath = join(tempRoot, `seed-${seed}.edited.xlsx`);
      await Workbook.fromEntries(mutatedEntries).save(inputPath);

      const roundtrip = await validateRoundtripFile(inputPath);
      assert.equal(roundtrip.ok, true, `roundtrip failed for seed ${seed}`);
      assert.deepEqual(roundtrip.diffs, [], `unexpected diffs for seed ${seed}`);

      const workbook = await Workbook.open(inputPath);
      const sheet = workbook.getSheet("Sheet1");
      const nextText = `Seed ${seed}`;

      assert.equal(sheet.getCell("A1"), "Hello", `cell read mismatch for seed ${seed}`);
      assert.equal(sheet.getStyleId("A1"), 1, `style read mismatch for seed ${seed}`);

      workbook.batch((currentWorkbook) => {
        const currentSheet = currentWorkbook.getSheet("Sheet1");
        currentSheet.setCell("A1", nextText);
        currentSheet.setFormula("B2", "LEN(A1)", { cachedValue: nextText.length });
      });

      await workbook.save(outputPath);

      const editedWorkbook = await Workbook.open(outputPath);
      const editedSheet = editedWorkbook.getSheet("Sheet1");
      const editedEntries = editedWorkbook.toEntries();

      assert.equal(editedSheet.getCell("A1"), nextText, `edited value mismatch for seed ${seed}`);
      assert.equal(editedSheet.getFormula("B2"), "LEN(A1)", `formula mismatch for seed ${seed}`);
      assert.equal(editedSheet.getStyleId("A1"), 1, `style preservation mismatch for seed ${seed}`);
      assert.equal(
        entryText(editedEntries, "xl/styles.xml"),
        entryText(mutatedEntries, "xl/styles.xml"),
        `styles.xml changed for seed ${seed}`,
      );
    }
  } finally {
    await rm(tempRoot, { recursive: true, force: true });
  }
});

async function loadFixtureEntries(rootDirectory: string): Promise<Array<{ path: string; data: Uint8Array }>> {
  const entries: Array<{ path: string; data: Uint8Array }> = [];
  const stack = [rootDirectory];

  while (stack.length > 0) {
    const current = stack.pop();
    if (!current) {
      continue;
    }

    const names = await readdir(current);

    for (const name of names) {
      const absolutePath = join(current, name);
      const info = await stat(absolutePath);

      if (info.isDirectory()) {
        stack.push(absolutePath);
        continue;
      }

      const relativePath = absolutePath.slice(rootDirectory.length + 1).replaceAll("\\", "/");
      entries.push({
        path: relativePath,
        data: await readFile(absolutePath),
      });
    }
  }

  entries.sort((left, right) => left.path.localeCompare(right.path));
  return entries;
}

function mutateEntriesFormatting(
  entries: Array<{ path: string; data: Uint8Array }>,
  seed: number,
  paths: string[],
): Array<{ path: string; data: Uint8Array }> {
  const encoder = new TextEncoder();
  const selected = new Set(paths);

  return entries.map((entry) => {
    if (!selected.has(entry.path)) {
      return entry;
    }

    const pathSeed = seed + hashText(entry.path);
    const text = Buffer.from(entry.data).toString("utf8");

    return {
      path: entry.path,
      data: encoder.encode(mutateXmlFormatting(text, pathSeed)),
    };
  });
}

function mutateXmlFormatting(xml: string, seed: number): string {
  return xml.replace(/<([A-Za-z_?:][\w:.-]*)([^<>]*?)(\/?)>/g, (source, tagName: string, attributesSource: string, selfClosing: string) => {
    if (tagName.startsWith("?") || tagName.startsWith("!")) {
      return source;
    }

    const attributes = parseRawAttributes(attributesSource);
    if (attributes.length === 0) {
      if (selfClosing !== "/") {
        return source;
      }

      return ((seed + tagName.length) & 1) === 0
        ? source.replace("/>", " />")
        : source;
    }

    const rotation = seed % attributes.length;
    const rotated = attributes.slice(rotation).concat(attributes.slice(0, rotation));
    const joiner = (seed & 1) === 0 ? " " : "\n    ";
    const rendered = rotated.map((attribute, index) => renderRawAttribute(attribute, seed, index)).join(joiner);
    const close = selfClosing === "/"
      ? ((seed + attributes.length) % 3 === 0 ? " />" : "/>")
      : ((seed + attributes.length) % 5 === 0 ? " >" : ">");

    return `<${tagName} ${rendered}${close}`;
  });
}

function parseRawAttributes(source: string): Array<{ name: string; quote: string; value: string }> {
  const attributes: Array<{ name: string; quote: string; value: string }> = [];

  for (const match of source.matchAll(/([A-Za-z_][\w:.-]*)\s*=\s*(["'])([\s\S]*?)\2/g)) {
    attributes.push({
      name: match[1]!,
      quote: match[2]!,
      value: match[3]!,
    });
  }

  return attributes;
}

function renderRawAttribute(
  attribute: { name: string; quote: string; value: string },
  seed: number,
  index: number,
): string {
  const preferSingleQuote = ((seed + index) & 1) === 0;
  const preferredQuote = preferSingleQuote ? "'" : "\"";
  const quote = attribute.value.includes(preferredQuote)
    ? attribute.quote
    : preferredQuote;
  const leftPadding = " ".repeat((seed + index) % 2);
  const rightPadding = " ".repeat((seed + index) % 3 === 0 ? 1 : 0);

  return `${attribute.name}${leftPadding}=${rightPadding}${quote}${attribute.value}${quote}`;
}

function hashText(text: string): number {
  let hash = 0;

  for (let index = 0; index < text.length; index += 1) {
    hash = (hash * 33 + text.charCodeAt(index)) >>> 0;
  }

  return hash;
}

function entryText(entries: Array<{ path: string; data: Uint8Array }>, path: string): string {
  const entry = entries.find((candidate) => candidate.path === path);
  if (!entry) {
    throw new Error(`Missing entry: ${path}`);
  }

  return Buffer.from(entry.data).toString("utf8");
}
