import { mkdir, readFile, writeFile } from "node:fs/promises";
import { dirname, join } from "node:path";

import { deflateRawSync } from "node:zlib";
import { unzipSync, Zip as ZipWriter, ZipPassThrough } from "fflate";

import type { ArchiveEntry } from "./types.js";

// Each archive entry is pushed once, with its complete uncompressed bytes.
// fflate owns CRC, UTF-8 names, headers, data descriptors, and central directory.
class NativeZipDeflate extends ZipPassThrough {
    override compression = 8;

    protected override process(chunk: Uint8Array, final: boolean): void {
        this.ondata(null, deflateRawSync(chunk, { level: 6 }), final);
    }
}

export class Zip {
    async readArchive(filePath: string): Promise<ArchiveEntry[]> {
        const archiveData = await readFile(filePath);
        return this.readArchiveData(new Uint8Array(archiveData));
    }

    async writeArchive(filePath: string, entries: Iterable<ArchiveEntry>): Promise<void> {
        const destinationDirectory = dirname(filePath);
        if (destinationDirectory !== ".") {
            await mkdir(destinationDirectory, { recursive: true });
        }

        await writeFile(filePath, this.writeArchiveData(entries));
    }

    readArchiveData(archiveData: Uint8Array | ArrayBuffer): ArchiveEntry[] {
        const entriesByPath = unzipSync(
            archiveData instanceof Uint8Array ? archiveData : new Uint8Array(archiveData)
        );

        return Object.keys(entriesByPath)
            .sort((left, right) => left.localeCompare(right))
            .map((path) => ({
                path,
                data: entriesByPath[path],
            }));
    }

    writeArchiveData(entries: Iterable<ArchiveEntry>): Uint8Array {
        const sources = new Map<string, Uint8Array>();
        for (const entry of entries) {
            sources.set(entry.path, entry.data);
        }

        const chunks: Uint8Array[] = [];
        let size = 0;
        const archive = new ZipWriter((error, chunk) => {
            if (error) {
                throw error;
            }
            chunks.push(chunk);
            size += chunk.length;
        });
        for (const [path, data] of sources) {
            const entry = new NativeZipDeflate(path);
            archive.add(entry);
            entry.push(data, true);
        }
        archive.end();

        const result = new Uint8Array(size);
        let offset = 0;
        for (const chunk of chunks) {
            result.set(chunk, offset);
            offset += chunk.length;
        }
        return result;
    }

    async readDirectoryEntries(
        directoryPath: string,
        root = directoryPath
    ): Promise<ArchiveEntry[]> {
        const collected: ArchiveEntry[] = [];

        const stack = [directoryPath];

        while (stack.length > 0) {
            const current = stack.pop();
            if (!current) {
                continue;
            }

            const children = await readDirNames(current);

            for (const child of children) {
                const absolutePath = join(current, child.name);

                if (child.isDirectory) {
                    stack.push(absolutePath);
                    continue;
                }

                const relativePath = absolutePath.slice(root.length + 1).replaceAll("\\", "/");
                collected.push({
                    path: relativePath,
                    data: await readFile(absolutePath),
                });
            }
        }

        collected.sort((left, right) => left.path.localeCompare(right.path));
        return collected;
    }
}

async function readDirNames(
    directoryPath: string
): Promise<Array<{ name: string; isDirectory: boolean }>> {
    const names = await (await import("node:fs/promises")).readdir(directoryPath);
    const rows: Array<{ name: string; isDirectory: boolean }> = [];

    for (const name of names as string[]) {
        const absolutePath = join(directoryPath, name);
        const stats = await (await import("node:fs/promises")).stat(absolutePath);
        rows.push({ name, isDirectory: stats.isDirectory() });
    }

    rows.sort((left, right) => left.name.localeCompare(right.name));
    return rows;
}
