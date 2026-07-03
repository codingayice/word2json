#!/usr/bin/env node
import { readFile, writeFile } from "node:fs/promises";
import { buildDocx } from "./docx-writer.js";
import { parseDocx } from "./docx-reader.js";
import type { DocumentJson } from "./schema.js";

async function main(argv: string[]): Promise<void> {
  const [command, inputPath, outputPath] = argv;

  if (!command || !inputPath || !outputPath) {
    printUsageAndExit();
  }

  if (command === "json2docx") {
    const document = JSON.parse(await readFile(inputPath, "utf8")) as DocumentJson;
    const docx = await buildDocx(document);
    await writeFile(outputPath, docx);
    return;
  }

  if (command === "docx2json") {
    const docx = await readFile(inputPath);
    const document = await parseDocx(docx);
    await writeFile(outputPath, `${JSON.stringify(document, null, 2)}\n`, "utf8");
    return;
  }

  printUsageAndExit();
}

function printUsageAndExit(): never {
  console.error("Usage: word2json <json2docx|docx2json> <input> <output>");
  process.exit(1);
}

main(process.argv.slice(2)).catch((error: unknown) => {
  console.error(error instanceof Error ? error.message : String(error));
  process.exit(1);
});
