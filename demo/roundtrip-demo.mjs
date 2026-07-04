#!/usr/bin/env node
import { execFile } from "node:child_process";
import { mkdir } from "node:fs/promises";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import { promisify } from "node:util";

const execFileAsync = promisify(execFile);
const repoRoot = resolve(dirname(fileURLToPath(import.meta.url)), "..");

const help = `Usage:
  node demo/roundtrip-demo.mjs [input.docx] [output-dir]

Examples:
  node demo/roundtrip-demo.mjs
  node demo/roundtrip-demo.mjs C:\\docs\\sample.docx demo/out/my-sample

Outputs:
  demo/out/demo-source.docx  created from demo/sample.json when no input is passed
  demo/out/imported.json     JSON imported from the source DOCX
  demo/out/exported.docx     DOCX exported from imported.json
`;

const args = process.argv.slice(2);

if (args.includes("--help") || args.includes("-h")) {
  process.stdout.write(help);
  process.exit(0);
}

const [inputArg, outputArg] = args;
const outputDir = resolve(repoRoot, outputArg ?? "demo/out");
const cliPath = resolve(repoRoot, "dist/cli.js");
const tscPath = resolve(repoRoot, "node_modules/typescript/bin/tsc");
const sampleJsonPath = resolve(repoRoot, "demo/sample.json");
const sourceDocxPath = inputArg ? resolve(process.cwd(), inputArg) : resolve(outputDir, "demo-source.docx");
const importedJsonPath = resolve(outputDir, "imported.json");
const exportedDocxPath = resolve(outputDir, "exported.docx");

await mkdir(outputDir, { recursive: true });
await run("node", [tscPath, "-p", "tsconfig.json"]);

if (!inputArg) {
  await run("node", [cliPath, "json2docx", sampleJsonPath, sourceDocxPath]);
}

await run("node", [cliPath, "docx2json", sourceDocxPath, importedJsonPath]);
await run("node", [cliPath, "json2docx", importedJsonPath, exportedDocxPath]);

process.stdout.write([
  "Roundtrip demo files:",
  `  Source DOCX:   ${sourceDocxPath}`,
  `  Imported JSON: ${importedJsonPath}`,
  `  Exported DOCX: ${exportedDocxPath}`,
  "",
].join("\n"));

async function run(command, args) {
  try {
    await execFileAsync(command, args, { cwd: repoRoot, windowsHide: true, maxBuffer: 10 * 1024 * 1024 });
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    const stdout = typeof error.stdout === "string" ? error.stdout : "";
    const stderr = typeof error.stderr === "string" ? error.stderr : "";
    throw new Error([`Command failed: ${command} ${args.join(" ")}`, message, stdout, stderr].filter(Boolean).join("\n"));
  }
}
