import { execFile } from "node:child_process";
import { stat } from "node:fs/promises";
import { resolve } from "node:path";
import { promisify } from "node:util";
import { describe, expect, test } from "vitest";

const execFileAsync = promisify(execFile);

describe("roundtrip demo", () => {
  test("prints usage for the demo workflow", async () => {
    const { stdout } = await execFileAsync("node", ["demo/roundtrip-demo.mjs", "--help"]);

    expect(stdout).toContain("Usage:");
    expect(stdout).toContain("node demo/roundtrip-demo.mjs [input.docx] [output-dir]");
    expect(stdout).toContain("demo/out/exported.docx");
  });

  test("runs the default sample roundtrip", async () => {
    const { stdout } = await execFileAsync("node", ["demo/roundtrip-demo.mjs"], { maxBuffer: 10 * 1024 * 1024 });

    expect(stdout).toContain("Roundtrip demo files:");
    await expectFile(resolve("demo/out/demo-source.docx"));
    await expectFile(resolve("demo/out/imported.json"));
    await expectFile(resolve("demo/out/exported.docx"));
  }, 20_000);
});

async function expectFile(path: string): Promise<void> {
  const info = await stat(path);
  expect(info.size).toBeGreaterThan(0);
}
