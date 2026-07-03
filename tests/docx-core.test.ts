import { mkdtemp, readFile, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { execa } from "execa";
import { describe, expect, it } from "vitest";
import JSZip from "jszip";
import { buildDocx, createDocumentJson, parseDocx } from "../src/index";

describe("DocumentJson schema", () => {
  it("creates a versioned document with sections and paragraphs", () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "Hello Word", bold: true }],
      },
    ]);

    expect(document.version).toBe("1.0");
    expect(document.sections).toHaveLength(1);
    expect(document.sections[0].blocks[0]).toMatchObject({
      type: "paragraph",
      runs: [{ text: "Hello Word", bold: true }],
    });
  });
});

describe("DOCX writer", () => {
  it("writes paragraphs and text run formatting into document.xml", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        alignment: "center",
        runs: [
          { text: "Hello ", bold: true },
          { text: "Word", italic: true, underline: true },
        ],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<w:jc w:val="center"/>');
    expect(xml).toContain("<w:b/>");
    expect(xml).toContain("<w:i/>");
    expect(xml).toContain('<w:u w:val="single"/>');
    expect(xml).toContain('<w:t xml:space="preserve">Hello </w:t>');
    expect(xml).toContain("<w:t>Word</w:t>");
  });
});

describe("DOCX reader", () => {
  it("parses supported paragraphs from generated docx back to json", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        alignment: "right",
        runs: [
          { text: "A", bold: true },
          { text: "B", italic: true, underline: true },
        ],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });
});

describe("CLI", () => {
  it("converts json to docx and back through the CLI", async () => {
    const dir = await mkdtemp(path.join(tmpdir(), "word2json-"));
    const jsonPath = path.join(dir, "input.json");
    const docxPath = path.join(dir, "output.docx");
    const roundTripPath = path.join(dir, "roundtrip.json");
    const document = createDocumentJson([
      { type: "paragraph", runs: [{ text: "CLI" }] },
    ]);

    await writeFile(jsonPath, JSON.stringify(document), "utf8");

    await execa("npx", ["tsx", "src/cli.ts", "json2docx", jsonPath, docxPath]);
    await execa("npx", ["tsx", "src/cli.ts", "docx2json", docxPath, roundTripPath]);

    expect(JSON.parse(await readFile(roundTripPath, "utf8"))).toEqual(document);
  });
});
