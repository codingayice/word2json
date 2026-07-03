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

  it("writes text styles and headings into document.xml", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        style: "heading1",
        runs: [
          {
            text: "Styled Heading",
            fontFamily: "Arial",
            fontSize: 16,
            color: "C00000",
          },
        ],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<w:pStyle w:val="Heading1"/>');
    expect(xml).toContain('<w:rFonts w:ascii="Arial" w:hAnsi="Arial"/>');
    expect(xml).toContain('<w:sz w:val="32"/>');
    expect(xml).toContain('<w:color w:val="C00000"/>');
  });

  it("writes page settings into section properties", async () => {
    const document = {
      version: "1.0" as const,
      sections: [
        {
          page: {
            width: 16840,
            height: 11900,
            orientation: "landscape" as const,
            margins: {
              top: 720,
              right: 900,
              bottom: 720,
              left: 900,
              header: 360,
              footer: 360,
              gutter: 0,
            },
          },
          blocks: [
            { type: "paragraph" as const, runs: [{ text: "Landscape" }] },
          ],
        },
      ],
    };

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<w:pgSz w:w="16840" w:h="11900" w:orient="landscape"/>');
    expect(xml).toContain('<w:pgMar w:top="720" w:right="900" w:bottom="720" w:left="900" w:header="360" w:footer="360" w:gutter="0"/>');
  });

  it("writes tables with widths borders and grid spans", async () => {
    const document = createDocumentJson([
      {
        type: "table",
        width: 9000,
        borders: "single",
        rows: [
          {
            cells: [
              {
                width: 3000,
                colSpan: 2,
                blocks: [{ type: "paragraph", runs: [{ text: "Merged" }] }],
              },
              {
                width: 3000,
                blocks: [{ type: "paragraph", runs: [{ text: "Cell" }] }],
              },
            ],
          },
        ],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain("<w:tbl>");
    expect(xml).toContain('<w:tblW w:w="9000" w:type="dxa"/>');
    expect(xml).toContain("<w:tblBorders>");
    expect(xml).toContain('<w:gridSpan w:val="2"/>');
    expect(xml).toContain("<w:t>Merged</w:t>");
    expect(xml).toContain("<w:t>Cell</w:t>");
  });

  it("writes lists with numbering definitions and paragraph numbering", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        list: { type: "bullet", level: 0 },
        runs: [{ text: "Bullet item" }],
      },
      {
        type: "paragraph",
        list: { type: "ordered", level: 1 },
        runs: [{ text: "Ordered item" }],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");
    const numbering = await zip.file("word/numbering.xml")!.async("string");

    expect(xml).toContain("<w:numPr><w:ilvl w:val=\"0\"/><w:numId w:val=\"1\"/></w:numPr>");
    expect(xml).toContain("<w:numPr><w:ilvl w:val=\"1\"/><w:numId w:val=\"2\"/></w:numPr>");
    expect(numbering).toContain('<w:numFmt w:val="bullet"/>');
    expect(numbering).toContain('<w:numFmt w:val="decimal"/>');
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

  it("round-trips text styles and headings", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        style: "heading1",
        runs: [
          {
            text: "Styled Heading",
            fontFamily: "Arial",
            fontSize: 16,
            color: "C00000",
          },
        ],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips page settings", async () => {
    const source = {
      version: "1.0" as const,
      sections: [
        {
          page: {
            width: 16840,
            height: 11900,
            orientation: "landscape" as const,
            margins: {
              top: 720,
              right: 900,
              bottom: 720,
              left: 900,
              header: 360,
              footer: 360,
              gutter: 0,
            },
          },
          blocks: [
            { type: "paragraph" as const, runs: [{ text: "Landscape" }] },
          ],
        },
      ],
    };

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips tables", async () => {
    const source = createDocumentJson([
      {
        type: "table",
        width: 9000,
        borders: "single",
        rows: [
          {
            cells: [
              {
                width: 3000,
                colSpan: 2,
                blocks: [{ type: "paragraph", runs: [{ text: "Merged" }] }],
              },
              {
                width: 3000,
                blocks: [{ type: "paragraph", runs: [{ text: "Cell" }] }],
              },
            ],
          },
        ],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips lists", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        list: { type: "bullet", level: 0 },
        runs: [{ text: "Bullet item" }],
      },
      {
        type: "paragraph",
        list: { type: "ordered", level: 1 },
        runs: [{ text: "Ordered item" }],
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

    await execa(process.execPath, ["--import", "tsx", "src/cli.ts", "json2docx", jsonPath, docxPath]);
    await execa(process.execPath, ["--import", "tsx", "src/cli.ts", "docx2json", docxPath, roundTripPath]);

    expect(JSON.parse(await readFile(roundTripPath, "utf8"))).toEqual(document);
  });
});
