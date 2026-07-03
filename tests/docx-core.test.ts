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

  it("writes hyperlinks with external relationships", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "OpenAI",
            link: { url: "https://example.com" },
          },
        ],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");
    const rels = await zip.file("word/_rels/document.xml.rels")!.async("string");

    expect(xml).toContain('<w:hyperlink r:id="rIdHyperlink1">');
    expect(xml).toContain("<w:t>OpenAI</w:t>");
    expect(rels).toContain('Id="rIdHyperlink1"');
    expect(rels).toContain('Target="https://example.com"');
    expect(rels).toContain('TargetMode="External"');
  });

  it("writes comments with ranges references and comments part", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "Needs review",
            comment: {
              author: "Ada",
              initials: "AL",
              date: "2026-07-04T00:00:00.000Z",
              text: "Please verify this clause.",
            },
          },
        ],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");
    const comments = await zip.file("word/comments.xml")!.async("string");
    const contentTypes = await zip.file("[Content_Types].xml")!.async("string");
    const rels = await zip.file("word/_rels/document.xml.rels")!.async("string");

    expect(xml).toContain('<w:commentRangeStart w:id="0"/>');
    expect(xml).toContain('<w:commentRangeEnd w:id="0"/>');
    expect(xml).toContain('<w:commentReference w:id="0"/>');
    expect(comments).toContain('<w:comment w:id="0" w:author="Ada" w:initials="AL" w:date="2026-07-04T00:00:00.000Z">');
    expect(comments).toContain("<w:t>Please verify this clause.</w:t>");
    expect(contentTypes).toContain('/word/comments.xml');
    expect(rels).toContain('Target="comments.xml"');
  });

  it("writes bookmarks and breaks", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          { text: "Anchor", bookmark: { name: "Clause1" } },
          { text: "", break: "line" },
          { text: "", break: "page" },
        ],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<w:bookmarkStart w:id="0" w:name="Clause1"/>');
    expect(xml).toContain('<w:bookmarkEnd w:id="0"/>');
    expect(xml).toContain("<w:br/>");
    expect(xml).toContain('<w:br w:type="page"/>');
  });

  it("writes images with media parts relationships and inline drawing metadata", async () => {
    const imageData = Buffer.from("fake-png").toString("base64");
    const document = createDocumentJson([
      {
        type: "image",
        data: imageData,
        contentType: "image/png",
        width: 120,
        height: 80,
        altText: "Logo",
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");
    const rels = await zip.file("word/_rels/document.xml.rels")!.async("string");
    const media = await zip.file("word/media/image1.png")!.async("nodebuffer");

    expect(xml).toContain('<wp:docPr id="1" name="Image 1" descr="Logo"/>');
    expect(xml).toContain('<a:ext cx="1143000" cy="762000"/>');
    expect(xml).toContain('<a:blip r:embed="rIdImage1"/>');
    expect(rels).toContain('Id="rIdImage1"');
    expect(rels).toContain('Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"');
    expect(rels).toContain('Target="media/image1.png"');
    expect(media.toString("base64")).toBe(imageData);
  });

  it("writes headers and footers with section relationships", async () => {
    const document = {
      version: "1.0" as const,
      sections: [
        {
          headers: {
            default: [{ type: "paragraph" as const, runs: [{ text: "Header text" }] }],
          },
          footers: {
            default: [{ type: "paragraph" as const, runs: [{ text: "Footer text" }] }],
          },
          blocks: [
            { type: "paragraph" as const, runs: [{ text: "Body" }] },
          ],
        },
      ],
    };

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");
    const header = await zip.file("word/header1.xml")!.async("string");
    const footer = await zip.file("word/footer1.xml")!.async("string");
    const rels = await zip.file("word/_rels/document.xml.rels")!.async("string");

    expect(xml).toContain('<w:headerReference w:type="default" r:id="rIdHeader1"/>');
    expect(xml).toContain('<w:footerReference w:type="default" r:id="rIdFooter1"/>');
    expect(header).toContain("<w:t>Header text</w:t>");
    expect(footer).toContain("<w:t>Footer text</w:t>");
    expect(rels).toContain('Id="rIdHeader1"');
    expect(rels).toContain('Target="header1.xml"');
    expect(rels).toContain('Id="rIdFooter1"');
    expect(rels).toContain('Target="footer1.xml"');
  });

  it("writes page fields in footer runs", async () => {
    const document = {
      version: "1.0" as const,
      sections: [
        {
          footers: {
            default: [
              {
                type: "paragraph" as const,
                runs: [
                  { text: "Page " },
                  { text: "", field: "page" as const },
                  { text: " of " },
                  { text: "", field: "numPages" as const },
                ],
              },
            ],
          },
          blocks: [{ type: "paragraph" as const, runs: [{ text: "Body" }] }],
        },
      ],
    };

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const footer = await zip.file("word/footer1.xml")!.async("string");

    expect(footer).toContain('<w:instrText xml:space="preserve">PAGE</w:instrText>');
    expect(footer).toContain('<w:instrText xml:space="preserve">NUMPAGES</w:instrText>');
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

  it("round-trips hyperlinks", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "OpenAI",
            link: { url: "https://example.com" },
          },
        ],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips comments", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "Needs review",
            comment: {
              author: "Ada",
              initials: "AL",
              date: "2026-07-04T00:00:00.000Z",
              text: "Please verify this clause.",
            },
          },
        ],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips bookmarks and breaks", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          { text: "Anchor", bookmark: { name: "Clause1" } },
          { text: "", break: "line" },
          { text: "", break: "page" },
        ],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips images", async () => {
    const imageData = Buffer.from("fake-png").toString("base64");
    const source = createDocumentJson([
      {
        type: "image",
        data: imageData,
        contentType: "image/png",
        width: 120,
        height: 80,
        altText: "Logo",
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips headers and footers", async () => {
    const source = {
      version: "1.0" as const,
      sections: [
        {
          headers: {
            default: [{ type: "paragraph" as const, runs: [{ text: "Header text" }] }],
          },
          footers: {
            default: [{ type: "paragraph" as const, runs: [{ text: "Footer text" }] }],
          },
          blocks: [
            { type: "paragraph" as const, runs: [{ text: "Body" }] },
          ],
        },
      ],
    };

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips page fields", async () => {
    const source = {
      version: "1.0" as const,
      sections: [
        {
          footers: {
            default: [
              {
                type: "paragraph" as const,
                runs: [
                  { text: "Page " },
                  { text: "", field: "page" as const },
                  { text: " of " },
                  { text: "", field: "numPages" as const },
                ],
              },
            ],
          },
          blocks: [{ type: "paragraph" as const, runs: [{ text: "Body" }] }],
        },
      ],
    };

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
