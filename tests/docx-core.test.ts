import { mkdtemp, readFile, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { execa } from "execa";
import { describe, expect, it } from "vitest";
import JSZip from "jszip";
import { buildDocx, createDocumentJson, parseDocx } from "../src/index";

function defaultThemeColors() {
  return {
    dark1: "000000",
    light1: "FFFFFF",
    dark2: "1F497D",
    light2: "EEECE1",
    accent2: "C0504D",
    accent3: "9BBB59",
    accent4: "8064A2",
    accent5: "4BACC6",
    accent6: "F79646",
    hyperlink: "0000FF",
    followedHyperlink: "800080",
  };
}

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

  it("writes text run underline none", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "Plain", underline: false }],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<w:rPr><w:u w:val="none"/></w:rPr><w:t>Plain</w:t>');
  });

  it("writes text run bold off", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "Plain", bold: false }],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<w:rPr><w:b w:val="0"/></w:rPr><w:t>Plain</w:t>');
  });

  it("writes text run italic off", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "Plain", italic: false }],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<w:rPr><w:i w:val="0"/></w:rPr><w:t>Plain</w:t>');
  });

  it("writes text run strike off", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "Plain", strike: false }],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<w:rPr><w:strike w:val="0"/></w:rPr><w:t>Plain</w:t>');
  });

  it("writes text run double strike off", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "Plain", doubleStrike: false }],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<w:rPr><w:dstrike w:val="0"/></w:rPr><w:t>Plain</w:t>');
  });

  it("writes text run small caps off", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "Plain", smallCaps: false }],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<w:rPr><w:smallCaps w:val="0"/></w:rPr><w:t>Plain</w:t>');
  });

  it("writes text run all caps off", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "Plain", allCaps: false }],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<w:rPr><w:caps w:val="0"/></w:rPr><w:t>Plain</w:t>');
  });

  it("writes text run shadow on", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "Shadow", shadow: true }],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain("<w:rPr><w:shadow/></w:rPr><w:t>Shadow</w:t>");
  });

  it("writes text run shadow off", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "Plain", shadow: false }],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<w:rPr><w:shadow w:val="0"/></w:rPr><w:t>Plain</w:t>');
  });

  it("writes text run outline on", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "Outline", outline: true }],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain("<w:rPr><w:outline/></w:rPr><w:t>Outline</w:t>");
  });

  it("writes text run outline off", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "Plain", outline: false }],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<w:rPr><w:outline w:val="0"/></w:rPr><w:t>Plain</w:t>');
  });

  it("writes text run emboss on", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "Emboss", emboss: true }],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain("<w:rPr><w:emboss/></w:rPr><w:t>Emboss</w:t>");
  });

  it("writes text run emboss off", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "Plain", emboss: false }],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<w:rPr><w:emboss w:val="0"/></w:rPr><w:t>Plain</w:t>');
  });

  it("writes text run imprint on", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "Imprint", imprint: true }],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain("<w:rPr><w:imprint/></w:rPr><w:t>Imprint</w:t>");
  });

  it("writes text run imprint off", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "Plain", imprint: false }],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<w:rPr><w:imprint w:val="0"/></w:rPr><w:t>Plain</w:t>');
  });

  it("writes text run rtl on", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "يمين", rtl: true }],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain("<w:rPr><w:rtl/></w:rPr><w:t>يمين</w:t>");
  });

  it("writes text run rtl off", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "Plain", rtl: false }],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<w:rPr><w:rtl w:val="0"/></w:rPr><w:t>Plain</w:t>');
  });

  it("writes text run complex script on", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "Complex", complexScript: true }],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain("<w:rPr><w:cs/></w:rPr><w:t>Complex</w:t>");
  });

  it("writes text run complex script off", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "Plain", complexScript: false }],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<w:rPr><w:cs w:val="0"/></w:rPr><w:t>Plain</w:t>');
  });

  it("writes text run spec vanish on", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "Hidden", specVanish: true }],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain("<w:rPr><w:specVanish/></w:rPr><w:t>Hidden</w:t>");
  });

  it("writes text run spec vanish off", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "Visible", specVanish: false }],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<w:rPr><w:specVanish w:val="0"/></w:rPr><w:t>Visible</w:t>');
  });

  it("writes text run hidden on", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "Hidden", hidden: true }],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain("<w:rPr><w:vanish/></w:rPr><w:t>Hidden</w:t>");
  });

  it("writes text run hidden off", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "Visible", hidden: false }],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<w:rPr><w:vanish w:val="0"/></w:rPr><w:t>Visible</w:t>');
  });

  it("writes text run web hidden on", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "WebHidden", webHidden: true }],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain("<w:rPr><w:webHidden/></w:rPr><w:t>WebHidden</w:t>");
  });

  it("writes text run web hidden off", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "Visible", webHidden: false }],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<w:rPr><w:webHidden w:val="0"/></w:rPr><w:t>Visible</w:t>');
  });

  it("writes text run fit text with id", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "Fit", fitText: { width: 1440, id: 7 } }],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<w:rPr><w:fitText w:val="1440" w:id="7"/></w:rPr><w:t>Fit</w:t>');
  });

  it("writes text run fit text without id", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "Fit", fitText: { width: 720 } }],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<w:rPr><w:fitText w:val="720"/></w:rPr><w:t>Fit</w:t>');
  });

  it("writes text run emphasis dot", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "Emphasis", emphasis: "dot" }],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<w:rPr><w:em w:val="dot"/></w:rPr><w:t>Emphasis</w:t>');
  });

  it("writes text run emphasis none", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "Plain", emphasis: "none" }],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<w:rPr><w:em w:val="none"/></w:rPr><w:t>Plain</w:t>');
  });

  it("writes text run snap to grid on", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "Grid", snapToGrid: true }],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain("<w:rPr><w:snapToGrid/></w:rPr><w:t>Grid</w:t>");
  });

  it("writes text run snap to grid off", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "Free", snapToGrid: false }],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<w:rPr><w:snapToGrid w:val="0"/></w:rPr><w:t>Free</w:t>');
  });

  it("writes text run no proof on", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "Unchecked", noProof: true }],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain("<w:rPr><w:noProof/></w:rPr><w:t>Unchecked</w:t>");
  });

  it("writes text run no proof off", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "Checked", noProof: false }],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<w:rPr><w:noProof w:val="0"/></w:rPr><w:t>Checked</w:t>');
  });

  it("writes text run office math on", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "Formula", officeMath: true }],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain("<w:rPr><w:oMath/></w:rPr><w:t>Formula</w:t>");
  });

  it("writes text run office math off", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "Plain", officeMath: false }],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<w:rPr><w:oMath w:val="0"/></w:rPr><w:t>Plain</w:t>');
  });

  it("writes text run language value", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "Hello", language: { value: "en-US" } }],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<w:rPr><w:lang w:val="en-US"/></w:rPr><w:t>Hello</w:t>');
  });

  it("writes text run language east asia and bidi", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "Mixed", language: { value: "en-US", eastAsia: "zh-CN", bidi: "ar-SA" } }],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<w:rPr><w:lang w:val="en-US" w:eastAsia="zh-CN" w:bidi="ar-SA"/></w:rPr><w:t>Mixed</w:t>');
  });

  it("writes text run character position raised", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "Raised", characterPosition: 4 }],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<w:rPr><w:position w:val="4"/></w:rPr><w:t>Raised</w:t>');
  });

  it("writes text run character position lowered", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "Lowered", characterPosition: -4 }],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<w:rPr><w:position w:val="-4"/></w:rPr><w:t>Lowered</w:t>');
  });

  it("writes text run kerning small threshold", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "Title", kerning: 24 }],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<w:rPr><w:kern w:val="24"/></w:rPr><w:t>Title</w:t>');
  });

  it("writes text run kerning large threshold", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "Display", kerning: 48 }],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<w:rPr><w:kern w:val="48"/></w:rPr><w:t>Display</w:t>');
  });

  it("writes text run complex script font size", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "مرحبا", complexScriptFontSize: 14 }],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<w:rPr><w:szCs w:val="28"/></w:rPr><w:t>مرحبا</w:t>');
  });

  it("writes text run complex script font size with latin size", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "Mixed", fontSize: 12, complexScriptFontSize: 16 }],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<w:rPr><w:sz w:val="24"/><w:szCs w:val="32"/></w:rPr><w:t>Mixed</w:t>');
  });

  it("writes text run complex script font family", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "مرحبا", complexScriptFontFamily: "Arial" }],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<w:rPr><w:rFonts w:cs="Arial"/></w:rPr><w:t>مرحبا</w:t>');
  });

  it("writes text run complex script font family with latin family", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "Mixed", fontFamily: "Aptos", complexScriptFontFamily: "Arial" }],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<w:rPr><w:rFonts w:ascii="Aptos" w:hAnsi="Aptos" w:cs="Arial"/></w:rPr><w:t>Mixed</w:t>');
  });

  it("writes text run east asia font family", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "你好", eastAsiaFontFamily: "SimSun" }],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<w:rPr><w:rFonts w:eastAsia="SimSun"/></w:rPr><w:t>你好</w:t>');
  });

  it("writes text run east asia font family with latin and complex script families", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "Mixed", fontFamily: "Aptos", eastAsiaFontFamily: "SimSun", complexScriptFontFamily: "Arial" }],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<w:rPr><w:rFonts w:ascii="Aptos" w:hAnsi="Aptos" w:eastAsia="SimSun" w:cs="Arial"/></w:rPr><w:t>Mixed</w:t>');
  });

  it("writes text run font hint east asia", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "你好", fontHint: "eastAsia" }],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr><w:t>你好</w:t>');
  });

  it("writes text run font hint with all font families", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "Mixed", fontFamily: "Aptos", eastAsiaFontFamily: "SimSun", complexScriptFontFamily: "Arial", fontHint: "cs" }],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<w:rPr><w:rFonts w:ascii="Aptos" w:hAnsi="Aptos" w:eastAsia="SimSun" w:cs="Arial" w:hint="cs"/></w:rPr><w:t>Mixed</w:t>');
  });

  it("writes text run theme fonts", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "Theme", fontTheme: "minorHAnsi", eastAsiaFontTheme: "minorEastAsia", complexScriptFontTheme: "minorBidi" }],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<w:rPr><w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi" w:eastAsiaTheme="minorEastAsia" w:cstheme="minorBidi"/></w:rPr><w:t>Theme</w:t>');
  });

  it("writes text run theme fonts with direct fonts and hint", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [{
          text: "Mixed",
          fontFamily: "Aptos",
          eastAsiaFontFamily: "SimSun",
          complexScriptFontFamily: "Arial",
          fontTheme: "majorHAnsi",
          eastAsiaFontTheme: "majorEastAsia",
          complexScriptFontTheme: "majorBidi",
          fontHint: "eastAsia",
        }],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<w:rPr><w:rFonts w:ascii="Aptos" w:hAnsi="Aptos" w:eastAsia="SimSun" w:cs="Arial" w:asciiTheme="majorHAnsi" w:hAnsiTheme="majorHAnsi" w:eastAsiaTheme="majorEastAsia" w:cstheme="majorBidi" w:hint="eastAsia"/></w:rPr><w:t>Mixed</w:t>');
  });

  it("writes inline office math runs", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          { text: "Equation: " },
          { text: "", math: { text: "x+1=2" } },
        ],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"');
    expect(xml).toContain("<m:oMath>");
    expect(xml).toContain("<m:r><m:t>x+1=2</m:t></m:r>");
    expect(xml).toContain("</m:oMath>");
  });

  it("writes structured office math runs", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "fraction" as const,
                  numerator: [{ type: "text" as const, text: "1" }],
                  denominator: [{ type: "text" as const, text: "2" }],
                },
                {
                  type: "superscript" as const,
                  base: [{ type: "text" as const, text: "x" }],
                  superscript: [{ type: "text" as const, text: "2" }],
                },
                {
                  type: "subscript" as const,
                  base: [{ type: "text" as const, text: "a" }],
                  subscript: [{ type: "text" as const, text: "i" }],
                },
              ],
            },
          },
        ],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain("<m:f><m:num><m:r><m:t>1</m:t></m:r></m:num><m:den><m:r><m:t>2</m:t></m:r></m:den></m:f>");
    expect(xml).toContain("<m:sSup><m:e><m:r><m:t>x</m:t></m:r></m:e><m:sup><m:r><m:t>2</m:t></m:r></m:sup></m:sSup>");
    expect(xml).toContain("<m:sSub><m:e><m:r><m:t>a</m:t></m:r></m:e><m:sub><m:r><m:t>i</m:t></m:r></m:sub></m:sSub>");
  });

  it("writes fraction control properties", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "fraction" as const,
                  controlProperties: {
                    bold: true,
                    italic: true,
                    fontFamily: "Cambria Math",
                    fontSize: 15,
                    color: "C55A11",
                    underline: true,
                    highlight: "blue",
                  },
                  numerator: [{ type: "text" as const, text: "1" }],
                  denominator: [{ type: "text" as const, text: "2" }],
                },
              ],
            },
          },
        ],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<m:f><m:fPr><m:ctrlPr><w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/><w:b/><w:i/><w:color w:val="C55A11"/><w:sz w:val="30"/><w:highlight w:val="blue"/><w:u w:val="single"/></w:rPr></m:ctrlPr></m:fPr><m:num><m:r><m:t>1</m:t></m:r></m:num><m:den><m:r><m:t>2</m:t></m:r></m:den></m:f>');
  });

  it("writes fraction type", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "fraction" as const,
                  fractionType: "skewed",
                  numerator: [{ type: "text" as const, text: "1" }],
                  denominator: [{ type: "text" as const, text: "2" }],
                },
              ],
            },
          },
        ],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<m:f><m:fPr><m:type m:val="skw"/></m:fPr><m:num><m:r><m:t>1</m:t></m:r></m:num><m:den><m:r><m:t>2</m:t></m:r></m:den></m:f>');
  });

  it("writes superscript control properties", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "superscript" as const,
                  controlProperties: {
                    bold: true,
                    italic: true,
                    fontFamily: "Cambria Math",
                    fontSize: 12,
                    color: "7030A0",
                    underline: true,
                    highlight: "green",
                  },
                  base: [{ type: "text" as const, text: "x" }],
                  superscript: [{ type: "text" as const, text: "2" }],
                },
              ],
            },
          },
        ],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<m:sSup><m:sSupPr><m:ctrlPr><w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/><w:b/><w:i/><w:color w:val="7030A0"/><w:sz w:val="24"/><w:highlight w:val="green"/><w:u w:val="single"/></w:rPr></m:ctrlPr></m:sSupPr><m:e><m:r><m:t>x</m:t></m:r></m:e><m:sup><m:r><m:t>2</m:t></m:r></m:sup></m:sSup>');
  });

  it("writes subscript control properties", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "subscript" as const,
                  controlProperties: {
                    bold: true,
                    italic: true,
                    fontFamily: "Cambria Math",
                    fontSize: 11,
                    color: "00B050",
                    underline: true,
                    highlight: "magenta",
                  },
                  base: [{ type: "text" as const, text: "a" }],
                  subscript: [{ type: "text" as const, text: "i" }],
                },
              ],
            },
          },
        ],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<m:sSub><m:sSubPr><m:ctrlPr><w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/><w:b/><w:i/><w:color w:val="00B050"/><w:sz w:val="22"/><w:highlight w:val="magenta"/><w:u w:val="single"/></w:rPr></m:ctrlPr></m:sSubPr><m:e><m:r><m:t>a</m:t></m:r></m:e><m:sub><m:r><m:t>i</m:t></m:r></m:sub></m:sSub>');
  });

  it("writes radical and n-ary office math runs", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "radical" as const,
                  degree: [{ type: "text" as const, text: "3" }],
                  content: [{ type: "text" as const, text: "x" }],
                },
                {
                  type: "nary" as const,
                  operator: "sum" as const,
                  lowerLimit: [{ type: "text" as const, text: "i=1" }],
                  upperLimit: [{ type: "text" as const, text: "n" }],
                  body: [{ type: "text" as const, text: "i" }],
                },
              ],
            },
          },
        ],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain("<m:rad><m:deg><m:r><m:t>3</m:t></m:r></m:deg><m:e><m:r><m:t>x</m:t></m:r></m:e></m:rad>");
    expect(xml).toContain('<m:nary><m:naryPr><m:chr m:val="∑"/></m:naryPr><m:sub><m:r><m:t>i=1</m:t></m:r></m:sub><m:sup><m:r><m:t>n</m:t></m:r></m:sup><m:e><m:r><m:t>i</m:t></m:r></m:e></m:nary>');
  });

  it("writes radical hide degree", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "radical" as const,
                  hideDegree: false,
                  degree: [{ type: "text" as const, text: "3" }],
                  content: [{ type: "text" as const, text: "x" }],
                },
              ],
            },
          },
        ],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<m:rad><m:radPr><m:degHide m:val="0"/></m:radPr><m:deg><m:r><m:t>3</m:t></m:r></m:deg><m:e><m:r><m:t>x</m:t></m:r></m:e></m:rad>');
  });

  it("writes radical control properties", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "radical" as const,
                  controlProperties: {
                    bold: true,
                    italic: true,
                    fontFamily: "Cambria Math",
                    fontSize: 16,
                    color: "8064A2",
                    underline: true,
                    highlight: "darkYellow",
                  },
                  degree: [{ type: "text" as const, text: "3" }],
                  content: [{ type: "text" as const, text: "x" }],
                },
              ],
            },
          },
        ],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<m:rad><m:radPr><m:ctrlPr><w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/><w:b/><w:i/><w:color w:val="8064A2"/><w:sz w:val="32"/><w:highlight w:val="darkYellow"/><w:u w:val="single"/></w:rPr></m:ctrlPr></m:radPr><m:deg><m:r><m:t>3</m:t></m:r></m:deg><m:e><m:r><m:t>x</m:t></m:r></m:e></m:rad>');
  });

  it("writes nary control properties", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "nary" as const,
                  operator: "sum" as const,
                  controlProperties: {
                    bold: true,
                    italic: true,
                    fontFamily: "Cambria Math",
                    fontSize: 16,
                    color: "8064A2",
                    underline: true,
                    highlight: "darkYellow",
                  },
                  lowerLimit: [{ type: "text" as const, text: "i=1" }],
                  upperLimit: [{ type: "text" as const, text: "n" }],
                  body: [{ type: "text" as const, text: "i" }],
                },
              ],
            },
          },
        ],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<m:nary><m:naryPr><m:chr m:val="∑"/><m:ctrlPr><w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/><w:b/><w:i/><w:color w:val="8064A2"/><w:sz w:val="32"/><w:highlight w:val="darkYellow"/><w:u w:val="single"/></w:rPr></m:ctrlPr></m:naryPr><m:sub><m:r><m:t>i=1</m:t></m:r></m:sub><m:sup><m:r><m:t>n</m:t></m:r></m:sup><m:e><m:r><m:t>i</m:t></m:r></m:e></m:nary>');
  });

  it("writes nary operator", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "nary" as const,
                  operator: "product" as const,
                  lowerLimit: [{ type: "text" as const, text: "i=1" }],
                  upperLimit: [{ type: "text" as const, text: "n" }],
                  body: [{ type: "text" as const, text: "i" }],
                },
              ],
            },
          },
        ],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<m:nary><m:naryPr><m:chr m:val="∏"/></m:naryPr><m:sub><m:r><m:t>i=1</m:t></m:r></m:sub><m:sup><m:r><m:t>n</m:t></m:r></m:sup><m:e><m:r><m:t>i</m:t></m:r></m:e></m:nary>');
  });

  it("writes nary custom operator character", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "nary" as const,
                  operator: "sum" as const,
                  operatorCharacter: "⊕",
                  lowerLimit: [{ type: "text" as const, text: "i=1" }],
                  upperLimit: [{ type: "text" as const, text: "n" }],
                  body: [{ type: "text" as const, text: "i" }],
                },
              ],
            },
          },
        ],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<m:nary><m:naryPr><m:chr m:val="⊕"/></m:naryPr><m:sub><m:r><m:t>i=1</m:t></m:r></m:sub><m:sup><m:r><m:t>n</m:t></m:r></m:sup><m:e><m:r><m:t>i</m:t></m:r></m:e></m:nary>');
  });

  it("writes nary limit location", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "nary" as const,
                  operator: "sum" as const,
                  limitLocation: "underOver" as const,
                  lowerLimit: [{ type: "text" as const, text: "i=1" }],
                  upperLimit: [{ type: "text" as const, text: "n" }],
                  body: [{ type: "text" as const, text: "i" }],
                },
              ],
            },
          },
        ],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<m:nary><m:naryPr><m:chr m:val="∑"/><m:limLoc m:val="undOvr"/></m:naryPr><m:sub><m:r><m:t>i=1</m:t></m:r></m:sub><m:sup><m:r><m:t>n</m:t></m:r></m:sup><m:e><m:r><m:t>i</m:t></m:r></m:e></m:nary>');
  });

  it("writes nary grow", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "nary" as const,
                  operator: "sum" as const,
                  grow: false,
                  lowerLimit: [{ type: "text" as const, text: "i=1" }],
                  upperLimit: [{ type: "text" as const, text: "n" }],
                  body: [{ type: "text" as const, text: "i" }],
                },
              ],
            },
          },
        ],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<m:nary><m:naryPr><m:chr m:val="∑"/><m:grow m:val="0"/></m:naryPr><m:sub><m:r><m:t>i=1</m:t></m:r></m:sub><m:sup><m:r><m:t>n</m:t></m:r></m:sup><m:e><m:r><m:t>i</m:t></m:r></m:e></m:nary>');
  });

  it("writes nary hide limits", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "nary" as const,
                  operator: "sum" as const,
                  hideLowerLimit: true,
                  hideUpperLimit: true,
                  lowerLimit: [{ type: "text" as const, text: "i=1" }],
                  upperLimit: [{ type: "text" as const, text: "n" }],
                  body: [{ type: "text" as const, text: "i" }],
                },
              ],
            },
          },
        ],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<m:nary><m:naryPr><m:chr m:val="∑"/><m:subHide m:val="1"/><m:supHide m:val="1"/></m:naryPr><m:sub><m:r><m:t>i=1</m:t></m:r></m:sub><m:sup><m:r><m:t>n</m:t></m:r></m:sup><m:e><m:r><m:t>i</m:t></m:r></m:e></m:nary>');
  });

  it("writes nary explicit visible limits", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "nary" as const,
                  operator: "sum" as const,
                  hideLowerLimit: false,
                  hideUpperLimit: false,
                  lowerLimit: [{ type: "text" as const, text: "i=1" }],
                  upperLimit: [{ type: "text" as const, text: "n" }],
                  body: [{ type: "text" as const, text: "i" }],
                },
              ],
            },
          },
        ],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<m:nary><m:naryPr><m:chr m:val="∑"/><m:subHide m:val="0"/><m:supHide m:val="0"/></m:naryPr><m:sub><m:r><m:t>i=1</m:t></m:r></m:sub><m:sup><m:r><m:t>n</m:t></m:r></m:sup><m:e><m:r><m:t>i</m:t></m:r></m:e></m:nary>');
  });

  it("writes matrix office math runs", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "matrix" as const,
                  rows: [
                    [[{ type: "text" as const, text: "a" }], [{ type: "text" as const, text: "b" }]],
                    [[{ type: "text" as const, text: "c" }], [{ type: "text" as const, text: "d" }]],
                  ],
                },
              ],
            },
          },
        ],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain("<m:m><m:mr><m:e><m:r><m:t>a</m:t></m:r></m:e><m:e><m:r><m:t>b</m:t></m:r></m:e></m:mr><m:mr><m:e><m:r><m:t>c</m:t></m:r></m:e><m:e><m:r><m:t>d</m:t></m:r></m:e></m:mr></m:m>");
  });

  it("writes matrix control properties", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "matrix" as const,
                  controlProperties: {
                    bold: true,
                    italic: true,
                    fontFamily: "Cambria Math",
                    fontSize: 16,
                    color: "8064A2",
                    underline: true,
                    highlight: "darkYellow",
                  },
                  rows: [
                    [[{ type: "text" as const, text: "a" }], [{ type: "text" as const, text: "b" }]],
                  ],
                },
              ],
            },
          },
        ],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<m:m><m:mPr><m:ctrlPr><w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/><w:b/><w:i/><w:color w:val="8064A2"/><w:sz w:val="32"/><w:highlight w:val="darkYellow"/><w:u w:val="single"/></w:rPr></m:ctrlPr></m:mPr><m:mr><m:e><m:r><m:t>a</m:t></m:r></m:e><m:e><m:r><m:t>b</m:t></m:r></m:e></m:mr></m:m>');
  });

  it("writes matrix base justification", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "matrix" as const,
                  baseJustification: "bottom",
                  rows: [
                    [
                      [{ type: "text" as const, text: "a" }],
                      [{ type: "text" as const, text: "b" }],
                    ],
                  ],
                },
              ],
            },
          },
        ],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<m:m><m:mPr><m:baseJc m:val="bot"/></m:mPr><m:mr><m:e><m:r><m:t>a</m:t></m:r></m:e><m:e><m:r><m:t>b</m:t></m:r></m:e></m:mr></m:m>');
  });

  it("writes matrix column justifications", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "matrix" as const,
                  columnJustifications: ["left", "center", "right"],
                  rows: [
                    [
                      [{ type: "text" as const, text: "a" }],
                      [{ type: "text" as const, text: "b" }],
                      [{ type: "text" as const, text: "c" }],
                    ],
                  ],
                },
              ],
            },
          },
        ],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<m:m><m:mPr><m:mcs><m:mc><m:mcPr><m:mcJc m:val="left"/></m:mcPr></m:mc><m:mc><m:mcPr><m:mcJc m:val="center"/></m:mcPr></m:mc><m:mc><m:mcPr><m:mcJc m:val="right"/></m:mcPr></m:mc></m:mcs></m:mPr><m:mr><m:e><m:r><m:t>a</m:t></m:r></m:e><m:e><m:r><m:t>b</m:t></m:r></m:e><m:e><m:r><m:t>c</m:t></m:r></m:e></m:mr></m:m>');
  });

  it("writes matrix column counts", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "matrix" as const,
                  columnCounts: [1, 2, 1],
                  rows: [
                    [
                      [{ type: "text" as const, text: "a" }],
                      [{ type: "text" as const, text: "b" }],
                      [{ type: "text" as const, text: "c" }],
                    ],
                  ],
                },
              ],
            },
          },
        ],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<m:m><m:mPr><m:mcs><m:mc><m:mcPr><m:count m:val="1"/></m:mcPr></m:mc><m:mc><m:mcPr><m:count m:val="2"/></m:mcPr></m:mc><m:mc><m:mcPr><m:count m:val="1"/></m:mcPr></m:mc></m:mcs></m:mPr><m:mr><m:e><m:r><m:t>a</m:t></m:r></m:e><m:e><m:r><m:t>b</m:t></m:r></m:e><m:e><m:r><m:t>c</m:t></m:r></m:e></m:mr></m:m>');
  });

  it("writes matrix row spacing", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "matrix" as const,
                  rowSpacing: 3,
                  rows: [
                    [[{ type: "text" as const, text: "a" }]],
                    [[{ type: "text" as const, text: "b" }]],
                  ],
                },
              ],
            },
          },
        ],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<m:m><m:mPr><m:rSp m:val="3"/></m:mPr><m:mr><m:e><m:r><m:t>a</m:t></m:r></m:e></m:mr><m:mr><m:e><m:r><m:t>b</m:t></m:r></m:e></m:mr></m:m>');
  });

  it("writes matrix row spacing rule", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "matrix" as const,
                  rowSpacing: 3,
                  rowSpacingRule: "exactly",
                  rows: [
                    [[{ type: "text" as const, text: "a" }]],
                    [[{ type: "text" as const, text: "b" }]],
                  ],
                },
              ],
            },
          },
        ],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<m:m><m:mPr><m:rSp m:val="3"/><m:rSpRule m:val="exactly"/></m:mPr><m:mr><m:e><m:r><m:t>a</m:t></m:r></m:e></m:mr><m:mr><m:e><m:r><m:t>b</m:t></m:r></m:e></m:mr></m:m>');
  });

  it("writes matrix column spacing", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "matrix" as const,
                  columnSpacing: 4,
                  rows: [
                    [
                      [{ type: "text" as const, text: "a" }],
                      [{ type: "text" as const, text: "b" }],
                    ],
                  ],
                },
              ],
            },
          },
        ],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<m:m><m:mPr><m:cSp m:val="4"/></m:mPr><m:mr><m:e><m:r><m:t>a</m:t></m:r></m:e><m:e><m:r><m:t>b</m:t></m:r></m:e></m:mr></m:m>');
  });

  it("writes matrix column spacing rule", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "matrix" as const,
                  columnSpacing: 4,
                  columnSpacingRule: "multiple",
                  rows: [
                    [
                      [{ type: "text" as const, text: "a" }],
                      [{ type: "text" as const, text: "b" }],
                    ],
                  ],
                },
              ],
            },
          },
        ],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<m:m><m:mPr><m:cSp m:val="4"/><m:cSpRule m:val="multiple"/></m:mPr><m:mr><m:e><m:r><m:t>a</m:t></m:r></m:e><m:e><m:r><m:t>b</m:t></m:r></m:e></m:mr></m:m>');
  });

  it("writes delimiter office math runs", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "delimiter" as const,
                  begin: "(",
                  end: ")",
                  content: [
                    {
                      type: "fraction" as const,
                      numerator: [{ type: "text" as const, text: "a" }],
                      denominator: [{ type: "text" as const, text: "b" }],
                    },
                  ],
                },
              ],
            },
          },
        ],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<m:d><m:dPr><m:begChr m:val="("/><m:endChr m:val=")"/></m:dPr><m:e><m:f><m:num><m:r><m:t>a</m:t></m:r></m:num><m:den><m:r><m:t>b</m:t></m:r></m:den></m:f></m:e></m:d>');
  });

  it("writes delimiter control properties", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "delimiter" as const,
                  begin: "[",
                  end: "]",
                  controlProperties: {
                    bold: true,
                    italic: true,
                    fontFamily: "Cambria Math",
                    fontSize: 16,
                    color: "8064A2",
                    underline: true,
                    highlight: "darkYellow",
                  },
                  content: [{ type: "text" as const, text: "x" }],
                },
              ],
            },
          },
        ],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<m:d><m:dPr><m:begChr m:val="["/><m:endChr m:val="]"/><m:ctrlPr><w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/><w:b/><w:i/><w:color w:val="8064A2"/><w:sz w:val="32"/><w:highlight w:val="darkYellow"/><w:u w:val="single"/></w:rPr></m:ctrlPr></m:dPr><m:e><m:r><m:t>x</m:t></m:r></m:e></m:d>');
  });

  it("writes delimiter grow", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "delimiter" as const,
                  begin: "(",
                  end: ")",
                  grow: false,
                  content: [{ type: "text" as const, text: "x" }],
                },
              ],
            },
          },
        ],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<m:d><m:dPr><m:begChr m:val="("/><m:endChr m:val=")"/><m:grow m:val="0"/></m:dPr><m:e><m:r><m:t>x</m:t></m:r></m:e></m:d>');
  });

  it("writes delimiter separator", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "delimiter" as const,
                  begin: "{",
                  end: "}",
                  separator: "|",
                  content: [{ type: "text" as const, text: "x" }],
                },
              ],
            },
          },
        ],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<m:d><m:dPr><m:begChr m:val="{"/><m:endChr m:val="}"/><m:sepChr m:val="|"/></m:dPr><m:e><m:r><m:t>x</m:t></m:r></m:e></m:d>');
  });

  it("writes accent office math runs", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "accent" as const,
                  mark: "¯",
                  content: [{ type: "text" as const, text: "x" }],
                },
              ],
            },
          },
        ],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<m:acc><m:accPr><m:chr m:val="¯"/></m:accPr><m:e><m:r><m:t>x</m:t></m:r></m:e></m:acc>');
  });

  it("writes accent control properties", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "accent" as const,
                  mark: "¯",
                  controlProperties: {
                    bold: true,
                    italic: true,
                    fontFamily: "Cambria Math",
                    fontSize: 16,
                    color: "8064A2",
                    underline: true,
                    highlight: "darkYellow",
                  },
                  content: [{ type: "text" as const, text: "x" }],
                },
              ],
            },
          },
        ],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<m:acc><m:accPr><m:chr m:val="¯"/><m:ctrlPr><w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/><w:b/><w:i/><w:color w:val="8064A2"/><w:sz w:val="32"/><w:highlight w:val="darkYellow"/><w:u w:val="single"/></w:rPr></m:ctrlPr></m:accPr><m:e><m:r><m:t>x</m:t></m:r></m:e></m:acc>');
  });

  it("writes bar office math runs", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "bar" as const,
                  position: "top",
                  content: [{ type: "text" as const, text: "x+y" }],
                },
              ],
            },
          },
        ],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<m:bar><m:barPr><m:pos m:val="top"/></m:barPr><m:e><m:r><m:t>x+y</m:t></m:r></m:e></m:bar>');
  });

  it("writes bar control properties", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "bar" as const,
                  position: "top" as const,
                  controlProperties: {
                    bold: true,
                    italic: true,
                    fontFamily: "Cambria Math",
                    fontSize: 16,
                    color: "8064A2",
                    underline: true,
                    highlight: "darkYellow",
                  },
                  content: [{ type: "text" as const, text: "x+y" }],
                },
              ],
            },
          },
        ],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<m:bar><m:barPr><m:pos m:val="top"/><m:ctrlPr><w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/><w:b/><w:i/><w:color w:val="8064A2"/><w:sz w:val="32"/><w:highlight w:val="darkYellow"/><w:u w:val="single"/></w:rPr></m:ctrlPr></m:barPr><m:e><m:r><m:t>x+y</m:t></m:r></m:e></m:bar>');
  });

  it("writes function office math runs", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "function" as const,
                  name: [{ type: "text" as const, text: "sin" }],
                  argument: [
                    {
                      type: "delimiter" as const,
                      begin: "(",
                      end: ")",
                      content: [{ type: "text" as const, text: "x" }],
                    },
                  ],
                },
              ],
            },
          },
        ],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<m:func><m:fName><m:r><m:t>sin</m:t></m:r></m:fName><m:e><m:d><m:dPr><m:begChr m:val="("/><m:endChr m:val=")"/></m:dPr><m:e><m:r><m:t>x</m:t></m:r></m:e></m:d></m:e></m:func>');
  });

  it("writes function control properties", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "function" as const,
                  controlProperties: {
                    bold: true,
                    italic: true,
                    fontFamily: "Cambria Math",
                    fontSize: 16,
                    color: "8064A2",
                    underline: true,
                    highlight: "darkYellow",
                  },
                  name: [{ type: "text" as const, text: "sin" }],
                  argument: [{ type: "text" as const, text: "x" }],
                },
              ],
            },
          },
        ],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<m:func><m:funcPr><m:ctrlPr><w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/><w:b/><w:i/><w:color w:val="8064A2"/><w:sz w:val="32"/><w:highlight w:val="darkYellow"/><w:u w:val="single"/></w:rPr></m:ctrlPr></m:funcPr><m:fName><m:r><m:t>sin</m:t></m:r></m:fName><m:e><m:r><m:t>x</m:t></m:r></m:e></m:func>');
  });

  it("writes limit office math runs", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "limitLower" as const,
                  base: [{ type: "text" as const, text: "lim" }],
                  limit: [{ type: "text" as const, text: "x→0" }],
                },
                {
                  type: "limitUpper" as const,
                  base: [{ type: "text" as const, text: "max" }],
                  limit: [{ type: "text" as const, text: "n" }],
                },
              ],
            },
          },
        ],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain("<m:limLow><m:e><m:r><m:t>lim</m:t></m:r></m:e><m:lim><m:r><m:t>x→0</m:t></m:r></m:lim></m:limLow>");
    expect(xml).toContain("<m:limUpp><m:e><m:r><m:t>max</m:t></m:r></m:e><m:lim><m:r><m:t>n</m:t></m:r></m:lim></m:limUpp>");
  });

  it("writes limit lower control properties", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "limitLower" as const,
                  controlProperties: {
                    bold: true,
                    italic: true,
                    fontFamily: "Cambria Math",
                    fontSize: 16,
                    color: "8064A2",
                    underline: true,
                    highlight: "darkYellow",
                  },
                  base: [{ type: "text" as const, text: "lim" }],
                  limit: [{ type: "text" as const, text: "x→0" }],
                },
              ],
            },
          },
        ],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<m:limLow><m:limLowPr><m:ctrlPr><w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/><w:b/><w:i/><w:color w:val="8064A2"/><w:sz w:val="32"/><w:highlight w:val="darkYellow"/><w:u w:val="single"/></w:rPr></m:ctrlPr></m:limLowPr><m:e><m:r><m:t>lim</m:t></m:r></m:e><m:lim><m:r><m:t>x→0</m:t></m:r></m:lim></m:limLow>');
  });

  it("writes limit upper control properties", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "limitUpper" as const,
                  controlProperties: {
                    bold: true,
                    italic: true,
                    fontFamily: "Cambria Math",
                    fontSize: 16,
                    color: "8064A2",
                    underline: true,
                    highlight: "darkYellow",
                  },
                  base: [{ type: "text" as const, text: "max" }],
                  limit: [{ type: "text" as const, text: "n" }],
                },
              ],
            },
          },
        ],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<m:limUpp><m:limUppPr><m:ctrlPr><w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/><w:b/><w:i/><w:color w:val="8064A2"/><w:sz w:val="32"/><w:highlight w:val="darkYellow"/><w:u w:val="single"/></w:rPr></m:ctrlPr></m:limUppPr><m:e><m:r><m:t>max</m:t></m:r></m:e><m:lim><m:r><m:t>n</m:t></m:r></m:lim></m:limUpp>');
  });

  it("writes equation array office math runs", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "equationArray" as const,
                  rows: [
                    [{ type: "text" as const, text: "x=1" }],
                    [
                      {
                        type: "fraction" as const,
                        numerator: [{ type: "text" as const, text: "a" }],
                        denominator: [{ type: "text" as const, text: "b" }],
                      },
                    ],
                  ],
                },
              ],
            },
          },
        ],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain("<m:eqArr><m:e><m:r><m:t>x=1</m:t></m:r></m:e><m:e><m:f><m:num><m:r><m:t>a</m:t></m:r></m:num><m:den><m:r><m:t>b</m:t></m:r></m:den></m:f></m:e></m:eqArr>");
  });

  it("writes equation array control properties", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "equationArray" as const,
                  controlProperties: {
                    bold: true,
                    italic: true,
                    fontFamily: "Cambria Math",
                    fontSize: 16,
                    color: "8064A2",
                    underline: true,
                    highlight: "darkYellow",
                  },
                  rows: [
                    [{ type: "text" as const, text: "x=1" }],
                    [{ type: "text" as const, text: "y=2" }],
                  ],
                },
              ],
            },
          },
        ],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<m:eqArr><m:eqArrPr><m:ctrlPr><w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/><w:b/><w:i/><w:color w:val="8064A2"/><w:sz w:val="32"/><w:highlight w:val="darkYellow"/><w:u w:val="single"/></w:rPr></m:ctrlPr></m:eqArrPr><m:e><m:r><m:t>x=1</m:t></m:r></m:e><m:e><m:r><m:t>y=2</m:t></m:r></m:e></m:eqArr>');
  });

  it("writes equation array row spacing", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "equationArray" as const,
                  rowSpacing: 3,
                  rows: [
                    [{ type: "text" as const, text: "x=1" }],
                    [{ type: "text" as const, text: "y=2" }],
                  ],
                },
              ],
            },
          },
        ],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<m:eqArr><m:eqArrPr><m:rSp m:val="3"/></m:eqArrPr><m:e><m:r><m:t>x=1</m:t></m:r></m:e><m:e><m:r><m:t>y=2</m:t></m:r></m:e></m:eqArr>');
  });

  it("writes equation array base justification", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "equationArray" as const,
                  baseJustification: "bottom",
                  rows: [
                    [{ type: "text" as const, text: "x=1" }],
                    [{ type: "text" as const, text: "y=2" }],
                  ],
                },
              ],
            },
          },
        ],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<m:eqArr><m:eqArrPr><m:baseJc m:val="bot"/></m:eqArrPr><m:e><m:r><m:t>x=1</m:t></m:r></m:e><m:e><m:r><m:t>y=2</m:t></m:r></m:e></m:eqArr>');
  });

  it("writes equation array vertical justification", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "equationArray" as const,
                  verticalJustification: "bottom",
                  rows: [
                    [{ type: "text" as const, text: "x=1" }],
                    [{ type: "text" as const, text: "y=2" }],
                  ],
                },
              ],
            },
          },
        ],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<m:eqArr><m:eqArrPr><m:vertJc m:val="bottom"/></m:eqArrPr><m:e><m:r><m:t>x=1</m:t></m:r></m:e><m:e><m:r><m:t>y=2</m:t></m:r></m:e></m:eqArr>');
  });

  it("writes equation array alignment", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "equationArray" as const,
                  alignment: false,
                  rows: [
                    [{ type: "text" as const, text: "x=1" }],
                    [{ type: "text" as const, text: "y=2" }],
                  ],
                },
              ],
            },
          },
        ],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<m:eqArr><m:eqArrPr><m:aln m:val="0"/></m:eqArrPr><m:e><m:r><m:t>x=1</m:t></m:r></m:e><m:e><m:r><m:t>y=2</m:t></m:r></m:e></m:eqArr>');
  });

  it("writes equation array row spacing rule", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "equationArray" as const,
                  rowSpacing: 3,
                  rowSpacingRule: "exactly",
                  rows: [
                    [{ type: "text" as const, text: "x=1" }],
                    [{ type: "text" as const, text: "y=2" }],
                  ],
                },
              ],
            },
          },
        ],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<m:eqArr><m:eqArrPr><m:rSp m:val="3"/><m:rSpRule m:val="exactly"/></m:eqArrPr><m:e><m:r><m:t>x=1</m:t></m:r></m:e><m:e><m:r><m:t>y=2</m:t></m:r></m:e></m:eqArr>');
  });

  it("writes equation array object distribution", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "equationArray" as const,
                  objectDistribution: false,
                  rows: [
                    [{ type: "text" as const, text: "x=1" }],
                    [{ type: "text" as const, text: "y=2" }],
                  ],
                },
              ],
            },
          },
        ],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<m:eqArr><m:eqArrPr><m:objDist m:val="0"/></m:eqArrPr><m:e><m:r><m:t>x=1</m:t></m:r></m:e><m:e><m:r><m:t>y=2</m:t></m:r></m:e></m:eqArr>');
  });

  it("writes equation array max distribution", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "equationArray" as const,
                  maxDistribution: false,
                  rows: [
                    [{ type: "text" as const, text: "x=1" }],
                    [{ type: "text" as const, text: "y=2" }],
                  ],
                },
              ],
            },
          },
        ],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<m:eqArr><m:eqArrPr><m:maxDist m:val="0"/></m:eqArrPr><m:e><m:r><m:t>x=1</m:t></m:r></m:e><m:e><m:r><m:t>y=2</m:t></m:r></m:e></m:eqArr>');
  });

  it("writes box office math runs", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "box" as const,
                  hideTop: true,
                  hideBottom: true,
                  content: [{ type: "text" as const, text: "x+y" }],
                },
              ],
            },
          },
        ],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<m:box><m:boxPr><m:hideTop m:val="1"/><m:hideBot m:val="1"/></m:boxPr><m:e><m:r><m:t>x+y</m:t></m:r></m:e></m:box>');
  });

  it("writes box explicit visible borders", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "box" as const,
                  hideTop: false,
                  hideBottom: false,
                  hideLeft: false,
                  hideRight: false,
                  content: [{ type: "text" as const, text: "x+y" }],
                },
              ],
            },
          },
        ],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<m:box><m:boxPr><m:hideTop m:val="0"/><m:hideBot m:val="0"/><m:hideLeft m:val="0"/><m:hideRight m:val="0"/></m:boxPr><m:e><m:r><m:t>x+y</m:t></m:r></m:e></m:box>');
  });

  it("writes box control properties", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "box" as const,
                  hideTop: true,
                  hideBottom: true,
                  controlProperties: {
                    bold: true,
                    italic: true,
                    fontFamily: "Cambria Math",
                    fontSize: 16,
                    color: "8064A2",
                    underline: true,
                    highlight: "darkYellow",
                  },
                  content: [{ type: "text" as const, text: "x+y" }],
                },
              ],
            },
          },
        ],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<m:box><m:boxPr><m:hideTop m:val="1"/><m:hideBot m:val="1"/><m:ctrlPr><w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/><w:b/><w:i/><w:color w:val="8064A2"/><w:sz w:val="32"/><w:highlight w:val="darkYellow"/><w:u w:val="single"/></w:rPr></m:ctrlPr></m:boxPr><m:e><m:r><m:t>x+y</m:t></m:r></m:e></m:box>');
  });

  it("writes border box office math runs", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "borderBox" as const,
                  hideTop: true,
                  hideBottom: true,
                  content: [{ type: "text" as const, text: "x+y" }],
                },
              ],
            },
          },
        ],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<m:borderBox><m:borderBoxPr><m:hideTop m:val="1"/><m:hideBot m:val="1"/></m:borderBoxPr><m:e><m:r><m:t>x+y</m:t></m:r></m:e></m:borderBox>');
  });

  it("writes border box explicit visible borders", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "borderBox" as const,
                  hideTop: false,
                  hideBottom: false,
                  hideLeft: false,
                  hideRight: false,
                  content: [{ type: "text" as const, text: "x+y" }],
                },
              ],
            },
          },
        ],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<m:borderBox><m:borderBoxPr><m:hideTop m:val="0"/><m:hideBot m:val="0"/><m:hideLeft m:val="0"/><m:hideRight m:val="0"/></m:borderBoxPr><m:e><m:r><m:t>x+y</m:t></m:r></m:e></m:borderBox>');
  });

  it("writes border box control properties", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "borderBox" as const,
                  hideTop: true,
                  hideBottom: true,
                  controlProperties: {
                    bold: true,
                    italic: true,
                    fontFamily: "Cambria Math",
                    fontSize: 16,
                    color: "8064A2",
                    underline: true,
                    highlight: "darkYellow",
                  },
                  content: [{ type: "text" as const, text: "x+y" }],
                },
              ],
            },
          },
        ],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<m:borderBox><m:borderBoxPr><m:hideTop m:val="1"/><m:hideBot m:val="1"/><m:ctrlPr><w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/><w:b/><w:i/><w:color w:val="8064A2"/><w:sz w:val="32"/><w:highlight w:val="darkYellow"/><w:u w:val="single"/></w:rPr></m:ctrlPr></m:borderBoxPr><m:e><m:r><m:t>x+y</m:t></m:r></m:e></m:borderBox>');
  });

  it("writes phantom office math runs", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "phantom" as const,
                  show: false,
                  zeroWidth: true,
                  content: [{ type: "text" as const, text: "x+y" }],
                },
              ],
            },
          },
        ],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<m:phant><m:phantPr><m:show m:val="0"/><m:zeroWid m:val="1"/></m:phantPr><m:e><m:r><m:t>x+y</m:t></m:r></m:e></m:phant>');
  });

  it("writes phantom explicit false properties", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "phantom" as const,
                  zeroWidth: false,
                  zeroAscent: false,
                  zeroDescent: false,
                  transparent: false,
                  content: [{ type: "text" as const, text: "x+y" }],
                },
              ],
            },
          },
        ],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<m:phant><m:phantPr><m:zeroWid m:val="0"/><m:zeroAsc m:val="0"/><m:zeroDesc m:val="0"/><m:transp m:val="0"/></m:phantPr><m:e><m:r><m:t>x+y</m:t></m:r></m:e></m:phant>');
  });

  it("writes phantom control properties", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "phantom" as const,
                  show: false,
                  zeroWidth: true,
                  controlProperties: {
                    bold: true,
                    italic: true,
                    fontFamily: "Cambria Math",
                    fontSize: 16,
                    color: "8064A2",
                    underline: true,
                    highlight: "darkYellow",
                  },
                  content: [{ type: "text" as const, text: "x+y" }],
                },
              ],
            },
          },
        ],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<m:phant><m:phantPr><m:show m:val="0"/><m:zeroWid m:val="1"/><m:ctrlPr><w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/><w:b/><w:i/><w:color w:val="8064A2"/><w:sz w:val="32"/><w:highlight w:val="darkYellow"/><w:u w:val="single"/></w:rPr></m:ctrlPr></m:phantPr><m:e><m:r><m:t>x+y</m:t></m:r></m:e></m:phant>');
  });

  it("writes group character office math runs", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "groupCharacter" as const,
                  character: "⏞",
                  position: "top",
                  verticalJustification: "bottom",
                  content: [{ type: "text" as const, text: "x+y" }],
                },
              ],
            },
          },
        ],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<m:groupChr><m:groupChrPr><m:chr m:val="⏞"/><m:pos m:val="top"/><m:vertJc m:val="bottom"/></m:groupChrPr><m:e><m:r><m:t>x+y</m:t></m:r></m:e></m:groupChr>');
  });

  it("writes group character control properties", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "groupCharacter" as const,
                  character: "⏞",
                  position: "top",
                  verticalJustification: "bottom",
                  controlProperties: {
                    bold: true,
                    italic: true,
                    fontFamily: "Cambria Math",
                    fontSize: 16,
                    color: "8064A2",
                    underline: true,
                    highlight: "darkYellow",
                  },
                  content: [{ type: "text" as const, text: "x+y" }],
                },
              ],
            },
          },
        ],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<m:groupChr><m:groupChrPr><m:chr m:val="⏞"/><m:pos m:val="top"/><m:vertJc m:val="bottom"/><m:ctrlPr><w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/><w:b/><w:i/><w:color w:val="8064A2"/><w:sz w:val="32"/><w:highlight w:val="darkYellow"/><w:u w:val="single"/></w:rPr></m:ctrlPr></m:groupChrPr><m:e><m:r><m:t>x+y</m:t></m:r></m:e></m:groupChr>');
  });

  it("writes pre sub sup office math runs", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "preSubSup" as const,
                  base: [{ type: "text" as const, text: "X" }],
                  subscript: [{ type: "text" as const, text: "i" }],
                  superscript: [{ type: "text" as const, text: "j" }],
                },
              ],
            },
          },
        ],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<m:preSubSup><m:e><m:r><m:t>X</m:t></m:r></m:e><m:sub><m:r><m:t>i</m:t></m:r></m:sub><m:sup><m:r><m:t>j</m:t></m:r></m:sup></m:preSubSup>');
  });

  it("writes pre sub sup control properties", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "preSubSup" as const,
                  controlProperties: {
                    bold: true,
                    italic: true,
                    fontFamily: "Cambria Math",
                    fontSize: 16,
                    color: "8064A2",
                    underline: true,
                    highlight: "darkYellow",
                  },
                  base: [{ type: "text" as const, text: "X" }],
                  subscript: [{ type: "text" as const, text: "i" }],
                  superscript: [{ type: "text" as const, text: "j" }],
                },
              ],
            },
          },
        ],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<m:preSubSup><m:preSubSupPr><m:ctrlPr><w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/><w:b/><w:i/><w:color w:val="8064A2"/><w:sz w:val="32"/><w:highlight w:val="darkYellow"/><w:u w:val="single"/></w:rPr></m:ctrlPr></m:preSubSupPr><m:e><m:r><m:t>X</m:t></m:r></m:e><m:sub><m:r><m:t>i</m:t></m:r></m:sub><m:sup><m:r><m:t>j</m:t></m:r></m:sup></m:preSubSup>');
  });

  it("writes sub sup office math runs", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "subSup" as const,
                  base: [{ type: "text" as const, text: "x" }],
                  subscript: [{ type: "text" as const, text: "i" }],
                  superscript: [{ type: "text" as const, text: "2" }],
                },
              ],
            },
          },
        ],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<m:sSubSup><m:e><m:r><m:t>x</m:t></m:r></m:e><m:sub><m:r><m:t>i</m:t></m:r></m:sub><m:sup><m:r><m:t>2</m:t></m:r></m:sup></m:sSubSup>');
  });

  it("writes sub sup control properties", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "subSup" as const,
                  controlProperties: {
                    bold: true,
                    italic: true,
                    fontFamily: "Cambria Math",
                    fontSize: 13,
                    color: "0070C0",
                    underline: true,
                    highlight: "cyan",
                  },
                  base: [{ type: "text" as const, text: "x" }],
                  subscript: [{ type: "text" as const, text: "i" }],
                  superscript: [{ type: "text" as const, text: "2" }],
                },
              ],
            },
          },
        ],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<m:sSubSup><m:sSubSupPr><m:ctrlPr><w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/><w:b/><w:i/><w:color w:val="0070C0"/><w:sz w:val="26"/><w:highlight w:val="cyan"/><w:u w:val="single"/></w:rPr></m:ctrlPr></m:sSubSupPr><m:e><m:r><m:t>x</m:t></m:r></m:e><m:sub><m:r><m:t>i</m:t></m:r></m:sub><m:sup><m:r><m:t>2</m:t></m:r></m:sup></m:sSubSup>');
  });

  it("writes s pre office math runs", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "sPre" as const,
                  base: [{ type: "text" as const, text: "X" }],
                  subscript: [{ type: "text" as const, text: "i" }],
                  superscript: [{ type: "text" as const, text: "j" }],
                },
              ],
            },
          },
        ],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<m:sPre><m:sub><m:r><m:t>i</m:t></m:r></m:sub><m:sup><m:r><m:t>j</m:t></m:r></m:sup><m:e><m:r><m:t>X</m:t></m:r></m:e></m:sPre>');
  });

  it("writes s pre control properties", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "sPre" as const,
                  controlProperties: {
                    bold: true,
                    italic: true,
                    fontFamily: "Cambria Math",
                    fontSize: 14,
                    color: "C00000",
                    underline: true,
                    highlight: "yellow",
                  },
                  base: [{ type: "text" as const, text: "X" }],
                  subscript: [{ type: "text" as const, text: "i" }],
                  superscript: [{ type: "text" as const, text: "j" }],
                },
              ],
            },
          },
        ],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<m:sPre><m:sPrePr><m:ctrlPr><w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/><w:b/><w:i/><w:color w:val="C00000"/><w:sz w:val="28"/><w:highlight w:val="yellow"/><w:u w:val="single"/></w:rPr></m:ctrlPr></m:sPrePr><m:sub><m:r><m:t>i</m:t></m:r></m:sub><m:sup><m:r><m:t>j</m:t></m:r></m:sup><m:e><m:r><m:t>X</m:t></m:r></m:e></m:sPre>');
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

  it("writes run highlight and strike", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          { text: "Marked", highlight: "yellow", strike: true, doubleStrike: true },
        ],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<w:highlight w:val="yellow"/>');
    expect(xml).toContain("<w:strike/>");
    expect(xml).toContain("<w:dstrike/>");
  });

  it("writes run caps and vertical align", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          { text: "Formula", smallCaps: true, allCaps: true, verticalAlign: "superscript" as const },
          { text: "2", verticalAlign: "subscript" as const },
        ],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain("<w:smallCaps/>");
    expect(xml).toContain("<w:caps/>");
    expect(xml).toContain('<w:vertAlign w:val="superscript"/>');
    expect(xml).toContain('<w:vertAlign w:val="subscript"/>');
  });

  it("writes run character spacing and scale", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          { text: "Tracked", characterSpacing: 20, scale: 90 },
        ],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<w:spacing w:val="20"/>');
    expect(xml).toContain('<w:w w:val="90"/>');
  });

  it("writes run border", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          { text: "Boxed", border: { style: "single", size: 6, color: "C00000", space: 1 } },
        ],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<w:bdr w:val="single" w:sz="6" w:space="1" w:color="C00000"/>');
  });

  it("writes inserted run revisions", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          { text: "Added", revision: { type: "insert", id: 1, author: "Ada", date: "2026-07-04T00:00:00.000Z" } },
        ],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<w:ins w:id="1" w:author="Ada" w:date="2026-07-04T00:00:00.000Z">');
    expect(xml).toContain("<w:t>Added</w:t>");
  });

  it("writes deleted run revisions", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          { text: "Removed", revision: { type: "delete", id: 2, author: "Lin", date: "2026-07-04T01:00:00.000Z" } },
        ],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<w:del w:id="2" w:author="Lin" w:date="2026-07-04T01:00:00.000Z">');
    expect(xml).toContain("<w:delText>Removed</w:delText>");
  });

  it("writes run move revisions", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          { text: "Moved away", revision: { type: "moveFrom", id: 30, author: "Ada", date: "2026-07-04T09:00:00.000Z" } },
          { text: "Moved here", revision: { type: "moveTo", id: 31, author: "Lin", date: "2026-07-04T10:00:00.000Z" } },
        ],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<w:moveFrom w:id="30" w:author="Ada" w:date="2026-07-04T09:00:00.000Z">');
    expect(xml).toContain("<w:delText>Moved away</w:delText>");
    expect(xml).toContain('<w:moveTo w:id="31" w:author="Lin" w:date="2026-07-04T10:00:00.000Z">');
    expect(xml).toContain("<w:t>Moved here</w:t>");
  });

  it("writes paragraph insert revisions", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        revision: { type: "insert", id: 10, author: "Ada", date: "2026-07-04T02:00:00.000Z" },
        runs: [{ text: "Inserted paragraph" }],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<w:ins w:id="10" w:author="Ada" w:date="2026-07-04T02:00:00.000Z"><w:p>');
    expect(xml).toContain("<w:t>Inserted paragraph</w:t>");
  });

  it("writes paragraph delete revisions", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        revision: { type: "delete", id: 11, author: "Lin", date: "2026-07-04T03:00:00.000Z" },
        runs: [{ text: "Deleted paragraph" }],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<w:del w:id="11" w:author="Lin" w:date="2026-07-04T03:00:00.000Z"><w:p>');
    expect(xml).toContain("<w:t>Deleted paragraph</w:t>");
  });

  it("writes paragraph move revisions", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        revision: { type: "moveFrom", id: 32, author: "Mira", date: "2026-07-04T11:00:00.000Z" },
        runs: [{ text: "Moved paragraph from" }],
      },
      {
        type: "paragraph",
        revision: { type: "moveTo", id: 33, author: "Noor", date: "2026-07-04T12:00:00.000Z" },
        runs: [{ text: "Moved paragraph to" }],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<w:moveFrom w:id="32" w:author="Mira" w:date="2026-07-04T11:00:00.000Z"><w:p>');
    expect(xml).toContain("<w:t>Moved paragraph from</w:t>");
    expect(xml).toContain('<w:moveTo w:id="33" w:author="Noor" w:date="2026-07-04T12:00:00.000Z"><w:p>');
    expect(xml).toContain("<w:t>Moved paragraph to</w:t>");
  });

  it("writes paragraph property change metadata", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        spacing: { before: 240 },
        propertyRevision: { id: 12, author: "Mira", date: "2026-07-04T04:00:00.000Z" },
        runs: [{ text: "Changed spacing" }],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<w:spacing w:before="240"/>');
    expect(xml).toContain('<w:pPrChange w:id="12" w:author="Mira" w:date="2026-07-04T04:00:00.000Z"><w:pPr/></w:pPrChange>');
  });

  it("writes block content controls", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        contentControl: { alias: "Customer Name", tag: "customer.name", lock: "sdtContentLocked" },
        runs: [{ text: "Acme Inc." }],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<w:sdt><w:sdtPr><w:alias w:val="Customer Name"/><w:tag w:val="customer.name"/><w:lock w:val="sdtContentLocked"/></w:sdtPr><w:sdtContent><w:p>');
    expect(xml).toContain("<w:t>Acme Inc.</w:t>");
    expect(xml).toContain("</w:p></w:sdtContent></w:sdt>");
  });

  it("writes run content controls", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          { text: "INV-001", contentControl: { alias: "Invoice Number", tag: "invoice.number" } },
        ],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<w:sdt><w:sdtPr><w:alias w:val="Invoice Number"/><w:tag w:val="invoice.number"/></w:sdtPr><w:sdtContent><w:r><w:t>INV-001</w:t></w:r></w:sdtContent></w:sdt>');
  });

  it("writes checkbox content controls", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "☒",
            contentControl: {
              alias: "Accepted",
              tag: "accepted",
              checkbox: { checked: true, checkedSymbol: "2612", uncheckedSymbol: "2610" },
            },
          },
        ],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<w:checkBox><w:checked w:val="1"/><w:checkedState w:val="2612"/><w:uncheckedState w:val="2610"/></w:checkBox>');
  });

  it("writes dropdown content controls", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "Gold",
            contentControl: {
              alias: "Plan",
              tag: "plan",
              dropdown: {
                items: [
                  { displayText: "Silver", value: "silver" },
                  { displayText: "Gold", value: "gold" },
                ],
              },
            },
          },
        ],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<w:dropDownList><w:listItem w:displayText="Silver" w:value="silver"/><w:listItem w:displayText="Gold" w:value="gold"/></w:dropDownList>');
  });

  it("writes combo box content controls", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "Custom",
            contentControl: {
              alias: "Choice",
              tag: "choice",
              comboBox: {
                items: [
                  { displayText: "Standard", value: "standard" },
                  { displayText: "Custom", value: "custom" },
                ],
              },
            },
          },
        ],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<w:comboBox><w:listItem w:displayText="Standard" w:value="standard"/><w:listItem w:displayText="Custom" w:value="custom"/></w:comboBox>');
  });

  it("writes date content controls", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "2026-07-04",
            contentControl: {
              alias: "Due Date",
              tag: "dueDate",
              date: { fullDate: "2026-07-04T00:00:00Z", format: "yyyy-MM-dd" },
            },
          },
        ],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<w:date><w:fullDate w:val="2026-07-04T00:00:00Z"/><w:dateFormat w:val="yyyy-MM-dd"/></w:date>');
  });

  it("writes content control placeholders", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "Click or tap here",
            contentControl: {
              alias: "Recipient",
              tag: "recipient",
              placeholder: { docPart: "DefaultPlaceholder_22610170" },
            },
          },
        ],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<w:placeholder><w:docPart w:val="DefaultPlaceholder_22610170"/></w:placeholder>');
  });

  it("writes content control data bindings", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "Ada Lovelace",
            contentControl: {
              alias: "Customer",
              tag: "customer",
              dataBinding: {
                storeItemId: "{11111111-2222-3333-4444-555555555555}",
                xpath: "/customer/name[1]",
                prefixMappings: "xmlns:crm='urn:crm'",
              },
            },
          },
        ],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<w:dataBinding w:storeItemID="{11111111-2222-3333-4444-555555555555}" w:xpath="/customer/name[1]" w:prefixMappings="xmlns:crm=&apos;urn:crm&apos;"/>');
  });

  it("writes content control appearance metadata", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "Click or tap here",
            contentControl: {
              alias: "Prompt",
              tag: "prompt",
              appearance: "tags",
              color: "2F5496",
              showingPlaceholder: true,
            },
          },
        ],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<w:appearance w:val="tags"/>');
    expect(xml).toContain('<w:color w:val="2F5496"/>');
    expect(xml).toContain("<w:showingPlcHdr/>");
  });

  it("writes custom xml parts with relationships", async () => {
    const document = {
      ...createDocumentJson([]),
      customXmlParts: [
        {
          path: "customXml/item1.xml",
          xml: "<customer><name>Ada Lovelace</name></customer>",
          properties: {
            path: "customXml/itemProps1.xml",
            storeItemId: "{11111111-2222-3333-4444-555555555555}",
          },
        },
      ],
    };

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const packageRels = await zip.file("_rels/.rels")!.async("string");
    const item = await zip.file("customXml/item1.xml")!.async("string");
    const itemRels = await zip.file("customXml/_rels/item1.xml.rels")!.async("string");
    const itemProps = await zip.file("customXml/itemProps1.xml")!.async("string");
    const contentTypes = await zip.file("[Content_Types].xml")!.async("string");

    expect(packageRels).toContain('Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml"');
    expect(packageRels).toContain('Target="customXml/item1.xml"');
    expect(item).toBe("<customer><name>Ada Lovelace</name></customer>");
    expect(itemRels).toContain('Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps"');
    expect(itemRels).toContain('Target="itemProps1.xml"');
    expect(itemProps).toContain('<ds:datastoreItem ds:itemID="{11111111-2222-3333-4444-555555555555}"');
    expect(contentTypes).toContain('PartName="/customXml/itemProps1.xml"');
    expect(contentTypes).toContain('ContentType="application/vnd.openxmlformats-officedocument.customXmlProperties+xml"');
  });

  it("writes custom xml schema refs", async () => {
    const document = {
      ...createDocumentJson([]),
      customXmlParts: [
        {
          path: "customXml/item1.xml",
          xml: "<customer><name>Ada Lovelace</name></customer>",
          properties: {
            path: "customXml/itemProps1.xml",
            storeItemId: "{11111111-2222-3333-4444-555555555555}",
            schemaRefs: ["urn:customer", "urn:crm"],
          },
        },
      ],
    };

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const itemProps = await zip.file("customXml/itemProps1.xml")!.async("string");

    expect(itemProps).toContain('<ds:schemaRef ds:uri="urn:customer"/>');
    expect(itemProps).toContain('<ds:schemaRef ds:uri="urn:crm"/>');
  });

  it("writes document core properties", async () => {
    const document = {
      ...createDocumentJson([]),
      properties: {
        core: {
          title: "Quarterly Report",
          subject: "Sales",
          creator: "Ada Lovelace",
          keywords: "sales,quarterly",
          description: "Executive summary",
          lastModifiedBy: "Grace Hopper",
          created: "2026-07-04T00:00:00Z",
          modified: "2026-07-04T01:00:00Z",
        },
      },
    };

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const core = await zip.file("docProps/core.xml")!.async("string");
    const packageRels = await zip.file("_rels/.rels")!.async("string");
    const contentTypes = await zip.file("[Content_Types].xml")!.async("string");

    expect(core).toContain("<dc:title>Quarterly Report</dc:title>");
    expect(core).toContain("<dc:creator>Ada Lovelace</dc:creator>");
    expect(core).toContain("<cp:lastModifiedBy>Grace Hopper</cp:lastModifiedBy>");
    expect(core).toContain('<dcterms:created xsi:type="dcterms:W3CDTF">2026-07-04T00:00:00Z</dcterms:created>');
    expect(packageRels).toContain('Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"');
    expect(packageRels).toContain('Target="docProps/core.xml"');
    expect(contentTypes).toContain('PartName="/docProps/core.xml"');
  });

  it("writes document app properties", async () => {
    const document = {
      ...createDocumentJson([]),
      properties: {
        app: {
          application: "word2json",
          company: "ACME",
          manager: "Mira",
          pages: 3,
          words: 1200,
          characters: 6400,
          lines: 80,
          paragraphs: 12,
        },
      },
    };

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const app = await zip.file("docProps/app.xml")!.async("string");
    const packageRels = await zip.file("_rels/.rels")!.async("string");
    const contentTypes = await zip.file("[Content_Types].xml")!.async("string");

    expect(app).toContain("<Application>word2json</Application>");
    expect(app).toContain("<Company>ACME</Company>");
    expect(app).toContain("<Manager>Mira</Manager>");
    expect(app).toContain("<Pages>3</Pages>");
    expect(packageRels).toContain('Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties"');
    expect(packageRels).toContain('Target="docProps/app.xml"');
    expect(contentTypes).toContain('PartName="/docProps/app.xml"');
  });

  it("writes custom document properties", async () => {
    const document = {
      ...createDocumentJson([]),
      properties: {
        custom: [
          { name: "ContractId", type: "string", value: "C-2026-001" },
          { name: "RiskScore", type: "number", value: 42 },
          { name: "Approved", type: "boolean", value: true },
          { name: "EffectiveDate", type: "date", value: "2026-07-04T00:00:00Z" },
        ],
      },
    };

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const custom = await zip.file("docProps/custom.xml")!.async("string");
    const packageRels = await zip.file("_rels/.rels")!.async("string");
    const contentTypes = await zip.file("[Content_Types].xml")!.async("string");

    expect(custom).toContain('name="ContractId"');
    expect(custom).toContain("<vt:lpwstr>C-2026-001</vt:lpwstr>");
    expect(custom).toContain('name="RiskScore"');
    expect(custom).toContain("<vt:i4>42</vt:i4>");
    expect(custom).toContain('name="Approved"');
    expect(custom).toContain("<vt:bool>true</vt:bool>");
    expect(custom).toContain('name="EffectiveDate"');
    expect(custom).toContain("<vt:filetime>2026-07-04T00:00:00Z</vt:filetime>");
    expect(packageRels).toContain('Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties"');
    expect(packageRels).toContain('Target="docProps/custom.xml"');
    expect(contentTypes).toContain('PartName="/docProps/custom.xml"');
  });

  it("writes web settings", async () => {
    const document = {
      ...createDocumentJson([]),
      settings: {
        web: {
          optimizeForBrowser: true,
          allowPng: true,
          doNotSaveAsSingleFile: true,
          pixelsPerInch: 120,
        },
      },
    };

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const webSettings = await zip.file("word/webSettings.xml")!.async("string");
    const rels = await zip.file("word/_rels/document.xml.rels")!.async("string");
    const contentTypes = await zip.file("[Content_Types].xml")!.async("string");

    expect(webSettings).toContain("<w:optimizeForBrowser/>");
    expect(webSettings).toContain("<w:allowPNG/>");
    expect(webSettings).toContain("<w:doNotSaveAsSingleFile/>");
    expect(webSettings).toContain('<w:pixelsPerInch w:val="120"/>');
    expect(rels).toContain('Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings"');
    expect(rels).toContain('Target="webSettings.xml"');
    expect(contentTypes).toContain('PartName="/word/webSettings.xml"');
  });

  it("writes font table", async () => {
    const document = {
      ...createDocumentJson([]),
      fonts: [
        { name: "Aptos", family: "swiss", pitch: "variable", charset: "00", panose1: "020F0502020204030204" },
        { name: "SimSun", family: "roman", pitch: "fixed", charset: "86" },
      ],
    };

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const fontTable = await zip.file("word/fontTable.xml")!.async("string");
    const rels = await zip.file("word/_rels/document.xml.rels")!.async("string");
    const contentTypes = await zip.file("[Content_Types].xml")!.async("string");

    expect(fontTable).toContain('<w:font w:name="Aptos">');
    expect(fontTable).toContain('<w:panose1 w:val="020F0502020204030204"/>');
    expect(fontTable).toContain('<w:charset w:val="00"/>');
    expect(fontTable).toContain('<w:family w:val="swiss"/>');
    expect(fontTable).toContain('<w:pitch w:val="variable"/>');
    expect(fontTable).toContain('<w:font w:name="SimSun">');
    expect(rels).toContain('Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable"');
    expect(rels).toContain('Target="fontTable.xml"');
    expect(contentTypes).toContain('PartName="/word/fontTable.xml"');
  });

  it("writes repeating section content controls", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        contentControl: {
          alias: "Line Items",
          tag: "lineItems",
          repeatingSection: { sectionTitle: "Item", doNotAllowInsertDeleteSection: true },
          repeatingSectionItem: { id: "{11111111-2222-3333-4444-555555555555}" },
        },
        runs: [{ text: "Widget" }],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<w:repeatingSection><w:sectionTitle w:val="Item"/><w:doNotAllowInsertDeleteSection/></w:repeatingSection>');
    expect(xml).toContain('<w:repeatingSectionItem><w:id w:val="{11111111-2222-3333-4444-555555555555}"/></w:repeatingSectionItem>');
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

  it("writes document settings", async () => {
    const document = {
      version: "1.0" as const,
      settings: { defaultTabStop: 720, evenAndOddHeaders: true, updateFields: true },
      sections: [{ blocks: [{ type: "paragraph" as const, runs: [{ text: "Settings" }] }] }],
    };

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const settings = await zip.file("word/settings.xml")!.async("string");
    const contentTypes = await zip.file("[Content_Types].xml")!.async("string");
    const rels = await zip.file("word/_rels/document.xml.rels")!.async("string");

    expect(settings).toContain('<w:defaultTabStop w:val="720"/>');
    expect(settings).toContain("<w:evenAndOddHeaders/>");
    expect(settings).toContain("<w:updateFields/>");
    expect(contentTypes).toContain("/word/settings.xml");
    expect(rels).toContain('Target="settings.xml"');
  });

  it("writes track revisions setting", async () => {
    const document = {
      version: "1.0" as const,
      settings: { trackRevisions: true },
      sections: [{ blocks: [{ type: "paragraph" as const, runs: [{ text: "Tracked" }] }] }],
    };

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const settings = await zip.file("word/settings.xml")!.async("string");

    expect(settings).toContain("<w:trackRevisions/>");
  });

  it("writes compatibility settings", async () => {
    const document = {
      version: "1.0" as const,
      settings: {
        compatibility: {
          compatMode: "15",
          settings: [
            { name: "overrideTableStyleFontSizeAndJustification", uri: "http://schemas.microsoft.com/office/word", value: "1" },
            { name: "useWord2013TrackBottomHyphenation", uri: "http://schemas.microsoft.com/office/word", value: "0" },
          ],
        },
      },
      sections: [{ blocks: [{ type: "paragraph" as const, runs: [{ text: "Compat" }] }] }],
    };

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const settings = await zip.file("word/settings.xml")!.async("string");

    expect(settings).toContain("<w:compat>");
    expect(settings).toContain('<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/>');
    expect(settings).toContain('<w:compatSetting w:name="overrideTableStyleFontSizeAndJustification" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/>');
    expect(settings).toContain('<w:compatSetting w:name="useWord2013TrackBottomHyphenation" w:uri="http://schemas.microsoft.com/office/word" w:val="0"/>');
    expect(settings).toContain("</w:compat>");
  });

  it("writes proofing settings", async () => {
    const document = {
      version: "1.0" as const,
      settings: {
        proofing: {
          spelling: "clean" as const,
          grammar: "dirty" as const,
          doNotHyphenateCaps: true,
          hyphenationZone: 360,
        },
      },
      sections: [{ blocks: [{ type: "paragraph" as const, runs: [{ text: "Proofing" }] }] }],
    };

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const settings = await zip.file("word/settings.xml")!.async("string");

    expect(settings).toContain('<w:proofState w:spelling="clean" w:grammar="dirty"/>');
    expect(settings).toContain("<w:doNotHyphenateCaps/>");
    expect(settings).toContain('<w:hyphenationZone w:val="360"/>');
  });

  it("writes view settings", async () => {
    const document = {
      version: "1.0" as const,
      settings: {
        view: {
          mode: "print" as const,
          zoom: { preset: "fullPage" as const, percent: 125 },
        },
      },
      sections: [{ blocks: [{ type: "paragraph" as const, runs: [{ text: "View" }] }] }],
    };

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const settings = await zip.file("word/settings.xml")!.async("string");

    expect(settings).toContain('<w:view w:val="print"/>');
    expect(settings).toContain('<w:zoom w:val="fullPage" w:percent="125"/>');
  });

  it("writes document protection settings", async () => {
    const document = {
      version: "1.0" as const,
      settings: {
        protection: {
          edit: "trackedChanges" as const,
          enforcement: true,
          cryptProviderType: "rsaFull",
          cryptAlgorithmClass: "hash",
          cryptAlgorithmType: "typeAny",
          cryptAlgorithmSid: 4,
          cryptSpinCount: 100000,
          hash: "ABCDEF0123456789",
          salt: "0123456789ABCDEF",
        },
      },
      sections: [{ blocks: [{ type: "paragraph" as const, runs: [{ text: "Protected" }] }] }],
    };

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const settings = await zip.file("word/settings.xml")!.async("string");

    expect(settings).toContain('<w:documentProtection w:edit="trackedChanges" w:enforcement="1" w:cryptProviderType="rsaFull" w:cryptAlgorithmClass="hash" w:cryptAlgorithmType="typeAny" w:cryptAlgorithmSid="4" w:cryptSpinCount="100000" w:hash="ABCDEF0123456789" w:salt="0123456789ABCDEF"/>');
  });

  it("writes mail merge settings", async () => {
    const document = {
      version: "1.0" as const,
      settings: {
        mailMerge: {
          mainDocumentType: "formLetters",
          dataType: "native",
          connectString: "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=contacts.xlsx;",
          query: "SELECT * FROM `Contacts$`",
          viewMergedData: true,
          activeRecord: 3,
          checkErrors: 1,
        },
      },
      sections: [{ blocks: [{ type: "paragraph" as const, runs: [{ text: "Mail merge" }] }] }],
    };

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const settings = await zip.file("word/settings.xml")!.async("string");

    expect(settings).toContain("<w:mailMerge>");
    expect(settings).toContain('<w:mainDocumentType w:val="formLetters"/>');
    expect(settings).toContain('<w:dataType w:val="native"/>');
    expect(settings).toContain('<w:connectString w:val="Provider=Microsoft.ACE.OLEDB.12.0;Data Source=contacts.xlsx;"/>');
    expect(settings).toContain('<w:query w:val="SELECT * FROM `Contacts$`"/>');
    expect(settings).toContain("<w:viewMergedData/>");
    expect(settings).toContain('<w:activeRecord w:val="3"/>');
    expect(settings).toContain('<w:checkErrors w:val="1"/>');
    expect(settings).toContain("</w:mailMerge>");
  });

  it("writes write protection settings", async () => {
    const document = {
      version: "1.0" as const,
      settings: {
        writeProtection: {
          recommended: true,
          cryptProviderType: "rsaFull",
          cryptAlgorithmClass: "hash",
          cryptAlgorithmType: "typeAny",
          cryptAlgorithmSid: 4,
          cryptSpinCount: 100000,
          hash: "FEDCBA9876543210",
          salt: "0011223344556677",
        },
      },
      sections: [{ blocks: [{ type: "paragraph" as const, runs: [{ text: "Write protected" }] }] }],
    };

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const settings = await zip.file("word/settings.xml")!.async("string");

    expect(settings).toContain('<w:writeProtection w:recommended="1" w:cryptProviderType="rsaFull" w:cryptAlgorithmClass="hash" w:cryptAlgorithmType="typeAny" w:cryptAlgorithmSid="4" w:cryptSpinCount="100000" w:hash="FEDCBA9876543210" w:salt="0011223344556677"/>');
  });

  it("writes math settings", async () => {
    const document = {
      version: "1.0" as const,
      settings: {
        math: {
          mathFont: "Cambria Math",
          breakBinary: "before" as const,
          smallFraction: true,
          displayDefaults: true,
        },
      },
      sections: [{ blocks: [{ type: "paragraph" as const, runs: [{ text: "Math settings" }] }] }],
    };

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const settings = await zip.file("word/settings.xml")!.async("string");

    expect(settings).toContain('xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"');
    expect(settings).toContain("<m:mathPr>");
    expect(settings).toContain('<m:mathFont m:val="Cambria Math"/>');
    expect(settings).toContain('<m:brkBin m:val="before"/>');
    expect(settings).toContain('<m:smallFrac m:val="1"/>');
    expect(settings).toContain("<m:dispDef/>");
    expect(settings).toContain("</m:mathPr>");
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

  it("writes detailed table borders", async () => {
    const document = createDocumentJson([
      {
        type: "table",
        borders: {
          top: { style: "double", size: 12, color: "4472C4", space: 2 },
          left: { style: "single", size: 8, color: "70AD47", space: 0 },
          bottom: { style: "dashed", size: 6, color: "C00000", space: 1 },
          right: { style: "dotted", size: 4, color: "7030A0", space: 0 },
          insideH: { style: "single", size: 4, color: "808080", space: 0 },
          insideV: { style: "nil" },
        },
        rows: [
          { cells: [{ blocks: [{ type: "paragraph", runs: [{ text: "Detailed borders" }] }] }] },
        ],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<w:tblBorders><w:top w:val="double" w:sz="12" w:space="2" w:color="4472C4"/><w:left w:val="single" w:sz="8" w:space="0" w:color="70AD47"/><w:bottom w:val="dashed" w:sz="6" w:space="1" w:color="C00000"/><w:right w:val="dotted" w:sz="4" w:space="0" w:color="7030A0"/><w:insideH w:val="single" w:sz="4" w:space="0" w:color="808080"/><w:insideV w:val="nil"/></w:tblBorders>');
  });

  it("writes table property change metadata", async () => {
    const document = createDocumentJson([
      {
        type: "table",
        width: 7200,
        propertyRevision: { id: 20, author: "Ada", date: "2026-07-04T05:00:00.000Z" },
        rows: [
          { cells: [{ blocks: [{ type: "paragraph", runs: [{ text: "Table changed" }] }] }] },
        ],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<w:tblW w:w="7200" w:type="dxa"/>');
    expect(xml).toContain('<w:tblPrChange w:id="20" w:author="Ada" w:date="2026-07-04T05:00:00.000Z"><w:tblPr/></w:tblPrChange>');
  });

  it("writes table row revisions", async () => {
    const document = createDocumentJson([
      {
        type: "table",
        rows: [
          {
            revision: { type: "insert", id: 21, author: "Lin", date: "2026-07-04T06:00:00.000Z" },
            cells: [{ blocks: [{ type: "paragraph", runs: [{ text: "Inserted row" }] }] }],
          },
          {
            revision: { type: "delete", id: 22, author: "Mira", date: "2026-07-04T07:00:00.000Z" },
            cells: [{ blocks: [{ type: "paragraph", runs: [{ text: "Deleted row" }] }] }],
          },
        ],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<w:trPr><w:ins w:id="21" w:author="Lin" w:date="2026-07-04T06:00:00.000Z"/></w:trPr>');
    expect(xml).toContain('<w:trPr><w:del w:id="22" w:author="Mira" w:date="2026-07-04T07:00:00.000Z"/></w:trPr>');
  });

  it("writes cell property change metadata", async () => {
    const document = createDocumentJson([
      {
        type: "table",
        rows: [
          {
            cells: [
              {
                shading: { fill: "D9EAF7" },
                propertyRevision: { id: 23, author: "Noor", date: "2026-07-04T08:00:00.000Z" },
                blocks: [{ type: "paragraph", runs: [{ text: "Cell changed" }] }],
              },
            ],
          },
        ],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<w:shd w:val="clear" w:fill="D9EAF7"/>');
    expect(xml).toContain('<w:tcPrChange w:id="23" w:author="Noor" w:date="2026-07-04T08:00:00.000Z"><w:tcPr/></w:tcPrChange>');
  });

  it("writes table alignment and cell spacing", async () => {
    const document = createDocumentJson([
      {
        type: "table",
        alignment: "center",
        cellSpacing: 120,
        rows: [
          { cells: [{ blocks: [{ type: "paragraph", runs: [{ text: "Centered" }] }] }] },
        ],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<w:jc w:val="center"/>');
    expect(xml).toContain('<w:tblCellSpacing w:w="120" w:type="dxa"/>');
  });

  it("writes table indentation", async () => {
    const document = createDocumentJson([
      {
        type: "table",
        indent: { width: 720 },
        rows: [
          { cells: [{ blocks: [{ type: "paragraph", runs: [{ text: "Indented" }] }] }] },
        ],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<w:tblInd w:w="720" w:type="dxa"/>');
  });

  it("writes floating table positioning", async () => {
    const document = createDocumentJson([
      {
        type: "table",
        position: {
          horizontalAnchor: "margin",
          verticalAnchor: "page",
          x: 720,
          y: 1440,
          leftFromText: 180,
          rightFromText: 180,
          topFromText: 120,
          bottomFromText: 120,
        },
        rows: [
          { cells: [{ blocks: [{ type: "paragraph", runs: [{ text: "Floating" }] }] }] },
        ],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<w:tblpPr w:leftFromText="180" w:rightFromText="180" w:topFromText="120" w:bottomFromText="120" w:horzAnchor="margin" w:vertAnchor="page" w:tblpX="720" w:tblpY="1440"/>');
  });

  it("writes floating table overlap control", async () => {
    const document = createDocumentJson([
      {
        type: "table",
        overlap: "never",
        rows: [
          { cells: [{ blocks: [{ type: "paragraph", runs: [{ text: "No overlap" }] }] }] },
        ],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<w:tblOverlap w:val="never"/>');
  });

  it("writes table caption and description", async () => {
    const document = createDocumentJson([
      {
        type: "table",
        caption: "Quarterly revenue",
        description: "Revenue by region and quarter",
        rows: [
          { cells: [{ blocks: [{ type: "paragraph", runs: [{ text: "Revenue" }] }] }] },
        ],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<w:tblCaption w:val="Quarterly revenue"/>');
    expect(xml).toContain('<w:tblDescription w:val="Revenue by region and quarter"/>');
  });

  it("writes percentage table width", async () => {
    const document = createDocumentJson([
      {
        type: "table",
        width: 5000,
        widthType: "pct",
        rows: [
          { cells: [{ blocks: [{ type: "paragraph", runs: [{ text: "Full width" }] }] }] },
        ],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<w:tblW w:w="5000" w:type="pct"/>');
  });

  it("writes table style look flags", async () => {
    const document = createDocumentJson([
      {
        type: "table",
        styleId: "TableGrid",
        look: {
          firstRow: true,
          lastRow: false,
          firstColumn: true,
          lastColumn: false,
          bandedRows: true,
          bandedColumns: false,
        },
        rows: [
          { cells: [{ blocks: [{ type: "paragraph", runs: [{ text: "Styled" }] }] }] },
        ],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<w:tblLook w:val="04A0"/>');
  });

  it("writes fixed table layout", async () => {
    const document = createDocumentJson([
      {
        type: "table",
        layout: "fixed",
        rows: [
          { cells: [{ blocks: [{ type: "paragraph", runs: [{ text: "Fixed width" }] }] }] },
        ],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<w:tblLayout w:type="fixed"/>');
  });

  it("writes table grid and row height", async () => {
    const document = createDocumentJson([
      {
        type: "table",
        grid: [2400, 3600],
        rows: [
          {
            height: { value: 480, rule: "exact" as const },
            cells: [
              { blocks: [{ type: "paragraph", runs: [{ text: "A" }] }] },
              { blocks: [{ type: "paragraph", runs: [{ text: "B" }] }] },
            ],
          },
        ],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="3600"/></w:tblGrid>');
    expect(xml).toContain('<w:trPr><w:trHeight w:val="480" w:hRule="exact"/></w:trPr>');
  });

  it("writes row table property exceptions", async () => {
    const document = createDocumentJson([
      {
        type: "table",
        rows: [
          {
            propertyExceptions: {
              width: 5000,
              widthType: "pct",
              cellSpacing: 120,
              indent: { width: 360 },
              layout: "fixed",
              look: {
                firstRow: true,
                lastRow: false,
                firstColumn: true,
                lastColumn: false,
                bandedRows: true,
                bandedColumns: false,
              },
            },
            cells: [{ blocks: [{ type: "paragraph", runs: [{ text: "Exception row" }] }] }],
          },
        ],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<w:trPr><w:tblPrEx><w:tblW w:w="5000" w:type="pct"/><w:tblCellSpacing w:w="120" w:type="dxa"/><w:tblInd w:w="360" w:type="dxa"/><w:tblLayout w:type="fixed"/><w:tblLook w:val="04A0"/></w:tblPrEx></w:trPr>');
  });

  it("writes repeating table header rows", async () => {
    const document = createDocumentJson([
      {
        type: "table",
        rows: [
          {
            repeatHeader: true,
            cells: [{ blocks: [{ type: "paragraph", runs: [{ text: "Header" }] }] }],
          },
          {
            cells: [{ blocks: [{ type: "paragraph", runs: [{ text: "Body" }] }] }],
          },
        ],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain("<w:trPr><w:tblHeader/></w:trPr>");
  });

  it("writes table rows that cannot split across pages", async () => {
    const document = createDocumentJson([
      {
        type: "table",
        rows: [
          {
            cantSplit: true,
            cells: [{ blocks: [{ type: "paragraph", runs: [{ text: "Keep together" }] }] }],
          },
          {
            cells: [{ blocks: [{ type: "paragraph", runs: [{ text: "Next row" }] }] }],
          },
        ],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain("<w:trPr><w:cantSplit/></w:trPr>");
  });

  it("writes cell vertical merge and alignment", async () => {
    const document = createDocumentJson([
      {
        type: "table",
        rows: [
          {
            cells: [
              {
                verticalMerge: "restart" as const,
                verticalAlignment: "center" as const,
                blocks: [{ type: "paragraph", runs: [{ text: "Merged" }] }],
              },
              { blocks: [{ type: "paragraph", runs: [{ text: "Top" }] }] },
            ],
          },
          {
            cells: [
              {
                verticalMerge: "continue" as const,
                blocks: [{ type: "paragraph", runs: [] }],
              },
              { blocks: [{ type: "paragraph", runs: [{ text: "Bottom" }] }] },
            ],
          },
        ],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<w:vMerge w:val="restart"/>');
    expect(xml).toContain('<w:vAlign w:val="center"/>');
    expect(xml).toContain('<w:vMerge w:val="continue"/>');
  });

  it("writes percentage cell width", async () => {
    const document = createDocumentJson([
      {
        type: "table",
        rows: [
          {
            cells: [
              {
                width: 2500,
                widthType: "pct",
                blocks: [{ type: "paragraph", runs: [{ text: "Half" }] }],
              },
            ],
          },
        ],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<w:tcW w:w="2500" w:type="pct"/>');
  });

  it("writes cell shading and margins", async () => {
    const document = createDocumentJson([
      {
        type: "table",
        rows: [
          {
            cells: [
              {
                shading: { fill: "D9EAF7" },
                margins: { top: 120, right: 180, bottom: 120, left: 180 },
                blocks: [{ type: "paragraph", runs: [{ text: "Header" }] }],
              },
            ],
          },
        ],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<w:shd w:val="clear" w:fill="D9EAF7"/>');
    expect(xml).toContain('<w:tcMar><w:top w:w="120" w:type="dxa"/><w:right w:w="180" w:type="dxa"/><w:bottom w:w="120" w:type="dxa"/><w:left w:w="180" w:type="dxa"/></w:tcMar>');
  });

  it("writes cell margin width types", async () => {
    const document = createDocumentJson([
      {
        type: "table",
        rows: [
          {
            cells: [
              {
                margins: {
                  top: { width: 240, type: "pct" },
                  right: 180,
                  bottom: { width: 0, type: "nil" },
                  left: { width: 120 },
                },
                blocks: [{ type: "paragraph", runs: [{ text: "Typed margins" }] }],
              },
            ],
          },
        ],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<w:tcMar><w:top w:w="240" w:type="pct"/><w:right w:w="180" w:type="dxa"/><w:bottom w:w="0" w:type="nil"/><w:left w:w="120" w:type="dxa"/></w:tcMar>');
  });

  it("writes cell no-wrap and fit-text controls", async () => {
    const document = createDocumentJson([
      {
        type: "table",
        rows: [
          {
            cells: [
              {
                noWrap: true,
                fitText: true,
                blocks: [{ type: "paragraph", runs: [{ text: "Compact code" }] }],
              },
            ],
          },
        ],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain("<w:noWrap/>");
    expect(xml).toContain("<w:tcFitText/>");
  });

  it("writes cell borders", async () => {
    const document = createDocumentJson([
      {
        type: "table",
        rows: [
          {
            cells: [
              {
                borders: {
                  top: { style: "single", size: 8, color: "4472C4", space: 0 },
                  bottom: { style: "single", size: 8, color: "4472C4", space: 0 },
                },
                blocks: [{ type: "paragraph", runs: [{ text: "Bordered cell" }] }],
              },
            ],
          },
        ],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain("<w:tcBorders>");
    expect(xml).toContain('<w:top w:val="single" w:sz="8" w:space="0" w:color="4472C4"/>');
    expect(xml).toContain('<w:bottom w:val="single" w:sz="8" w:space="0" w:color="4472C4"/>');
  });

  it("writes detailed cell borders", async () => {
    const document = createDocumentJson([
      {
        type: "table",
        rows: [
          {
            cells: [
              {
                borders: {
                  top: { style: "double", size: 12, color: "4472C4", space: 2 },
                  left: { style: "dashed", size: 6, color: "70AD47", space: 1 },
                  bottom: { style: "dotted", size: 4, color: "C00000", space: 0 },
                  right: { style: "single", size: 8, color: "7030A0", space: 0 },
                  insideH: { style: "single", size: 4, color: "808080", space: 0 },
                  insideV: { style: "nil" },
                },
                blocks: [{ type: "paragraph", runs: [{ text: "Detailed cell borders" }] }],
              },
            ],
          },
        ],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<w:tcBorders><w:top w:val="double" w:sz="12" w:space="2" w:color="4472C4"/><w:left w:val="dashed" w:sz="6" w:space="1" w:color="70AD47"/><w:bottom w:val="dotted" w:sz="4" w:space="0" w:color="C00000"/><w:right w:val="single" w:sz="8" w:space="0" w:color="7030A0"/><w:insideH w:val="single" w:sz="4" w:space="0" w:color="808080"/><w:insideV w:val="nil"/></w:tcBorders>');
  });

  it("writes cell text direction", async () => {
    const document = createDocumentJson([
      {
        type: "table",
        rows: [
          {
            cells: [
              {
                textDirection: "btLr",
                blocks: [{ type: "paragraph", runs: [{ text: "Vertical" }] }],
              },
            ],
          },
        ],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<w:textDirection w:val="btLr"/>');
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

  it("writes custom numbering definitions", async () => {
    const document = {
      version: "1.0" as const,
      numbering: {
        abstractNums: [
          {
            id: 10,
            levels: [
              { level: 0, format: "decimal" as const, text: "%1.", start: 1, left: 720, hanging: 360 },
              { level: 1, format: "lowerLetter" as const, text: "%2)", start: 1, left: 1440, hanging: 360 },
            ],
          },
        ],
        nums: [{ id: 10, abstractId: 10 }],
      },
      sections: [
        {
          blocks: [
            {
              type: "paragraph" as const,
              list: { type: "ordered" as const, level: 1, numberingId: 10 },
              runs: [{ text: "Nested clause" }],
            },
          ],
        },
      ],
    };

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");
    const numbering = await zip.file("word/numbering.xml")!.async("string");

    expect(xml).toContain('<w:numPr><w:ilvl w:val="1"/><w:numId w:val="10"/></w:numPr>');
    expect(numbering).toContain('<w:abstractNum w:abstractNumId="10">');
    expect(numbering).toContain('<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl>');
    expect(numbering).toContain('<w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="%2)"/><w:pPr><w:ind w:left="1440" w:hanging="360"/></w:pPr></w:lvl>');
    expect(numbering).toContain('<w:num w:numId="10"><w:abstractNumId w:val="10"/></w:num>');
  });

  it("writes numbering level suffix restart and legal controls", async () => {
    const document = {
      version: "1.0" as const,
      numbering: {
        abstractNums: [
          {
            id: 20,
            levels: [
              {
                level: 0,
                format: "decimal" as const,
                text: "%1)",
                start: 3,
                suffix: "space" as const,
                restart: 2,
                legal: true,
                left: 720,
                hanging: 360,
              },
              {
                level: 1,
                format: "lowerLetter" as const,
                text: "%2.",
                start: 1,
                legal: false,
              },
            ],
          },
        ],
        nums: [{ id: 20, abstractId: 20 }],
      },
      sections: [{ blocks: [] }],
    };

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const numbering = await zip.file("word/numbering.xml")!.async("string");

    expect(numbering).toContain('<w:lvl w:ilvl="0"><w:start w:val="3"/><w:numFmt w:val="decimal"/><w:lvlRestart w:val="2"/><w:isLgl/><w:suff w:val="space"/><w:lvlText w:val="%1)"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl>');
    expect(numbering).toContain('<w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="lowerLetter"/><w:isLgl w:val="0"/><w:lvlText w:val="%2."/></w:lvl>');
  });

  it("writes numbering level style alignment and run properties", async () => {
    const document = {
      version: "1.0" as const,
      numbering: {
        abstractNums: [
          {
            id: 30,
            levels: [
              {
                level: 0,
                format: "decimal" as const,
                text: "%1.",
                start: 1,
                styleId: "LegalClause",
                alignment: "right" as const,
                left: 720,
                hanging: 360,
                run: {
                  bold: true,
                  italic: false,
                  fontFamily: "Aptos",
                  fontSize: 11,
                  color: "C00000",
                },
              },
            ],
          },
        ],
        nums: [{ id: 30, abstractId: 30 }],
      },
      sections: [{ blocks: [] }],
    };

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const numbering = await zip.file("word/numbering.xml")!.async("string");

    expect(numbering).toContain('<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:pStyle w:val="LegalClause"/><w:lvlText w:val="%1."/><w:lvlJc w:val="right"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr><w:rPr><w:b/><w:i w:val="0"/><w:rFonts w:ascii="Aptos" w:hAnsi="Aptos"/><w:sz w:val="22"/><w:color w:val="C00000"/></w:rPr></w:lvl>');
  });

  it("writes abstract numbering identity and style links", async () => {
    const document = {
      version: "1.0" as const,
      numbering: {
        abstractNums: [
          {
            id: 40,
            nsid: "5E2A1C9B",
            multiLevelType: "hybridMultilevel" as const,
            templateCode: "03A54D6C",
            styleLink: "LegalList",
            numberingStyleLink: "LegalListNumbering",
            levels: [
              { level: 0, format: "decimal" as const, text: "%1.", start: 1 },
            ],
          },
        ],
        nums: [{ id: 40, abstractId: 40 }],
      },
      sections: [{ blocks: [] }],
    };

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const numbering = await zip.file("word/numbering.xml")!.async("string");

    expect(numbering).toContain('<w:abstractNum w:abstractNumId="40"><w:nsid w:val="5E2A1C9B"/><w:multiLevelType w:val="hybridMultilevel"/><w:tmpl w:val="03A54D6C"/><w:styleLink w:val="LegalList"/><w:numStyleLink w:val="LegalListNumbering"/><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum>');
  });

  it("writes numbering instance level overrides", async () => {
    const document = {
      version: "1.0" as const,
      numbering: {
        abstractNums: [
          {
            id: 50,
            levels: [
              { level: 0, format: "decimal" as const, text: "%1.", start: 1 },
              { level: 1, format: "lowerLetter" as const, text: "%2)", start: 1 },
            ],
          },
        ],
        nums: [
          {
            id: 50,
            abstractId: 50,
            overrides: [
              { level: 0, start: 7 },
              {
                level: 1,
                definition: {
                  level: 1,
                  format: "upperRoman" as const,
                  text: "%2.",
                  start: 3,
                  left: 1440,
                  hanging: 360,
                },
              },
            ],
          },
        ],
      },
      sections: [{ blocks: [] }],
    };

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const numbering = await zip.file("word/numbering.xml")!.async("string");

    expect(numbering).toContain('<w:num w:numId="50"><w:abstractNumId w:val="50"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="7"/></w:lvlOverride><w:lvlOverride w:ilvl="1"><w:lvl w:ilvl="1"><w:start w:val="3"/><w:numFmt w:val="upperRoman"/><w:lvlText w:val="%2."/><w:pPr><w:ind w:left="1440" w:hanging="360"/></w:pPr></w:lvl></w:lvlOverride></w:num>');
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

  it("writes internal hyperlinks", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          { text: "Jump", link: { anchor: "Clause1" } },
        ],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");
    const rels = await zip.file("word/_rels/document.xml.rels")!.async("string");

    expect(xml).toContain('<w:hyperlink w:anchor="Clause1">');
    expect(xml).not.toContain('Target="Clause1"');
    expect(rels).not.toContain("rIdHyperlink1");
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

  it("writes comment ids", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "Reviewed",
            comment: {
              id: 42,
              author: "Ada",
              initials: "AL",
              date: "2026-07-04T13:00:00.000Z",
              text: "Stable comment id.",
            },
          },
        ],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");
    const comments = await zip.file("word/comments.xml")!.async("string");

    expect(xml).toContain('<w:commentRangeStart w:id="42"/>');
    expect(xml).toContain('<w:commentRangeEnd w:id="42"/>');
    expect(xml).toContain('<w:commentReference w:id="42"/>');
    expect(comments).toContain('<w:comment w:id="42" w:author="Ada" w:initials="AL" w:date="2026-07-04T13:00:00.000Z">');
  });

  it("writes multi-run comments as one range", async () => {
    const comment = { id: 50, author: "Ada", text: "One range." };
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          { text: "First ", comment },
          { text: "second", bold: true, comment },
        ],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");
    const comments = await zip.file("word/comments.xml")!.async("string");

    expect(xml.match(/<w:commentRangeStart w:id="50"\/>/g)).toHaveLength(1);
    expect(xml.match(/<w:commentRangeEnd w:id="50"\/>/g)).toHaveLength(1);
    expect(xml.match(/<w:commentReference w:id="50"\/>/g)).toHaveLength(1);
    expect(xml).toContain('<w:commentRangeStart w:id="50"/><w:r><w:t xml:space="preserve">First </w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>second</w:t></w:r><w:commentRangeEnd w:id="50"/>');
    expect(comments.match(/<w:comment w:id="50"/g)).toHaveLength(1);
  });

  it("writes cross-paragraph comments", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        commentRangeStart: { id: 60, author: "Ada", text: "Across paragraphs." },
        runs: [{ text: "First paragraph" }],
      },
      {
        type: "paragraph",
        commentRangeEnd: { id: 60 },
        runs: [{ text: "Second paragraph" }],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");
    const comments = await zip.file("word/comments.xml")!.async("string");

    expect(xml).toContain('<w:commentRangeStart w:id="60"/><w:p><w:r><w:t>First paragraph</w:t></w:r></w:p>');
    expect(xml).toContain('<w:p><w:r><w:t>Second paragraph</w:t></w:r></w:p><w:commentRangeEnd w:id="60"/><w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="60"/></w:r>');
    expect(comments.match(/<w:comment w:id="60"/g)).toHaveLength(1);
    expect(comments).toContain("<w:t>Across paragraphs.</w:t>");
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

  it("writes image crop", async () => {
    const imageData = Buffer.from("fake-png").toString("base64");
    const document = createDocumentJson([
      {
        type: "image",
        data: imageData,
        contentType: "image/png",
        width: 120,
        height: 80,
        crop: { left: 1000, top: 2000, right: 3000, bottom: 4000 },
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<a:srcRect l="1000" t="2000" r="3000" b="4000"/>');
  });

  it("writes image rotation", async () => {
    const imageData = Buffer.from("fake-png").toString("base64");
    const document = createDocumentJson([
      {
        type: "image",
        data: imageData,
        contentType: "image/png",
        width: 120,
        height: 80,
        rotation: 15,
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<a:xfrm rot="900000">');
  });

  it("writes floating image layout", async () => {
    const imageData = Buffer.from("fake-png").toString("base64");
    const document = createDocumentJson([
      {
        type: "image",
        data: imageData,
        contentType: "image/png",
        width: 120,
        height: 80,
        floating: { wrap: "square", horizontalOffset: 1440, verticalOffset: 720 },
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain("<wp:anchor ");
    expect(xml).toContain("<wp:wrapSquare/>");
    expect(xml).toContain('<wp:positionH relativeFrom="page"><wp:posOffset>1440</wp:posOffset></wp:positionH>');
    expect(xml).toContain('<wp:positionV relativeFrom="page"><wp:posOffset>720</wp:posOffset></wp:positionV>');
  });

  it("writes floating image wrap variants", async () => {
    const imageData = Buffer.from("fake-png").toString("base64");
    const document = createDocumentJson([
      {
        type: "image",
        data: imageData,
        contentType: "image/png",
        width: 120,
        height: 80,
        floating: { wrap: "topAndBottom", horizontalOffset: 1440, verticalOffset: 720 },
      },
      {
        type: "image",
        data: imageData,
        contentType: "image/png",
        width: 120,
        height: 80,
        floating: { wrap: "none", horizontalOffset: 2880, verticalOffset: 1440 },
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain("<wp:wrapTopAndBottom/>");
    expect(xml).toContain("<wp:wrapNone/>");
  });

  it("writes floating image anchor references", async () => {
    const imageData = Buffer.from("fake-png").toString("base64");
    const document = createDocumentJson([
      {
        type: "image",
        data: imageData,
        contentType: "image/png",
        width: 120,
        height: 80,
        floating: {
          wrap: "square",
          horizontalOffset: 1440,
          verticalOffset: 720,
          horizontalRelativeFrom: "margin",
          verticalRelativeFrom: "paragraph",
        },
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<wp:positionH relativeFrom="margin"><wp:posOffset>1440</wp:posOffset></wp:positionH>');
    expect(xml).toContain('<wp:positionV relativeFrom="paragraph"><wp:posOffset>720</wp:posOffset></wp:positionV>');
  });

  it("writes floating image alignment positioning", async () => {
    const imageData = Buffer.from("fake-png").toString("base64");
    const document = createDocumentJson([
      {
        type: "image",
        data: imageData,
        contentType: "image/png",
        width: 120,
        height: 80,
        floating: {
          wrap: "square",
          horizontalOffset: 0,
          verticalOffset: 0,
          horizontalRelativeFrom: "margin",
          horizontalAlign: "center",
          verticalAlign: "bottom",
        },
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<wp:positionH relativeFrom="margin"><wp:align>center</wp:align></wp:positionH>');
    expect(xml).toContain('<wp:positionV relativeFrom="page"><wp:align>bottom</wp:align></wp:positionV>');
  });

  it("writes floating image distances and layout flags", async () => {
    const imageData = Buffer.from("fake-png").toString("base64");
    const document = createDocumentJson([
      {
        type: "image",
        data: imageData,
        contentType: "image/png",
        width: 120,
        height: 80,
        floating: {
          wrap: "square",
          horizontalOffset: 1440,
          verticalOffset: 720,
          distanceTop: 10,
          distanceBottom: 20,
          distanceLeft: 30,
          distanceRight: 40,
          behindDoc: true,
          allowOverlap: false,
          layoutInCell: false,
        },
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('behindDoc="1" locked="0" layoutInCell="0" allowOverlap="0"');
    expect(xml).toContain('distT="10" distB="20" distL="30" distR="40"');
  });

  it("writes floating image anchor controls", async () => {
    const imageData = Buffer.from("fake-png").toString("base64");
    const document = createDocumentJson([
      {
        type: "image",
        data: imageData,
        contentType: "image/png",
        width: 120,
        height: 80,
        floating: {
          wrap: "square",
          horizontalOffset: 1440,
          verticalOffset: 720,
          simplePosition: { x: 123, y: 456 },
          relativeHeight: 251659264,
          locked: true,
        },
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('simplePos="1" relativeHeight="251659264"');
    expect(xml).toContain('locked="1" layoutInCell="1" allowOverlap="1"');
    expect(xml).toContain('<wp:simplePos x="123" y="456"/>');
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

  it("writes first and even page headers and footers", async () => {
    const document = {
      version: "1.0" as const,
      sections: [
        {
          titlePage: true,
          headers: {
            default: [{ type: "paragraph" as const, runs: [{ text: "Default header" }] }],
            first: [{ type: "paragraph" as const, runs: [{ text: "First header" }] }],
            even: [{ type: "paragraph" as const, runs: [{ text: "Even header" }] }],
          },
          footers: {
            default: [{ type: "paragraph" as const, runs: [{ text: "Default footer" }] }],
            first: [{ type: "paragraph" as const, runs: [{ text: "First footer" }] }],
            even: [{ type: "paragraph" as const, runs: [{ text: "Even footer" }] }],
          },
          blocks: [{ type: "paragraph" as const, runs: [{ text: "Body" }] }],
        },
      ],
    };

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");
    const rels = await zip.file("word/_rels/document.xml.rels")!.async("string");
    const firstHeader = await zip.file("word/header2.xml")!.async("string");
    const evenFooter = await zip.file("word/footer3.xml")!.async("string");

    expect(xml).toContain('<w:headerReference w:type="default" r:id="rIdHeader1"/>');
    expect(xml).toContain('<w:headerReference w:type="first" r:id="rIdHeader2"/>');
    expect(xml).toContain('<w:headerReference w:type="even" r:id="rIdHeader3"/>');
    expect(xml).toContain('<w:footerReference w:type="default" r:id="rIdFooter1"/>');
    expect(xml).toContain('<w:footerReference w:type="first" r:id="rIdFooter2"/>');
    expect(xml).toContain('<w:footerReference w:type="even" r:id="rIdFooter3"/>');
    expect(xml).toContain("<w:titlePg/>");
    expect(xml.indexOf("<w:titlePg/>")).toBeGreaterThan(xml.indexOf("<w:pgMar"));
    expect(firstHeader).toContain("<w:t>First header</w:t>");
    expect(evenFooter).toContain("<w:t>Even footer</w:t>");
    expect(rels).toContain('Target="header2.xml"');
    expect(rels).toContain('Target="footer3.xml"');
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

  it("writes multiple sections with section break types", async () => {
    const document = {
      version: "1.0" as const,
      sections: [
        {
          breakType: "nextPage" as const,
          page: {
            width: 12240,
            height: 15840,
            margins: { top: 1440, right: 1440, bottom: 1440, left: 1440, header: 720, footer: 720, gutter: 0 },
          },
          blocks: [{ type: "paragraph" as const, runs: [{ text: "Section one" }] }],
        },
        {
          page: {
            width: 16840,
            height: 11900,
            orientation: "landscape" as const,
            margins: { top: 720, right: 900, bottom: 720, left: 900, header: 360, footer: 360, gutter: 0 },
          },
          blocks: [{ type: "paragraph" as const, runs: [{ text: "Section two" }] }],
        },
      ],
    };

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml.match(/<w:sectPr>/g)).toHaveLength(2);
    expect(xml).toContain('<w:type w:val="nextPage"/>');
    expect(xml).toContain("<w:t>Section one</w:t>");
    expect(xml).toContain("<w:t>Section two</w:t>");
    expect(xml).toContain('<w:pgSz w:w="16840" w:h="11900" w:orient="landscape"/>');
  });

  it("writes section page numbering settings", async () => {
    const document = {
      version: "1.0" as const,
      sections: [
        {
          pageNumbering: {
            start: 3,
            format: "lowerRoman" as const,
            chapterStyle: 1,
            chapterSeparator: "hyphen" as const,
          },
          blocks: [{ type: "paragraph" as const, runs: [{ text: "Preface" }] }],
        },
      ],
    };

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<w:pgNumType w:start="3" w:fmt="lowerRoman" w:chapStyle="1" w:chapSep="hyphen"/>');
  });

  it("writes section line numbering settings", async () => {
    const document = {
      version: "1.0" as const,
      sections: [
        {
          lineNumbering: {
            start: 5,
            countBy: 2,
            distance: 360,
            restart: "newPage" as const,
          },
          blocks: [{ type: "paragraph" as const, runs: [{ text: "Draft" }] }],
        },
      ],
    };

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<w:lnNumType w:start="5" w:countBy="2" w:distance="360" w:restart="newPage"/>');
  });

  it("writes section document grid settings", async () => {
    const document = {
      version: "1.0" as const,
      sections: [
        {
          documentGrid: {
            type: "linesAndChars" as const,
            linePitch: 360,
            charSpace: 180,
          },
          blocks: [{ type: "paragraph" as const, runs: [{ text: "Grid" }] }],
        },
      ],
    };

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<w:docGrid w:type="linesAndChars" w:linePitch="360" w:charSpace="180"/>');
  });

  it("writes section vertical alignment", async () => {
    const document = {
      version: "1.0" as const,
      sections: [
        {
          verticalAlignment: "center" as const,
          blocks: [{ type: "paragraph" as const, runs: [{ text: "Centered" }] }],
        },
      ],
    };

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<w:vAlign w:val="center"/>');
  });

  it("writes section text direction", async () => {
    const document = {
      version: "1.0" as const,
      sections: [
        {
          textDirection: "tbRl" as const,
          blocks: [{ type: "paragraph" as const, runs: [{ text: "Vertical" }] }],
        },
      ],
    };

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<w:textDirection w:val="tbRl"/>');
  });

  it("writes section rtl gutter", async () => {
    const document = {
      version: "1.0" as const,
      sections: [
        {
          rtlGutter: true,
          blocks: [{ type: "paragraph" as const, runs: [{ text: "RTL gutter" }] }],
        },
      ],
    };

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain("<w:rtlGutter/>");
  });

  it("writes section bidi layout", async () => {
    const document = {
      version: "1.0" as const,
      sections: [
        {
          bidi: true,
          blocks: [{ type: "paragraph" as const, runs: [{ text: "Bidi section" }] }],
        },
      ],
    };

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain("<w:bidi/>");
  });

  it("writes section no endnote", async () => {
    const document = {
      version: "1.0" as const,
      sections: [
        {
          noEndnote: true,
          blocks: [{ type: "paragraph" as const, runs: [{ text: "No endnote here" }] }],
        },
      ],
    };

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain("<w:noEndnote/>");
  });

  it("writes section mirror margins", async () => {
    const document = {
      version: "1.0" as const,
      sections: [
        {
          mirrorMargins: true,
          blocks: [{ type: "paragraph" as const, runs: [{ text: "Booklet" }] }],
        },
      ],
    };

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain("<w:mirrorMargins/>");
  });

  it("writes text watermarks into section headers", async () => {
    const document = {
      version: "1.0" as const,
      sections: [
        {
          watermark: {
            text: "IMPORT QA",
            color: "C0C0C0",
            opacity: 0.15,
            rotation: 315,
            fontFamily: "Calibri",
          },
          blocks: [{ type: "paragraph" as const, runs: [{ text: "Body" }] }],
        },
      ],
    };

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");
    const header = await zip.file("word/header1.xml")!.async("string");

    expect(xml).toContain('<w:headerReference w:type="default" r:id="rIdHeader1"/>');
    expect(header).toContain('<v:shape id="Word2JsonWatermark"');
    expect(header).toContain('fillcolor="#C0C0C0"');
    expect(header).toContain("rotation:315");
    expect(header).toContain('<v:fill opacity="0.15"/>');
    expect(header).toContain('string="IMPORT QA"');
  });

  it("writes section footnote and endnote properties", async () => {
    const document = {
      version: "1.0" as const,
      sections: [
        {
          footnoteProperties: {
            position: "beneathText" as const,
            numbering: { format: "lowerRoman" as const, start: 2, restart: "eachSect" as const },
          },
          endnoteProperties: {
            position: "sectEnd" as const,
            numbering: { format: "upperRoman" as const, start: 4, restart: "continuous" as const },
          },
          blocks: [{ type: "paragraph" as const, runs: [{ text: "Notes" }] }],
        },
      ],
    };

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<w:footnotePr><w:pos w:val="beneathText"/><w:numFmt w:val="lowerRoman"/><w:numStart w:val="2"/><w:numRestart w:val="eachSect"/></w:footnotePr>');
    expect(xml).toContain('<w:endnotePr><w:pos w:val="sectEnd"/><w:numFmt w:val="upperRoman"/><w:numStart w:val="4"/><w:numRestart w:val="continuous"/></w:endnotePr>');
  });

  it("writes section columns", async () => {
    const document = {
      version: "1.0" as const,
      sections: [
        {
          columns: { count: 2, space: 720 },
          blocks: [{ type: "paragraph" as const, runs: [{ text: "Columns" }] }],
        },
      ],
    };

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<w:cols w:num="2" w:space="720"/>');
  });

  it("writes paragraph pagination controls", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        pagination: { keepNext: true, keepLines: true, pageBreakBefore: true },
        runs: [{ text: "Controlled paragraph" }],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain("<w:keepNext/>");
    expect(xml).toContain("<w:keepLines/>");
    expect(xml).toContain("<w:pageBreakBefore/>");
  });

  it("writes explicit off paragraph pagination controls", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        pagination: { keepNext: false, keepLines: false, pageBreakBefore: false },
        runs: [{ text: "Loose paragraph" }],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<w:keepNext w:val="0"/>');
    expect(xml).toContain('<w:keepLines w:val="0"/>');
    expect(xml).toContain('<w:pageBreakBefore w:val="0"/>');
  });

  it("writes extra paragraph pagination toggles", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        pagination: { widowControl: true, suppressLineNumbers: true, suppressAutoHyphens: false },
        runs: [{ text: "Controlled typography" }],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain("<w:widowControl/>");
    expect(xml).toContain("<w:suppressLineNumbers/>");
    expect(xml).toContain('<w:suppressAutoHyphens w:val="0"/>');
  });

  it("writes paragraph layout toggles", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        pagination: { contextualSpacing: true, mirrorIndents: false, overflowPunct: true, topLinePunct: false },
        runs: [{ text: "Layout-sensitive paragraph" }],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain("<w:contextualSpacing/>");
    expect(xml).toContain('<w:mirrorIndents w:val="0"/>');
    expect(xml).toContain("<w:overflowPunct/>");
    expect(xml).toContain('<w:topLinePunct w:val="0"/>');
  });

  it("writes paragraph text flow properties", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        pagination: {
          textAlignment: "center",
          textDirection: "tbRl",
          adjustRightInd: true,
          autoSpaceDE: false,
          autoSpaceDN: true,
        },
        runs: [{ text: "Vertical paragraph" }],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<w:textAlignment w:val="center"/>');
    expect(xml).toContain('<w:textDirection w:val="tbRl"/>');
    expect(xml).toContain("<w:adjustRightInd/>");
    expect(xml).toContain('<w:autoSpaceDE w:val="0"/>');
    expect(xml).toContain("<w:autoSpaceDN/>");
  });

  it("writes paragraph frame properties", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        frame: {
          width: 2880,
          height: 1440,
          x: 720,
          y: 360,
          horizontalAnchor: "margin",
          verticalAnchor: "page",
          xAlign: "center",
          yAlign: "top",
          wrap: "around",
          dropCap: "drop",
          lines: 3,
          anchorLock: true,
          heightRule: "exact",
        },
        runs: [{ text: "Framed paragraph" }],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<w:framePr w:w="2880" w:h="1440" w:x="720" w:y="360" w:hAnchor="margin" w:vAnchor="page" w:xAlign="center" w:yAlign="top" w:wrap="around" w:dropCap="drop" w:lines="3" w:anchorLock="1" w:hRule="exact"/>');
  });

  it("writes paragraph tab stops", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        tabs: [
          { value: "left", position: 720 },
          { value: "right", position: 4320, leader: "dot" },
        ],
        runs: [{ text: "Label\tValue" }],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<w:tabs><w:tab w:val="left" w:pos="720"/><w:tab w:val="right" w:pos="4320" w:leader="dot"/></w:tabs>');
  });

  it("writes paragraph spacing", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        spacing: { before: 240, after: 120, line: 360, lineRule: "auto" as const },
        runs: [{ text: "Spaced paragraph" }],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<w:spacing w:before="240" w:after="120" w:line="360" w:lineRule="auto"/>');
  });

  it("writes paragraph shading", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        shading: { fill: "FFF2CC" },
        runs: [{ text: "Highlighted paragraph" }],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<w:shd w:val="clear" w:fill="FFF2CC"/>');
  });

  it("writes paragraph borders", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        borders: {
          top: { style: "single", size: 8, color: "4472C4", space: 2 },
          bottom: { style: "single", size: 8, color: "4472C4", space: 2 },
        },
        runs: [{ text: "Bordered paragraph" }],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain("<w:pBdr>");
    expect(xml).toContain('<w:top w:val="single" w:sz="8" w:space="2" w:color="4472C4"/>');
    expect(xml).toContain('<w:bottom w:val="single" w:sz="8" w:space="2" w:color="4472C4"/>');
  });

  it("writes paragraph between and bar borders", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        borders: {
          between: { style: "dashed", size: 6, color: "70AD47", space: 1 },
          bar: { style: "double", size: 12, color: "C00000", space: 0 },
        },
        runs: [{ text: "Paragraph side borders" }],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<w:pBdr><w:between w:val="dashed" w:sz="6" w:space="1" w:color="70AD47"/><w:bar w:val="double" w:sz="12" w:space="0" w:color="C00000"/></w:pBdr>');
  });

  it("writes paragraph indentation", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        indent: { left: 720, right: 360, firstLine: 240, hanging: 120 },
        runs: [{ text: "Indented paragraph" }],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<w:ind w:left="720" w:right="360" w:firstLine="240" w:hanging="120"/>');
  });

  it("writes footnotes with references relationships and notes part", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          { text: "Body" },
          { text: "", footnote: { blocks: [{ type: "paragraph", runs: [{ text: "Footnote text" }] }] } },
        ],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");
    const footnotes = await zip.file("word/footnotes.xml")!.async("string");
    const rels = await zip.file("word/_rels/document.xml.rels")!.async("string");
    const contentTypes = await zip.file("[Content_Types].xml")!.async("string");

    expect(xml).toContain('<w:footnoteReference w:id="1"/>');
    expect(footnotes).toContain('<w:footnote w:id="1">');
    expect(footnotes).toContain("<w:t>Footnote text</w:t>");
    expect(rels).toContain('Target="footnotes.xml"');
    expect(contentTypes).toContain('/word/footnotes.xml');
  });

  it("writes endnotes with references relationships and notes part", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          { text: "Body" },
          { text: "", endnote: { blocks: [{ type: "paragraph", runs: [{ text: "Endnote text" }] }] } },
        ],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");
    const endnotes = await zip.file("word/endnotes.xml")!.async("string");
    const rels = await zip.file("word/_rels/document.xml.rels")!.async("string");
    const contentTypes = await zip.file("[Content_Types].xml")!.async("string");

    expect(xml).toContain('<w:endnoteReference w:id="1"/>');
    expect(endnotes).toContain('<w:endnote w:id="1">');
    expect(endnotes).toContain("<w:t>Endnote text</w:t>");
    expect(rels).toContain('Target="endnotes.xml"');
    expect(contentTypes).toContain('/word/endnotes.xml');
  });

  it("writes reference fields", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          { text: "See " },
          { text: "", field: { type: "ref", target: "Clause1" } },
          { text: " on page " },
          { text: "", field: { type: "pageRef", target: "Clause1" } },
        ],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<w:instrText xml:space="preserve">REF Clause1</w:instrText>');
    expect(xml).toContain('<w:instrText xml:space="preserve">PAGEREF Clause1</w:instrText>');
  });

  it("writes field result text", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          { text: "", field: { type: "page", result: "3" } },
        ],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<w:instrText xml:space="preserve">PAGE</w:instrText>');
    expect(xml).toContain("<w:t>3</w:t>");
  });

  it("writes table of contents field", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          { text: "", field: { type: "toc", switches: 'o "1-3" h z u', result: "Table of Contents" } },
        ],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<w:instrText xml:space="preserve">TOC \\o "1-3" \\h \\z \\u</w:instrText>');
    expect(xml).toContain("<w:t>Table of Contents</w:t>");
  });

  it("writes paragraph style definitions and paragraph style ids", async () => {
    const document = {
      version: "1.0" as const,
      styles: {
        paragraph: [
          { id: "ContractTitle", name: "Contract Title", basedOn: "Normal", next: "Normal" },
        ],
      },
      sections: [
        {
          blocks: [
            { type: "paragraph" as const, styleId: "ContractTitle", runs: [{ text: "Agreement" }] },
          ],
        },
      ],
    };

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");
    const styles = await zip.file("word/styles.xml")!.async("string");

    expect(xml).toContain('<w:pStyle w:val="ContractTitle"/>');
    expect(styles).toContain('<w:style w:type="paragraph" w:styleId="ContractTitle">');
    expect(styles).toContain('<w:name w:val="Contract Title"/>');
    expect(styles).toContain('<w:basedOn w:val="Normal"/>');
    expect(styles).toContain('<w:next w:val="Normal"/>');
  });

  it("writes run document defaults", async () => {
    const document = {
      version: "1.0" as const,
      styles: {
        defaults: {
          run: { fontFamily: "Aptos", fontSize: 11, color: "1F1F1F" },
        },
      },
      sections: [{ blocks: [{ type: "paragraph" as const, runs: [{ text: "Defaults" }] }] }],
    };

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const styles = await zip.file("word/styles.xml")!.async("string");

    expect(styles).toContain("<w:docDefaults>");
    expect(styles).toContain('<w:rPrDefault><w:rPr><w:rFonts w:ascii="Aptos" w:hAnsi="Aptos"/><w:sz w:val="22"/><w:color w:val="1F1F1F"/></w:rPr></w:rPrDefault>');
  });

  it("writes paragraph document defaults", async () => {
    const document = {
      version: "1.0" as const,
      styles: {
        defaults: {
          paragraph: {
            spacing: { after: 160, line: 276, lineRule: "auto" as const },
            indent: { firstLine: 420 },
          },
        },
      },
      sections: [{ blocks: [{ type: "paragraph" as const, runs: [{ text: "Paragraph defaults" }] }] }],
    };

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const styles = await zip.file("word/styles.xml")!.async("string");

    expect(styles).toContain('<w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="276" w:lineRule="auto"/><w:ind w:firstLine="420"/></w:pPr></w:pPrDefault>');
  });

  it("writes paragraph style properties", async () => {
    const document = {
      version: "1.0" as const,
      styles: {
        paragraph: [
          {
            id: "ContractTitle",
            name: "Contract Title",
            basedOn: "Normal",
            next: "Normal",
            paragraph: { alignment: "center" as const },
            run: { bold: true, fontFamily: "Aptos Display", fontSize: 18, color: "1F4E79" },
          },
        ],
      },
      sections: [
        {
          blocks: [
            { type: "paragraph" as const, styleId: "ContractTitle", runs: [{ text: "Agreement" }] },
          ],
        },
      ],
    };

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const styles = await zip.file("word/styles.xml")!.async("string");

    expect(styles).toContain('<w:pPr><w:jc w:val="center"/></w:pPr>');
    expect(styles).toContain('<w:rPr><w:b/><w:rFonts w:ascii="Aptos Display" w:hAnsi="Aptos Display"/><w:sz w:val="36"/><w:color w:val="1F4E79"/></w:rPr>');
  });

  it("writes paragraph style layout properties", async () => {
    const document = {
      version: "1.0" as const,
      styles: {
        paragraph: [
          {
            id: "BodyText",
            name: "Body Text",
            paragraph: {
              spacing: { before: 120, after: 120 },
              indent: { left: 360, hanging: 180 },
            },
          },
        ],
      },
      sections: [
        {
          blocks: [
            { type: "paragraph" as const, styleId: "BodyText", runs: [{ text: "Styled body" }] },
          ],
        },
      ],
    };

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const styles = await zip.file("word/styles.xml")!.async("string");

    expect(styles).toContain('<w:pPr><w:spacing w:before="120" w:after="120"/><w:ind w:left="360" w:hanging="180"/></w:pPr>');
  });

  it("writes paragraph style pagination properties", async () => {
    const document = {
      version: "1.0" as const,
      styles: {
        paragraph: [{
          id: "KeepHeading",
          name: "Keep Heading",
          paragraph: {
            pagination: { keepNext: true, keepLines: true, pageBreakBefore: true },
          },
        }],
      },
      sections: [{ blocks: [{ type: "paragraph" as const, styleId: "KeepHeading", runs: [{ text: "Heading" }] }] }],
    };

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const styles = await zip.file("word/styles.xml")!.async("string");

    expect(styles).toContain('<w:style w:type="paragraph" w:styleId="KeepHeading"><w:name w:val="Keep Heading"/><w:pPr><w:keepNext/><w:keepLines/><w:pageBreakBefore/></w:pPr></w:style>');
  });

  it("writes explicit off paragraph style pagination properties", async () => {
    const document = {
      version: "1.0" as const,
      styles: {
        paragraph: [{
          id: "LooseHeading",
          name: "Loose Heading",
          paragraph: {
            pagination: { keepNext: false, keepLines: false, pageBreakBefore: false },
          },
        }],
      },
      sections: [{ blocks: [{ type: "paragraph" as const, styleId: "LooseHeading", runs: [{ text: "Heading" }] }] }],
    };

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const styles = await zip.file("word/styles.xml")!.async("string");

    expect(styles).toContain('<w:style w:type="paragraph" w:styleId="LooseHeading"><w:name w:val="Loose Heading"/><w:pPr><w:keepNext w:val="0"/><w:keepLines w:val="0"/><w:pageBreakBefore w:val="0"/></w:pPr></w:style>');
  });

  it("writes extra paragraph style pagination toggles", async () => {
    const document = {
      version: "1.0" as const,
      styles: {
        paragraph: [{
          id: "ControlledBody",
          name: "Controlled Body",
          paragraph: {
            pagination: { widowControl: false, suppressLineNumbers: true, suppressAutoHyphens: true },
          },
        }],
      },
      sections: [{ blocks: [{ type: "paragraph" as const, styleId: "ControlledBody", runs: [{ text: "Body" }] }] }],
    };

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const styles = await zip.file("word/styles.xml")!.async("string");

    expect(styles).toContain('<w:style w:type="paragraph" w:styleId="ControlledBody"><w:name w:val="Controlled Body"/><w:pPr><w:widowControl w:val="0"/><w:suppressLineNumbers/><w:suppressAutoHyphens/></w:pPr></w:style>');
  });

  it("writes paragraph style layout toggles", async () => {
    const document = {
      version: "1.0" as const,
      styles: {
        paragraph: [{
          id: "LayoutBody",
          name: "Layout Body",
          paragraph: {
            pagination: { contextualSpacing: false, mirrorIndents: true, overflowPunct: false, topLinePunct: true },
          },
        }],
      },
      sections: [{ blocks: [{ type: "paragraph" as const, styleId: "LayoutBody", runs: [{ text: "Body" }] }] }],
    };

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const styles = await zip.file("word/styles.xml")!.async("string");

    expect(styles).toContain('<w:style w:type="paragraph" w:styleId="LayoutBody"><w:name w:val="Layout Body"/><w:pPr><w:contextualSpacing w:val="0"/><w:mirrorIndents/><w:overflowPunct w:val="0"/><w:topLinePunct/></w:pPr></w:style>');
  });

  it("writes paragraph style text flow properties", async () => {
    const document = {
      version: "1.0" as const,
      styles: {
        paragraph: [{
          id: "VerticalBody",
          name: "Vertical Body",
          paragraph: {
            pagination: {
              textAlignment: "baseline",
              textDirection: "btLr",
              adjustRightInd: false,
              autoSpaceDE: true,
              autoSpaceDN: false,
            },
          },
        }],
      },
      sections: [{ blocks: [{ type: "paragraph" as const, styleId: "VerticalBody", runs: [{ text: "Body" }] }] }],
    };

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const styles = await zip.file("word/styles.xml")!.async("string");

    expect(styles).toContain('<w:style w:type="paragraph" w:styleId="VerticalBody"><w:name w:val="Vertical Body"/><w:pPr><w:textAlignment w:val="baseline"/><w:textDirection w:val="btLr"/><w:adjustRightInd w:val="0"/><w:autoSpaceDE/><w:autoSpaceDN w:val="0"/></w:pPr></w:style>');
  });

  it("writes paragraph style borders and shading", async () => {
    const document = {
      version: "1.0" as const,
      styles: {
        paragraph: [
          {
            id: "Callout",
            name: "Callout",
            paragraph: {
              shading: { fill: "E2F0D9" },
              borders: { left: { style: "single", size: 12, color: "70AD47", space: 4 } },
            },
          },
        ],
      },
      sections: [
        {
          blocks: [{ type: "paragraph" as const, styleId: "Callout", runs: [{ text: "Styled callout" }] }],
        },
      ],
    };

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const styles = await zip.file("word/styles.xml")!.async("string");

    expect(styles).toContain('<w:shd w:val="clear" w:fill="E2F0D9"/>');
    expect(styles).toContain('<w:left w:val="single" w:sz="12" w:space="4" w:color="70AD47"/>');
  });

  it("writes paragraph style frame properties", async () => {
    const document = {
      version: "1.0" as const,
      styles: {
        paragraph: [{
          id: "Sidebar",
          name: "Sidebar",
          paragraph: {
            frame: {
              width: 2160,
              horizontalAnchor: "page",
              verticalAnchor: "margin",
              xAlign: "right",
              yAlign: "bottom",
              wrap: "notBeside",
              anchorLock: false,
            },
          },
        }],
      },
      sections: [{ blocks: [{ type: "paragraph" as const, styleId: "Sidebar", runs: [{ text: "Sidebar" }] }] }],
    };

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const styles = await zip.file("word/styles.xml")!.async("string");

    expect(styles).toContain('<w:style w:type="paragraph" w:styleId="Sidebar"><w:name w:val="Sidebar"/><w:pPr><w:framePr w:w="2160" w:hAnchor="page" w:vAnchor="margin" w:xAlign="right" w:yAlign="bottom" w:wrap="notBeside" w:anchorLock="0"/></w:pPr></w:style>');
  });

  it("writes paragraph style tab stops", async () => {
    const document = {
      version: "1.0" as const,
      styles: {
        paragraph: [{
          id: "Tabular",
          name: "Tabular",
          paragraph: {
            tabs: [
              { value: "center", position: 2160, leader: "hyphen" },
              { value: "clear", position: 3600 },
            ],
          },
        }],
      },
      sections: [{ blocks: [{ type: "paragraph" as const, styleId: "Tabular", runs: [{ text: "A\tB" }] }] }],
    };

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const styles = await zip.file("word/styles.xml")!.async("string");

    expect(styles).toContain('<w:style w:type="paragraph" w:styleId="Tabular"><w:name w:val="Tabular"/><w:pPr><w:tabs><w:tab w:val="center" w:pos="2160" w:leader="hyphen"/><w:tab w:val="clear" w:pos="3600"/></w:tabs></w:pPr></w:style>');
  });

  it("writes character and table styles", async () => {
    const document = {
      version: "1.0" as const,
      styles: {
        character: [
          { id: "DefinedTerm", name: "Defined Term", basedOn: "DefaultParagraphFont" },
        ],
        table: [
          { id: "ContractTable", name: "Contract Table", basedOn: "TableNormal" },
        ],
      },
      sections: [
        {
          blocks: [
            {
              type: "paragraph" as const,
              runs: [{ text: "Term", styleId: "DefinedTerm" }],
            },
            {
              type: "table" as const,
              styleId: "ContractTable",
              rows: [
                {
                  cells: [
                    { blocks: [{ type: "paragraph" as const, runs: [{ text: "Cell" }] }] },
                  ],
                },
              ],
            },
          ],
        },
      ],
    };

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");
    const styles = await zip.file("word/styles.xml")!.async("string");

    expect(xml).toContain('<w:rStyle w:val="DefinedTerm"/>');
    expect(xml).toContain('<w:tblStyle w:val="ContractTable"/>');
    expect(styles).toContain('<w:style w:type="character" w:styleId="DefinedTerm">');
    expect(styles).toContain('<w:name w:val="Defined Term"/>');
    expect(styles).toContain('<w:basedOn w:val="DefaultParagraphFont"/>');
    expect(styles).toContain('<w:style w:type="table" w:styleId="ContractTable">');
    expect(styles).toContain('<w:name w:val="Contract Table"/>');
    expect(styles).toContain('<w:basedOn w:val="TableNormal"/>');
  });

  it("writes character style properties", async () => {
    const document = {
      version: "1.0" as const,
      styles: {
        character: [
          {
            id: "DefinedTerm",
            name: "Defined Term",
            basedOn: "DefaultParagraphFont",
            run: { italic: true, underline: true, color: "C00000" },
          },
        ],
      },
      sections: [
        {
          blocks: [
            {
              type: "paragraph" as const,
              runs: [{ text: "Term", styleId: "DefinedTerm" }],
            },
          ],
        },
      ],
    };

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const styles = await zip.file("word/styles.xml")!.async("string");

    expect(styles).toContain('<w:style w:type="character" w:styleId="DefinedTerm">');
    expect(styles).toContain('<w:rPr><w:i/><w:u w:val="single"/><w:color w:val="C00000"/></w:rPr>');
  });

  it("writes style run fonts with direct theme and hint attributes", async () => {
    const document = {
      version: "1.0" as const,
      styles: {
        character: [{
          id: "DefinedTerm",
          name: "Defined Term",
          run: {
            fontFamily: "Aptos",
            eastAsiaFontFamily: "SimSun",
            complexScriptFontFamily: "Arial",
            fontTheme: "majorHAnsi",
            eastAsiaFontTheme: "majorEastAsia",
            complexScriptFontTheme: "majorBidi",
            fontHint: "eastAsia",
          },
        }],
      },
      sections: [
        {
          blocks: [
            {
              type: "paragraph" as const,
              runs: [{ text: "Term", styleId: "DefinedTerm" }],
            },
          ],
        },
      ],
    };

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const styles = await zip.file("word/styles.xml")!.async("string");

    expect(styles).toContain('<w:style w:type="character" w:styleId="DefinedTerm"><w:name w:val="Defined Term"/><w:rPr><w:rFonts w:ascii="Aptos" w:hAnsi="Aptos" w:eastAsia="SimSun" w:cs="Arial" w:asciiTheme="majorHAnsi" w:hAnsiTheme="majorHAnsi" w:eastAsiaTheme="majorEastAsia" w:cstheme="majorBidi" w:hint="eastAsia"/></w:rPr></w:style>');
  });

  it("writes character style underline none", async () => {
    const document = {
      version: "1.0" as const,
      styles: {
        character: [
          {
            id: "PlainTerm",
            name: "Plain Term",
            basedOn: "DefaultParagraphFont",
            run: { underline: false },
          },
        ],
      },
      sections: [
        {
          blocks: [
            {
              type: "paragraph" as const,
              runs: [{ text: "Term", styleId: "PlainTerm" }],
            },
          ],
        },
      ],
    };

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const styles = await zip.file("word/styles.xml")!.async("string");

    expect(styles).toContain('<w:style w:type="character" w:styleId="PlainTerm">');
    expect(styles).toContain('<w:rPr><w:u w:val="none"/></w:rPr>');
  });

  it("writes style run advanced formatting", async () => {
    const document = {
      version: "1.0" as const,
      styles: {
        character: [
          {
            id: "WarningText",
            name: "Warning Text",
            run: { highlight: "yellow" as const, strike: true, verticalAlign: "superscript" as const, characterSpacing: 20, scale: 90 },
          },
        ],
      },
      sections: [
        {
          blocks: [
            {
              type: "paragraph" as const,
              runs: [{ text: "Warning", styleId: "WarningText" }],
            },
          ],
        },
      ],
    };

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const styles = await zip.file("word/styles.xml")!.async("string");

    expect(styles).toContain('<w:style w:type="character" w:styleId="WarningText">');
    expect(styles).toContain('<w:rPr><w:highlight w:val="yellow"/><w:strike/><w:vertAlign w:val="superscript"/><w:spacing w:val="20"/><w:w w:val="90"/></w:rPr>');
  });

  it("writes style run advanced typography properties", async () => {
    const document = {
      version: "1.0" as const,
      styles: {
        character: [{
          id: "EmphasisText",
          name: "Emphasis Text",
          run: {
            complexScriptFontSize: 14,
            language: { value: "en-US", eastAsia: "zh-CN", bidi: "ar-SA" },
            characterPosition: 4,
            kerning: 28,
          },
        }],
      },
      sections: [{ blocks: [{ type: "paragraph" as const, runs: [{ text: "Term", styleId: "EmphasisText" }] }] }],
    };

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const styles = await zip.file("word/styles.xml")!.async("string");

    expect(styles).toContain('<w:rPr><w:szCs w:val="28"/><w:lang w:val="en-US" w:eastAsia="zh-CN" w:bidi="ar-SA"/><w:position w:val="4"/><w:kern w:val="28"/></w:rPr>');
  });

  it("writes style run boolean effect properties", async () => {
    const document = {
      version: "1.0" as const,
      styles: {
        character: [{
          id: "HiddenEffectText",
          name: "Hidden Effect Text",
          run: {
            shadow: true,
            outline: true,
            emboss: true,
            imprint: true,
            rtl: true,
            complexScript: true,
            specVanish: true,
            hidden: true,
            webHidden: true,
            snapToGrid: true,
            noProof: true,
            officeMath: true,
          },
        }],
      },
      sections: [{ blocks: [{ type: "paragraph" as const, runs: [{ text: "Hidden", styleId: "HiddenEffectText" }] }] }],
    };

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const styles = await zip.file("word/styles.xml")!.async("string");

    expect(styles).toContain('<w:rPr><w:shadow/><w:outline/><w:emboss/><w:imprint/><w:rtl/><w:cs/><w:specVanish/><w:vanish/><w:webHidden/><w:snapToGrid/><w:noProof/><w:oMath/></w:rPr>');
  });

  it("writes style run fit text and emphasis properties", async () => {
    const document = {
      version: "1.0" as const,
      styles: {
        character: [{
          id: "CompressedText",
          name: "Compressed Text",
          run: {
            fitText: { width: 1440, id: 11 },
            emphasis: "underDot" as const,
          },
        }],
      },
      sections: [{ blocks: [{ type: "paragraph" as const, runs: [{ text: "Fit", styleId: "CompressedText" }] }] }],
    };

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const styles = await zip.file("word/styles.xml")!.async("string");

    expect(styles).toContain('<w:rPr><w:fitText w:val="1440" w:id="11"/><w:em w:val="underDot"/></w:rPr>');
  });

  it("writes style run explicit off boolean properties", async () => {
    const document = {
      version: "1.0" as const,
      styles: {
        character: [{
          id: "ExplicitOffText",
          name: "Explicit Off Text",
          run: {
            bold: false,
            italic: false,
            strike: false,
            doubleStrike: false,
            smallCaps: false,
            allCaps: false,
            shadow: false,
            outline: false,
            emboss: false,
            imprint: false,
            rtl: false,
            complexScript: false,
            specVanish: false,
            hidden: false,
            webHidden: false,
            snapToGrid: false,
            noProof: false,
            officeMath: false,
          },
        }],
      },
      sections: [{ blocks: [{ type: "paragraph" as const, runs: [{ text: "Off", styleId: "ExplicitOffText" }] }] }],
    };

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const styles = await zip.file("word/styles.xml")!.async("string");

    expect(styles).toContain('<w:rPr><w:b w:val="0"/><w:i w:val="0"/><w:strike w:val="0"/><w:dstrike w:val="0"/><w:smallCaps w:val="0"/><w:caps w:val="0"/><w:shadow w:val="0"/><w:outline w:val="0"/><w:emboss w:val="0"/><w:imprint w:val="0"/><w:rtl w:val="0"/><w:cs w:val="0"/><w:specVanish w:val="0"/><w:vanish w:val="0"/><w:webHidden w:val="0"/><w:snapToGrid w:val="0"/><w:noProof w:val="0"/><w:oMath w:val="0"/></w:rPr>');
  });

  it("writes table style properties", async () => {
    const document = {
      version: "1.0" as const,
      styles: {
        table: [
          {
            id: "ContractTable",
            name: "Contract Table",
            basedOn: "TableNormal",
            table: { borders: "single" as const },
          },
        ],
      },
      sections: [
        {
          blocks: [
            {
              type: "table" as const,
              styleId: "ContractTable",
              rows: [
                {
                  cells: [
                    { blocks: [{ type: "paragraph" as const, runs: [{ text: "Cell" }] }] },
                  ],
                },
              ],
            },
          ],
        },
      ],
    };

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const styles = await zip.file("word/styles.xml")!.async("string");

    expect(styles).toContain('<w:style w:type="table" w:styleId="ContractTable">');
    expect(styles).toContain("<w:tblPr><w:tblBorders>");
  });

  it("writes theme part", async () => {
    const document = {
      version: "1.0" as const,
      theme: {
        name: "Contract Theme",
        fonts: { major: "Aptos Display", minor: "Aptos" },
        colors: { accent1: "4472C4" },
      },
      sections: [
        {
          blocks: [{ type: "paragraph" as const, runs: [{ text: "Themed" }] }],
        },
      ],
    };

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const theme = await zip.file("word/theme/theme1.xml")!.async("string");
    const rels = await zip.file("word/_rels/document.xml.rels")!.async("string");
    const contentTypes = await zip.file("[Content_Types].xml")!.async("string");

    expect(theme).toContain('<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Contract Theme">');
    expect(theme).toContain('<a:latin typeface="Aptos Display"/>');
    expect(theme).toContain('<a:latin typeface="Aptos"/>');
    expect(theme).toContain('<a:srgbClr val="4472C4"/>');
    expect(rels).toContain('Id="rIdTheme"');
    expect(rels).toContain('Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"');
    expect(rels).toContain('Target="theme/theme1.xml"');
    expect(contentTypes).toContain('/word/theme/theme1.xml');
  });

  it("writes rich theme color scheme", async () => {
    const document = {
      version: "1.0" as const,
      theme: {
        name: "Contract Theme",
        fonts: { major: "Aptos Display", minor: "Aptos" },
        colors: {
          dark1: "000000",
          light1: "FFFFFF",
          dark2: "1F2937",
          light2: "F8FAFC",
          accent1: "4472C4",
          accent2: "ED7D31",
          accent3: "A5A5A5",
          accent4: "FFC000",
          accent5: "5B9BD5",
          accent6: "70AD47",
          hyperlink: "0563C1",
          followedHyperlink: "954F72",
        },
      },
      sections: [
        {
          blocks: [{ type: "paragraph" as const, runs: [{ text: "Themed" }] }],
        },
      ],
    };

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const theme = await zip.file("word/theme/theme1.xml")!.async("string");

    expect(theme).toContain('<a:dk1><a:srgbClr val="000000"/></a:dk1>');
    expect(theme).toContain('<a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>');
    expect(theme).toContain('<a:dk2><a:srgbClr val="1F2937"/></a:dk2>');
    expect(theme).toContain('<a:lt2><a:srgbClr val="F8FAFC"/></a:lt2>');
    expect(theme).toContain('<a:accent1><a:srgbClr val="4472C4"/></a:accent1>');
    expect(theme).toContain('<a:accent2><a:srgbClr val="ED7D31"/></a:accent2>');
    expect(theme).toContain('<a:accent3><a:srgbClr val="A5A5A5"/></a:accent3>');
    expect(theme).toContain('<a:accent4><a:srgbClr val="FFC000"/></a:accent4>');
    expect(theme).toContain('<a:accent5><a:srgbClr val="5B9BD5"/></a:accent5>');
    expect(theme).toContain('<a:accent6><a:srgbClr val="70AD47"/></a:accent6>');
    expect(theme).toContain('<a:hlink><a:srgbClr val="0563C1"/></a:hlink>');
    expect(theme).toContain('<a:folHlink><a:srgbClr val="954F72"/></a:folHlink>');
  });

  it("writes theme script fonts", async () => {
    const document = {
      version: "1.0" as const,
      theme: {
        name: "Multilingual Theme",
        fonts: {
          major: "Aptos Display",
          minor: "Aptos",
          majorEastAsia: "SimSun",
          majorComplexScript: "Arial",
          minorEastAsia: "Microsoft YaHei",
          minorComplexScript: "Arial",
        },
        colors: { accent1: "4472C4" },
      },
      sections: [
        {
          blocks: [{ type: "paragraph" as const, runs: [{ text: "Themed" }] }],
        },
      ],
    };

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const theme = await zip.file("word/theme/theme1.xml")!.async("string");

    expect(theme).toContain('<a:majorFont><a:latin typeface="Aptos Display"/><a:ea typeface="SimSun"/><a:cs typeface="Arial"/></a:majorFont>');
    expect(theme).toContain('<a:minorFont><a:latin typeface="Aptos"/><a:ea typeface="Microsoft YaHei"/><a:cs typeface="Arial"/></a:minorFont>');
  });

  it("writes theme supplemental fonts", async () => {
    const document = {
      version: "1.0" as const,
      theme: {
        name: "Multilingual Theme",
        fonts: {
          major: "Aptos Display",
          minor: "Aptos",
          supplemental: [
            { group: "major" as const, script: "Hans", typeface: "SimSun" },
            { group: "major" as const, script: "Jpan", typeface: "Yu Gothic" },
            { group: "minor" as const, script: "Hans", typeface: "Microsoft YaHei" },
            { group: "minor" as const, script: "Hang", typeface: "Malgun Gothic" },
          ],
        },
        colors: { accent1: "4472C4" },
      },
      sections: [
        {
          blocks: [{ type: "paragraph" as const, runs: [{ text: "Themed" }] }],
        },
      ],
    };

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const theme = await zip.file("word/theme/theme1.xml")!.async("string");

    expect(theme).toContain('<a:font script="Hans" typeface="SimSun"/>');
    expect(theme).toContain('<a:font script="Jpan" typeface="Yu Gothic"/>');
    expect(theme).toContain('<a:font script="Hans" typeface="Microsoft YaHei"/>');
    expect(theme).toContain('<a:font script="Hang" typeface="Malgun Gothic"/>');
  });

  it("writes theme format scheme", async () => {
    const document = {
      version: "1.0" as const,
      theme: {
        name: "Visual Theme",
        fonts: { major: "Aptos Display", minor: "Aptos" },
        colors: { accent1: "4472C4" },
        formatScheme: {
          name: "Visual Formats",
          fillStyleColors: ["FFFFFF", "F2F2F2"],
          lineStyleColors: ["4472C4", "70AD47"],
          effectStyleColors: ["808080"],
          backgroundFillStyleColors: ["000000", "1F2937"],
        },
      },
      sections: [
        {
          blocks: [{ type: "paragraph" as const, runs: [{ text: "Themed" }] }],
        },
      ],
    };

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const theme = await zip.file("word/theme/theme1.xml")!.async("string");

    expect(theme).toContain('<a:fmtScheme name="Visual Formats">');
    expect(theme).toContain("<a:fillStyleLst>");
    expect(theme).toContain('<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>');
    expect(theme).toContain('<a:solidFill><a:srgbClr val="F2F2F2"/></a:solidFill>');
    expect(theme).toContain("<a:lnStyleLst>");
    expect(theme).toContain('<a:ln w="9525"><a:solidFill><a:srgbClr val="4472C4"/></a:solidFill></a:ln>');
    expect(theme).toContain('<a:ln w="9525"><a:solidFill><a:srgbClr val="70AD47"/></a:solidFill></a:ln>');
    expect(theme).toContain("<a:effectStyleLst>");
    expect(theme).toContain('<a:outerShdw><a:srgbClr val="808080"/></a:outerShdw>');
    expect(theme).toContain("<a:bgFillStyleLst>");
    expect(theme).toContain('<a:solidFill><a:srgbClr val="000000"/></a:solidFill>');
    expect(theme).toContain('<a:solidFill><a:srgbClr val="1F2937"/></a:solidFill>');
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

  it("round-trips text run underline none", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "Plain", underline: false }],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips text run bold off", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "Plain", bold: false }],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips text run italic off", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "Plain", italic: false }],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips text run strike off", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "Plain", strike: false }],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips text run double strike off", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "Plain", doubleStrike: false }],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips text run small caps off", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "Plain", smallCaps: false }],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips text run all caps off", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "Plain", allCaps: false }],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips text run shadow on", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "Shadow", shadow: true }],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips text run shadow off", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "Plain", shadow: false }],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips text run outline on", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "Outline", outline: true }],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips text run outline off", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "Plain", outline: false }],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips text run emboss on", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "Emboss", emboss: true }],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips text run emboss off", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "Plain", emboss: false }],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips text run imprint on", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "Imprint", imprint: true }],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips text run imprint off", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "Plain", imprint: false }],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips text run rtl on", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "يمين", rtl: true }],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips text run rtl off", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "Plain", rtl: false }],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips text run complex script on", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "Complex", complexScript: true }],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips text run complex script off", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "Plain", complexScript: false }],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips text run spec vanish on", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "Hidden", specVanish: true }],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips text run spec vanish off", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "Visible", specVanish: false }],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips text run hidden on", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "Hidden", hidden: true }],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips text run hidden off", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "Visible", hidden: false }],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips text run web hidden on", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "WebHidden", webHidden: true }],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips text run web hidden off", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "Visible", webHidden: false }],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips text run fit text with id", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "Fit", fitText: { width: 1440, id: 7 } }],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips text run fit text without id", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "Fit", fitText: { width: 720 } }],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips text run emphasis dot", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "Emphasis", emphasis: "dot" }],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips text run emphasis none", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "Plain", emphasis: "none" }],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips text run snap to grid on", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "Grid", snapToGrid: true }],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips text run snap to grid off", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "Free", snapToGrid: false }],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips text run no proof on", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "Unchecked", noProof: true }],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips text run no proof off", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "Checked", noProof: false }],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips text run office math on", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "Formula", officeMath: true }],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips text run office math off", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "Plain", officeMath: false }],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips text run language value", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "Hello", language: { value: "en-US" } }],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips text run language east asia and bidi", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "Mixed", language: { value: "en-US", eastAsia: "zh-CN", bidi: "ar-SA" } }],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips text run character position raised", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "Raised", characterPosition: 4 }],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips text run character position lowered", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "Lowered", characterPosition: -4 }],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips text run kerning small threshold", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "Title", kerning: 24 }],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips text run kerning large threshold", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "Display", kerning: 48 }],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips text run complex script font size", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "مرحبا", complexScriptFontSize: 14 }],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips text run complex script font size with latin size", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "Mixed", fontSize: 12, complexScriptFontSize: 16 }],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips text run complex script font family", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "مرحبا", complexScriptFontFamily: "Arial" }],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips text run complex script font family with latin family", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "Mixed", fontFamily: "Aptos", complexScriptFontFamily: "Arial" }],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips text run east asia font family", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "你好", eastAsiaFontFamily: "SimSun" }],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips text run east asia font family with latin and complex script families", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "Mixed", fontFamily: "Aptos", eastAsiaFontFamily: "SimSun", complexScriptFontFamily: "Arial" }],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips text run font hint east asia", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "你好", fontHint: "eastAsia" }],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips text run font hint with all font families", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "Mixed", fontFamily: "Aptos", eastAsiaFontFamily: "SimSun", complexScriptFontFamily: "Arial", fontHint: "cs" }],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips text run theme fonts", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "Theme", fontTheme: "minorHAnsi", eastAsiaFontTheme: "minorEastAsia", complexScriptFontTheme: "minorBidi" }],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips text run theme fonts with direct fonts and hint", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [{
          text: "Mixed",
          fontFamily: "Aptos",
          eastAsiaFontFamily: "SimSun",
          complexScriptFontFamily: "Arial",
          fontTheme: "majorHAnsi",
          eastAsiaFontTheme: "majorEastAsia",
          complexScriptFontTheme: "majorBidi",
          fontHint: "eastAsia",
        }],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips inline office math runs", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          { text: "Equation: " },
          { text: "", math: { text: "x+1=2" } },
          { text: " solved" },
        ],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips structured office math runs", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "fraction" as const,
                  numerator: [{ type: "text" as const, text: "1" }],
                  denominator: [{ type: "text" as const, text: "2" }],
                },
                {
                  type: "superscript" as const,
                  base: [{ type: "text" as const, text: "x" }],
                  superscript: [{ type: "text" as const, text: "2" }],
                },
                {
                  type: "subscript" as const,
                  base: [{ type: "text" as const, text: "a" }],
                  subscript: [{ type: "text" as const, text: "i" }],
                },
              ],
            },
          },
        ],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips radical and n-ary office math runs", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "radical" as const,
                  degree: [{ type: "text" as const, text: "3" }],
                  content: [{ type: "text" as const, text: "x" }],
                },
                {
                  type: "nary" as const,
                  operator: "sum" as const,
                  lowerLimit: [{ type: "text" as const, text: "i=1" }],
                  upperLimit: [{ type: "text" as const, text: "n" }],
                  body: [{ type: "text" as const, text: "i" }],
                },
              ],
            },
          },
        ],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips radical hide degree", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "radical" as const,
                  hideDegree: false,
                  degree: [{ type: "text" as const, text: "3" }],
                  content: [{ type: "text" as const, text: "x" }],
                },
              ],
            },
          },
        ],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips radical control properties", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "radical" as const,
                  controlProperties: {
                    bold: true,
                    italic: true,
                    fontFamily: "Cambria Math",
                    fontSize: 16,
                    color: "8064A2",
                    underline: true,
                    highlight: "darkYellow",
                  },
                  degree: [{ type: "text" as const, text: "3" }],
                  content: [{ type: "text" as const, text: "x" }],
                },
              ],
            },
          },
        ],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips nary control properties", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "nary" as const,
                  operator: "sum" as const,
                  controlProperties: {
                    bold: true,
                    italic: true,
                    fontFamily: "Cambria Math",
                    fontSize: 16,
                    color: "8064A2",
                    underline: true,
                    highlight: "darkYellow",
                  },
                  lowerLimit: [{ type: "text" as const, text: "i=1" }],
                  upperLimit: [{ type: "text" as const, text: "n" }],
                  body: [{ type: "text" as const, text: "i" }],
                },
              ],
            },
          },
        ],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips nary operator", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "nary" as const,
                  operator: "integral" as const,
                  lowerLimit: [{ type: "text" as const, text: "0" }],
                  upperLimit: [{ type: "text" as const, text: "1" }],
                  body: [{ type: "text" as const, text: "f(x)" }],
                },
              ],
            },
          },
        ],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips nary custom operator character", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "nary" as const,
                  operator: "sum" as const,
                  operatorCharacter: "⊕",
                  lowerLimit: [{ type: "text" as const, text: "i=1" }],
                  upperLimit: [{ type: "text" as const, text: "n" }],
                  body: [{ type: "text" as const, text: "i" }],
                },
              ],
            },
          },
        ],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips nary limit location", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "nary" as const,
                  operator: "sum" as const,
                  limitLocation: "subSup" as const,
                  lowerLimit: [{ type: "text" as const, text: "i=1" }],
                  upperLimit: [{ type: "text" as const, text: "n" }],
                  body: [{ type: "text" as const, text: "i" }],
                },
              ],
            },
          },
        ],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips nary grow", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "nary" as const,
                  operator: "sum" as const,
                  grow: false,
                  lowerLimit: [{ type: "text" as const, text: "i=1" }],
                  upperLimit: [{ type: "text" as const, text: "n" }],
                  body: [{ type: "text" as const, text: "i" }],
                },
              ],
            },
          },
        ],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips nary hide limits", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "nary" as const,
                  operator: "sum" as const,
                  hideLowerLimit: true,
                  hideUpperLimit: true,
                  lowerLimit: [{ type: "text" as const, text: "i=1" }],
                  upperLimit: [{ type: "text" as const, text: "n" }],
                  body: [{ type: "text" as const, text: "i" }],
                },
              ],
            },
          },
        ],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips nary explicit visible limits", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "nary" as const,
                  operator: "sum" as const,
                  hideLowerLimit: false,
                  hideUpperLimit: false,
                  lowerLimit: [{ type: "text" as const, text: "i=1" }],
                  upperLimit: [{ type: "text" as const, text: "n" }],
                  body: [{ type: "text" as const, text: "i" }],
                },
              ],
            },
          },
        ],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips matrix office math runs", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "matrix" as const,
                  rows: [
                    [[{ type: "text" as const, text: "a" }], [{ type: "text" as const, text: "b" }]],
                    [[{ type: "text" as const, text: "c" }], [{ type: "text" as const, text: "d" }]],
                  ],
                },
              ],
            },
          },
        ],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips matrix control properties", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "matrix" as const,
                  controlProperties: {
                    bold: true,
                    italic: true,
                    fontFamily: "Cambria Math",
                    fontSize: 16,
                    color: "8064A2",
                    underline: true,
                    highlight: "darkYellow",
                  },
                  rows: [
                    [[{ type: "text" as const, text: "a" }], [{ type: "text" as const, text: "b" }]],
                  ],
                },
              ],
            },
          },
        ],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips matrix base justification", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "matrix" as const,
                  baseJustification: "bottom",
                  rows: [
                    [
                      [{ type: "text" as const, text: "a" }],
                      [{ type: "text" as const, text: "b" }],
                    ],
                  ],
                },
              ],
            },
          },
        ],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips matrix column justifications", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "matrix" as const,
                  columnJustifications: ["left", "center", "right"],
                  rows: [
                    [
                      [{ type: "text" as const, text: "a" }],
                      [{ type: "text" as const, text: "b" }],
                      [{ type: "text" as const, text: "c" }],
                    ],
                  ],
                },
              ],
            },
          },
        ],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips matrix column counts", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "matrix" as const,
                  columnCounts: [1, 2, 1],
                  rows: [
                    [
                      [{ type: "text" as const, text: "a" }],
                      [{ type: "text" as const, text: "b" }],
                      [{ type: "text" as const, text: "c" }],
                    ],
                  ],
                },
              ],
            },
          },
        ],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips matrix row spacing", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "matrix" as const,
                  rowSpacing: 3,
                  rows: [
                    [[{ type: "text" as const, text: "a" }]],
                    [[{ type: "text" as const, text: "b" }]],
                  ],
                },
              ],
            },
          },
        ],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips matrix row spacing rule", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "matrix" as const,
                  rowSpacing: 3,
                  rowSpacingRule: "exactly",
                  rows: [
                    [[{ type: "text" as const, text: "a" }]],
                    [[{ type: "text" as const, text: "b" }]],
                  ],
                },
              ],
            },
          },
        ],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips matrix column spacing", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "matrix" as const,
                  columnSpacing: 4,
                  rows: [
                    [
                      [{ type: "text" as const, text: "a" }],
                      [{ type: "text" as const, text: "b" }],
                    ],
                  ],
                },
              ],
            },
          },
        ],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips matrix column spacing rule", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "matrix" as const,
                  columnSpacing: 4,
                  columnSpacingRule: "multiple",
                  rows: [
                    [
                      [{ type: "text" as const, text: "a" }],
                      [{ type: "text" as const, text: "b" }],
                    ],
                  ],
                },
              ],
            },
          },
        ],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips delimiter office math runs", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "delimiter" as const,
                  begin: "|",
                  end: "|",
                  content: [
                    {
                      type: "superscript" as const,
                      base: [{ type: "text" as const, text: "x" }],
                      superscript: [{ type: "text" as const, text: "2" }],
                    },
                  ],
                },
              ],
            },
          },
        ],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips delimiter control properties", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "delimiter" as const,
                  begin: "[",
                  end: "]",
                  controlProperties: {
                    bold: true,
                    italic: true,
                    fontFamily: "Cambria Math",
                    fontSize: 16,
                    color: "8064A2",
                    underline: true,
                    highlight: "darkYellow",
                  },
                  content: [{ type: "text" as const, text: "x" }],
                },
              ],
            },
          },
        ],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips delimiter grow", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "delimiter" as const,
                  begin: "(",
                  end: ")",
                  grow: false,
                  content: [{ type: "text" as const, text: "x" }],
                },
              ],
            },
          },
        ],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips delimiter separator", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "delimiter" as const,
                  begin: "{",
                  end: "}",
                  separator: "|",
                  content: [{ type: "text" as const, text: "x" }],
                },
              ],
            },
          },
        ],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips accent office math runs", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "accent" as const,
                  mark: "ˆ",
                  content: [
                    {
                      type: "subscript" as const,
                      base: [{ type: "text" as const, text: "v" }],
                      subscript: [{ type: "text" as const, text: "i" }],
                    },
                  ],
                },
              ],
            },
          },
        ],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips accent control properties", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "accent" as const,
                  mark: "¯",
                  controlProperties: {
                    bold: true,
                    italic: true,
                    fontFamily: "Cambria Math",
                    fontSize: 16,
                    color: "8064A2",
                    underline: true,
                    highlight: "darkYellow",
                  },
                  content: [{ type: "text" as const, text: "x" }],
                },
              ],
            },
          },
        ],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips bar office math runs", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "bar" as const,
                  position: "bottom" as const,
                  content: [
                    {
                      type: "fraction" as const,
                      numerator: [{ type: "text" as const, text: "a" }],
                      denominator: [{ type: "text" as const, text: "b" }],
                    },
                  ],
                },
              ],
            },
          },
        ],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips bar control properties", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "bar" as const,
                  position: "top" as const,
                  controlProperties: {
                    bold: true,
                    italic: true,
                    fontFamily: "Cambria Math",
                    fontSize: 16,
                    color: "8064A2",
                    underline: true,
                    highlight: "darkYellow",
                  },
                  content: [{ type: "text" as const, text: "x+y" }],
                },
              ],
            },
          },
        ],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips function office math runs", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "function" as const,
                  name: [{ type: "text" as const, text: "log" }],
                  argument: [
                    {
                      type: "fraction" as const,
                      numerator: [{ type: "text" as const, text: "a" }],
                      denominator: [{ type: "text" as const, text: "b" }],
                    },
                  ],
                },
              ],
            },
          },
        ],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips function control properties", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "function" as const,
                  controlProperties: {
                    bold: true,
                    italic: true,
                    fontFamily: "Cambria Math",
                    fontSize: 16,
                    color: "8064A2",
                    underline: true,
                    highlight: "darkYellow",
                  },
                  name: [{ type: "text" as const, text: "sin" }],
                  argument: [{ type: "text" as const, text: "x" }],
                },
              ],
            },
          },
        ],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips limit office math runs", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "limitLower" as const,
                  base: [{ type: "text" as const, text: "lim" }],
                  limit: [
                    {
                      type: "subscript" as const,
                      base: [{ type: "text" as const, text: "x" }],
                      subscript: [{ type: "text" as const, text: "0" }],
                    },
                  ],
                },
                {
                  type: "limitUpper" as const,
                  base: [{ type: "text" as const, text: "sup" }],
                  limit: [{ type: "text" as const, text: "n" }],
                },
              ],
            },
          },
        ],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips limit lower control properties", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "limitLower" as const,
                  controlProperties: {
                    bold: true,
                    italic: true,
                    fontFamily: "Cambria Math",
                    fontSize: 16,
                    color: "8064A2",
                    underline: true,
                    highlight: "darkYellow",
                  },
                  base: [{ type: "text" as const, text: "lim" }],
                  limit: [{ type: "text" as const, text: "x→0" }],
                },
              ],
            },
          },
        ],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips limit upper control properties", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "limitUpper" as const,
                  controlProperties: {
                    bold: true,
                    italic: true,
                    fontFamily: "Cambria Math",
                    fontSize: 16,
                    color: "8064A2",
                    underline: true,
                    highlight: "darkYellow",
                  },
                  base: [{ type: "text" as const, text: "max" }],
                  limit: [{ type: "text" as const, text: "n" }],
                },
              ],
            },
          },
        ],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips equation array office math runs", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "equationArray" as const,
                  rows: [
                    [{ type: "text" as const, text: "x=1" }],
                    [
                      {
                        type: "superscript" as const,
                        base: [{ type: "text" as const, text: "y" }],
                        superscript: [{ type: "text" as const, text: "2" }],
                      },
                    ],
                  ],
                },
              ],
            },
          },
        ],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips equation array control properties", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "equationArray" as const,
                  controlProperties: {
                    bold: true,
                    italic: true,
                    fontFamily: "Cambria Math",
                    fontSize: 16,
                    color: "8064A2",
                    underline: true,
                    highlight: "darkYellow",
                  },
                  rows: [
                    [{ type: "text" as const, text: "x=1" }],
                    [{ type: "text" as const, text: "y=2" }],
                  ],
                },
              ],
            },
          },
        ],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips equation array row spacing", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "equationArray" as const,
                  rowSpacing: 3,
                  rows: [
                    [{ type: "text" as const, text: "x=1" }],
                    [{ type: "text" as const, text: "y=2" }],
                  ],
                },
              ],
            },
          },
        ],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips equation array base justification", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "equationArray" as const,
                  baseJustification: "bottom",
                  rows: [
                    [{ type: "text" as const, text: "x=1" }],
                    [{ type: "text" as const, text: "y=2" }],
                  ],
                },
              ],
            },
          },
        ],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips equation array vertical justification", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "equationArray" as const,
                  verticalJustification: "bottom",
                  rows: [
                    [{ type: "text" as const, text: "x=1" }],
                    [{ type: "text" as const, text: "y=2" }],
                  ],
                },
              ],
            },
          },
        ],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips equation array alignment", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "equationArray" as const,
                  alignment: false,
                  rows: [
                    [{ type: "text" as const, text: "x=1" }],
                    [{ type: "text" as const, text: "y=2" }],
                  ],
                },
              ],
            },
          },
        ],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips equation array row spacing rule", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "equationArray" as const,
                  rowSpacing: 3,
                  rowSpacingRule: "exactly",
                  rows: [
                    [{ type: "text" as const, text: "x=1" }],
                    [{ type: "text" as const, text: "y=2" }],
                  ],
                },
              ],
            },
          },
        ],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips equation array object distribution", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "equationArray" as const,
                  objectDistribution: false,
                  rows: [
                    [{ type: "text" as const, text: "x=1" }],
                    [{ type: "text" as const, text: "y=2" }],
                  ],
                },
              ],
            },
          },
        ],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips equation array max distribution", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "equationArray" as const,
                  maxDistribution: false,
                  rows: [
                    [{ type: "text" as const, text: "x=1" }],
                    [{ type: "text" as const, text: "y=2" }],
                  ],
                },
              ],
            },
          },
        ],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips box office math runs", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "box" as const,
                  hideLeft: true,
                  hideRight: true,
                  content: [
                    {
                      type: "superscript" as const,
                      base: [{ type: "text" as const, text: "x" }],
                      superscript: [{ type: "text" as const, text: "2" }],
                    },
                  ],
                },
              ],
            },
          },
        ],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips box explicit visible borders", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "box" as const,
                  hideTop: false,
                  hideBottom: false,
                  hideLeft: false,
                  hideRight: false,
                  content: [
                    {
                      type: "superscript" as const,
                      base: [{ type: "text" as const, text: "x" }],
                      superscript: [{ type: "text" as const, text: "2" }],
                    },
                  ],
                },
              ],
            },
          },
        ],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips box control properties", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "box" as const,
                  hideTop: true,
                  hideBottom: true,
                  controlProperties: {
                    bold: true,
                    italic: true,
                    fontFamily: "Cambria Math",
                    fontSize: 16,
                    color: "8064A2",
                    underline: true,
                    highlight: "darkYellow",
                  },
                  content: [{ type: "text" as const, text: "x+y" }],
                },
              ],
            },
          },
        ],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips border box office math runs", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "borderBox" as const,
                  hideLeft: true,
                  hideRight: true,
                  content: [
                    {
                      type: "fraction" as const,
                      numerator: [{ type: "text" as const, text: "a" }],
                      denominator: [{ type: "text" as const, text: "b" }],
                    },
                  ],
                },
              ],
            },
          },
        ],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips border box explicit visible borders", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "borderBox" as const,
                  hideTop: false,
                  hideBottom: false,
                  hideLeft: false,
                  hideRight: false,
                  content: [
                    {
                      type: "fraction" as const,
                      numerator: [{ type: "text" as const, text: "a" }],
                      denominator: [{ type: "text" as const, text: "b" }],
                    },
                  ],
                },
              ],
            },
          },
        ],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips border box control properties", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "borderBox" as const,
                  hideTop: true,
                  hideBottom: true,
                  controlProperties: {
                    bold: true,
                    italic: true,
                    fontFamily: "Cambria Math",
                    fontSize: 16,
                    color: "8064A2",
                    underline: true,
                    highlight: "darkYellow",
                  },
                  content: [{ type: "text" as const, text: "x+y" }],
                },
              ],
            },
          },
        ],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips phantom office math runs", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "phantom" as const,
                  show: false,
                  zeroAscent: true,
                  zeroDescent: true,
                  content: [
                    {
                      type: "fraction" as const,
                      numerator: [{ type: "text" as const, text: "a" }],
                      denominator: [{ type: "text" as const, text: "b" }],
                    },
                  ],
                },
              ],
            },
          },
        ],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips phantom explicit false properties", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "phantom" as const,
                  zeroWidth: false,
                  zeroAscent: false,
                  zeroDescent: false,
                  transparent: false,
                  content: [
                    {
                      type: "fraction" as const,
                      numerator: [{ type: "text" as const, text: "a" }],
                      denominator: [{ type: "text" as const, text: "b" }],
                    },
                  ],
                },
              ],
            },
          },
        ],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips phantom control properties", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "phantom" as const,
                  show: false,
                  zeroWidth: true,
                  controlProperties: {
                    bold: true,
                    italic: true,
                    fontFamily: "Cambria Math",
                    fontSize: 16,
                    color: "8064A2",
                    underline: true,
                    highlight: "darkYellow",
                  },
                  content: [{ type: "text" as const, text: "x+y" }],
                },
              ],
            },
          },
        ],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips group character office math runs", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "groupCharacter" as const,
                  character: "⏟",
                  position: "bottom" as const,
                  verticalJustification: "top" as const,
                  content: [
                    {
                      type: "fraction" as const,
                      numerator: [{ type: "text" as const, text: "a" }],
                      denominator: [{ type: "text" as const, text: "b" }],
                    },
                  ],
                },
              ],
            },
          },
        ],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips group character control properties", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "groupCharacter" as const,
                  character: "⏟",
                  position: "bottom" as const,
                  verticalJustification: "top" as const,
                  controlProperties: {
                    bold: true,
                    italic: true,
                    fontFamily: "Cambria Math",
                    fontSize: 16,
                    color: "8064A2",
                    underline: true,
                    highlight: "darkYellow",
                  },
                  content: [{ type: "text" as const, text: "x+y" }],
                },
              ],
            },
          },
        ],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips pre sub sup office math runs", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "preSubSup" as const,
                  base: [{ type: "text" as const, text: "T" }],
                  subscript: [
                    {
                      type: "fraction" as const,
                      numerator: [{ type: "text" as const, text: "i" }],
                      denominator: [{ type: "text" as const, text: "n" }],
                    },
                  ],
                  superscript: [{ type: "text" as const, text: "j" }],
                },
              ],
            },
          },
        ],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips pre sub sup control properties", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "preSubSup" as const,
                  controlProperties: {
                    bold: true,
                    italic: true,
                    fontFamily: "Cambria Math",
                    fontSize: 16,
                    color: "8064A2",
                    underline: true,
                    highlight: "darkYellow",
                  },
                  base: [{ type: "text" as const, text: "X" }],
                  subscript: [{ type: "text" as const, text: "i" }],
                  superscript: [{ type: "text" as const, text: "j" }],
                },
              ],
            },
          },
        ],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips sub sup office math runs", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "subSup" as const,
                  base: [{ type: "text" as const, text: "x" }],
                  subscript: [{ type: "text" as const, text: "i" }],
                  superscript: [
                    {
                      type: "fraction" as const,
                      numerator: [{ type: "text" as const, text: "m" }],
                      denominator: [{ type: "text" as const, text: "n" }],
                    },
                  ],
                },
              ],
            },
          },
        ],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips superscript control properties", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "superscript" as const,
                  controlProperties: {
                    bold: true,
                    italic: true,
                    fontFamily: "Cambria Math",
                    fontSize: 12,
                    color: "7030A0",
                    underline: true,
                    highlight: "green",
                  },
                  base: [{ type: "text" as const, text: "x" }],
                  superscript: [{ type: "text" as const, text: "2" }],
                },
              ],
            },
          },
        ],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips subscript control properties", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "subscript" as const,
                  controlProperties: {
                    bold: true,
                    italic: true,
                    fontFamily: "Cambria Math",
                    fontSize: 11,
                    color: "00B050",
                    underline: true,
                    highlight: "magenta",
                  },
                  base: [{ type: "text" as const, text: "a" }],
                  subscript: [{ type: "text" as const, text: "i" }],
                },
              ],
            },
          },
        ],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips fraction control properties", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "fraction" as const,
                  controlProperties: {
                    bold: true,
                    italic: true,
                    fontFamily: "Cambria Math",
                    fontSize: 15,
                    color: "C55A11",
                    underline: true,
                    highlight: "blue",
                  },
                  numerator: [{ type: "text" as const, text: "1" }],
                  denominator: [{ type: "text" as const, text: "2" }],
                },
              ],
            },
          },
        ],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips fraction type", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "fraction" as const,
                  fractionType: "linear",
                  numerator: [{ type: "text" as const, text: "1" }],
                  denominator: [{ type: "text" as const, text: "2" }],
                },
              ],
            },
          },
        ],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips sub sup control properties", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "subSup" as const,
                  controlProperties: {
                    bold: true,
                    italic: true,
                    fontFamily: "Cambria Math",
                    fontSize: 13,
                    color: "0070C0",
                    underline: true,
                    highlight: "cyan",
                  },
                  base: [{ type: "text" as const, text: "x" }],
                  subscript: [{ type: "text" as const, text: "i" }],
                  superscript: [{ type: "text" as const, text: "2" }],
                },
              ],
            },
          },
        ],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips s pre office math runs", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "sPre" as const,
                  base: [{ type: "text" as const, text: "T" }],
                  subscript: [
                    {
                      type: "fraction" as const,
                      numerator: [{ type: "text" as const, text: "i" }],
                      denominator: [{ type: "text" as const, text: "n" }],
                    },
                  ],
                  superscript: [{ type: "text" as const, text: "j" }],
                },
              ],
            },
          },
        ],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips s pre control properties", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "sPre" as const,
                  controlProperties: {
                    bold: true,
                    italic: true,
                    fontFamily: "Cambria Math",
                    fontSize: 14,
                    color: "C00000",
                    underline: true,
                    highlight: "yellow",
                  },
                  base: [{ type: "text" as const, text: "T" }],
                  subscript: [{ type: "text" as const, text: "i" }],
                  superscript: [{ type: "text" as const, text: "j" }],
                },
              ],
            },
          },
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

  it("round-trips run highlight and strike", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          { text: "Marked", highlight: "yellow", strike: true, doubleStrike: true },
        ],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips run caps and vertical align", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          { text: "Formula", smallCaps: true, allCaps: true, verticalAlign: "superscript" as const },
          { text: "2", verticalAlign: "subscript" as const },
        ],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips run character spacing and scale", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          { text: "Tracked", characterSpacing: 20, scale: 90 },
        ],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips run border", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          { text: "Boxed", border: { style: "single", size: 6, color: "C00000", space: 1 } },
        ],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips inserted run revisions", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          { text: "Added", revision: { type: "insert", id: 1, author: "Ada", date: "2026-07-04T00:00:00.000Z" } },
        ],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips deleted run revisions", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          { text: "Removed", revision: { type: "delete", id: 2, author: "Lin", date: "2026-07-04T01:00:00.000Z" } },
        ],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips run move revisions", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          { text: "Moved away", revision: { type: "moveFrom", id: 30, author: "Ada", date: "2026-07-04T09:00:00.000Z" } },
          { text: "Moved here", revision: { type: "moveTo", id: 31, author: "Lin", date: "2026-07-04T10:00:00.000Z" } },
        ],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips table property change metadata", async () => {
    const source = createDocumentJson([
      {
        type: "table",
        width: 7200,
        propertyRevision: { id: 20, author: "Ada", date: "2026-07-04T05:00:00.000Z" },
        rows: [
          { cells: [{ blocks: [{ type: "paragraph", runs: [{ text: "Table changed" }] }] }] },
        ],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips table row revisions", async () => {
    const source = createDocumentJson([
      {
        type: "table",
        rows: [
          {
            revision: { type: "insert", id: 21, author: "Lin", date: "2026-07-04T06:00:00.000Z" },
            cells: [{ blocks: [{ type: "paragraph", runs: [{ text: "Inserted row" }] }] }],
          },
          {
            revision: { type: "delete", id: 22, author: "Mira", date: "2026-07-04T07:00:00.000Z" },
            cells: [{ blocks: [{ type: "paragraph", runs: [{ text: "Deleted row" }] }] }],
          },
        ],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips cell property change metadata", async () => {
    const source = createDocumentJson([
      {
        type: "table",
        rows: [
          {
            cells: [
              {
                shading: { fill: "D9EAF7" },
                propertyRevision: { id: 23, author: "Noor", date: "2026-07-04T08:00:00.000Z" },
                blocks: [{ type: "paragraph", runs: [{ text: "Cell changed" }] }],
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

  it("round-trips paragraph insert revisions", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        revision: { type: "insert", id: 10, author: "Ada", date: "2026-07-04T02:00:00.000Z" },
        runs: [{ text: "Inserted paragraph" }],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips paragraph delete revisions", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        revision: { type: "delete", id: 11, author: "Lin", date: "2026-07-04T03:00:00.000Z" },
        runs: [{ text: "Deleted paragraph" }],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips paragraph move revisions", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        revision: { type: "moveFrom", id: 32, author: "Mira", date: "2026-07-04T11:00:00.000Z" },
        runs: [{ text: "Moved paragraph from" }],
      },
      {
        type: "paragraph",
        revision: { type: "moveTo", id: 33, author: "Noor", date: "2026-07-04T12:00:00.000Z" },
        runs: [{ text: "Moved paragraph to" }],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips paragraph property change metadata", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        spacing: { before: 240 },
        propertyRevision: { id: 12, author: "Mira", date: "2026-07-04T04:00:00.000Z" },
        runs: [{ text: "Changed spacing" }],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips block content controls", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        contentControl: { alias: "Customer Name", tag: "customer.name", lock: "sdtContentLocked" },
        runs: [{ text: "Acme Inc." }],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips run content controls", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          { text: "INV-001", contentControl: { alias: "Invoice Number", tag: "invoice.number" } },
        ],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips checkbox content controls", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "☒",
            contentControl: {
              alias: "Accepted",
              tag: "accepted",
              checkbox: { checked: true, checkedSymbol: "2612", uncheckedSymbol: "2610" },
            },
          },
        ],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips dropdown content controls", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "Gold",
            contentControl: {
              alias: "Plan",
              tag: "plan",
              dropdown: {
                items: [
                  { displayText: "Silver", value: "silver" },
                  { displayText: "Gold", value: "gold" },
                ],
              },
            },
          },
        ],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips combo box content controls", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "Custom",
            contentControl: {
              alias: "Choice",
              tag: "choice",
              comboBox: {
                items: [
                  { displayText: "Standard", value: "standard" },
                  { displayText: "Custom", value: "custom" },
                ],
              },
            },
          },
        ],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips date content controls", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "2026-07-04",
            contentControl: {
              alias: "Due Date",
              tag: "dueDate",
              date: { fullDate: "2026-07-04T00:00:00Z", format: "yyyy-MM-dd" },
            },
          },
        ],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips content control placeholders", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "Click or tap here",
            contentControl: {
              alias: "Recipient",
              tag: "recipient",
              placeholder: { docPart: "DefaultPlaceholder_22610170" },
            },
          },
        ],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips content control data bindings", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "Ada Lovelace",
            contentControl: {
              alias: "Customer",
              tag: "customer",
              dataBinding: {
                storeItemId: "{11111111-2222-3333-4444-555555555555}",
                xpath: "/customer/name[1]",
                prefixMappings: "xmlns:crm='urn:crm'",
              },
            },
          },
        ],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips content control appearance metadata", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "Click or tap here",
            contentControl: {
              alias: "Prompt",
              tag: "prompt",
              appearance: "tags",
              color: "2F5496",
              showingPlaceholder: true,
            },
          },
        ],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips custom xml parts", async () => {
    const source = {
      ...createDocumentJson([]),
      customXmlParts: [
        {
          path: "customXml/item1.xml",
          xml: "<customer><name>Ada Lovelace</name></customer>",
          properties: {
            path: "customXml/itemProps1.xml",
            storeItemId: "{11111111-2222-3333-4444-555555555555}",
          },
        },
      ],
    };

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed.customXmlParts).toEqual(source.customXmlParts);
  });

  it("round-trips custom xml schema refs", async () => {
    const source = {
      ...createDocumentJson([]),
      customXmlParts: [
        {
          path: "customXml/item1.xml",
          xml: "<customer><name>Ada Lovelace</name></customer>",
          properties: {
            path: "customXml/itemProps1.xml",
            storeItemId: "{11111111-2222-3333-4444-555555555555}",
            schemaRefs: ["urn:customer", "urn:crm"],
          },
        },
      ],
    };

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed.customXmlParts).toEqual(source.customXmlParts);
  });

  it("round-trips document core properties", async () => {
    const source = {
      ...createDocumentJson([]),
      properties: {
        core: {
          title: "Quarterly Report",
          subject: "Sales",
          creator: "Ada Lovelace",
          keywords: "sales,quarterly",
          description: "Executive summary",
          lastModifiedBy: "Grace Hopper",
          created: "2026-07-04T00:00:00Z",
          modified: "2026-07-04T01:00:00Z",
        },
      },
    };

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed.properties).toEqual(source.properties);
  });

  it("round-trips document app properties", async () => {
    const source = {
      ...createDocumentJson([]),
      properties: {
        app: {
          application: "word2json",
          company: "ACME",
          manager: "Mira",
          pages: 3,
          words: 1200,
          characters: 6400,
          lines: 80,
          paragraphs: 12,
        },
      },
    };

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed.properties).toEqual(source.properties);
  });

  it("round-trips custom document properties", async () => {
    const source = {
      ...createDocumentJson([]),
      properties: {
        custom: [
          { name: "ContractId", type: "string", value: "C-2026-001" },
          { name: "RiskScore", type: "number", value: 42 },
          { name: "Approved", type: "boolean", value: true },
          { name: "EffectiveDate", type: "date", value: "2026-07-04T00:00:00Z" },
        ],
      },
    };

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed.properties).toEqual(source.properties);
  });

  it("round-trips web settings", async () => {
    const source = {
      ...createDocumentJson([]),
      settings: {
        web: {
          optimizeForBrowser: true,
          allowPng: true,
          doNotSaveAsSingleFile: true,
          pixelsPerInch: 120,
        },
      },
    };

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed.settings).toEqual(source.settings);
  });

  it("round-trips font table", async () => {
    const source = {
      ...createDocumentJson([]),
      fonts: [
        { name: "Aptos", family: "swiss", pitch: "variable", charset: "00", panose1: "020F0502020204030204" },
        { name: "SimSun", family: "roman", pitch: "fixed", charset: "86" },
      ],
    };

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed.fonts).toEqual(source.fonts);
  });

  it("round-trips repeating section content controls", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        contentControl: {
          alias: "Line Items",
          tag: "lineItems",
          repeatingSection: { sectionTitle: "Item", doNotAllowInsertDeleteSection: true },
          repeatingSectionItem: { id: "{11111111-2222-3333-4444-555555555555}" },
        },
        runs: [{ text: "Widget" }],
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

  it("round-trips document settings", async () => {
    const source = {
      version: "1.0" as const,
      settings: { defaultTabStop: 720, evenAndOddHeaders: true, updateFields: true },
      sections: [{ blocks: [{ type: "paragraph" as const, runs: [{ text: "Settings" }] }] }],
    };

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips track revisions setting", async () => {
    const source = {
      version: "1.0" as const,
      settings: { trackRevisions: true },
      sections: [{ blocks: [{ type: "paragraph" as const, runs: [{ text: "Tracked" }] }] }],
    };

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips compatibility settings", async () => {
    const source = {
      version: "1.0" as const,
      settings: {
        compatibility: {
          compatMode: "15",
          settings: [
            { name: "overrideTableStyleFontSizeAndJustification", uri: "http://schemas.microsoft.com/office/word", value: "1" },
            { name: "useWord2013TrackBottomHyphenation", uri: "http://schemas.microsoft.com/office/word", value: "0" },
          ],
        },
      },
      sections: [{ blocks: [{ type: "paragraph" as const, runs: [{ text: "Compat" }] }] }],
    };

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips proofing settings", async () => {
    const source = {
      version: "1.0" as const,
      settings: {
        proofing: {
          spelling: "clean" as const,
          grammar: "dirty" as const,
          doNotHyphenateCaps: true,
          hyphenationZone: 360,
        },
      },
      sections: [{ blocks: [{ type: "paragraph" as const, runs: [{ text: "Proofing" }] }] }],
    };

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips view settings", async () => {
    const source = {
      version: "1.0" as const,
      settings: {
        view: {
          mode: "print" as const,
          zoom: { preset: "fullPage" as const, percent: 125 },
        },
      },
      sections: [{ blocks: [{ type: "paragraph" as const, runs: [{ text: "View" }] }] }],
    };

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips document protection settings", async () => {
    const source = {
      version: "1.0" as const,
      settings: {
        protection: {
          edit: "trackedChanges" as const,
          enforcement: true,
          cryptProviderType: "rsaFull",
          cryptAlgorithmClass: "hash",
          cryptAlgorithmType: "typeAny",
          cryptAlgorithmSid: 4,
          cryptSpinCount: 100000,
          hash: "ABCDEF0123456789",
          salt: "0123456789ABCDEF",
        },
      },
      sections: [{ blocks: [{ type: "paragraph" as const, runs: [{ text: "Protected" }] }] }],
    };

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips mail merge settings", async () => {
    const source = {
      version: "1.0" as const,
      settings: {
        mailMerge: {
          mainDocumentType: "formLetters",
          dataType: "native",
          connectString: "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=contacts.xlsx;",
          query: "SELECT * FROM `Contacts$`",
          viewMergedData: true,
          activeRecord: 3,
          checkErrors: 1,
        },
      },
      sections: [{ blocks: [{ type: "paragraph" as const, runs: [{ text: "Mail merge" }] }] }],
    };

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips write protection settings", async () => {
    const source = {
      version: "1.0" as const,
      settings: {
        writeProtection: {
          recommended: true,
          cryptProviderType: "rsaFull",
          cryptAlgorithmClass: "hash",
          cryptAlgorithmType: "typeAny",
          cryptAlgorithmSid: 4,
          cryptSpinCount: 100000,
          hash: "FEDCBA9876543210",
          salt: "0011223344556677",
        },
      },
      sections: [{ blocks: [{ type: "paragraph" as const, runs: [{ text: "Write protected" }] }] }],
    };

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips math settings", async () => {
    const source = {
      version: "1.0" as const,
      settings: {
        math: {
          mathFont: "Cambria Math",
          breakBinary: "before" as const,
          smallFraction: true,
          displayDefaults: true,
        },
      },
      sections: [{ blocks: [{ type: "paragraph" as const, runs: [{ text: "Math settings" }] }] }],
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

  it("round-trips detailed table borders", async () => {
    const source = createDocumentJson([
      {
        type: "table",
        borders: {
          top: { style: "double", size: 12, color: "4472C4", space: 2 },
          left: { style: "single", size: 8, color: "70AD47", space: 0 },
          bottom: { style: "dashed", size: 6, color: "C00000", space: 1 },
          right: { style: "dotted", size: 4, color: "7030A0", space: 0 },
          insideH: { style: "single", size: 4, color: "808080", space: 0 },
          insideV: { style: "nil" },
        },
        rows: [
          { cells: [{ blocks: [{ type: "paragraph", runs: [{ text: "Detailed borders" }] }] }] },
        ],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips table alignment and cell spacing", async () => {
    const source = createDocumentJson([
      {
        type: "table",
        alignment: "center",
        cellSpacing: 120,
        rows: [
          { cells: [{ blocks: [{ type: "paragraph", runs: [{ text: "Centered" }] }] }] },
        ],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips table indentation", async () => {
    const source = createDocumentJson([
      {
        type: "table",
        indent: { width: 720 },
        rows: [
          { cells: [{ blocks: [{ type: "paragraph", runs: [{ text: "Indented" }] }] }] },
        ],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips floating table positioning", async () => {
    const source = createDocumentJson([
      {
        type: "table",
        position: {
          horizontalAnchor: "margin",
          verticalAnchor: "page",
          x: 720,
          y: 1440,
          leftFromText: 180,
          rightFromText: 180,
          topFromText: 120,
          bottomFromText: 120,
        },
        rows: [
          { cells: [{ blocks: [{ type: "paragraph", runs: [{ text: "Floating" }] }] }] },
        ],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips floating table overlap control", async () => {
    const source = createDocumentJson([
      {
        type: "table",
        overlap: "never",
        rows: [
          { cells: [{ blocks: [{ type: "paragraph", runs: [{ text: "No overlap" }] }] }] },
        ],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips table caption and description", async () => {
    const source = createDocumentJson([
      {
        type: "table",
        caption: "Quarterly revenue",
        description: "Revenue by region and quarter",
        rows: [
          { cells: [{ blocks: [{ type: "paragraph", runs: [{ text: "Revenue" }] }] }] },
        ],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips percentage table width", async () => {
    const source = createDocumentJson([
      {
        type: "table",
        width: 5000,
        widthType: "pct",
        rows: [
          { cells: [{ blocks: [{ type: "paragraph", runs: [{ text: "Full width" }] }] }] },
        ],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips table style look flags", async () => {
    const source = createDocumentJson([
      {
        type: "table",
        styleId: "TableGrid",
        look: {
          firstRow: true,
          lastRow: false,
          firstColumn: true,
          lastColumn: false,
          bandedRows: true,
          bandedColumns: false,
        },
        rows: [
          { cells: [{ blocks: [{ type: "paragraph", runs: [{ text: "Styled" }] }] }] },
        ],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips fixed table layout", async () => {
    const source = createDocumentJson([
      {
        type: "table",
        layout: "fixed",
        rows: [
          { cells: [{ blocks: [{ type: "paragraph", runs: [{ text: "Fixed width" }] }] }] },
        ],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips table grid and row height", async () => {
    const source = createDocumentJson([
      {
        type: "table",
        grid: [2400, 3600],
        rows: [
          {
            height: { value: 480, rule: "exact" as const },
            cells: [
              { blocks: [{ type: "paragraph", runs: [{ text: "A" }] }] },
              { blocks: [{ type: "paragraph", runs: [{ text: "B" }] }] },
            ],
          },
        ],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips row table property exceptions", async () => {
    const source = createDocumentJson([
      {
        type: "table",
        rows: [
          {
            propertyExceptions: {
              width: 5000,
              widthType: "pct",
              cellSpacing: 120,
              indent: { width: 360 },
              layout: "fixed",
              look: {
                firstRow: true,
                lastRow: false,
                firstColumn: true,
                lastColumn: false,
                bandedRows: true,
                bandedColumns: false,
              },
            },
            cells: [{ blocks: [{ type: "paragraph", runs: [{ text: "Exception row" }] }] }],
          },
        ],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips repeating table header rows", async () => {
    const source = createDocumentJson([
      {
        type: "table",
        rows: [
          {
            repeatHeader: true,
            cells: [{ blocks: [{ type: "paragraph", runs: [{ text: "Header" }] }] }],
          },
          {
            cells: [{ blocks: [{ type: "paragraph", runs: [{ text: "Body" }] }] }],
          },
        ],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips table rows that cannot split across pages", async () => {
    const source = createDocumentJson([
      {
        type: "table",
        rows: [
          {
            cantSplit: true,
            cells: [{ blocks: [{ type: "paragraph", runs: [{ text: "Keep together" }] }] }],
          },
          {
            cells: [{ blocks: [{ type: "paragraph", runs: [{ text: "Next row" }] }] }],
          },
        ],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips cell vertical merge and alignment", async () => {
    const source = createDocumentJson([
      {
        type: "table",
        rows: [
          {
            cells: [
              {
                verticalMerge: "restart" as const,
                verticalAlignment: "center" as const,
                blocks: [{ type: "paragraph", runs: [{ text: "Merged" }] }],
              },
              { blocks: [{ type: "paragraph", runs: [{ text: "Top" }] }] },
            ],
          },
          {
            cells: [
              {
                verticalMerge: "continue" as const,
                blocks: [{ type: "paragraph", runs: [] }],
              },
              { blocks: [{ type: "paragraph", runs: [{ text: "Bottom" }] }] },
            ],
          },
        ],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips percentage cell width", async () => {
    const source = createDocumentJson([
      {
        type: "table",
        rows: [
          {
            cells: [
              {
                width: 2500,
                widthType: "pct",
                blocks: [{ type: "paragraph", runs: [{ text: "Half" }] }],
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

  it("round-trips cell shading and margins", async () => {
    const source = createDocumentJson([
      {
        type: "table",
        rows: [
          {
            cells: [
              {
                shading: { fill: "D9EAF7" },
                margins: { top: 120, right: 180, bottom: 120, left: 180 },
                blocks: [{ type: "paragraph", runs: [{ text: "Header" }] }],
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

  it("round-trips cell margin width types", async () => {
    const source = createDocumentJson([
      {
        type: "table",
        rows: [
          {
            cells: [
              {
                margins: {
                  top: { width: 240, type: "pct" },
                  right: 180,
                  bottom: { width: 0, type: "nil" },
                  left: 120,
                },
                blocks: [{ type: "paragraph", runs: [{ text: "Typed margins" }] }],
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

  it("round-trips cell borders", async () => {
    const source = createDocumentJson([
      {
        type: "table",
        rows: [
          {
            cells: [
              {
                borders: {
                  top: { style: "single", size: 8, color: "4472C4", space: 0 },
                  bottom: { style: "single", size: 8, color: "4472C4", space: 0 },
                },
                blocks: [{ type: "paragraph", runs: [{ text: "Bordered cell" }] }],
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

  it("round-trips detailed cell borders", async () => {
    const source = createDocumentJson([
      {
        type: "table",
        rows: [
          {
            cells: [
              {
                borders: {
                  top: { style: "double", size: 12, color: "4472C4", space: 2 },
                  left: { style: "dashed", size: 6, color: "70AD47", space: 1 },
                  bottom: { style: "dotted", size: 4, color: "C00000", space: 0 },
                  right: { style: "single", size: 8, color: "7030A0", space: 0 },
                  insideH: { style: "single", size: 4, color: "808080", space: 0 },
                  insideV: { style: "nil" },
                },
                blocks: [{ type: "paragraph", runs: [{ text: "Detailed cell borders" }] }],
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

  it("round-trips cell text direction", async () => {
    const source = createDocumentJson([
      {
        type: "table",
        rows: [
          {
            cells: [
              {
                textDirection: "btLr",
                blocks: [{ type: "paragraph", runs: [{ text: "Vertical" }] }],
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

  it("round-trips cell no-wrap and fit-text controls", async () => {
    const source = createDocumentJson([
      {
        type: "table",
        rows: [
          {
            cells: [
              {
                noWrap: true,
                fitText: true,
                blocks: [{ type: "paragraph", runs: [{ text: "Compact code" }] }],
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

  it("round-trips custom numbering definitions", async () => {
    const source = {
      version: "1.0" as const,
      numbering: {
        abstractNums: [
          {
            id: 10,
            levels: [
              { level: 0, format: "decimal" as const, text: "%1.", start: 1, left: 720, hanging: 360 },
              { level: 1, format: "lowerLetter" as const, text: "%2)", start: 1, left: 1440, hanging: 360 },
            ],
          },
        ],
        nums: [{ id: 10, abstractId: 10 }],
      },
      sections: [
        {
          blocks: [
            {
              type: "paragraph" as const,
              list: { type: "ordered" as const, level: 1, numberingId: 10 },
              runs: [{ text: "Nested clause" }],
            },
          ],
        },
      ],
    };

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips numbering level suffix restart and legal controls", async () => {
    const source = {
      version: "1.0" as const,
      numbering: {
        abstractNums: [
          {
            id: 20,
            levels: [
              {
                level: 0,
                format: "decimal" as const,
                text: "%1)",
                start: 3,
                suffix: "space" as const,
                restart: 2,
                legal: true,
                left: 720,
                hanging: 360,
              },
              {
                level: 1,
                format: "lowerLetter" as const,
                text: "%2.",
                start: 1,
                legal: false,
              },
            ],
          },
        ],
        nums: [{ id: 20, abstractId: 20 }],
      },
      sections: [{ blocks: [] }],
    };

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips numbering level style alignment and run properties", async () => {
    const source = {
      version: "1.0" as const,
      numbering: {
        abstractNums: [
          {
            id: 30,
            levels: [
              {
                level: 0,
                format: "decimal" as const,
                text: "%1.",
                start: 1,
                styleId: "LegalClause",
                alignment: "right" as const,
                left: 720,
                hanging: 360,
                run: {
                  bold: true,
                  italic: false,
                  fontFamily: "Aptos",
                  fontSize: 11,
                  color: "C00000",
                },
              },
            ],
          },
        ],
        nums: [{ id: 30, abstractId: 30 }],
      },
      sections: [{ blocks: [] }],
    };

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips abstract numbering identity and style links", async () => {
    const source = {
      version: "1.0" as const,
      numbering: {
        abstractNums: [
          {
            id: 40,
            nsid: "5E2A1C9B",
            multiLevelType: "hybridMultilevel" as const,
            templateCode: "03A54D6C",
            styleLink: "LegalList",
            numberingStyleLink: "LegalListNumbering",
            levels: [
              { level: 0, format: "decimal" as const, text: "%1.", start: 1 },
            ],
          },
        ],
        nums: [{ id: 40, abstractId: 40 }],
      },
      sections: [{ blocks: [] }],
    };

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips numbering instance level overrides", async () => {
    const source = {
      version: "1.0" as const,
      numbering: {
        abstractNums: [
          {
            id: 50,
            levels: [
              { level: 0, format: "decimal" as const, text: "%1.", start: 1 },
              { level: 1, format: "lowerLetter" as const, text: "%2)", start: 1 },
            ],
          },
        ],
        nums: [
          {
            id: 50,
            abstractId: 50,
            overrides: [
              { level: 0, start: 7 },
              {
                level: 1,
                definition: {
                  level: 1,
                  format: "upperRoman" as const,
                  text: "%2.",
                  start: 3,
                  left: 1440,
                  hanging: 360,
                },
              },
            ],
          },
        ],
      },
      sections: [{ blocks: [] }],
    };

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

  it("round-trips internal hyperlinks", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          { text: "Jump", link: { anchor: "Clause1" } },
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

  it("round-trips comment ids", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "Reviewed",
            comment: {
              id: 42,
              author: "Ada",
              initials: "AL",
              date: "2026-07-04T13:00:00.000Z",
              text: "Stable comment id.",
            },
          },
        ],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips multi-run comments", async () => {
    const comment = { id: 50, author: "Ada", text: "One range." };
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          { text: "First ", comment },
          { text: "second", bold: true, comment },
        ],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips cross-paragraph comments", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        commentRangeStart: { id: 60, author: "Ada", text: "Across paragraphs." },
        runs: [{ text: "First paragraph" }],
      },
      {
        type: "paragraph",
        commentRangeEnd: { id: 60 },
        runs: [{ text: "Second paragraph" }],
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

  it("round-trips image crop", async () => {
    const imageData = Buffer.from("fake-png").toString("base64");
    const source = createDocumentJson([
      {
        type: "image",
        data: imageData,
        contentType: "image/png",
        width: 120,
        height: 80,
        crop: { left: 1000, top: 2000, right: 3000, bottom: 4000 },
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips image rotation", async () => {
    const imageData = Buffer.from("fake-png").toString("base64");
    const source = createDocumentJson([
      {
        type: "image",
        data: imageData,
        contentType: "image/png",
        width: 120,
        height: 80,
        rotation: 15,
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips floating image layout", async () => {
    const imageData = Buffer.from("fake-png").toString("base64");
    const source = createDocumentJson([
      {
        type: "image",
        data: imageData,
        contentType: "image/png",
        width: 120,
        height: 80,
        floating: { wrap: "square", horizontalOffset: 1440, verticalOffset: 720 },
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips floating image wrap variants", async () => {
    const imageData = Buffer.from("fake-png").toString("base64");
    const source = createDocumentJson([
      {
        type: "image",
        data: imageData,
        contentType: "image/png",
        width: 120,
        height: 80,
        floating: { wrap: "topAndBottom", horizontalOffset: 1440, verticalOffset: 720 },
      },
      {
        type: "image",
        data: imageData,
        contentType: "image/png",
        width: 120,
        height: 80,
        floating: { wrap: "none", horizontalOffset: 2880, verticalOffset: 1440 },
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips floating image anchor references", async () => {
    const imageData = Buffer.from("fake-png").toString("base64");
    const source = createDocumentJson([
      {
        type: "image",
        data: imageData,
        contentType: "image/png",
        width: 120,
        height: 80,
        floating: {
          wrap: "square",
          horizontalOffset: 1440,
          verticalOffset: 720,
          horizontalRelativeFrom: "margin",
          verticalRelativeFrom: "paragraph",
        },
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips floating image alignment positioning", async () => {
    const imageData = Buffer.from("fake-png").toString("base64");
    const source = createDocumentJson([
      {
        type: "image",
        data: imageData,
        contentType: "image/png",
        width: 120,
        height: 80,
        floating: {
          wrap: "square",
          horizontalOffset: 0,
          verticalOffset: 0,
          horizontalRelativeFrom: "margin",
          horizontalAlign: "center",
          verticalAlign: "bottom",
        },
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips floating image distances and layout flags", async () => {
    const imageData = Buffer.from("fake-png").toString("base64");
    const source = createDocumentJson([
      {
        type: "image",
        data: imageData,
        contentType: "image/png",
        width: 120,
        height: 80,
        floating: {
          wrap: "square",
          horizontalOffset: 1440,
          verticalOffset: 720,
          distanceTop: 10,
          distanceBottom: 20,
          distanceLeft: 30,
          distanceRight: 40,
          behindDoc: true,
          allowOverlap: false,
          layoutInCell: false,
        },
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips floating image anchor controls", async () => {
    const imageData = Buffer.from("fake-png").toString("base64");
    const source = createDocumentJson([
      {
        type: "image",
        data: imageData,
        contentType: "image/png",
        width: 120,
        height: 80,
        floating: {
          wrap: "square",
          horizontalOffset: 1440,
          verticalOffset: 720,
          simplePosition: { x: 123, y: 456 },
          relativeHeight: 251659264,
          locked: true,
        },
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

  it("round-trips first and even page headers and footers", async () => {
    const source = {
      version: "1.0" as const,
      sections: [
        {
          titlePage: true,
          headers: {
            default: [{ type: "paragraph" as const, runs: [{ text: "Default header" }] }],
            first: [{ type: "paragraph" as const, runs: [{ text: "First header" }] }],
            even: [{ type: "paragraph" as const, runs: [{ text: "Even header" }] }],
          },
          footers: {
            default: [{ type: "paragraph" as const, runs: [{ text: "Default footer" }] }],
            first: [{ type: "paragraph" as const, runs: [{ text: "First footer" }] }],
            even: [{ type: "paragraph" as const, runs: [{ text: "Even footer" }] }],
          },
          blocks: [{ type: "paragraph" as const, runs: [{ text: "Body" }] }],
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

  it("round-trips multiple sections", async () => {
    const source = {
      version: "1.0" as const,
      sections: [
        {
          breakType: "nextPage" as const,
          page: {
            width: 12240,
            height: 15840,
            margins: { top: 1440, right: 1440, bottom: 1440, left: 1440, header: 720, footer: 720, gutter: 0 },
          },
          blocks: [{ type: "paragraph" as const, runs: [{ text: "Section one" }] }],
        },
        {
          page: {
            width: 16840,
            height: 11900,
            orientation: "landscape" as const,
            margins: { top: 720, right: 900, bottom: 720, left: 900, header: 360, footer: 360, gutter: 0 },
          },
          blocks: [{ type: "paragraph" as const, runs: [{ text: "Section two" }] }],
        },
      ],
    };

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips section page numbering settings", async () => {
    const source = {
      version: "1.0" as const,
      sections: [
        {
          pageNumbering: {
            start: 3,
            format: "lowerRoman" as const,
            chapterStyle: 1,
            chapterSeparator: "hyphen" as const,
          },
          blocks: [{ type: "paragraph" as const, runs: [{ text: "Preface" }] }],
        },
      ],
    };

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips section line numbering settings", async () => {
    const source = {
      version: "1.0" as const,
      sections: [
        {
          lineNumbering: {
            start: 5,
            countBy: 2,
            distance: 360,
            restart: "newPage" as const,
          },
          blocks: [{ type: "paragraph" as const, runs: [{ text: "Draft" }] }],
        },
      ],
    };

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips section document grid settings", async () => {
    const source = {
      version: "1.0" as const,
      sections: [
        {
          documentGrid: {
            type: "linesAndChars" as const,
            linePitch: 360,
            charSpace: 180,
          },
          blocks: [{ type: "paragraph" as const, runs: [{ text: "Grid" }] }],
        },
      ],
    };

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips section vertical alignment", async () => {
    const source = {
      version: "1.0" as const,
      sections: [
        {
          verticalAlignment: "both" as const,
          blocks: [{ type: "paragraph" as const, runs: [{ text: "Distributed" }] }],
        },
      ],
    };

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips section text direction", async () => {
    const source = {
      version: "1.0" as const,
      sections: [
        {
          textDirection: "btLr" as const,
          blocks: [{ type: "paragraph" as const, runs: [{ text: "Rotated" }] }],
        },
      ],
    };

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips section rtl gutter", async () => {
    const source = {
      version: "1.0" as const,
      sections: [
        {
          rtlGutter: true,
          blocks: [{ type: "paragraph" as const, runs: [{ text: "RTL gutter" }] }],
        },
      ],
    };

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips section bidi layout", async () => {
    const source = {
      version: "1.0" as const,
      sections: [
        {
          bidi: true,
          blocks: [{ type: "paragraph" as const, runs: [{ text: "Bidi section" }] }],
        },
      ],
    };

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips section no endnote", async () => {
    const source = {
      version: "1.0" as const,
      sections: [
        {
          noEndnote: true,
          blocks: [{ type: "paragraph" as const, runs: [{ text: "No endnote here" }] }],
        },
      ],
    };

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips section mirror margins", async () => {
    const source = {
      version: "1.0" as const,
      sections: [
        {
          mirrorMargins: true,
          blocks: [{ type: "paragraph" as const, runs: [{ text: "Booklet" }] }],
        },
      ],
    };

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips text watermarks", async () => {
    const source = {
      version: "1.0" as const,
      sections: [
        {
          watermark: {
            text: "IMPORT QA",
            color: "C0C0C0",
            opacity: 0.15,
            rotation: 315,
            fontFamily: "Calibri",
          },
          blocks: [{ type: "paragraph" as const, runs: [{ text: "Body" }] }],
        },
      ],
    };

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips section footnote and endnote properties", async () => {
    const source = {
      version: "1.0" as const,
      sections: [
        {
          footnoteProperties: {
            position: "beneathText" as const,
            numbering: { format: "lowerRoman" as const, start: 2, restart: "eachSect" as const },
          },
          endnoteProperties: {
            position: "sectEnd" as const,
            numbering: { format: "upperRoman" as const, start: 4, restart: "continuous" as const },
          },
          blocks: [{ type: "paragraph" as const, runs: [{ text: "Notes" }] }],
        },
      ],
    };

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips columns", async () => {
    const source = {
      version: "1.0" as const,
      sections: [
        {
          columns: { count: 2, space: 720 },
          blocks: [{ type: "paragraph" as const, runs: [{ text: "Columns" }] }],
        },
      ],
    };

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips pagination controls", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        pagination: { keepNext: true, keepLines: true, pageBreakBefore: true },
        runs: [{ text: "Controlled paragraph" }],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips explicit off paragraph pagination controls", async () => {
    const source = {
      version: "1.0" as const,
      styles: {
        paragraph: [{
          id: "LooseHeading",
          name: "Loose Heading",
          paragraph: {
            pagination: { keepNext: false, keepLines: false, pageBreakBefore: false },
          },
        }],
      },
      sections: [{
        blocks: [{
          type: "paragraph" as const,
          styleId: "LooseHeading",
          pagination: { keepNext: false, keepLines: false, pageBreakBefore: false },
          runs: [{ text: "Loose paragraph" }],
        }],
      }],
    };

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips extra paragraph pagination toggles", async () => {
    const source = {
      version: "1.0" as const,
      styles: {
        paragraph: [{
          id: "ControlledBody",
          name: "Controlled Body",
          paragraph: {
            pagination: { widowControl: false, suppressLineNumbers: true, suppressAutoHyphens: true },
          },
        }],
      },
      sections: [{
        blocks: [{
          type: "paragraph" as const,
          styleId: "ControlledBody",
          pagination: { widowControl: true, suppressLineNumbers: true, suppressAutoHyphens: false },
          runs: [{ text: "Controlled typography" }],
        }],
      }],
    };

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips paragraph layout toggles", async () => {
    const source = {
      version: "1.0" as const,
      styles: {
        paragraph: [{
          id: "LayoutBody",
          name: "Layout Body",
          paragraph: {
            pagination: { contextualSpacing: false, mirrorIndents: true, overflowPunct: false, topLinePunct: true },
          },
        }],
      },
      sections: [{
        blocks: [{
          type: "paragraph" as const,
          styleId: "LayoutBody",
          pagination: { contextualSpacing: true, mirrorIndents: false, overflowPunct: true, topLinePunct: false },
          runs: [{ text: "Layout-sensitive paragraph" }],
        }],
      }],
    };

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips paragraph text flow properties", async () => {
    const source = {
      version: "1.0" as const,
      styles: {
        paragraph: [{
          id: "VerticalBody",
          name: "Vertical Body",
          paragraph: {
            pagination: {
              textAlignment: "baseline",
              textDirection: "btLr",
              adjustRightInd: false,
              autoSpaceDE: true,
              autoSpaceDN: false,
            },
          },
        }],
      },
      sections: [{
        blocks: [{
          type: "paragraph" as const,
          styleId: "VerticalBody",
          pagination: {
            textAlignment: "center",
            textDirection: "tbRl",
            adjustRightInd: true,
            autoSpaceDE: false,
            autoSpaceDN: true,
          },
          runs: [{ text: "Vertical paragraph" }],
        }],
      }],
    };

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips paragraph frame properties", async () => {
    const source = {
      version: "1.0" as const,
      styles: {
        paragraph: [{
          id: "Sidebar",
          name: "Sidebar",
          paragraph: {
            frame: {
              width: 2160,
              horizontalAnchor: "page",
              verticalAnchor: "margin",
              xAlign: "right",
              yAlign: "bottom",
              wrap: "notBeside",
              anchorLock: false,
            },
          },
        }],
      },
      sections: [{
        blocks: [{
          type: "paragraph" as const,
          styleId: "Sidebar",
          frame: {
            width: 2880,
            height: 1440,
            x: 720,
            y: 360,
            horizontalAnchor: "margin",
            verticalAnchor: "page",
            xAlign: "center",
            yAlign: "top",
            wrap: "around",
            dropCap: "drop",
            lines: 3,
            anchorLock: true,
            heightRule: "exact",
          },
          runs: [{ text: "Framed paragraph" }],
        }],
      }],
    };

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips paragraph tab stops", async () => {
    const source = {
      version: "1.0" as const,
      styles: {
        paragraph: [{
          id: "Tabular",
          name: "Tabular",
          paragraph: {
            tabs: [
              { value: "center", position: 2160, leader: "hyphen" },
              { value: "clear", position: 3600 },
            ],
          },
        }],
      },
      sections: [{
        blocks: [{
          type: "paragraph" as const,
          styleId: "Tabular",
          tabs: [
            { value: "left", position: 720 },
            { value: "right", position: 4320, leader: "dot" },
          ],
          runs: [{ text: "Label\tValue" }],
        }],
      }],
    };

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips paragraph spacing", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        spacing: { before: 240, after: 120, line: 360, lineRule: "auto" as const },
        runs: [{ text: "Spaced paragraph" }],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips paragraph shading", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        shading: { fill: "FFF2CC" },
        runs: [{ text: "Highlighted paragraph" }],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips shading theme attributes", async () => {
    const source = {
      version: "1.0" as const,
      styles: {
        paragraph: [
          {
            id: "ThemedShade",
            name: "Themed Shade",
            paragraph: {
              shading: {
                fill: "E2F0D9",
                themeFill: "accent3",
                themeFillTint: "66",
                themeColor: "accent1",
                themeTint: "33",
              },
            },
          },
        ],
      },
      sections: [
        {
          blocks: [
            {
              type: "paragraph" as const,
              shading: {
                fill: "FFF2CC",
                themeFill: "accent2",
                themeFillShade: "80",
                color: "auto",
                themeColor: "text1",
                themeShade: "40",
              },
              runs: [{ text: "Themed paragraph shading" }],
            },
            {
              type: "table" as const,
              rows: [
                {
                  cells: [
                    {
                      shading: {
                        fill: "D9EAF7",
                        themeFill: "accent5",
                        themeFillTint: "99",
                        color: "auto",
                        themeColor: "background1",
                      },
                      blocks: [{ type: "paragraph" as const, runs: [{ text: "Themed cell shading" }] }],
                    },
                  ],
                },
              ],
            },
            { type: "paragraph" as const, styleId: "ThemedShade", runs: [{ text: "Styled shading" }] },
          ],
        },
      ],
    };

    const docx = await buildDocx(source);
    const zip = await JSZip.loadAsync(docx);
    const documentXml = await zip.file("word/document.xml")!.async("string");
    const stylesXml = await zip.file("word/styles.xml")!.async("string");
    const parsed = await parseDocx(docx);

    expect(documentXml).toContain('<w:shd w:val="clear" w:color="auto" w:fill="FFF2CC" w:themeFill="accent2" w:themeFillShade="80" w:themeColor="text1" w:themeShade="40"/>');
    expect(documentXml).toContain('<w:shd w:val="clear" w:color="auto" w:fill="D9EAF7" w:themeFill="accent5" w:themeFillTint="99" w:themeColor="background1"/>');
    expect(stylesXml).toContain('<w:shd w:val="clear" w:fill="E2F0D9" w:themeFill="accent3" w:themeFillTint="66" w:themeColor="accent1" w:themeTint="33"/>');
    expect(parsed).toEqual(source);
  });

  it("round-trips paragraph borders", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        borders: {
          top: { style: "single", size: 8, color: "4472C4", space: 2 },
          bottom: { style: "single", size: 8, color: "4472C4", space: 2 },
        },
        runs: [{ text: "Bordered paragraph" }],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips paragraph between and bar borders", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        borders: {
          between: { style: "dashed", size: 6, color: "70AD47", space: 1 },
          bar: { style: "double", size: 12, color: "C00000", space: 0 },
        },
        runs: [{ text: "Paragraph side borders" }],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips additional border styles", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        borders: {
          top: { style: "thick", size: 18, color: "4472C4", space: 2 },
          bottom: { style: "dotDash", size: 8, color: "70AD47", space: 1 },
          bar: { style: "wave", size: 12, color: "C00000", space: 0 },
        },
        runs: [{ text: "Additional paragraph border styles" }],
      },
      {
        type: "table",
        borders: {
          top: { style: "triple", size: 12, color: "7030A0", space: 1 },
          insideH: { style: "dotDotDash", size: 6, color: "808080", space: 0 },
        },
        rows: [
          {
            cells: [
              {
                borders: {
                  left: { style: "thick", size: 16, color: "4472C4", space: 2 },
                  right: { style: "wave", size: 10, color: "C00000", space: 0 },
                  insideV: { style: "dotDash", size: 4, color: "70AD47", space: 1 },
                },
                blocks: [{ type: "paragraph", runs: [{ text: "Additional cell border styles" }] }],
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

  it("round-trips border theme color attributes", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        borders: {
          top: { style: "single", size: 8, color: "4472C4", themeColor: "accent1", themeTint: "66", themeShade: "33" },
        },
        runs: [
          {
            text: "Themed run border",
            border: { style: "double", size: 6, color: "70AD47", themeColor: "accent2", themeTint: "99" },
          },
        ],
      },
      {
        type: "table",
        borders: {
          top: { style: "single", size: 4, color: "7030A0", themeColor: "accent4", themeShade: "80" },
        },
        rows: [
          {
            cells: [
              {
                borders: {
                  bottom: { style: "dashed", size: 8, color: "C00000", themeColor: "accent6", themeTint: "40" },
                },
                blocks: [{ type: "paragraph", runs: [{ text: "Themed cell border" }] }],
              },
            ],
          },
        ],
      },
    ]);

    const docx = await buildDocx(source);
    const zip = await JSZip.loadAsync(docx);
    const xml = await zip.file("word/document.xml")!.async("string");
    const parsed = await parseDocx(docx);

    expect(xml).toContain('w:themeColor="accent1" w:themeTint="66" w:themeShade="33"');
    expect(xml).toContain('w:themeColor="accent2" w:themeTint="99"');
    expect(xml).toContain('w:themeColor="accent4" w:themeShade="80"');
    expect(xml).toContain('w:themeColor="accent6" w:themeTint="40"');
    expect(parsed).toEqual(source);
  });

  it("round-trips paragraph indentation", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        indent: { left: 720, right: 360, firstLine: 240, hanging: 120 },
        runs: [{ text: "Indented paragraph" }],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips footnotes", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          { text: "Body" },
          { text: "", footnote: { blocks: [{ type: "paragraph", runs: [{ text: "Footnote text" }] }] } },
        ],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips endnotes", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          { text: "Body" },
          { text: "", endnote: { blocks: [{ type: "paragraph", runs: [{ text: "Endnote text" }] }] } },
        ],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips reference fields", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          { text: "See " },
          { text: "", field: { type: "ref", target: "Clause1" } },
          { text: " on page " },
          { text: "", field: { type: "pageRef", target: "Clause1" } },
        ],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips field result text", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          { text: "", field: { type: "page", result: "3" } },
        ],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips table of contents field", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          { text: "", field: { type: "toc", switches: 'o "1-3" h z u', result: "Table of Contents" } },
        ],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips paragraph style definitions", async () => {
    const source = {
      version: "1.0" as const,
      styles: {
        paragraph: [
          { id: "ContractTitle", name: "Contract Title", basedOn: "Normal", next: "Normal" },
        ],
      },
      sections: [
        {
          blocks: [
            { type: "paragraph" as const, styleId: "ContractTitle", runs: [{ text: "Agreement" }] },
          ],
        },
      ],
    };

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips run document defaults", async () => {
    const source = {
      version: "1.0" as const,
      styles: {
        defaults: {
          run: { fontFamily: "Aptos", fontSize: 11, color: "1F1F1F" },
        },
      },
      sections: [{ blocks: [{ type: "paragraph" as const, runs: [{ text: "Defaults" }] }] }],
    };

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips paragraph document defaults", async () => {
    const source = {
      version: "1.0" as const,
      styles: {
        defaults: {
          paragraph: {
            spacing: { after: 160, line: 276, lineRule: "auto" as const },
            indent: { firstLine: 420 },
          },
        },
      },
      sections: [{ blocks: [{ type: "paragraph" as const, runs: [{ text: "Paragraph defaults" }] }] }],
    };

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips paragraph style properties", async () => {
    const source = {
      version: "1.0" as const,
      styles: {
        paragraph: [
          {
            id: "ContractTitle",
            name: "Contract Title",
            basedOn: "Normal",
            next: "Normal",
            paragraph: { alignment: "center" as const },
            run: { bold: true, fontFamily: "Aptos Display", fontSize: 18, color: "1F4E79" },
          },
        ],
      },
      sections: [
        {
          blocks: [
            { type: "paragraph" as const, styleId: "ContractTitle", runs: [{ text: "Agreement" }] },
          ],
        },
      ],
    };

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips paragraph style pagination properties", async () => {
    const source = {
      version: "1.0" as const,
      styles: {
        paragraph: [{
          id: "KeepHeading",
          name: "Keep Heading",
          paragraph: {
            pagination: { keepNext: true, keepLines: true, pageBreakBefore: true },
          },
        }],
      },
      sections: [{ blocks: [{ type: "paragraph" as const, styleId: "KeepHeading", runs: [{ text: "Heading" }] }] }],
    };

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips paragraph style layout properties", async () => {
    const source = {
      version: "1.0" as const,
      styles: {
        paragraph: [
          {
            id: "BodyText",
            name: "Body Text",
            paragraph: {
              spacing: { before: 120, after: 120 },
              indent: { left: 360, hanging: 180 },
            },
          },
        ],
      },
      sections: [
        {
          blocks: [
            { type: "paragraph" as const, styleId: "BodyText", runs: [{ text: "Styled body" }] },
          ],
        },
      ],
    };

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips paragraph style borders and shading", async () => {
    const source = {
      version: "1.0" as const,
      styles: {
        paragraph: [
          {
            id: "Callout",
            name: "Callout",
            paragraph: {
              shading: { fill: "E2F0D9" },
              borders: { left: { style: "single", size: 12, color: "70AD47", space: 4 } },
            },
          },
        ],
      },
      sections: [
        {
          blocks: [{ type: "paragraph" as const, styleId: "Callout", runs: [{ text: "Styled callout" }] }],
        },
      ],
    };

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips character and table styles", async () => {
    const source = {
      version: "1.0" as const,
      styles: {
        character: [
          { id: "DefinedTerm", name: "Defined Term", basedOn: "DefaultParagraphFont" },
        ],
        table: [
          { id: "ContractTable", name: "Contract Table", basedOn: "TableNormal" },
        ],
      },
      sections: [
        {
          blocks: [
            {
              type: "paragraph" as const,
              runs: [{ text: "Term", styleId: "DefinedTerm" }],
            },
            {
              type: "table" as const,
              styleId: "ContractTable",
              rows: [
                {
                  cells: [
                    { blocks: [{ type: "paragraph" as const, runs: [{ text: "Cell" }] }] },
                  ],
                },
              ],
            },
          ],
        },
      ],
    };

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips character style properties", async () => {
    const source = {
      version: "1.0" as const,
      styles: {
        character: [
          {
            id: "DefinedTerm",
            name: "Defined Term",
            basedOn: "DefaultParagraphFont",
            run: { italic: true, underline: true, color: "C00000" },
          },
        ],
      },
      sections: [
        {
          blocks: [
            {
              type: "paragraph" as const,
              runs: [{ text: "Term", styleId: "DefinedTerm" }],
            },
          ],
        },
      ],
    };

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips character style underline none", async () => {
    const source = {
      version: "1.0" as const,
      styles: {
        character: [
          {
            id: "PlainTerm",
            name: "Plain Term",
            basedOn: "DefaultParagraphFont",
            run: { underline: false },
          },
        ],
      },
      sections: [
        {
          blocks: [
            {
              type: "paragraph" as const,
              runs: [{ text: "Term", styleId: "PlainTerm" }],
            },
          ],
        },
      ],
    };

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips style run fonts with direct theme and hint attributes", async () => {
    const source = {
      version: "1.0" as const,
      styles: {
        character: [{
          id: "DefinedTerm",
          name: "Defined Term",
          run: {
            fontFamily: "Aptos",
            eastAsiaFontFamily: "SimSun",
            complexScriptFontFamily: "Arial",
            fontTheme: "majorHAnsi",
            eastAsiaFontTheme: "majorEastAsia",
            complexScriptFontTheme: "majorBidi",
            fontHint: "eastAsia",
          },
        }],
      },
      sections: [
        {
          blocks: [
            {
              type: "paragraph" as const,
              runs: [{ text: "Term", styleId: "DefinedTerm" }],
            },
          ],
        },
      ],
    };

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips style run advanced formatting", async () => {
    const source = {
      version: "1.0" as const,
      styles: {
        character: [
          {
            id: "WarningText",
            name: "Warning Text",
            run: { highlight: "yellow" as const, strike: true, verticalAlign: "superscript" as const, characterSpacing: 20, scale: 90 },
          },
        ],
      },
      sections: [
        {
          blocks: [
            {
              type: "paragraph" as const,
              runs: [{ text: "Warning", styleId: "WarningText" }],
            },
          ],
        },
      ],
    };

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips style run advanced typography properties", async () => {
    const source = {
      version: "1.0" as const,
      styles: {
        character: [{
          id: "EmphasisText",
          name: "Emphasis Text",
          run: {
            complexScriptFontSize: 14,
            language: { value: "en-US", eastAsia: "zh-CN", bidi: "ar-SA" },
            characterPosition: 4,
            kerning: 28,
          },
        }],
      },
      sections: [{ blocks: [{ type: "paragraph" as const, runs: [{ text: "Term", styleId: "EmphasisText" }] }] }],
    };

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips style run boolean effect properties", async () => {
    const source = {
      version: "1.0" as const,
      styles: {
        character: [{
          id: "HiddenEffectText",
          name: "Hidden Effect Text",
          run: {
            shadow: true,
            outline: true,
            emboss: true,
            imprint: true,
            rtl: true,
            complexScript: true,
            specVanish: true,
            hidden: true,
            webHidden: true,
            snapToGrid: true,
            noProof: true,
            officeMath: true,
          },
        }],
      },
      sections: [{ blocks: [{ type: "paragraph" as const, runs: [{ text: "Hidden", styleId: "HiddenEffectText" }] }] }],
    };

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips style run fit text and emphasis properties", async () => {
    const source = {
      version: "1.0" as const,
      styles: {
        character: [{
          id: "CompressedText",
          name: "Compressed Text",
          run: {
            fitText: { width: 1440, id: 11 },
            emphasis: "underDot" as const,
          },
        }],
      },
      sections: [{ blocks: [{ type: "paragraph" as const, runs: [{ text: "Fit", styleId: "CompressedText" }] }] }],
    };

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips style run explicit off boolean properties", async () => {
    const source = {
      version: "1.0" as const,
      styles: {
        character: [{
          id: "ExplicitOffText",
          name: "Explicit Off Text",
          run: {
            bold: false,
            italic: false,
            strike: false,
            doubleStrike: false,
            smallCaps: false,
            allCaps: false,
            shadow: false,
            outline: false,
            emboss: false,
            imprint: false,
            rtl: false,
            complexScript: false,
            specVanish: false,
            hidden: false,
            webHidden: false,
            snapToGrid: false,
            noProof: false,
            officeMath: false,
          },
        }],
      },
      sections: [{ blocks: [{ type: "paragraph" as const, runs: [{ text: "Off", styleId: "ExplicitOffText" }] }] }],
    };

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips table style properties", async () => {
    const source = {
      version: "1.0" as const,
      styles: {
        table: [
          {
            id: "ContractTable",
            name: "Contract Table",
            basedOn: "TableNormal",
            table: { borders: "single" as const },
          },
        ],
      },
      sections: [
        {
          blocks: [
            {
              type: "table" as const,
              styleId: "ContractTable",
              rows: [
                {
                  cells: [
                    { blocks: [{ type: "paragraph" as const, runs: [{ text: "Cell" }] }] },
                  ],
                },
              ],
            },
          ],
        },
      ],
    };

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips theme part", async () => {
    const source = {
      version: "1.0" as const,
      theme: {
        name: "Contract Theme",
        fonts: { major: "Aptos Display", minor: "Aptos" },
        colors: { ...defaultThemeColors(), accent1: "4472C4" },
      },
      sections: [
        {
          blocks: [{ type: "paragraph" as const, runs: [{ text: "Themed" }] }],
        },
      ],
    };

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips rich theme color scheme", async () => {
    const source = {
      version: "1.0" as const,
      theme: {
        name: "Contract Theme",
        fonts: { major: "Aptos Display", minor: "Aptos" },
        colors: {
          dark1: "000000",
          light1: "FFFFFF",
          dark2: "1F2937",
          light2: "F8FAFC",
          accent1: "4472C4",
          accent2: "ED7D31",
          accent3: "A5A5A5",
          accent4: "FFC000",
          accent5: "5B9BD5",
          accent6: "70AD47",
          hyperlink: "0563C1",
          followedHyperlink: "954F72",
        },
      },
      sections: [
        {
          blocks: [{ type: "paragraph" as const, runs: [{ text: "Themed" }] }],
        },
      ],
    };

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips theme script fonts", async () => {
    const source = {
      version: "1.0" as const,
      theme: {
        name: "Multilingual Theme",
        fonts: {
          major: "Aptos Display",
          minor: "Aptos",
          majorEastAsia: "SimSun",
          majorComplexScript: "Arial",
          minorEastAsia: "Microsoft YaHei",
          minorComplexScript: "Arial",
        },
        colors: { ...defaultThemeColors(), accent1: "4472C4" },
      },
      sections: [
        {
          blocks: [{ type: "paragraph" as const, runs: [{ text: "Themed" }] }],
        },
      ],
    };

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips theme supplemental fonts", async () => {
    const source = {
      version: "1.0" as const,
      theme: {
        name: "Multilingual Theme",
        fonts: {
          major: "Aptos Display",
          minor: "Aptos",
          supplemental: [
            { group: "major" as const, script: "Hans", typeface: "SimSun" },
            { group: "major" as const, script: "Jpan", typeface: "Yu Gothic" },
            { group: "minor" as const, script: "Hans", typeface: "Microsoft YaHei" },
            { group: "minor" as const, script: "Hang", typeface: "Malgun Gothic" },
          ],
        },
        colors: { ...defaultThemeColors(), accent1: "4472C4" },
      },
      sections: [
        {
          blocks: [{ type: "paragraph" as const, runs: [{ text: "Themed" }] }],
        },
      ],
    };

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips theme format scheme", async () => {
    const source = {
      version: "1.0" as const,
      theme: {
        name: "Visual Theme",
        fonts: { major: "Aptos Display", minor: "Aptos" },
        colors: { ...defaultThemeColors(), accent1: "4472C4" },
        formatScheme: {
          name: "Visual Formats",
          fillStyleColors: ["FFFFFF", "F2F2F2", "FFFFFF"],
          lineStyleColors: ["4472C4", "70AD47", "000000"],
          effectStyleColors: ["808080", "000000", "808080"],
          backgroundFillStyleColors: ["000000", "1F2937", "FFFFFF"],
        },
      },
      sections: [
        {
          blocks: [{ type: "paragraph" as const, runs: [{ text: "Themed" }] }],
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
