import JSZip from "jszip";
import type { DocumentJson, ParagraphNode, TextRun } from "./schema.js";

export async function buildDocx(document: DocumentJson): Promise<Buffer> {
  const zip = new JSZip();

  zip.file("[Content_Types].xml", contentTypesXml());
  zip.folder("_rels")!.file(".rels", packageRelsXml());
  zip.folder("word")!.file("document.xml", documentXml(document));
  zip.folder("word")!.file("styles.xml", stylesXml());
  zip.folder("word")!.folder("_rels")!.file("document.xml.rels", documentRelsXml());

  return zip.generateAsync({ type: "nodebuffer" });
}

function documentXml(document: DocumentJson): string {
  const body = document.sections
    .flatMap((section) => section.blocks)
    .map((block) => {
      if (block.type === "paragraph") {
        return paragraphXml(block);
      }

      return "";
    })
    .join("");

  return xmlDeclaration(
    `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
      `<w:body>${body}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body>` +
      `</w:document>`,
  );
}

function paragraphXml(paragraph: ParagraphNode): string {
  const properties = paragraph.alignment
    ? `<w:pPr><w:jc w:val="${paragraph.alignment}"/></w:pPr>`
    : "";
  const runs = paragraph.runs.map(runXml).join("");

  return `<w:p>${properties}${runs}</w:p>`;
}

function runXml(run: TextRun): string {
  const properties = runPropertiesXml(run);
  const textSpace = /^\s|\s$/.test(run.text) ? ' xml:space="preserve"' : "";

  return `<w:r>${properties}<w:t${textSpace}>${escapeXml(run.text)}</w:t></w:r>`;
}

function runPropertiesXml(run: TextRun): string {
  const properties = [
    run.bold ? "<w:b/>" : "",
    run.italic ? "<w:i/>" : "",
    run.underline ? '<w:u w:val="single"/>' : "",
  ].join("");

  return properties ? `<w:rPr>${properties}</w:rPr>` : "";
}

function contentTypesXml(): string {
  return xmlDeclaration(
    `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
      `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
      `<Default Extension="xml" ContentType="application/xml"/>` +
      `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>` +
      `<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>` +
      `</Types>`,
  );
}

function packageRelsXml(): string {
  return xmlDeclaration(
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>` +
      `</Relationships>`,
  );
}

function documentRelsXml(): string {
  return xmlDeclaration(
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`,
  );
}

function stylesXml(): string {
  return xmlDeclaration(
    `<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
      `<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>` +
      `</w:styles>`,
  );
}

function xmlDeclaration(xml: string): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${xml}`;
}

function escapeXml(value: string): string {
  return value
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&apos;");
}
