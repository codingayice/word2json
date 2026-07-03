import { XMLParser } from "fast-xml-parser";
import JSZip from "jszip";
import type { DocumentJson, ParagraphAlignment, ParagraphNode, TextRun } from "./schema.js";

type XmlNode = Record<string, unknown>;

const parser = new XMLParser({
  attributeNamePrefix: "",
  ignoreAttributes: false,
  removeNSPrefix: true,
});

export async function parseDocx(buffer: Buffer | Uint8Array): Promise<DocumentJson> {
  const zip = await JSZip.loadAsync(buffer);
  const documentFile = zip.file("word/document.xml");

  if (!documentFile) {
    throw new Error("Invalid DOCX package: word/document.xml is missing.");
  }

  const xml = await documentFile.async("string");
  const parsed = parser.parse(xml) as XmlNode;
  const documentNode = asObject(parsed.document);
  const body = asObject(documentNode.body);
  const paragraphs = asArray(body.p);

  return {
    version: "1.0",
    sections: [
      {
        blocks: paragraphs.map(parseParagraph),
      },
    ],
  };
}

function parseParagraph(value: unknown): ParagraphNode {
  const paragraph = asObject(value);
  const properties = asObject(paragraph.pPr);
  const alignmentNode = asObject(properties.jc);
  const alignment = typeof alignmentNode.val === "string"
    ? (alignmentNode.val as ParagraphAlignment)
    : undefined;

  return {
    type: "paragraph",
    ...(alignment ? { alignment } : {}),
    runs: asArray(paragraph.r).map(parseRun),
  };
}

function parseRun(value: unknown): TextRun {
  const run = asObject(value);
  const properties = asObject(run.rPr);

  return {
    text: parseText(run.t),
    ...(properties.b !== undefined ? { bold: true } : {}),
    ...(properties.i !== undefined ? { italic: true } : {}),
    ...(properties.u !== undefined ? { underline: true } : {}),
  };
}

function parseText(value: unknown): string {
  if (typeof value === "string") {
    return value;
  }

  const node = asObject(value);
  const text = node["#text"];
  return typeof text === "string" ? text : "";
}

function asArray(value: unknown): unknown[] {
  if (value === undefined) {
    return [];
  }

  return Array.isArray(value) ? value : [value];
}

function asObject(value: unknown): XmlNode {
  return value && typeof value === "object" && !Array.isArray(value)
    ? (value as XmlNode)
    : {};
}
