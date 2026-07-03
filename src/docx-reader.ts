import { XMLParser } from "fast-xml-parser";
import JSZip from "jszip";
import type {
  DocumentJson,
  PageSettings,
  ParagraphAlignment,
  ParagraphNode,
  ParagraphStyle,
  TableCellNode,
  TableNode,
  TextRun,
} from "./schema.js";

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
  const page = parsePageSettings(body.sectPr);
  const section = {
    ...(page ? { page } : {}),
    blocks: extractBlockXml(xml).map(parseBlockXml),
  };

  return {
    version: "1.0",
    sections: [section],
  };
}

function parseBlockXml(xml: string): ParagraphNode | TableNode {
  const parsed = parser.parse(xml) as XmlNode;

  if (parsed.p !== undefined) {
    return parseParagraph(parsed.p);
  }

  return parseTable(parsed.tbl);
}

function extractBlockXml(xml: string): string[] {
  const bodyMatch = xml.match(/<w:body>([\s\S]*?)<w:sectPr>/);
  const body = bodyMatch?.[1] ?? "";
  const blocks: string[] = [];
  let index = 0;

  while (index < body.length) {
    const paragraphIndex = body.indexOf("<w:p", index);
    const tableIndex = body.indexOf("<w:tbl", index);
    const blockIndex = nextBlockIndex(paragraphIndex, tableIndex);

    if (blockIndex === -1) {
      break;
    }

    if (blockIndex === paragraphIndex) {
      const end = body.indexOf("</w:p>", blockIndex);
      blocks.push(body.slice(blockIndex, end + "</w:p>".length));
      index = end + "</w:p>".length;
    } else {
      const end = body.indexOf("</w:tbl>", blockIndex);
      blocks.push(body.slice(blockIndex, end + "</w:tbl>".length));
      index = end + "</w:tbl>".length;
    }
  }

  return blocks;
}

function nextBlockIndex(paragraphIndex: number, tableIndex: number): number {
  if (paragraphIndex === -1) {
    return tableIndex;
  }

  if (tableIndex === -1) {
    return paragraphIndex;
  }

  return Math.min(paragraphIndex, tableIndex);
}

function parsePageSettings(value: unknown): PageSettings | undefined {
  const sectionProperties = asObject(value);
  const size = asObject(sectionProperties.pgSz);
  const margins = asObject(sectionProperties.pgMar);

  if (size.w === undefined || size.h === undefined || margins.top === undefined) {
    return undefined;
  }

  const page = {
    width: parseNumber(size.w),
    height: parseNumber(size.h),
    ...(typeof size.orient === "string" ? { orientation: size.orient as PageSettings["orientation"] } : {}),
    margins: {
      top: parseNumber(margins.top),
      right: parseNumber(margins.right),
      bottom: parseNumber(margins.bottom),
      left: parseNumber(margins.left),
      header: parseNumber(margins.header),
      footer: parseNumber(margins.footer),
      gutter: parseNumber(margins.gutter),
    },
  };

  return isDefaultPageSettings(page) ? undefined : page;
}

function parseParagraph(value: unknown): ParagraphNode {
  const paragraph = asObject(value);
  const properties = asObject(paragraph.pPr);
  const styleNode = asObject(properties.pStyle);
  const alignmentNode = asObject(properties.jc);
  const numbering = parseListSettings(properties.numPr);
  const style = typeof styleNode.val === "string"
    ? paragraphStyleFromId(styleNode.val)
    : undefined;
  const alignment = typeof alignmentNode.val === "string"
    ? (alignmentNode.val as ParagraphAlignment)
    : undefined;

  return {
    type: "paragraph",
    ...(style ? { style } : {}),
    ...(alignment ? { alignment } : {}),
    ...(numbering ? { list: numbering } : {}),
    runs: asArray(paragraph.r).map(parseRun),
  };
}

function parseListSettings(value: unknown): ParagraphNode["list"] | undefined {
  const numbering = asObject(value);
  const level = asObject(numbering.ilvl);
  const numId = asObject(numbering.numId);

  if (level.val === undefined || numId.val === undefined) {
    return undefined;
  }

  return {
    type: parseNumber(numId.val) === 1 ? "bullet" : "ordered",
    level: parseNumber(level.val),
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
    ...parseRunFont(properties),
  };
}

function parseRunFont(properties: XmlNode): Partial<TextRun> {
  const fonts = asObject(properties.rFonts);
  const size = asObject(properties.sz);
  const color = asObject(properties.color);

  return {
    ...(typeof fonts.ascii === "string" ? { fontFamily: fonts.ascii } : {}),
    ...(typeof size.val === "number" ? { fontSize: size.val / 2 } : {}),
    ...(typeof size.val === "string" ? { fontSize: Number.parseInt(size.val, 10) / 2 } : {}),
    ...(typeof color.val === "string" ? { color: color.val } : {}),
  };
}

function parseTable(value: unknown): TableNode {
  const table = asObject(value);
  const properties = asObject(table.tblPr);
  const width = asObject(properties.tblW);
  const borders = asObject(properties.tblBorders);

  return {
    type: "table",
    ...(width.w !== undefined ? { width: parseNumber(width.w) } : {}),
    ...(borders.top !== undefined ? { borders: "single" as const } : {}),
    rows: asArray(table.tr).map((rowValue) => {
      const row = asObject(rowValue);

      return {
        cells: asArray(row.tc).map(parseTableCell),
      };
    }),
  };
}

function parseTableCell(value: unknown): TableCellNode {
  const cell = asObject(value);
  const properties = asObject(cell.tcPr);
  const width = asObject(properties.tcW);
  const gridSpan = asObject(properties.gridSpan);

  return {
    ...(width.w !== undefined ? { width: parseNumber(width.w) } : {}),
    ...(gridSpan.val !== undefined ? { colSpan: parseNumber(gridSpan.val) } : {}),
    blocks: asArray(cell.p).map(parseParagraph),
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

function paragraphStyleFromId(styleId: string): ParagraphStyle | undefined {
  if (styleId === "Heading1") {
    return "heading1";
  }

  if (styleId === "Heading2") {
    return "heading2";
  }

  if (styleId === "Heading3") {
    return "heading3";
  }

  if (styleId === "Normal") {
    return "normal";
  }

  return undefined;
}

function parseNumber(value: unknown): number {
  if (typeof value === "number") {
    return value;
  }

  if (typeof value === "string") {
    return Number.parseInt(value, 10);
  }

  return 0;
}

function isDefaultPageSettings(page: PageSettings): boolean {
  return page.width === 12240 &&
    page.height === 15840 &&
    page.orientation === undefined &&
    page.margins.top === 1440 &&
    page.margins.right === 1440 &&
    page.margins.bottom === 1440 &&
    page.margins.left === 1440 &&
    page.margins.header === 720 &&
    page.margins.footer === 720 &&
    page.margins.gutter === 0;
}
