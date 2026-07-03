import { XMLParser } from "fast-xml-parser";
import JSZip from "jszip";
import type {
  DocumentJson,
  ImageNode,
  PageSettings,
  ParagraphAlignment,
  ParagraphNode,
  ParagraphStyle,
  TableCellNode,
  TableNode,
  TextRun,
} from "./schema.js";

type XmlNode = Record<string, unknown>;
type RelationshipMap = Record<string, string>;
type CommentMap = Record<string, TextRun["comment"]>;
type MediaMap = Record<string, Pick<ImageNode, "data" | "contentType">>;

const parser = new XMLParser({
  attributeNamePrefix: "",
  ignoreAttributes: false,
  removeNSPrefix: true,
  trimValues: false,
});

export async function parseDocx(buffer: Buffer | Uint8Array): Promise<DocumentJson> {
  const zip = await JSZip.loadAsync(buffer);
  const documentFile = zip.file("word/document.xml");

  if (!documentFile) {
    throw new Error("Invalid DOCX package: word/document.xml is missing.");
  }

  const xml = await documentFile.async("string");
  const relationships = await parseDocumentRelationships(zip);
  const media = await parseMedia(zip, relationships);
  const comments = await parseComments(zip);
  const parsed = parser.parse(xml) as XmlNode;
  const documentNode = asObject(parsed.document);
  const body = asObject(documentNode.body);
  const page = parsePageSettings(body.sectPr);
  const headerFooter = await parseHeaderFooterContent(zip, body.sectPr, relationships, comments);
  const section = {
    ...(page ? { page } : {}),
    ...headerFooter,
    blocks: extractBlockXml(xml).map((blockXml) => parseBlockXml(blockXml, relationships, comments, media)),
  };

  return {
    version: "1.0",
    sections: [section],
  };
}

async function parseHeaderFooterContent(
  zip: JSZip,
  sectionPropertiesValue: unknown,
  relationships: RelationshipMap,
  comments: CommentMap,
): Promise<Pick<import("./schema.js").SectionNode, "headers" | "footers">> {
  const sectionProperties = asObject(sectionPropertiesValue);
  const headerReference = asObject(sectionProperties.headerReference);
  const footerReference = asObject(sectionProperties.footerReference);
  const headers = await parseHeaderFooterReference(zip, headerReference, relationships, comments);
  const footers = await parseHeaderFooterReference(zip, footerReference, relationships, comments);

  return {
    ...(headers ? { headers: { default: headers } } : {}),
    ...(footers ? { footers: { default: footers } } : {}),
  };
}

async function parseHeaderFooterReference(
  zip: JSZip,
  reference: XmlNode,
  relationships: RelationshipMap,
  comments: CommentMap,
): Promise<ParagraphNode[] | undefined> {
  if (typeof reference.id !== "string") {
    return undefined;
  }

  const target = relationships[reference.id];
  if (!target) {
    return undefined;
  }

  const file = zip.file(`word/${target}`);
  if (!file) {
    return undefined;
  }

  const xml = await file.async("string");
  const parsed = parser.parse(xml) as XmlNode;
  const root = asObject(parsed.hdr ?? parsed.ftr);

  return asArray(root.p).map((paragraph) => parseParagraph(paragraph, relationships, comments));
}

async function parseMedia(zip: JSZip, relationships: RelationshipMap): Promise<MediaMap> {
  const entries = await Promise.all(Object.entries(relationships)
    .filter(([, target]) => target.startsWith("media/"))
    .map(async ([id, target]) => {
      const file = zip.file(`word/${target}`);
      const data = file ? (await file.async("nodebuffer")).toString("base64") : "";
      const contentType = target.endsWith(".png") ? "image/png" : "image/jpeg";

      return [id, { data, contentType }] as const;
    }));

  return Object.fromEntries(entries);
}

async function parseDocumentRelationships(zip: JSZip): Promise<RelationshipMap> {
  const relsFile = zip.file("word/_rels/document.xml.rels");

  if (!relsFile) {
    return {};
  }

  const xml = await relsFile.async("string");
  const parsed = parser.parse(xml) as XmlNode;
  const relationships = asArray(asObject(parsed.Relationships).Relationship);

  return Object.fromEntries(relationships
    .map((value) => asObject(value))
    .filter((relationship) => relationship.Id !== undefined && relationship.Target !== undefined)
    .map((relationship) => [String(relationship.Id), String(relationship.Target)]));
}

async function parseComments(zip: JSZip): Promise<CommentMap> {
  const commentsFile = zip.file("word/comments.xml");

  if (!commentsFile) {
    return {};
  }

  const xml = await commentsFile.async("string");
  const parsed = parser.parse(xml) as XmlNode;
  const comments = asArray(asObject(parsed.comments).comment);

  return Object.fromEntries(comments.map((commentValue) => {
    const comment = asObject(commentValue);
    const paragraph = asObject(asObject(comment.p).r);

    return [String(comment.id), {
      author: String(comment.author ?? ""),
      ...(typeof comment.initials === "string" ? { initials: comment.initials } : {}),
      ...(typeof comment.date === "string" ? { date: comment.date } : {}),
      text: parseText(paragraph.t),
    }];
  }));
}

function parseBlockXml(
  xml: string,
  relationships: RelationshipMap,
  comments: CommentMap,
  media: MediaMap,
): ParagraphNode | TableNode | ImageNode {
  const parsed = parser.parse(xml) as XmlNode;

  if (parsed.p !== undefined) {
    if (xml.includes("<w:drawing>")) {
      return parseImageBlock(parsed.p, media);
    }

    return parseParagraph(parsed.p, relationships, comments);
  }

  return parseTable(parsed.tbl, relationships, comments);
}

function parseImageBlock(value: unknown, media: MediaMap): ImageNode {
  const paragraph = asObject(value);
  const run = asObject(paragraph.r);
  const drawing = asObject(run.drawing);
  const inline = asObject(drawing.inline);
  const extent = asObject(inline.extent);
  const docPr = asObject(inline.docPr);
  const relationshipId = parseImageRelationshipId(inline);
  const image = relationshipId ? media[relationshipId] : undefined;

  return {
    type: "image",
    data: image?.data ?? "",
    contentType: image?.contentType ?? "image/png",
    width: parseNumber(extent.cx) / 9525,
    height: parseNumber(extent.cy) / 9525,
    ...(typeof docPr.descr === "string" ? { altText: docPr.descr } : {}),
  };
}

function parseImageRelationshipId(inline: XmlNode): string | undefined {
  const graphic = asObject(inline.graphic);
  const graphicData = asObject(graphic.graphicData);
  const picture = asObject(graphicData.pic);
  const blipFill = asObject(picture.blipFill);
  const blip = asObject(blipFill.blip);

  return typeof blip.embed === "string" ? blip.embed : undefined;
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

function parseParagraph(value: unknown, relationships: RelationshipMap = {}, comments: CommentMap = {}): ParagraphNode {
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
    runs: parseParagraphRuns(paragraph, relationships, comments),
  };
}

function parseParagraphRuns(
  paragraph: XmlNode,
  relationships: RelationshipMap,
  comments: CommentMap,
): TextRun[] {
  const commentRanges = commentRangesByText(paragraph);
  const bookmarkRanges = bookmarkRangesByText(paragraph);
  const normalRuns = asArray(paragraph.r)
    .map((run) => parseRun(run))
    .filter((run) => run.text !== "" || run.break !== undefined || run.field !== undefined)
    .map((run) => withMatchingBookmark(run, bookmarkRanges))
    .map((run) => withMatchingComment(run, commentRanges, comments));
  const hyperlinkRuns = asArray(paragraph.hyperlink)
    .flatMap((hyperlink) => parseHyperlink(hyperlink, relationships, comments));

  return [...normalRuns, ...hyperlinkRuns];
}

function commentRangesByText(paragraph: XmlNode): Map<string, string> {
  const starts = asArray(paragraph.commentRangeStart)
    .map((value) => asObject(value))
    .filter((value) => value.id !== undefined);
  const ranges = new Map<string, string>();
  const runs = asArray(paragraph.r).map((run) => parseRun(run)).filter((run) => run.text !== "");

  starts.forEach((start, index) => {
    const run = runs[index];
    if (run) {
      ranges.set(run.text, String(start.id));
    }
  });

  return ranges;
}

function withMatchingComment(run: TextRun, ranges: Map<string, string>, comments: CommentMap): TextRun {
  const commentId = ranges.get(run.text);
  const comment = commentId ? comments[commentId] : undefined;

  return comment ? { ...run, comment } : run;
}

function bookmarkRangesByText(paragraph: XmlNode): Map<string, string> {
  const starts = asArray(paragraph.bookmarkStart)
    .map((value) => asObject(value))
    .filter((value) => value.name !== undefined);
  const ranges = new Map<string, string>();
  const runs = asArray(paragraph.r).map((run) => parseRun(run)).filter((run) => run.text !== "");

  starts.forEach((start, index) => {
    const run = runs[index];
    if (run) {
      ranges.set(run.text, String(start.name));
    }
  });

  return ranges;
}

function withMatchingBookmark(run: TextRun, ranges: Map<string, string>): TextRun {
  const name = ranges.get(run.text);

  return name ? { ...run, bookmark: { name } } : run;
}

function parseHyperlink(value: unknown, relationships: RelationshipMap, comments: CommentMap): TextRun[] {
  const hyperlink = asObject(value);
  const url = typeof hyperlink.id === "string" ? relationships[hyperlink.id] : undefined;
  const ranges = commentRangesByText(hyperlink);

  return asArray(hyperlink.r).map((run) => ({
    ...withMatchingComment(parseRun(run), ranges, comments),
    ...(url ? { link: { url } } : {}),
  })).filter((run) => run.text !== "");
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
  const breakNode = asObject(run.br);

  if (run.br !== undefined) {
    return {
      text: "",
      break: breakNode.type === "page" ? "page" : "line",
    };
  }

  if (run.instrText !== undefined) {
    return {
      text: "",
      field: parseField(run.instrText),
    };
  }

  return {
    text: parseText(run.t),
    ...(properties.b !== undefined ? { bold: true } : {}),
    ...(properties.i !== undefined ? { italic: true } : {}),
    ...(properties.u !== undefined ? { underline: true } : {}),
    ...parseRunFont(properties),
  };
}

function parseField(value: unknown): TextRun["field"] {
  const instruction = parseText(value).trim();
  return instruction === "NUMPAGES" ? "numPages" : "page";
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

function parseTable(value: unknown, relationships: RelationshipMap, comments: CommentMap): TableNode {
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
        cells: asArray(row.tc).map((cell) => parseTableCell(cell, relationships, comments)),
      };
    }),
  };
}

function parseTableCell(value: unknown, relationships: RelationshipMap, comments: CommentMap): TableCellNode {
  const cell = asObject(value);
  const properties = asObject(cell.tcPr);
  const width = asObject(properties.tcW);
  const gridSpan = asObject(properties.gridSpan);

  return {
    ...(width.w !== undefined ? { width: parseNumber(width.w) } : {}),
    ...(gridSpan.val !== undefined ? { colSpan: parseNumber(gridSpan.val) } : {}),
    blocks: asArray(cell.p).map((paragraph) => parseParagraph(paragraph, relationships, comments)),
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
