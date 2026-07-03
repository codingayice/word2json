import { XMLParser } from "fast-xml-parser";
import JSZip from "jszip";
import type {
  DocumentJson,
  DocumentStyles,
  DocumentTheme,
  ImageNode,
  PageSettings,
  ParagraphAlignment,
  ParagraphNode,
  ParagraphStyle,
  StyleParagraphProperties,
  StyleRunProperties,
  TableStyleDefinition,
  TableCellNode,
  TableNode,
  TextRun,
} from "./schema.js";

type XmlNode = Record<string, unknown>;
type RelationshipMap = Record<string, string>;
type CommentMap = Record<string, TextRun["comment"]>;
type MediaMap = Record<string, Pick<ImageNode, "data" | "contentType">>;
type NoteMap = Record<string, NonNullable<TextRun["footnote"]>>;

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
  const footnotes = await parseNotes(zip, "footnotes", "footnote");
  const endnotes = await parseNotes(zip, "endnotes", "endnote");
  const styles = await parseStyles(zip);
  const theme = await parseTheme(zip);
  const parsed = parser.parse(xml) as XmlNode;
  const documentNode = asObject(parsed.document);
  const body = asObject(documentNode.body);

  return {
    version: "1.0",
    ...(theme ? { theme } : {}),
    ...(styles ? { styles } : {}),
    sections: await parseSections(zip, xml, body, relationships, comments, media, footnotes, endnotes),
  };
}

async function parseTheme(zip: JSZip): Promise<DocumentTheme | undefined> {
  const themeFile = zip.file("word/theme/theme1.xml");

  if (!themeFile) {
    return undefined;
  }

  const xml = await themeFile.async("string");
  const parsed = parser.parse(xml) as XmlNode;
  const theme = asObject(parsed.theme);
  const elements = asObject(theme.themeElements);
  const fontScheme = asObject(elements.fontScheme);
  const colorScheme = asObject(elements.clrScheme);
  const majorLatin = asObject(asObject(fontScheme.majorFont).latin);
  const minorLatin = asObject(asObject(fontScheme.minorFont).latin);
  const accent = asObject(asObject(colorScheme.accent1).srgbClr);

  if (typeof majorLatin.typeface !== "string" || typeof minorLatin.typeface !== "string" || typeof accent.val !== "string") {
    return undefined;
  }

  return {
    name: typeof theme.name === "string" ? theme.name : "Theme",
    fonts: {
      major: majorLatin.typeface,
      minor: minorLatin.typeface,
    },
    colors: {
      accent1: accent.val,
    },
  };
}

async function parseStyles(zip: JSZip): Promise<DocumentStyles | undefined> {
  const stylesFile = zip.file("word/styles.xml");

  if (!stylesFile) {
    return undefined;
  }

  const xml = await stylesFile.async("string");
  const parsed = parser.parse(xml) as XmlNode;
  const stylesRoot = asObject(parsed.styles);
  const styleNodes = asArray(stylesRoot.style).map((styleValue) => asObject(styleValue));
  const paragraph = styleNodes
    .filter((style) => style.type === "paragraph" && typeof style.styleId === "string" && !isBuiltInParagraphStyleId(style.styleId))
    .map((style) => {
      const name = asObject(style.name);
      const basedOn = asObject(style.basedOn);
      const next = asObject(style.next);
      const paragraph = parseStyleParagraphProperties(style.pPr);
      const run = parseStyleRunProperties(style.rPr);

      return {
        id: String(style.styleId),
        name: typeof name.val === "string" ? name.val : String(style.styleId),
        ...(typeof basedOn.val === "string" ? { basedOn: basedOn.val } : {}),
        ...(typeof next.val === "string" ? { next: next.val } : {}),
        ...(paragraph ? { paragraph } : {}),
        ...(run ? { run } : {}),
      };
    });
  const character = parseStyleDefinitions(styleNodes, "character");
  const table = parseTableStyleDefinitions(styleNodes);
  const styles: DocumentStyles = {
    ...(paragraph.length > 0 ? { paragraph } : {}),
    ...(character.length > 0 ? { character } : {}),
    ...(table.length > 0 ? { table } : {}),
  };

  return Object.keys(styles).length > 0 ? styles : undefined;
}

function parseStyleDefinitions(styleNodes: XmlNode[], type: "character" | "table"): NonNullable<DocumentStyles["character"]> {
  return styleNodes
    .filter((style) => style.type === type && typeof style.styleId === "string")
    .map((style) => {
      const name = asObject(style.name);
      const basedOn = asObject(style.basedOn);
      const run = parseStyleRunProperties(style.rPr);

      return {
        id: String(style.styleId),
        name: typeof name.val === "string" ? name.val : String(style.styleId),
        ...(typeof basedOn.val === "string" ? { basedOn: basedOn.val } : {}),
        ...(run ? { run } : {}),
      };
    });
}

function parseTableStyleDefinitions(styleNodes: XmlNode[]): TableStyleDefinition[] {
  return styleNodes
    .filter((style) => style.type === "table" && typeof style.styleId === "string")
    .map((style) => {
      const name = asObject(style.name);
      const basedOn = asObject(style.basedOn);
      const run = parseStyleRunProperties(style.rPr);
      const table = parseStyleTableProperties(style.tblPr);

      return {
        id: String(style.styleId),
        name: typeof name.val === "string" ? name.val : String(style.styleId),
        ...(typeof basedOn.val === "string" ? { basedOn: basedOn.val } : {}),
        ...(run ? { run } : {}),
        ...(table ? { table } : {}),
      };
    });
}

function parseStyleParagraphProperties(value: unknown): StyleParagraphProperties | undefined {
  const properties = asObject(value);
  const alignment = asObject(properties.jc);
  const parsed = {
    ...(typeof alignment.val === "string" ? { alignment: alignment.val as ParagraphAlignment } : {}),
  };

  return Object.keys(parsed).length > 0 ? parsed : undefined;
}

function parseStyleRunProperties(value: unknown): StyleRunProperties | undefined {
  const properties = asObject(value);
  const fonts = asObject(properties.rFonts);
  const size = asObject(properties.sz);
  const color = asObject(properties.color);
  const parsed = {
    ...(properties.b !== undefined ? { bold: true } : {}),
    ...(properties.i !== undefined ? { italic: true } : {}),
    ...(properties.u !== undefined ? { underline: true } : {}),
    ...(typeof fonts.ascii === "string" ? { fontFamily: fonts.ascii } : {}),
    ...(typeof size.val === "number" ? { fontSize: size.val / 2 } : {}),
    ...(typeof size.val === "string" ? { fontSize: Number.parseInt(size.val, 10) / 2 } : {}),
    ...(typeof color.val === "string" ? { color: color.val } : {}),
  };

  return Object.keys(parsed).length > 0 ? parsed : undefined;
}

function parseStyleTableProperties(value: unknown): TableStyleDefinition["table"] | undefined {
  const properties = asObject(value);
  const borders = asObject(properties.tblBorders);

  return borders.top !== undefined ? { borders: "single" } : undefined;
}

async function parseSections(
  zip: JSZip,
  documentXml: string,
  body: XmlNode,
  relationships: RelationshipMap,
  comments: CommentMap,
  media: MediaMap,
  footnotes: NoteMap,
  endnotes: NoteMap,
): Promise<DocumentJson["sections"]> {
  const bodyContent = documentXml.match(/<w:body>([\s\S]*?)<\/w:body>/)?.[1] ?? "";
  const breakPattern = /<w:p><w:pPr>(<w:sectPr>[\s\S]*?<\/w:sectPr>)<\/w:pPr><\/w:p>/g;
  const sectionParts: Array<{ content: string; sectPrXml: string; preserveDefaultPage: boolean }> = [];
  let cursor = 0;
  let match: RegExpExecArray | null;

  while ((match = breakPattern.exec(bodyContent)) !== null) {
    sectionParts.push({
      content: bodyContent.slice(cursor, match.index),
      sectPrXml: match[1],
      preserveDefaultPage: true,
    });
    cursor = match.index + match[0].length;
  }

  const remainingBody = bodyContent.slice(cursor);
  const finalSectionXml = remainingBody.match(/(<w:sectPr>[\s\S]*?<\/w:sectPr>)\s*$/)?.[1];

  if (finalSectionXml) {
    sectionParts.push({
      content: remainingBody.slice(0, remainingBody.lastIndexOf(finalSectionXml)),
      sectPrXml: finalSectionXml,
      preserveDefaultPage: sectionParts.length > 0,
    });
  }

  if (sectionParts.length === 0) {
    const page = parsePageSettings(body.sectPr, false);
    const headerFooter = await parseHeaderFooterContent(zip, body.sectPr, relationships, comments);
    return [{
      ...(page ? { page } : {}),
      ...headerFooter,
      blocks: extractBlockXml(documentXml).map((blockXml) => parseBlockXml(blockXml, relationships, comments, media, footnotes, endnotes)),
    }];
  }

  return Promise.all(sectionParts.map(async (part) => {
    const sectPr = parseSectPrXml(part.sectPrXml);
    const page = parsePageSettings(sectPr, part.preserveDefaultPage);
    const headerFooter = await parseHeaderFooterContent(zip, sectPr, relationships, comments);
    const breakType = parseSectionBreakType(sectPr);
    const columns = parseColumns(sectPr);

    return {
      ...(breakType ? { breakType } : {}),
      ...(page ? { page } : {}),
      ...headerFooter,
      ...(columns ? { columns } : {}),
      blocks: extractBlockXmlFromContent(part.content).map((blockXml) => parseBlockXml(blockXml, relationships, comments, media, footnotes, endnotes)),
    };
  }));
}

function parseSectPrXml(xml: string): XmlNode {
  const namespacedXml = xml.replace("<w:sectPr>", '<w:sectPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">');
  const parsed = parser.parse(namespacedXml) as XmlNode;
  return asObject(parsed.sectPr);
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

async function parseNotes(zip: JSZip, rootName: "footnotes" | "endnotes", itemName: "footnote" | "endnote"): Promise<NoteMap> {
  const file = zip.file(`word/${rootName}.xml`);

  if (!file) {
    return {};
  }

  const xml = await file.async("string");
  const parsed = parser.parse(xml) as XmlNode;
  const root = asObject(parsed[rootName]);
  const notes = asArray(root[itemName]);

  return Object.fromEntries(notes
    .map((noteValue) => asObject(noteValue))
    .filter((note) => note.id !== undefined && Number.parseInt(String(note.id), 10) > 0)
    .map((note) => [String(note.id), {
      blocks: asArray(note.p).map((paragraph) => parseParagraph(paragraph)),
    }]));
}

function parseBlockXml(
  xml: string,
  relationships: RelationshipMap,
  comments: CommentMap,
  media: MediaMap,
  footnotes: NoteMap,
  endnotes: NoteMap,
): ParagraphNode | TableNode | ImageNode {
  const parsed = parser.parse(xml) as XmlNode;

  if (parsed.p !== undefined) {
    if (xml.includes("<w:drawing>")) {
      return parseImageBlock(parsed.p, media);
    }

    return parseParagraph(parsed.p, relationships, comments, footnotes, endnotes);
  }

  return parseTable(parsed.tbl, relationships, comments, footnotes, endnotes);
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
  return extractBlockXmlFromContent(bodyMatch?.[1] ?? "");
}

function extractBlockXmlFromContent(body: string): string[] {
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

function parsePageSettings(value: unknown, preserveDefault: boolean): PageSettings | undefined {
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

  return !preserveDefault && isDefaultPageSettings(page) ? undefined : page;
}

function parseSectionBreakType(sectionPropertiesValue: unknown): import("./schema.js").SectionBreakType | undefined {
  const type = asObject(asObject(sectionPropertiesValue).type);

  if (type.val === "continuous") {
    return "continuous";
  }

  if (type.val === "evenPage") {
    return "evenPage";
  }

  if (type.val === "oddPage") {
    return "oddPage";
  }

  if (type.val === "nextPage") {
    return "nextPage";
  }

  return undefined;
}

function parseColumns(sectionPropertiesValue: unknown): import("./schema.js").ColumnSettings | undefined {
  const columns = asObject(asObject(sectionPropertiesValue).cols);

  if (columns.num === undefined) {
    return undefined;
  }

  return {
    count: parseNumber(columns.num),
    ...(columns.space !== undefined ? { space: parseNumber(columns.space) } : {}),
  };
}

function parseParagraph(
  value: unknown,
  relationships: RelationshipMap = {},
  comments: CommentMap = {},
  footnotes: NoteMap = {},
  endnotes: NoteMap = {},
): ParagraphNode {
  const paragraph = asObject(value);
  const properties = asObject(paragraph.pPr);
  const styleNode = asObject(properties.pStyle);
  const alignmentNode = asObject(properties.jc);
  const numbering = parseListSettings(properties.numPr);
  const pagination = parsePagination(properties);
  const style = typeof styleNode.val === "string"
    ? paragraphStyleFromId(styleNode.val)
    : undefined;
  const styleId = typeof styleNode.val === "string" && !style
    ? styleNode.val
    : undefined;
  const alignment = typeof alignmentNode.val === "string"
    ? (alignmentNode.val as ParagraphAlignment)
    : undefined;

  return {
    type: "paragraph",
    ...(style ? { style } : {}),
    ...(styleId ? { styleId } : {}),
    ...(alignment ? { alignment } : {}),
    ...(numbering ? { list: numbering } : {}),
    ...(pagination ? { pagination } : {}),
    runs: parseParagraphRuns(paragraph, relationships, comments, footnotes, endnotes),
  };
}

function parsePagination(properties: XmlNode): ParagraphNode["pagination"] | undefined {
  const pagination = {
    ...(properties.keepNext !== undefined ? { keepNext: true } : {}),
    ...(properties.keepLines !== undefined ? { keepLines: true } : {}),
    ...(properties.pageBreakBefore !== undefined ? { pageBreakBefore: true } : {}),
  };

  return Object.keys(pagination).length > 0 ? pagination : undefined;
}

function parseParagraphRuns(
  paragraph: XmlNode,
  relationships: RelationshipMap,
  comments: CommentMap,
  footnotes: NoteMap,
  endnotes: NoteMap,
): TextRun[] {
  const commentRanges = commentRangesByText(paragraph);
  const bookmarkRanges = bookmarkRangesByText(paragraph);
  const normalRuns = asArray(paragraph.r)
    .map((run) => parseRun(run))
    .filter((run) => run.text !== "" || run.break !== undefined || run.field !== undefined || run.footnote !== undefined || run.endnote !== undefined)
    .map((run) => withMatchingNotes(run, footnotes, endnotes))
    .map((run) => withMatchingBookmark(run, bookmarkRanges))
    .map((run) => withMatchingComment(run, commentRanges, comments));
  const hyperlinkRuns = asArray(paragraph.hyperlink)
    .flatMap((hyperlink) => parseHyperlink(hyperlink, relationships, comments));

  return [...normalRuns, ...hyperlinkRuns];
}

function withMatchingNotes(run: TextRun, footnotes: NoteMap, endnotes: NoteMap): TextRun {
  if (run.footnote && "id" in run.footnote) {
    const note = footnotes[String(run.footnote.id)];
    return note ? { text: "", footnote: note } : { text: "" };
  }

  if (run.endnote && "id" in run.endnote) {
    const note = endnotes[String(run.endnote.id)];
    return note ? { text: "", endnote: note } : { text: "" };
  }

  return run;
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

  const footnoteReference = asObject(run.footnoteReference);
  if (run.footnoteReference !== undefined) {
    return {
      text: "",
      footnote: { id: String(footnoteReference.id), blocks: [] } as unknown as TextRun["footnote"],
    };
  }

  const endnoteReference = asObject(run.endnoteReference);
  if (run.endnoteReference !== undefined) {
    return {
      text: "",
      endnote: { id: String(endnoteReference.id), blocks: [] } as unknown as TextRun["endnote"],
    };
  }

  return {
    text: parseText(run.t),
    ...parseRunStyle(properties),
    ...(properties.b !== undefined ? { bold: true } : {}),
    ...(properties.i !== undefined ? { italic: true } : {}),
    ...(properties.u !== undefined ? { underline: true } : {}),
    ...parseRunFont(properties),
  };
}

function parseRunStyle(properties: XmlNode): Partial<TextRun> {
  const style = asObject(properties.rStyle);

  return typeof style.val === "string" ? { styleId: style.val } : {};
}

function parseField(value: unknown): TextRun["field"] {
  const instruction = parseText(value).trim();

  if (instruction.startsWith("PAGEREF ")) {
    return { type: "pageRef", target: instruction.slice("PAGEREF ".length) };
  }

  if (instruction.startsWith("REF ")) {
    return { type: "ref", target: instruction.slice("REF ".length) };
  }

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

function parseTable(value: unknown, relationships: RelationshipMap, comments: CommentMap, footnotes: NoteMap, endnotes: NoteMap): TableNode {
  const table = asObject(value);
  const properties = asObject(table.tblPr);
  const style = asObject(properties.tblStyle);
  const width = asObject(properties.tblW);
  const borders = asObject(properties.tblBorders);

  return {
    type: "table",
    ...(typeof style.val === "string" ? { styleId: style.val } : {}),
    ...(width.w !== undefined ? { width: parseNumber(width.w) } : {}),
    ...(borders.top !== undefined ? { borders: "single" as const } : {}),
    rows: asArray(table.tr).map((rowValue) => {
      const row = asObject(rowValue);

      return {
        cells: asArray(row.tc).map((cell) => parseTableCell(cell, relationships, comments, footnotes, endnotes)),
      };
    }),
  };
}

function parseTableCell(value: unknown, relationships: RelationshipMap, comments: CommentMap, footnotes: NoteMap, endnotes: NoteMap): TableCellNode {
  const cell = asObject(value);
  const properties = asObject(cell.tcPr);
  const width = asObject(properties.tcW);
  const gridSpan = asObject(properties.gridSpan);

  return {
    ...(width.w !== undefined ? { width: parseNumber(width.w) } : {}),
    ...(gridSpan.val !== undefined ? { colSpan: parseNumber(gridSpan.val) } : {}),
    blocks: asArray(cell.p).map((paragraph) => parseParagraph(paragraph, relationships, comments, footnotes, endnotes)),
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

function isBuiltInParagraphStyleId(styleId: unknown): boolean {
  return styleId === "Normal" ||
    styleId === "Heading1" ||
    styleId === "Heading2" ||
    styleId === "Heading3";
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
