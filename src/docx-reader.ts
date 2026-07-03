import { XMLParser } from "fast-xml-parser";
import JSZip from "jszip";
import type {
  AbstractNumberingDefinition,
  DocumentJson,
  DocumentNumbering,
  DocumentStyles,
  DocumentTheme,
  ImageNode,
  PageSettings,
  ParagraphAlignment,
  ParagraphNode,
  ParagraphStyle,
  NumberingFormat,
  StyleParagraphProperties,
  StyleRunProperties,
  TableStyleDefinition,
  TableCellNode,
  TableNode,
  TextRun,
} from "./schema.js";

type XmlNode = Record<string, unknown>;
type RelationshipMap = Record<string, string>;
type RelationshipEntry = {
  id: string;
  type: string;
  target: string;
};
type CommentMap = Record<string, TextRun["comment"]>;
type MediaMap = Record<string, Pick<ImageNode, "data" | "contentType">>;
type NoteMap = Record<string, NonNullable<TextRun["footnote"]>>;
type NumberingContext = {
  numbering?: DocumentNumbering;
  listTypes: Map<number, "bullet" | "ordered">;
};

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
  const numberingContext = await parseNumbering(zip);
  const styles = await parseStyles(zip);
  const theme = await parseTheme(zip);
  const settings = await parseSettings(zip);
  const properties = await parseDocumentProperties(zip);
  const customXmlParts = await parseCustomXmlParts(zip);
  const parsed = parser.parse(xml) as XmlNode;
  const documentNode = asObject(parsed.document);
  const body = asObject(documentNode.body);

  return {
    version: "1.0",
    ...(settings ? { settings } : {}),
    ...(properties ? { properties } : {}),
    ...(theme ? { theme } : {}),
    ...(styles ? { styles } : {}),
    ...(numberingContext.numbering ? { numbering: numberingContext.numbering } : {}),
    ...(customXmlParts.length > 0 ? { customXmlParts } : {}),
    sections: await parseSections(zip, xml, body, relationships, comments, media, footnotes, endnotes, numberingContext),
  };
}

async function parseDocumentProperties(zip: JSZip): Promise<DocumentJson["properties"] | undefined> {
  const core = await parseCoreProperties(zip);
  const app = await parseAppProperties(zip);
  const custom = await parseCustomProperties(zip);
  const properties = {
    ...(core ? { core } : {}),
    ...(app ? { app } : {}),
    ...(custom.length > 0 ? { custom } : {}),
  };

  return Object.keys(properties).length > 0 ? properties : undefined;
}

async function parseCoreProperties(zip: JSZip): Promise<NonNullable<DocumentJson["properties"]>["core"] | undefined> {
  const coreFile = zip.file("docProps/core.xml");

  if (!coreFile) {
    return undefined;
  }

  const xml = await coreFile.async("string");
  const parsed = parser.parse(xml) as XmlNode;
  const coreProperties = asObject(parsed.coreProperties);
  const core = {
    ...(typeof coreProperties.title === "string" ? { title: coreProperties.title } : {}),
    ...(typeof coreProperties.subject === "string" ? { subject: coreProperties.subject } : {}),
    ...(typeof coreProperties.creator === "string" ? { creator: coreProperties.creator } : {}),
    ...(typeof coreProperties.keywords === "string" ? { keywords: coreProperties.keywords } : {}),
    ...(typeof coreProperties.description === "string" ? { description: coreProperties.description } : {}),
    ...(typeof coreProperties.lastModifiedBy === "string" ? { lastModifiedBy: coreProperties.lastModifiedBy } : {}),
    ...(xmlText(coreProperties.created) ? { created: xmlText(coreProperties.created) } : {}),
    ...(xmlText(coreProperties.modified) ? { modified: xmlText(coreProperties.modified) } : {}),
  };

  return Object.keys(core).length > 0 ? core : undefined;
}

async function parseAppProperties(zip: JSZip): Promise<NonNullable<DocumentJson["properties"]>["app"] | undefined> {
  const appFile = zip.file("docProps/app.xml");

  if (!appFile) {
    return undefined;
  }

  const xml = await appFile.async("string");
  const parsed = parser.parse(xml) as XmlNode;
  const appProperties = asObject(parsed.Properties);
  const app = {
    ...(typeof appProperties.Application === "string" ? { application: appProperties.Application } : {}),
    ...(typeof appProperties.Company === "string" ? { company: appProperties.Company } : {}),
    ...(typeof appProperties.Manager === "string" ? { manager: appProperties.Manager } : {}),
    ...(appProperties.Pages !== undefined ? { pages: parseNumber(appProperties.Pages) } : {}),
    ...(appProperties.Words !== undefined ? { words: parseNumber(appProperties.Words) } : {}),
    ...(appProperties.Characters !== undefined ? { characters: parseNumber(appProperties.Characters) } : {}),
    ...(appProperties.Lines !== undefined ? { lines: parseNumber(appProperties.Lines) } : {}),
    ...(appProperties.Paragraphs !== undefined ? { paragraphs: parseNumber(appProperties.Paragraphs) } : {}),
  };

  return Object.keys(app).length > 0 ? app : undefined;
}

async function parseCustomProperties(zip: JSZip): Promise<NonNullable<NonNullable<DocumentJson["properties"]>["custom"]>> {
  const customFile = zip.file("docProps/custom.xml");

  if (!customFile) {
    return [];
  }

  const xml = await customFile.async("string");
  const parsed = parser.parse(xml) as XmlNode;
  const properties = asArray(asObject(parsed.Properties).property)
    .map((property) => parseCustomProperty(asObject(property)))
    .filter((property): property is NonNullable<NonNullable<DocumentJson["properties"]>["custom"]>[number] => property !== undefined);

  return properties;
}

function parseCustomProperty(property: XmlNode): NonNullable<NonNullable<DocumentJson["properties"]>["custom"]>[number] | undefined {
  if (typeof property.name !== "string") {
    return undefined;
  }

  if (property.lpwstr !== undefined) {
    return { name: property.name, type: "string", value: String(property.lpwstr) };
  }

  if (property.i4 !== undefined) {
    return { name: property.name, type: "number", value: parseNumber(property.i4) };
  }

  if (property.bool !== undefined) {
    return { name: property.name, type: "boolean", value: property.bool === true || property.bool === "true" };
  }

  if (property.filetime !== undefined) {
    return { name: property.name, type: "date", value: String(property.filetime) };
  }

  return undefined;
}

async function parseSettings(zip: JSZip): Promise<DocumentJson["settings"] | undefined> {
  const settingsFile = zip.file("word/settings.xml");

  if (!settingsFile) {
    return undefined;
  }

  const xml = await settingsFile.async("string");
  const parsed = parser.parse(xml) as XmlNode;
  const settings = asObject(parsed.settings);
  const defaultTabStop = asObject(settings.defaultTabStop);
  const result = {
    ...(defaultTabStop.val !== undefined ? { defaultTabStop: parseNumber(defaultTabStop.val) } : {}),
    ...(settings.evenAndOddHeaders !== undefined ? { evenAndOddHeaders: true } : {}),
    ...(settings.updateFields !== undefined ? { updateFields: true } : {}),
    ...(settings.trackRevisions !== undefined ? { trackRevisions: true } : {}),
  };

  return Object.keys(result).length > 0 ? result : undefined;
}

async function parseNumbering(zip: JSZip): Promise<NumberingContext> {
  const numberingFile = zip.file("word/numbering.xml");
  const builtInListTypes = new Map<number, "bullet" | "ordered">([[1, "bullet"], [2, "ordered"]]);

  if (!numberingFile) {
    return { listTypes: builtInListTypes };
  }

  const xml = await numberingFile.async("string");
  const parsed = parser.parse(xml) as XmlNode;
  const numberingRoot = asObject(parsed.numbering);
  const abstractNums = asArray(numberingRoot.abstractNum)
    .map((abstractNumValue) => asObject(abstractNumValue))
    .filter((abstractNum) => abstractNum.abstractNumId !== undefined)
    .map((abstractNum) => parseAbstractNumberingDefinition(abstractNum));
  const nums = asArray(numberingRoot.num)
    .map((numValue) => asObject(numValue))
    .filter((num) => num.numId !== undefined)
    .map((num) => ({
      id: parseNumber(num.numId),
      abstractId: parseNumber(asObject(num.abstractNumId).val),
    }));
  const customAbstractNums = abstractNums.filter((abstractNum) => !isBuiltInNumberingId(abstractNum.id));
  const customNums = nums.filter((num) => !isBuiltInNumberingId(num.id));
  const abstractTypes = new Map(abstractNums.map((abstractNum) => [
    abstractNum.id,
    abstractNum.levels.some((level) => level.format === "bullet") ? "bullet" as const : "ordered" as const,
  ]));
  const listTypes = new Map(builtInListTypes);

  nums.forEach((num) => {
    listTypes.set(num.id, abstractTypes.get(num.abstractId) ?? "ordered");
  });

  return {
    listTypes,
    ...(customAbstractNums.length > 0 || customNums.length > 0
      ? { numbering: { abstractNums: customAbstractNums, nums: customNums } }
      : {}),
  };
}

function parseAbstractNumberingDefinition(value: XmlNode): AbstractNumberingDefinition {
  return {
    id: parseNumber(value.abstractNumId),
    levels: asArray(value.lvl).map((levelValue) => {
      const level = asObject(levelValue);
      const start = asObject(level.start);
      const format = asObject(level.numFmt);
      const text = asObject(level.lvlText);
      const indentation = asObject(asObject(level.pPr).ind);

      return {
        level: parseNumber(level.ilvl),
        format: parseNumberingFormat(format.val),
        text: typeof text.val === "string" ? text.val : "",
        ...(start.val !== undefined ? { start: parseNumber(start.val) } : {}),
        ...(indentation.left !== undefined ? { left: parseNumber(indentation.left) } : {}),
        ...(indentation.hanging !== undefined ? { hanging: parseNumber(indentation.hanging) } : {}),
      };
    }),
  };
}

function parseNumberingFormat(value: unknown): NumberingFormat {
  if (value === "bullet" || value === "decimal" || value === "lowerLetter" || value === "upperLetter" || value === "lowerRoman" || value === "upperRoman") {
    return value;
  }

  return "decimal";
}

function isBuiltInNumberingId(id: number): boolean {
  return id === 1 || id === 2;
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
  const defaults = parseDocDefaults(stylesRoot.docDefaults);
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
    ...(defaults ? { defaults } : {}),
    ...(paragraph.length > 0 ? { paragraph } : {}),
    ...(character.length > 0 ? { character } : {}),
    ...(table.length > 0 ? { table } : {}),
  };

  return Object.keys(styles).length > 0 ? styles : undefined;
}

function parseDocDefaults(value: unknown): DocumentStyles["defaults"] | undefined {
  const defaults = asObject(value);
  const run = parseStyleRunProperties(asObject(defaults.rPrDefault).rPr);
  const paragraph = parseStyleParagraphProperties(asObject(defaults.pPrDefault).pPr);
  const parsed = {
    ...(run ? { run } : {}),
    ...(paragraph ? { paragraph } : {}),
  };

  return Object.keys(parsed).length > 0 ? parsed : undefined;
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
  const spacing = parseParagraphSpacing(properties.spacing);
  const indent = parseParagraphIndent(properties.ind);
  const shading = parseShading(properties.shd);
  const borders = parseParagraphBorders(properties.pBdr);
  const parsed = {
    ...(typeof alignment.val === "string" ? { alignment: alignment.val as ParagraphAlignment } : {}),
    ...(spacing ? { spacing } : {}),
    ...(indent ? { indent } : {}),
    ...(shading ? { shading } : {}),
    ...(borders ? { borders } : {}),
  };

  return Object.keys(parsed).length > 0 ? parsed : undefined;
}

function parseStyleRunProperties(value: unknown): StyleRunProperties | undefined {
  const properties = asObject(value);
  const fonts = asObject(properties.rFonts);
  const size = asObject(properties.sz);
  const color = asObject(properties.color);
  const highlight = asObject(properties.highlight);
  const verticalAlign = asObject(properties.vertAlign);
  const characterSpacing = asObject(properties.spacing);
  const scale = asObject(properties.w);
  const border = parseBorder(properties.bdr);
  const parsed = {
    ...(properties.b !== undefined ? { bold: true } : {}),
    ...(properties.i !== undefined ? { italic: true } : {}),
    ...(properties.u !== undefined ? { underline: true } : {}),
    ...(typeof fonts.ascii === "string" ? { fontFamily: fonts.ascii } : {}),
    ...(typeof size.val === "number" ? { fontSize: size.val / 2 } : {}),
    ...(typeof size.val === "string" ? { fontSize: Number.parseInt(size.val, 10) / 2 } : {}),
    ...(typeof color.val === "string" ? { color: color.val } : {}),
    ...(typeof highlight.val === "string" ? { highlight: highlight.val as NonNullable<StyleRunProperties["highlight"]> } : {}),
    ...(properties.strike !== undefined ? { strike: true } : {}),
    ...(properties.dstrike !== undefined ? { doubleStrike: true } : {}),
    ...(properties.smallCaps !== undefined ? { smallCaps: true } : {}),
    ...(properties.caps !== undefined ? { allCaps: true } : {}),
    ...(typeof verticalAlign.val === "string" ? { verticalAlign: verticalAlign.val as NonNullable<StyleRunProperties["verticalAlign"]> } : {}),
    ...(characterSpacing.val !== undefined ? { characterSpacing: parseNumber(characterSpacing.val) } : {}),
    ...(scale.val !== undefined ? { scale: parseNumber(scale.val) } : {}),
    ...(border ? { border } : {}),
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
  numberingContext: NumberingContext,
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
      blocks: extractBlockXml(documentXml).map((blockXml) => parseBlockXml(blockXml, relationships, comments, media, footnotes, endnotes, numberingContext)),
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
      blocks: extractBlockXmlFromContent(part.content).map((blockXml) => parseBlockXml(blockXml, relationships, comments, media, footnotes, endnotes, numberingContext)),
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
  const relationships = await parseRelationships(zip, "word/_rels/document.xml.rels");

  return Object.fromEntries(relationships.map((relationship) => [relationship.id, relationship.target]));
}

async function parseRelationships(zip: JSZip, path: string): Promise<RelationshipEntry[]> {
  const relsFile = zip.file(path);

  if (!relsFile) {
    return [];
  }
  const xml = await relsFile.async("string");
  const parsed = parser.parse(xml) as XmlNode;
  const relationships = asArray(asObject(parsed.Relationships).Relationship);

  return relationships
    .map((value) => asObject(value))
    .filter((relationship) => relationship.Id !== undefined && relationship.Target !== undefined)
    .map((relationship) => ({
      id: String(relationship.Id),
      type: relationship.Type !== undefined ? String(relationship.Type) : "",
      target: String(relationship.Target),
    }));
}

async function parseCustomXmlParts(zip: JSZip): Promise<NonNullable<DocumentJson["customXmlParts"]>> {
  const packageRelationships = await parseRelationships(zip, "_rels/.rels");
  const customXmlRelationships = packageRelationships.filter((relationship) =>
    relationship.type === "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml" &&
    relationship.target.startsWith("customXml/"));
  const parts = await Promise.all(customXmlRelationships.map(async (relationship) => {
    const file = zip.file(relationship.target);
    const xml = file ? await file.async("string") : "";
    const properties = await parseCustomXmlPartProperties(zip, relationship.target);

    return {
      path: relationship.target,
      xml,
      ...(properties ? { properties } : {}),
    };
  }));

  return parts;
}

async function parseCustomXmlPartProperties(zip: JSZip, partPath: string): Promise<NonNullable<NonNullable<DocumentJson["customXmlParts"]>[number]["properties"]> | undefined> {
  const relationships = await parseRelationships(zip, `${pathDirname(partPath)}/_rels/${pathBasename(partPath)}.rels`);
  const propertiesRelationship = relationships.find((relationship) =>
    relationship.type === "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps");

  if (!propertiesRelationship) {
    return undefined;
  }

  const propertiesPath = `${pathDirname(partPath)}/${propertiesRelationship.target}`;
  const propertiesFile = zip.file(propertiesPath);
  const parsedProperties = {
    path: propertiesPath,
    ...(propertiesFile ? await parseCustomXmlPropertiesFile(propertiesFile) : {}),
  };

  return parsedProperties;
}

async function parseCustomXmlPropertiesFile(file: JSZip.JSZipObject): Promise<Pick<NonNullable<NonNullable<DocumentJson["customXmlParts"]>[number]["properties"]>, "storeItemId">> {
  const xml = await file.async("string");
  const parsed = parser.parse(xml) as XmlNode;
  const datastoreItem = asObject(parsed.datastoreItem);
  const schemaRefs = asArray(asObject(datastoreItem.schemaRefs).schemaRef)
    .map((schemaRef) => asObject(schemaRef).uri)
    .filter((uri): uri is string => typeof uri === "string");

  return {
    ...(typeof datastoreItem.itemID === "string" ? { storeItemId: datastoreItem.itemID } : {}),
    ...(schemaRefs.length > 0 ? { schemaRefs } : {}),
  };
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
    const id = parseNumber(comment.id);

    return [String(comment.id), {
      ...(id !== 0 ? { id } : {}),
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
  numberingContext: NumberingContext,
): ParagraphNode | TableNode | ImageNode {
  const parsed = parser.parse(xml) as XmlNode;

  if (parsed.ins !== undefined) {
    return parseRevisedParagraph(parsed.ins, "insert", relationships, comments, footnotes, endnotes, numberingContext);
  }

  if (parsed.del !== undefined) {
    return parseRevisedParagraph(parsed.del, "delete", relationships, comments, footnotes, endnotes, numberingContext);
  }

  if (parsed.moveFrom !== undefined) {
    return parseRevisedParagraph(parsed.moveFrom, "moveFrom", relationships, comments, footnotes, endnotes, numberingContext);
  }

  if (parsed.moveTo !== undefined) {
    return parseRevisedParagraph(parsed.moveTo, "moveTo", relationships, comments, footnotes, endnotes, numberingContext);
  }

  if (parsed.sdt !== undefined) {
    return parseSdtBlock(parsed.sdt, relationships, comments, footnotes, endnotes, numberingContext, xml);
  }

  if (parsed.p !== undefined) {
    if (xml.includes("<w:drawing>")) {
      return parseImageBlock(parsed.p, media);
    }

    const paragraph = parseParagraph(parsed.p, relationships, comments, footnotes, endnotes, numberingContext, xml);
    const commentRangeStart = parseParagraphCommentRangeStart(xml, comments);
    const commentRangeEnd = parseParagraphCommentRangeEnd(xml);
    return {
      ...paragraph,
      ...(commentRangeStart ? { commentRangeStart } : {}),
      ...(commentRangeEnd ? { commentRangeEnd } : {}),
    };
  }

  return parseTable(parsed.tbl, relationships, comments, footnotes, endnotes, numberingContext);
}

function parseSdtBlock(
  value: unknown,
  relationships: RelationshipMap,
  comments: CommentMap,
  footnotes: NoteMap,
  endnotes: NoteMap,
  numberingContext: NumberingContext,
  xml: string,
): ParagraphNode | TableNode | ImageNode {
  const sdt = asObject(value);
  const content = asObject(sdt.sdtContent);
  const contentControl = parseContentControl(sdt.sdtPr);
  const paragraph = parseParagraph(content.p, relationships, comments, footnotes, endnotes, numberingContext, xml.match(/<w:p\b[\s\S]*<\/w:p>/)?.[0]);

  return {
    ...paragraph,
    ...(contentControl ? { contentControl } : {}),
  };
}

function parseRevisedParagraph(
  value: unknown,
  type: NonNullable<ParagraphNode["revision"]>["type"],
  relationships: RelationshipMap,
  comments: CommentMap,
  footnotes: NoteMap,
  endnotes: NoteMap,
  numberingContext: NumberingContext,
): ParagraphNode {
  const revision = asObject(value);
  const paragraph = parseParagraph(revision.p, relationships, comments, footnotes, endnotes, numberingContext);

  return {
    ...paragraph,
    revision: {
      type,
      id: parseNumber(revision.id),
      author: String(revision.author ?? ""),
      ...(typeof revision.date === "string" ? { date: revision.date } : {}),
    },
  };
}

function parseImageBlock(value: unknown, media: MediaMap): ImageNode {
  const paragraph = asObject(value);
  const run = asObject(paragraph.r);
  const drawing = asObject(run.drawing);
  const container = imageDrawingContainer(drawing);
  const extent = asObject(container.extent);
  const docPr = asObject(container.docPr);
  const relationshipId = parseImageRelationshipId(container);
  const crop = parseImageCrop(container);
  const rotation = parseImageRotation(container);
  const floating = parseImageFloating(drawing);
  const image = relationshipId ? media[relationshipId] : undefined;

  return {
    type: "image",
    data: image?.data ?? "",
    contentType: image?.contentType ?? "image/png",
    width: parseNumber(extent.cx) / 9525,
    height: parseNumber(extent.cy) / 9525,
    ...(typeof docPr.descr === "string" ? { altText: docPr.descr } : {}),
    ...(crop ? { crop } : {}),
    ...(rotation !== undefined ? { rotation } : {}),
    ...(floating ? { floating } : {}),
  };
}

function imageDrawingContainer(drawing: XmlNode): XmlNode {
  return drawing.anchor !== undefined ? asObject(drawing.anchor) : asObject(drawing.inline);
}

function parseImageFloating(drawing: XmlNode): ImageNode["floating"] | undefined {
  const anchor = asObject(drawing.anchor);

  if (drawing.anchor === undefined) {
    return undefined;
  }

  return {
    wrap: "square",
    horizontalOffset: parseNumber(asObject(anchor.positionH).posOffset),
    verticalOffset: parseNumber(asObject(anchor.positionV).posOffset),
  };
}

function parseImageRotation(inline: XmlNode): number | undefined {
  const graphic = asObject(inline.graphic);
  const graphicData = asObject(graphic.graphicData);
  const picture = asObject(graphicData.pic);
  const shapeProperties = asObject(picture.spPr);
  const transform = asObject(shapeProperties.xfrm);

  return transform.rot !== undefined ? parseNumber(transform.rot) / 60000 : undefined;
}

function parseImageCrop(inline: XmlNode): ImageNode["crop"] | undefined {
  const graphic = asObject(inline.graphic);
  const graphicData = asObject(graphic.graphicData);
  const picture = asObject(graphicData.pic);
  const blipFill = asObject(picture.blipFill);
  const sourceRect = asObject(blipFill.srcRect);
  const crop = {
    ...(sourceRect.l !== undefined ? { left: parseNumber(sourceRect.l) } : {}),
    ...(sourceRect.t !== undefined ? { top: parseNumber(sourceRect.t) } : {}),
    ...(sourceRect.r !== undefined ? { right: parseNumber(sourceRect.r) } : {}),
    ...(sourceRect.b !== undefined ? { bottom: parseNumber(sourceRect.b) } : {}),
  };

  return Object.keys(crop).length > 0 ? crop : undefined;
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
    const sdtIndex = body.indexOf("<w:sdt", index);
    const insertIndex = body.indexOf("<w:ins", index);
    const deleteIndex = body.indexOf("<w:del", index);
    const moveFromIndex = body.indexOf("<w:moveFrom", index);
    const moveToIndex = body.indexOf("<w:moveTo", index);
    const blockIndex = nextBlockIndex(paragraphIndex, tableIndex, sdtIndex, insertIndex, deleteIndex, moveFromIndex, moveToIndex);

    if (blockIndex === -1) {
      break;
    }

    if (blockIndex === sdtIndex) {
      const end = body.indexOf("</w:sdt>", blockIndex);
      blocks.push(body.slice(blockIndex, end + "</w:sdt>".length));
      index = end + "</w:sdt>".length;
    } else if (blockIndex === insertIndex) {
      const end = body.indexOf("</w:ins>", blockIndex);
      blocks.push(body.slice(blockIndex, end + "</w:ins>".length));
      index = end + "</w:ins>".length;
    } else if (blockIndex === deleteIndex) {
      const end = body.indexOf("</w:del>", blockIndex);
      blocks.push(body.slice(blockIndex, end + "</w:del>".length));
      index = end + "</w:del>".length;
    } else if (blockIndex === moveFromIndex) {
      const end = body.indexOf("</w:moveFrom>", blockIndex);
      blocks.push(body.slice(blockIndex, end + "</w:moveFrom>".length));
      index = end + "</w:moveFrom>".length;
    } else if (blockIndex === moveToIndex) {
      const end = body.indexOf("</w:moveTo>", blockIndex);
      blocks.push(body.slice(blockIndex, end + "</w:moveTo>".length));
      index = end + "</w:moveTo>".length;
    } else if (blockIndex === paragraphIndex) {
      const end = body.indexOf("</w:p>", blockIndex);
      const start = paragraphBlockStart(body, index, blockIndex);
      const blockEnd = paragraphBlockEnd(body, end + "</w:p>".length);
      blocks.push(body.slice(start, blockEnd));
      index = blockEnd;
    } else {
      const end = body.indexOf("</w:tbl>", blockIndex);
      blocks.push(body.slice(blockIndex, end + "</w:tbl>".length));
      index = end + "</w:tbl>".length;
    }
  }

  return blocks;
}

function paragraphBlockStart(body: string, index: number, paragraphIndex: number): number {
  const prefix = body.slice(index, paragraphIndex);
  return /^(\s*<w:commentRangeStart\b[^>]*\/>)*\s*$/.test(prefix) ? index : paragraphIndex;
}

function paragraphBlockEnd(body: string, paragraphEnd: number): number {
  const suffixPattern = /^(\s*<w:commentRangeEnd\b[^>]*\/>)?(\s*<w:r\b[\s\S]*?<w:commentReference\b[^>]*\/>[\s\S]*?<\/w:r>)?/;
  const suffix = body.slice(paragraphEnd);
  const match = suffix.match(suffixPattern);

  return paragraphEnd + (match?.[0].length ?? 0);
}

function nextBlockIndex(...indexes: number[]): number {
  const presentIndexes = indexes.filter((index) => index !== -1);
  return presentIndexes.length > 0 ? Math.min(...presentIndexes) : -1;
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
  numberingContext: NumberingContext = { listTypes: new Map([[1, "bullet"], [2, "ordered"]]) },
  paragraphXml?: string,
): ParagraphNode {
  const paragraph = asObject(value);
  const properties = asObject(paragraph.pPr);
  const styleNode = asObject(properties.pStyle);
  const alignmentNode = asObject(properties.jc);
  const spacing = parseParagraphSpacing(properties.spacing);
  const indent = parseParagraphIndent(properties.ind);
  const shading = parseShading(properties.shd);
  const borders = parseParagraphBorders(properties.pBdr);
  const numbering = parseListSettings(properties.numPr, numberingContext);
  const pagination = parsePagination(properties);
  const propertyRevision = parsePropertyRevision(properties.pPrChange);
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
    ...(spacing ? { spacing } : {}),
    ...(indent ? { indent } : {}),
    ...(shading ? { shading } : {}),
    ...(borders ? { borders } : {}),
    ...(numbering ? { list: numbering } : {}),
    ...(pagination ? { pagination } : {}),
    ...(propertyRevision ? { propertyRevision } : {}),
    runs: parseParagraphRuns(paragraph, relationships, comments, footnotes, endnotes, paragraphInnerXml(paragraphXml)),
  };
}

function paragraphInnerXml(xml?: string): string | undefined {
  return xml?.match(/<w:p\b[\s\S]*<\/w:p>/)?.[0];
}

function parseParagraphCommentRangeStart(xml: string, comments: CommentMap): ParagraphNode["commentRangeStart"] | undefined {
  const id = xml.match(/^\s*<w:commentRangeStart\b[^>]*w:id="([^"]+)"[^>]*\/>\s*<w:p/)?.[1];
  if (id === undefined) {
    return undefined;
  }

  const comment = comments[id];
  return comment ? { ...comment, id: parseNumber(id) } : { id: parseNumber(id), author: "", text: "" };
}

function parseParagraphCommentRangeEnd(xml: string): ParagraphNode["commentRangeEnd"] | undefined {
  const id = xml.match(/<\/w:p>\s*<w:commentRangeEnd\b[^>]*w:id="([^"]+)"[^>]*\/>/)?.[1];
  return id !== undefined ? { id: parseNumber(id) } : undefined;
}

function parsePropertyRevision(value: unknown): ParagraphNode["propertyRevision"] | undefined {
  const revision = asObject(value);

  if (revision.id === undefined && revision.author === undefined) {
    return undefined;
  }

  return {
    id: parseNumber(revision.id),
    author: String(revision.author ?? ""),
    ...(typeof revision.date === "string" ? { date: revision.date } : {}),
  };
}

function parseTableRowRevision(properties: XmlNode): TableNode["rows"][number]["revision"] | undefined {
  const insertedRevision = parseTableRowRevisionMarker(properties.ins, "insert");

  if (insertedRevision) {
    return insertedRevision;
  }

  return parseTableRowRevisionMarker(properties.del, "delete");
}

function parseTableRowRevisionMarker(
  value: unknown,
  type: NonNullable<TableNode["rows"][number]["revision"]>["type"],
): TableNode["rows"][number]["revision"] | undefined {
  const revision = asObject(value);

  if (revision.id === undefined && revision.author === undefined) {
    return undefined;
  }

  return {
    type,
    id: parseNumber(revision.id),
    author: String(revision.author ?? ""),
    ...(typeof revision.date === "string" ? { date: revision.date } : {}),
  };
}

function parseShading(value: unknown): NonNullable<ParagraphNode["shading"]> | undefined {
  const shading = asObject(value);

  return typeof shading.fill === "string" ? { fill: shading.fill } : undefined;
}

function parseParagraphBorders(value: unknown): NonNullable<ParagraphNode["borders"]> | undefined {
  const borders = asObject(value);
  const parsed = {
    ...parseParagraphBorderSide(borders.top, "top"),
    ...parseParagraphBorderSide(borders.left, "left"),
    ...parseParagraphBorderSide(borders.bottom, "bottom"),
    ...parseParagraphBorderSide(borders.right, "right"),
  };

  return Object.keys(parsed).length > 0 ? parsed : undefined;
}

function parseParagraphBorderSide(value: unknown, side: keyof NonNullable<ParagraphNode["borders"]>): Partial<NonNullable<ParagraphNode["borders"]>> {
  const border = parseBorder(value);

  return border ? { [side]: border } : {};
}

function parseBorder(value: unknown): NonNullable<TextRun["border"]> | undefined {
  const border = asObject(value);

  if (border.val !== "single") {
    return undefined;
  }

  return {
    style: "single",
    ...(border.sz !== undefined ? { size: parseNumber(border.sz) } : {}),
    ...(typeof border.color === "string" ? { color: border.color } : {}),
    ...(border.space !== undefined ? { space: parseNumber(border.space) } : {}),
  };
}

function parseParagraphSpacing(value: unknown): ParagraphNode["spacing"] | undefined {
  const spacing = asObject(value);
  const parsed = {
    ...(spacing.before !== undefined ? { before: parseNumber(spacing.before) } : {}),
    ...(spacing.after !== undefined ? { after: parseNumber(spacing.after) } : {}),
    ...(spacing.line !== undefined ? { line: parseNumber(spacing.line) } : {}),
    ...(typeof spacing.lineRule === "string" ? { lineRule: spacing.lineRule as NonNullable<ParagraphNode["spacing"]>["lineRule"] } : {}),
  };

  return Object.keys(parsed).length > 0 ? parsed : undefined;
}

function parseParagraphIndent(value: unknown): ParagraphNode["indent"] | undefined {
  const indent = asObject(value);
  const parsed = {
    ...(indent.left !== undefined ? { left: parseNumber(indent.left) } : {}),
    ...(indent.right !== undefined ? { right: parseNumber(indent.right) } : {}),
    ...(indent.firstLine !== undefined ? { firstLine: parseNumber(indent.firstLine) } : {}),
    ...(indent.hanging !== undefined ? { hanging: parseNumber(indent.hanging) } : {}),
  };

  return Object.keys(parsed).length > 0 ? parsed : undefined;
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
  paragraphXml?: string,
): TextRun[] {
  const commentRanges = paragraphXml ? commentRangesByRunIndex(paragraphXml) : commentRangesByText(paragraph);
  const bookmarkRanges = bookmarkRangesByText(paragraph);
  const revisionRuns = [
    ...asArray(paragraph.ins).map((ins) => ({ ins })),
    ...asArray(paragraph.del).map((del) => ({ del })),
    ...asArray(paragraph.moveFrom).map((moveFrom) => ({ moveFrom })),
    ...asArray(paragraph.moveTo).map((moveTo) => ({ moveTo })),
    ...asArray(paragraph.sdt).map((sdt) => ({ sdt })),
  ];
  const normalRuns = parseComplexFieldRuns([...asArray(paragraph.r), ...revisionRuns])
    .filter((run) => run.text !== "" || run.break !== undefined || run.field !== undefined || run.footnote !== undefined || run.endnote !== undefined)
    .map((run) => withMatchingNotes(run, footnotes, endnotes))
    .map((run) => withMatchingBookmark(run, bookmarkRanges))
    .map((run, index) => withMatchingComment(run, commentRanges, comments, index));
  const hyperlinkRuns = asArray(paragraph.hyperlink)
    .flatMap((hyperlink) => parseHyperlink(hyperlink, relationships, comments));

  return [...normalRuns, ...hyperlinkRuns];
}

function parseComplexFieldRuns(runValues: unknown[]): TextRun[] {
  const runs: TextRun[] = [];

  for (let index = 0; index < runValues.length; index += 1) {
    const run = asObject(runValues[index]);
    if (run.ins !== undefined) {
      runs.push(parseInsertedRevision(run.ins));
      continue;
    }

    if (run.del !== undefined) {
      runs.push(parseDeletedRevision(run.del));
      continue;
    }

    if (run.moveFrom !== undefined) {
      runs.push(parseMoveRevision(run.moveFrom, "moveFrom"));
      continue;
    }

    if (run.moveTo !== undefined) {
      runs.push(parseMoveRevision(run.moveTo, "moveTo"));
      continue;
    }

    if (run.sdt !== undefined) {
      runs.push(parseSdtRun(run.sdt));
      continue;
    }

    const fieldChar = asObject(run.fldChar);

    if (fieldChar.fldCharType !== "begin") {
      runs.push(parseRun(run));
      continue;
    }

    const instructionRun = asObject(runValues[index + 1]);
    const separateRun = asObject(runValues[index + 2]);
    const resultRun = asObject(runValues[index + 3]);
    const endRun = asObject(runValues[index + 4]);

    if (
      instructionRun.instrText === undefined ||
      asObject(separateRun.fldChar).fldCharType !== "separate" ||
      asObject(endRun.fldChar).fldCharType !== "end"
    ) {
      runs.push(parseRun(run));
      continue;
    }

    const field = parseField(instructionRun.instrText);
    const result = parseText(resultRun.t);
    runs.push({
      text: "",
      field: result ? fieldWithResult(field, result) : field,
    });
    index += 4;
  }

  return runs;
}

function parseInsertedRevision(value: unknown): TextRun {
  const revision = asObject(value);
  return {
    ...parseRun(revision.r),
    revision: {
      type: "insert",
      id: parseNumber(revision.id),
      author: String(revision.author ?? ""),
      ...(typeof revision.date === "string" ? { date: revision.date } : {}),
    },
  };
}

function parseDeletedRevision(value: unknown): TextRun {
  const revision = asObject(value);
  return {
    ...parseRun(revision.r),
    revision: {
      type: "delete",
      id: parseNumber(revision.id),
      author: String(revision.author ?? ""),
      ...(typeof revision.date === "string" ? { date: revision.date } : {}),
    },
  };
}

function parseMoveRevision(value: unknown, type: "moveFrom" | "moveTo"): TextRun {
  const revision = asObject(value);
  return {
    ...parseRun(revision.r),
    revision: {
      type,
      id: parseNumber(revision.id),
      author: String(revision.author ?? ""),
      ...(typeof revision.date === "string" ? { date: revision.date } : {}),
    },
  };
}

function parseSdtRun(value: unknown): TextRun {
  const sdt = asObject(value);
  const content = asObject(sdt.sdtContent);
  const contentControl = parseContentControl(sdt.sdtPr);

  return {
    ...parseRun(content.r),
    ...(contentControl ? { contentControl } : {}),
  };
}

function parseContentControl(value: unknown): ParagraphNode["contentControl"] | undefined {
  const properties = asObject(value);
  const alias = asObject(properties.alias);
  const tag = asObject(properties.tag);
  const lock = asObject(properties.lock);
  const appearance = asObject(properties.appearance);
  const color = asObject(properties.color);
  const dataBinding = parseDataBindingContentControl(properties.dataBinding);
  const placeholder = parsePlaceholderContentControl(properties.placeholder);
  const checkbox = parseCheckboxContentControl(properties.checkBox);
  const dropdown = parseDropdownContentControl(properties.dropDownList);
  const comboBox = parseComboBoxContentControl(properties.comboBox);
  const date = parseDateContentControl(properties.date);
  const repeatingSection = parseRepeatingSectionContentControl(properties.repeatingSection);
  const repeatingSectionItem = parseRepeatingSectionItemContentControl(properties.repeatingSectionItem);
  const parsed = {
    ...(typeof alias.val === "string" ? { alias: alias.val } : {}),
    ...(typeof tag.val === "string" ? { tag: tag.val } : {}),
    ...(typeof lock.val === "string" ? { lock: lock.val as NonNullable<ParagraphNode["contentControl"]>["lock"] } : {}),
    ...(typeof appearance.val === "string" ? { appearance: appearance.val as NonNullable<ParagraphNode["contentControl"]>["appearance"] } : {}),
    ...(typeof color.val === "string" ? { color: color.val } : {}),
    ...(properties.showingPlcHdr !== undefined ? { showingPlaceholder: true } : {}),
    ...(dataBinding ? { dataBinding } : {}),
    ...(placeholder ? { placeholder } : {}),
    ...(checkbox ? { checkbox } : {}),
    ...(dropdown ? { dropdown } : {}),
    ...(comboBox ? { comboBox } : {}),
    ...(date ? { date } : {}),
    ...(repeatingSection ? { repeatingSection } : {}),
    ...(repeatingSectionItem ? { repeatingSectionItem } : {}),
  };

  return Object.keys(parsed).length > 0 ? parsed : undefined;
}

function parseDataBindingContentControl(value: unknown): NonNullable<NonNullable<ParagraphNode["contentControl"]>["dataBinding"]> | undefined {
  const dataBinding = asObject(value);
  const parsed = {
    ...(typeof dataBinding.storeItemID === "string" ? { storeItemId: dataBinding.storeItemID } : {}),
    ...(typeof dataBinding.xpath === "string" ? { xpath: dataBinding.xpath } : {}),
    ...(typeof dataBinding.prefixMappings === "string" ? { prefixMappings: dataBinding.prefixMappings } : {}),
  };

  return Object.keys(parsed).length > 0 ? parsed : undefined;
}

function parsePlaceholderContentControl(value: unknown): NonNullable<NonNullable<ParagraphNode["contentControl"]>["placeholder"]> | undefined {
  const placeholder = asObject(value);
  const docPart = asObject(placeholder.docPart);

  return typeof docPart.val === "string" ? { docPart: docPart.val } : undefined;
}

function parseCheckboxContentControl(value: unknown): NonNullable<NonNullable<ParagraphNode["contentControl"]>["checkbox"]> | undefined {
  const checkbox = asObject(value);

  if (Object.keys(checkbox).length === 0) {
    return undefined;
  }

  const checked = asObject(checkbox.checked);
  const checkedState = asObject(checkbox.checkedState);
  const uncheckedState = asObject(checkbox.uncheckedState);

  return {
    checked: checked.val === 1 || checked.val === "1" || checked.val === true || checked.val === "true",
    ...(typeof checkedState.val === "string" ? { checkedSymbol: checkedState.val } : {}),
    ...(typeof uncheckedState.val === "string" ? { uncheckedSymbol: uncheckedState.val } : {}),
  };
}

function parseDropdownContentControl(value: unknown): NonNullable<NonNullable<ParagraphNode["contentControl"]>["dropdown"]> | undefined {
  const dropdown = asObject(value);
  const items = asArray(dropdown.listItem)
    .map((item) => asObject(item))
    .filter((item) => item.displayText !== undefined || item.value !== undefined)
    .map((item) => ({
      displayText: String(item.displayText ?? ""),
      value: String(item.value ?? ""),
    }));

  return items.length > 0 ? { items } : undefined;
}

function parseComboBoxContentControl(value: unknown): NonNullable<NonNullable<ParagraphNode["contentControl"]>["comboBox"]> | undefined {
  const comboBox = asObject(value);
  const items = asArray(comboBox.listItem)
    .map((item) => asObject(item))
    .filter((item) => item.displayText !== undefined || item.value !== undefined)
    .map((item) => ({
      displayText: String(item.displayText ?? ""),
      value: String(item.value ?? ""),
    }));

  return items.length > 0 ? { items } : undefined;
}

function parseDateContentControl(value: unknown): NonNullable<NonNullable<ParagraphNode["contentControl"]>["date"]> | undefined {
  const date = asObject(value);

  if (Object.keys(date).length === 0) {
    return undefined;
  }

  const fullDate = asObject(date.fullDate);
  const format = asObject(date.dateFormat);
  const parsed = {
    ...(typeof fullDate.val === "string" ? { fullDate: fullDate.val } : {}),
    ...(typeof format.val === "string" ? { format: format.val } : {}),
  };

  return Object.keys(parsed).length > 0 ? parsed : undefined;
}

function parseRepeatingSectionContentControl(value: unknown): NonNullable<NonNullable<ParagraphNode["contentControl"]>["repeatingSection"]> | undefined {
  const repeatingSection = asObject(value);

  if (Object.keys(repeatingSection).length === 0) {
    return undefined;
  }

  const sectionTitle = asObject(repeatingSection.sectionTitle);
  const parsed = {
    ...(typeof sectionTitle.val === "string" ? { sectionTitle: sectionTitle.val } : {}),
    ...(repeatingSection.doNotAllowInsertDeleteSection !== undefined ? { doNotAllowInsertDeleteSection: true } : {}),
  };

  return Object.keys(parsed).length > 0 ? parsed : undefined;
}

function parseRepeatingSectionItemContentControl(value: unknown): NonNullable<NonNullable<ParagraphNode["contentControl"]>["repeatingSectionItem"]> | undefined {
  const repeatingSectionItem = asObject(value);

  if (Object.keys(repeatingSectionItem).length === 0) {
    return undefined;
  }

  const id = asObject(repeatingSectionItem.id);
  const parsed = {
    ...(typeof id.val === "string" ? { id: id.val } : {}),
  };

  return Object.keys(parsed).length > 0 ? parsed : undefined;
}

function fieldWithResult(field: TextRun["field"], result: string): TextRun["field"] {
  if (field === "page" || field === "numPages") {
    return { type: field, result };
  }

  if (typeof field === "object") {
    return { ...field, result };
  }

  return field;
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

function commentRangesByRunIndex(paragraphXml: string): Map<number, string> {
  const ranges = new Map<number, string>();
  const activeCommentIds: string[] = [];
  const tokenPattern = /<w:commentRangeStart\b[^>]*w:id="([^"]+)"[^>]*\/>|<w:commentRangeEnd\b[^>]*w:id="([^"]+)"[^>]*\/>|<w:r\b[\s\S]*?<\/w:r>/g;
  let runIndex = 0;
  let match: RegExpExecArray | null;

  while ((match = tokenPattern.exec(paragraphXml)) !== null) {
    if (match[1] !== undefined) {
      activeCommentIds.push(match[1]);
      continue;
    }

    if (match[2] !== undefined) {
      const activeIndex = activeCommentIds.lastIndexOf(match[2]);
      if (activeIndex !== -1) {
        activeCommentIds.splice(activeIndex, 1);
      }
      continue;
    }

    const runXml = match[0];
    if (runXml.includes("<w:commentReference")) {
      continue;
    }

    const commentId = activeCommentIds[activeCommentIds.length - 1];
    if (commentId !== undefined) {
      ranges.set(runIndex, commentId);
    }
    runIndex += 1;
  }

  return ranges;
}

function withMatchingComment(run: TextRun, ranges: Map<string, string> | Map<number, string>, comments: CommentMap, index?: number): TextRun {
  const commentId = index !== undefined ? ranges.get(index as never) : ranges.get(run.text as never);
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
  const anchor = typeof hyperlink.anchor === "string" ? hyperlink.anchor : undefined;
  const ranges = commentRangesByText(hyperlink);

  return asArray(hyperlink.r).map((run) => ({
    ...withMatchingComment(parseRun(run), ranges, comments),
    ...(url ? { link: { url } } : {}),
    ...(anchor ? { link: { anchor } } : {}),
  })).filter((run) => run.text !== "");
}

function parseListSettings(value: unknown, numberingContext: NumberingContext): ParagraphNode["list"] | undefined {
  const numbering = asObject(value);
  const level = asObject(numbering.ilvl);
  const numId = asObject(numbering.numId);

  if (level.val === undefined || numId.val === undefined) {
    return undefined;
  }

  const parsedNumId = parseNumber(numId.val);

  return {
    type: numberingContext.listTypes.get(parsedNumId) ?? (parsedNumId === 1 ? "bullet" : "ordered"),
    level: parseNumber(level.val),
    ...(!isBuiltInNumberingId(parsedNumId) ? { numberingId: parsedNumId } : {}),
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
    text: parseText(run.t ?? run.delText),
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

  if (instruction.startsWith("TOC")) {
    return parseTocField(instruction);
  }

  if (instruction.startsWith("REF ")) {
    return { type: "ref", target: instruction.slice("REF ".length) };
  }

  return instruction === "NUMPAGES" ? "numPages" : "page";
}

function parseTocField(instruction: string): TextRun["field"] {
  const switches = instruction
    .slice("TOC".length)
    .trim()
    .replaceAll("\\", "")
    .trim();

  return {
    type: "toc",
    ...(switches ? { switches } : {}),
  };
}

function parseRunFont(properties: XmlNode): Partial<TextRun> {
  const fonts = asObject(properties.rFonts);
  const size = asObject(properties.sz);
  const color = asObject(properties.color);
  const highlight = asObject(properties.highlight);
  const verticalAlign = asObject(properties.vertAlign);
  const characterSpacing = asObject(properties.spacing);
  const scale = asObject(properties.w);
  const border = parseBorder(properties.bdr);

  return {
    ...(typeof fonts.ascii === "string" ? { fontFamily: fonts.ascii } : {}),
    ...(typeof size.val === "number" ? { fontSize: size.val / 2 } : {}),
    ...(typeof size.val === "string" ? { fontSize: Number.parseInt(size.val, 10) / 2 } : {}),
    ...(typeof color.val === "string" ? { color: color.val } : {}),
    ...(typeof highlight.val === "string" ? { highlight: highlight.val as NonNullable<TextRun["highlight"]> } : {}),
    ...(properties.strike !== undefined ? { strike: true } : {}),
    ...(properties.dstrike !== undefined ? { doubleStrike: true } : {}),
    ...(properties.smallCaps !== undefined ? { smallCaps: true } : {}),
    ...(properties.caps !== undefined ? { allCaps: true } : {}),
    ...(typeof verticalAlign.val === "string" ? { verticalAlign: verticalAlign.val as NonNullable<TextRun["verticalAlign"]> } : {}),
    ...(characterSpacing.val !== undefined ? { characterSpacing: parseNumber(characterSpacing.val) } : {}),
    ...(scale.val !== undefined ? { scale: parseNumber(scale.val) } : {}),
    ...(border ? { border } : {}),
  };
}

function parseTable(value: unknown, relationships: RelationshipMap, comments: CommentMap, footnotes: NoteMap, endnotes: NoteMap, numberingContext: NumberingContext): TableNode {
  const table = asObject(value);
  const properties = asObject(table.tblPr);
  const style = asObject(properties.tblStyle);
  const width = asObject(properties.tblW);
  const borders = asObject(properties.tblBorders);
  const alignment = asObject(properties.jc);
  const cellSpacing = asObject(properties.tblCellSpacing);
  const propertyRevision = parsePropertyRevision(properties.tblPrChange);
  const grid = parseTableGrid(table.tblGrid);

  return {
    type: "table",
    ...(typeof style.val === "string" ? { styleId: style.val } : {}),
    ...(grid ? { grid } : {}),
    ...(width.w !== undefined ? { width: parseNumber(width.w) } : {}),
    ...(borders.top !== undefined ? { borders: "single" as const } : {}),
    ...(typeof alignment.val === "string" ? { alignment: alignment.val as NonNullable<TableNode["alignment"]> } : {}),
    ...(cellSpacing.w !== undefined ? { cellSpacing: parseNumber(cellSpacing.w) } : {}),
    ...(propertyRevision ? { propertyRevision } : {}),
    rows: asArray(table.tr).map((rowValue) => {
      const row = asObject(rowValue);
      const rowProperties = asObject(row.trPr);
      const revision = parseTableRowRevision(rowProperties);
      const height = parseTableRowHeight(rowProperties);

      return {
        ...(revision ? { revision } : {}),
        ...(height ? { height } : {}),
        cells: asArray(row.tc).map((cell) => parseTableCell(cell, relationships, comments, footnotes, endnotes, numberingContext)),
      };
    }),
  };
}

function parseTableGrid(value: unknown): number[] | undefined {
  const grid = asObject(value);
  const columns = asArray(grid.gridCol)
    .map((column) => asObject(column))
    .filter((column) => column.w !== undefined)
    .map((column) => parseNumber(column.w));

  return columns.length > 0 ? columns : undefined;
}

function parseTableRowHeight(value: unknown): TableNode["rows"][number]["height"] | undefined {
  const properties = asObject(value);
  const height = asObject(properties.trHeight);

  if (height.val === undefined) {
    return undefined;
  }

  return {
    value: parseNumber(height.val),
    ...(typeof height.hRule === "string" ? { rule: height.hRule as NonNullable<TableNode["rows"][number]["height"]>["rule"] } : {}),
  };
}

function parseTableCell(value: unknown, relationships: RelationshipMap, comments: CommentMap, footnotes: NoteMap, endnotes: NoteMap, numberingContext: NumberingContext): TableCellNode {
  const cell = asObject(value);
  const properties = asObject(cell.tcPr);
  const width = asObject(properties.tcW);
  const gridSpan = asObject(properties.gridSpan);
  const verticalMerge = asObject(properties.vMerge);
  const verticalAlignment = asObject(properties.vAlign);
  const shading = parseTableCellShading(properties.shd);
  const borders = parseParagraphBorders(properties.tcBorders);
  const textDirection = asObject(properties.textDirection);
  const margins = parseTableCellMargins(properties.tcMar);
  const propertyRevision = parsePropertyRevision(properties.tcPrChange);

  return {
    ...(width.w !== undefined ? { width: parseNumber(width.w) } : {}),
    ...(gridSpan.val !== undefined ? { colSpan: parseNumber(gridSpan.val) } : {}),
    ...(typeof verticalMerge.val === "string" ? { verticalMerge: verticalMerge.val as NonNullable<TableCellNode["verticalMerge"]> } : {}),
    ...(typeof verticalAlignment.val === "string" ? { verticalAlignment: verticalAlignment.val as NonNullable<TableCellNode["verticalAlignment"]> } : {}),
    ...(shading ? { shading } : {}),
    ...(borders ? { borders } : {}),
    ...(typeof textDirection.val === "string" ? { textDirection: textDirection.val as NonNullable<TableCellNode["textDirection"]> } : {}),
    ...(margins ? { margins } : {}),
    ...(propertyRevision ? { propertyRevision } : {}),
    blocks: asArray(cell.p).map((paragraph) => parseParagraph(paragraph, relationships, comments, footnotes, endnotes, numberingContext)),
  };
}

function parseTableCellShading(value: unknown): TableCellNode["shading"] | undefined {
  const shading = asObject(value);

  return typeof shading.fill === "string" ? { fill: shading.fill } : undefined;
}

function parseTableCellMargins(value: unknown): TableCellNode["margins"] | undefined {
  const margins = asObject(value);
  const parsed = {
    ...parseTableCellMarginSide(margins.top, "top"),
    ...parseTableCellMarginSide(margins.right, "right"),
    ...parseTableCellMarginSide(margins.bottom, "bottom"),
    ...parseTableCellMarginSide(margins.left, "left"),
  };

  return Object.keys(parsed).length > 0 ? parsed : undefined;
}

function parseTableCellMarginSide(value: unknown, side: keyof NonNullable<TableCellNode["margins"]>): Partial<NonNullable<TableCellNode["margins"]>> {
  const margin = asObject(value);

  return margin.w !== undefined ? { [side]: parseNumber(margin.w) } : {};
}

function parseText(value: unknown): string {
  if (typeof value === "string") {
    return value;
  }

  if (typeof value === "number") {
    return String(value);
  }

  const node = asObject(value);
  const text = node["#text"];
  if (typeof text === "number") {
    return String(text);
  }

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

function pathBasename(path: string): string {
  return path.split("/").pop() ?? path;
}

function pathDirname(path: string): string {
  const parts = path.split("/");
  parts.pop();

  return parts.join("/");
}

function xmlText(value: unknown): string | undefined {
  if (typeof value === "string") {
    return value;
  }

  const object = asObject(value);
  const text = object["#text"];

  return typeof text === "string" ? text : undefined;
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
