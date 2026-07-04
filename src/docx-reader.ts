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
  NumberingLevelDefinition,
  NumberingInstance,
  NumberingFormat,
  StyleParagraphProperties,
  StyleRunProperties,
  TableStyleDefinition,
  TableCellNode,
  TableNode,
  TextRun,
  MathNode,
  MathControlProperties,
  RunLanguage,
  SectionNode,
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
  const fonts = await parseFontTable(zip);
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
    ...(fonts.length > 0 ? { fonts } : {}),
    sections: await parseSections(zip, xml, body, relationships, comments, media, footnotes, endnotes, numberingContext),
  };
}

async function parseFontTable(zip: JSZip): Promise<NonNullable<DocumentJson["fonts"]>> {
  const fontTableFile = zip.file("word/fontTable.xml");

  if (!fontTableFile) {
    return [];
  }

  const xml = await fontTableFile.async("string");
  const parsed = parser.parse(xml) as XmlNode;
  const fontsRoot = asObject(parsed.fonts);

  return asArray(fontsRoot.font)
    .map((fontValue) => asObject(fontValue))
    .filter((font) => typeof font.name === "string")
    .map((font) => {
      const panose1 = asObject(font.panose1);
      const charset = asObject(font.charset);
      const family = asObject(font.family);
      const pitch = asObject(font.pitch);

      return {
        name: String(font.name),
        ...(typeof family.val === "string" ? { family: family.val as NonNullable<DocumentJson["fonts"]>[number]["family"] } : {}),
        ...(typeof pitch.val === "string" ? { pitch: pitch.val as NonNullable<DocumentJson["fonts"]>[number]["pitch"] } : {}),
        ...(typeof charset.val === "string" ? { charset: charset.val } : {}),
        ...(typeof panose1.val === "string" ? { panose1: panose1.val } : {}),
      };
    });
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
  const web = await parseWebSettings(zip);

  if (!settingsFile) {
    return web ? { web } : undefined;
  }

  const xml = await settingsFile.async("string");
  const parsed = parser.parse(xml) as XmlNode;
  const settings = asObject(parsed.settings);
  const defaultTabStop = asObject(settings.defaultTabStop);
  const compatibility = parseCompatibilitySettings(settings);
  const proofing = parseProofingSettings(settings);
  const protection = parseDocumentProtection(settings);
  const mailMerge = parseMailMergeSettings(settings);
  const writeProtection = parseWriteProtection(settings);
  const math = parseMathSettings(settings);
  const view = parseViewSettings(settings);
  const result = {
    ...(defaultTabStop.val !== undefined ? { defaultTabStop: parseNumber(defaultTabStop.val) } : {}),
    ...(settings.evenAndOddHeaders !== undefined ? { evenAndOddHeaders: true } : {}),
    ...(settings.updateFields !== undefined ? { updateFields: true } : {}),
    ...(settings.trackRevisions !== undefined ? { trackRevisions: true } : {}),
    ...(compatibility ? { compatibility } : {}),
    ...(proofing ? { proofing } : {}),
    ...(protection ? { protection } : {}),
    ...(mailMerge ? { mailMerge } : {}),
    ...(writeProtection ? { writeProtection } : {}),
    ...(math ? { math } : {}),
    ...(view ? { view } : {}),
    ...(web ? { web } : {}),
  };

  return Object.keys(result).length > 0 ? result : undefined;
}

function parseMathSettings(settings: XmlNode): NonNullable<DocumentJson["settings"]>["math"] | undefined {
  const mathPr = asObject(settings.mathPr);
  const mathFont = asObject(mathPr.mathFont);
  const breakBinary = asObject(mathPr.brkBin);
  const smallFraction = asObject(mathPr.smallFrac);
  const breakBinaryValue = parseBreakBinary(breakBinary.val);
  const result = {
    ...(typeof mathFont.val === "string" ? { mathFont: mathFont.val } : {}),
    ...(breakBinaryValue ? { breakBinary: breakBinaryValue } : {}),
    ...(smallFraction.val !== undefined ? { smallFraction: parseOnOff(smallFraction.val) } : {}),
    ...(mathPr.dispDef !== undefined ? { displayDefaults: true } : {}),
  };

  return Object.keys(result).length > 0 ? result : undefined;
}

function parseBreakBinary(value: unknown): NonNullable<NonNullable<DocumentJson["settings"]>["math"]>["breakBinary"] | undefined {
  return value === "before" || value === "after" || value === "repeat" ? value : undefined;
}

function parseWriteProtection(settings: XmlNode): NonNullable<DocumentJson["settings"]>["writeProtection"] | undefined {
  const writeProtection = asObject(settings.writeProtection);
  const result = {
    ...(writeProtection.recommended !== undefined ? { recommended: parseOnOff(writeProtection.recommended) } : {}),
    ...(typeof writeProtection.cryptProviderType === "string" ? { cryptProviderType: writeProtection.cryptProviderType } : {}),
    ...(typeof writeProtection.cryptAlgorithmClass === "string" ? { cryptAlgorithmClass: writeProtection.cryptAlgorithmClass } : {}),
    ...(typeof writeProtection.cryptAlgorithmType === "string" ? { cryptAlgorithmType: writeProtection.cryptAlgorithmType } : {}),
    ...(writeProtection.cryptAlgorithmSid !== undefined ? { cryptAlgorithmSid: parseNumber(writeProtection.cryptAlgorithmSid) } : {}),
    ...(writeProtection.cryptSpinCount !== undefined ? { cryptSpinCount: parseNumber(writeProtection.cryptSpinCount) } : {}),
    ...(typeof writeProtection.hash === "string" ? { hash: writeProtection.hash } : {}),
    ...(typeof writeProtection.salt === "string" ? { salt: writeProtection.salt } : {}),
  };

  return Object.keys(result).length > 0 ? result : undefined;
}

function parseMailMergeSettings(settings: XmlNode): NonNullable<DocumentJson["settings"]>["mailMerge"] | undefined {
  const mailMerge = asObject(settings.mailMerge);
  const mainDocumentType = asObject(mailMerge.mainDocumentType);
  const dataType = asObject(mailMerge.dataType);
  const connectString = asObject(mailMerge.connectString);
  const query = asObject(mailMerge.query);
  const activeRecord = asObject(mailMerge.activeRecord);
  const checkErrors = asObject(mailMerge.checkErrors);
  const result = {
    ...(typeof mainDocumentType.val === "string" ? { mainDocumentType: mainDocumentType.val } : {}),
    ...(typeof dataType.val === "string" ? { dataType: dataType.val } : {}),
    ...(typeof connectString.val === "string" ? { connectString: connectString.val } : {}),
    ...(typeof query.val === "string" ? { query: query.val } : {}),
    ...(mailMerge.viewMergedData !== undefined ? { viewMergedData: true } : {}),
    ...(activeRecord.val !== undefined ? { activeRecord: parseNumber(activeRecord.val) } : {}),
    ...(checkErrors.val !== undefined ? { checkErrors: parseNumber(checkErrors.val) } : {}),
  };

  return Object.keys(result).length > 0 ? result : undefined;
}

function parseDocumentProtection(settings: XmlNode): NonNullable<DocumentJson["settings"]>["protection"] | undefined {
  const protection = asObject(settings.documentProtection);
  const edit = protectionEditValue(protection.edit);
  const result = {
    ...(edit ? { edit } : {}),
    ...(protection.enforcement !== undefined ? { enforcement: parseOnOff(protection.enforcement) } : {}),
    ...(typeof protection.cryptProviderType === "string" ? { cryptProviderType: protection.cryptProviderType } : {}),
    ...(typeof protection.cryptAlgorithmClass === "string" ? { cryptAlgorithmClass: protection.cryptAlgorithmClass } : {}),
    ...(typeof protection.cryptAlgorithmType === "string" ? { cryptAlgorithmType: protection.cryptAlgorithmType } : {}),
    ...(protection.cryptAlgorithmSid !== undefined ? { cryptAlgorithmSid: parseNumber(protection.cryptAlgorithmSid) } : {}),
    ...(protection.cryptSpinCount !== undefined ? { cryptSpinCount: parseNumber(protection.cryptSpinCount) } : {}),
    ...(typeof protection.hash === "string" ? { hash: protection.hash } : {}),
    ...(typeof protection.salt === "string" ? { salt: protection.salt } : {}),
  };

  return Object.keys(result).length > 0 ? result : undefined;
}

function protectionEditValue(value: unknown): NonNullable<NonNullable<DocumentJson["settings"]>["protection"]>["edit"] | undefined {
  return value === "none" || value === "readOnly" || value === "comments" || value === "trackedChanges" || value === "forms" ? value : undefined;
}

function parseOnOff(value: unknown): boolean {
  return value === true || value === "true" || value === "1" || value === "on";
}

function parseViewSettings(settings: XmlNode): NonNullable<DocumentJson["settings"]>["view"] | undefined {
  const view = asObject(settings.view);
  const zoom = asObject(settings.zoom);
  const mode = viewModeValue(view.val);
  const preset = zoomPresetValue(zoom.val);
  const zoomValue = {
    ...(preset ? { preset } : {}),
    ...(zoom.percent !== undefined ? { percent: parseNumber(zoom.percent) } : {}),
  };
  const result = {
    ...(mode ? { mode } : {}),
    ...(Object.keys(zoomValue).length > 0 ? { zoom: zoomValue } : {}),
  };

  return Object.keys(result).length > 0 ? result : undefined;
}

function viewModeValue(value: unknown): NonNullable<NonNullable<DocumentJson["settings"]>["view"]>["mode"] | undefined {
  return value === "none" || value === "print" || value === "outline" || value === "masterPages" || value === "normal" || value === "web" ? value : undefined;
}

function zoomPresetValue(value: unknown): NonNullable<NonNullable<NonNullable<DocumentJson["settings"]>["view"]>["zoom"]>["preset"] | undefined {
  return value === "none" || value === "fullPage" || value === "bestFit" || value === "textFit" ? value : undefined;
}

function parseProofingSettings(settings: XmlNode): NonNullable<DocumentJson["settings"]>["proofing"] | undefined {
  const proofState = asObject(settings.proofState);
  const spelling = proofStateValue(proofState.spelling);
  const grammar = proofStateValue(proofState.grammar);
  const hyphenationZone = asObject(settings.hyphenationZone);
  const result = {
    ...(spelling ? { spelling } : {}),
    ...(grammar ? { grammar } : {}),
    ...(settings.doNotHyphenateCaps !== undefined ? { doNotHyphenateCaps: true } : {}),
    ...(hyphenationZone.val !== undefined ? { hyphenationZone: parseNumber(hyphenationZone.val) } : {}),
  };

  return Object.keys(result).length > 0 ? result : undefined;
}

function proofStateValue(value: unknown): "clean" | "dirty" | undefined {
  return value === "clean" || value === "dirty" ? value : undefined;
}

function parseCompatibilitySettings(settings: XmlNode): NonNullable<DocumentJson["settings"]>["compatibility"] | undefined {
  const compat = asObject(settings.compat);
  const compatSettings = asArray(compat.compatSetting)
    .map((setting) => asObject(setting))
    .filter((setting) => typeof setting.name === "string" && typeof setting.uri === "string" && setting.val !== undefined);
  const compatMode = compatSettings.find((setting) => setting.name === "compatibilityMode");
  const otherSettings = compatSettings
    .filter((setting) => setting.name !== "compatibilityMode")
    .map((setting) => ({
      name: setting.name as string,
      uri: setting.uri as string,
      value: String(setting.val),
    }));

  if (!compatMode && otherSettings.length === 0) {
    return undefined;
  }

  return {
    ...(compatMode ? { compatMode: String(compatMode.val) } : {}),
    ...(otherSettings.length > 0 ? { settings: otherSettings } : {}),
  };
}

async function parseWebSettings(zip: JSZip): Promise<NonNullable<NonNullable<DocumentJson["settings"]>["web"]> | undefined> {
  const webSettingsFile = zip.file("word/webSettings.xml");

  if (!webSettingsFile) {
    return undefined;
  }

  const xml = await webSettingsFile.async("string");
  const parsed = parser.parse(xml) as XmlNode;
  const webSettings = asObject(parsed.webSettings);
  const pixelsPerInch = asObject(webSettings.pixelsPerInch);
  const result = {
    ...(webSettings.optimizeForBrowser !== undefined ? { optimizeForBrowser: true } : {}),
    ...(webSettings.allowPNG !== undefined ? { allowPng: true } : {}),
    ...(webSettings.doNotSaveAsSingleFile !== undefined ? { doNotSaveAsSingleFile: true } : {}),
    ...(pixelsPerInch.val !== undefined ? { pixelsPerInch: parseNumber(pixelsPerInch.val) } : {}),
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
    .map((num) => parseNumberingInstance(num));
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

function parseNumberingInstance(value: XmlNode): NumberingInstance {
  const overrides = asArray(value.lvlOverride).map((overrideValue) => parseNumberingLevelOverride(overrideValue));

  return {
    id: parseNumber(value.numId),
    abstractId: parseNumber(asObject(value.abstractNumId).val),
    ...(overrides.length > 0 ? { overrides } : {}),
  };
}

function parseNumberingLevelOverride(value: unknown): NonNullable<NumberingInstance["overrides"]>[number] {
  const override = asObject(value);
  const start = asObject(override.startOverride);
  const level = asObject(override.lvl);

  return {
    level: parseNumber(override.ilvl),
    ...(start.val !== undefined ? { start: parseNumber(start.val) } : {}),
    ...(override.lvl !== undefined ? { definition: parseNumberingLevelDefinition(level) } : {}),
  };
}

function parseAbstractNumberingDefinition(value: XmlNode): AbstractNumberingDefinition {
  const nsid = asObject(value.nsid);
  const multiLevelType = asObject(value.multiLevelType);
  const template = asObject(value.tmpl);
  const styleLink = asObject(value.styleLink);
  const numberingStyleLink = asObject(value.numStyleLink);

  return {
    id: parseNumber(value.abstractNumId),
    ...(typeof nsid.val === "string" ? { nsid: nsid.val } : {}),
    ...(isMultiLevelType(multiLevelType.val) ? { multiLevelType: multiLevelType.val } : {}),
    ...(typeof template.val === "string" ? { templateCode: template.val } : {}),
    ...(typeof styleLink.val === "string" ? { styleLink: styleLink.val } : {}),
    ...(typeof numberingStyleLink.val === "string" ? { numberingStyleLink: numberingStyleLink.val } : {}),
    levels: asArray(value.lvl).map((levelValue) => parseNumberingLevelDefinition(levelValue)),
  };
}

function parseNumberingLevelDefinition(value: unknown): NumberingLevelDefinition {
  const level = asObject(value);
  const start = asObject(level.start);
  const style = asObject(level.pStyle);
  const format = asObject(level.numFmt);
  const text = asObject(level.lvlText);
  const suffix = asObject(level.suff);
  const restart = asObject(level.lvlRestart);
  const legal = parseOptionalOnOff(level.isLgl);
  const alignment = asObject(level.lvlJc);
  const run = parseStyleRunProperties(level.rPr);
  const indentation = asObject(asObject(level.pPr).ind);

  return {
    level: parseNumber(level.ilvl),
    format: parseNumberingFormat(format.val),
    text: typeof text.val === "string" ? text.val : "",
    ...(start.val !== undefined ? { start: parseNumber(start.val) } : {}),
    ...(typeof style.val === "string" ? { styleId: style.val } : {}),
    ...(isNumberingSuffix(suffix.val) ? { suffix: suffix.val } : {}),
    ...(restart.val !== undefined ? { restart: parseNumber(restart.val) } : {}),
    ...(legal !== undefined ? { legal } : {}),
    ...(typeof alignment.val === "string" ? { alignment: alignment.val as ParagraphAlignment } : {}),
    ...(run ? { run } : {}),
    ...(indentation.left !== undefined ? { left: parseNumber(indentation.left) } : {}),
    ...(indentation.hanging !== undefined ? { hanging: parseNumber(indentation.hanging) } : {}),
  };
}

function isMultiLevelType(value: unknown): value is NonNullable<AbstractNumberingDefinition["multiLevelType"]> {
  return value === "singleLevel" || value === "multilevel" || value === "hybridMultilevel";
}

function parseNumberingFormat(value: unknown): NumberingFormat {
  if (value === "bullet" || value === "decimal" || value === "lowerLetter" || value === "upperLetter" || value === "lowerRoman" || value === "upperRoman") {
    return value;
  }

  return "decimal";
}

function isNumberingSuffix(value: unknown): value is NonNullable<NumberingLevelDefinition["suffix"]> {
  return value === "nothing" || value === "space" || value === "tab";
}

function parseOptionalOnOff(value: unknown): boolean | undefined {
  if (value === undefined) {
    return undefined;
  }

  const rawValue = asObject(value).val;
  return rawValue === "0" || rawValue === false || rawValue === "false" || rawValue === "off" ? false : true;
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
  const accent1 = themeColorValue(colorScheme, "accent1");

  if (typeof majorLatin.typeface !== "string" || typeof minorLatin.typeface !== "string" || typeof accent1 !== "string") {
    return undefined;
  }

  const formatScheme = parseThemeFormatScheme(elements);

  return {
    name: typeof theme.name === "string" ? theme.name : "Theme",
    fonts: themeFonts(fontScheme, majorLatin.typeface, minorLatin.typeface),
    colors: themeColors(colorScheme, accent1),
    ...(formatScheme ? { formatScheme } : {}),
  };
}

function parseThemeFormatScheme(elements: XmlNode): DocumentTheme["formatScheme"] | undefined {
  const formatScheme = asObject(elements.fmtScheme);

  if (Object.keys(formatScheme).length === 0) {
    return undefined;
  }

  const fillStyleColors = parseSolidFillColors(asObject(formatScheme.fillStyleLst).solidFill);
  const lineStyleColors = asArray(asObject(formatScheme.lnStyleLst).ln)
    .map((line) => themeNestedColorValue(line, ["solidFill", "srgbClr"]))
    .filter((color): color is string => typeof color === "string");
  const effectStyleColors = asArray(asObject(formatScheme.effectStyleLst).effectStyle)
    .map((effect) => themeNestedColorValue(effect, ["effectLst", "outerShdw", "srgbClr"]))
    .filter((color): color is string => typeof color === "string");
  const backgroundFillStyleColors = parseSolidFillColors(asObject(formatScheme.bgFillStyleLst).solidFill);

  return {
    name: typeof formatScheme.name === "string" ? formatScheme.name : "Format Scheme",
    ...(fillStyleColors.length > 0 ? { fillStyleColors } : {}),
    ...(lineStyleColors.length > 0 ? { lineStyleColors } : {}),
    ...(effectStyleColors.length > 0 ? { effectStyleColors } : {}),
    ...(backgroundFillStyleColors.length > 0 ? { backgroundFillStyleColors } : {}),
  };
}

function parseSolidFillColors(value: unknown): string[] {
  return asArray(value)
    .map((fill) => themeNestedColorValue(fill, ["srgbClr"]))
    .filter((color): color is string => typeof color === "string");
}

function themeNestedColorValue(value: unknown, path: string[]): string | undefined {
  const color = path.reduce<unknown>((current, key) => asObject(current)[key], value);
  const colorObject = asObject(color);
  return typeof colorObject.val === "string" ? colorObject.val : undefined;
}

function themeFontValue(fontScheme: XmlNode, group: "majorFont" | "minorFont", slot: "ea" | "cs"): string | undefined {
  const font = asObject(asObject(fontScheme[group])[slot]);
  return typeof font.typeface === "string" ? font.typeface : undefined;
}

function themeFonts(fontScheme: XmlNode, major: string, minor: string): DocumentTheme["fonts"] {
  const majorEastAsia = themeFontValue(fontScheme, "majorFont", "ea");
  const majorComplexScript = themeFontValue(fontScheme, "majorFont", "cs");
  const minorEastAsia = themeFontValue(fontScheme, "minorFont", "ea");
  const minorComplexScript = themeFontValue(fontScheme, "minorFont", "cs");
  const supplemental = [
    ...themeSupplementalFonts(fontScheme, "majorFont", "major"),
    ...themeSupplementalFonts(fontScheme, "minorFont", "minor"),
  ];

  return {
    major,
    minor,
    ...(majorEastAsia ? { majorEastAsia } : {}),
    ...(majorComplexScript ? { majorComplexScript } : {}),
    ...(minorEastAsia ? { minorEastAsia } : {}),
    ...(minorComplexScript ? { minorComplexScript } : {}),
    ...(supplemental.length > 0 ? { supplemental } : {}),
  };
}

function themeSupplementalFonts(fontScheme: XmlNode, groupTag: "majorFont" | "minorFont", group: "major" | "minor"): NonNullable<DocumentTheme["fonts"]["supplemental"]> {
  return asArray(asObject(fontScheme[groupTag]).font)
    .map((font) => asObject(font))
    .filter((font) => typeof font.script === "string" && typeof font.typeface === "string")
    .map((font) => ({
      group,
      script: font.script as string,
      typeface: font.typeface as string,
    }));
}

function themeColorValue(colorScheme: XmlNode, slot: string): string | undefined {
  const color = asObject(asObject(colorScheme[slot]).srgbClr);
  return typeof color.val === "string" ? color.val : undefined;
}

function themeColors(colorScheme: XmlNode, accent1: string): DocumentTheme["colors"] {
  const dark1 = themeColorValue(colorScheme, "dk1");
  const light1 = themeColorValue(colorScheme, "lt1");
  const dark2 = themeColorValue(colorScheme, "dk2");
  const light2 = themeColorValue(colorScheme, "lt2");
  const accent2 = themeColorValue(colorScheme, "accent2");
  const accent3 = themeColorValue(colorScheme, "accent3");
  const accent4 = themeColorValue(colorScheme, "accent4");
  const accent5 = themeColorValue(colorScheme, "accent5");
  const accent6 = themeColorValue(colorScheme, "accent6");
  const hyperlink = themeColorValue(colorScheme, "hlink");
  const followedHyperlink = themeColorValue(colorScheme, "folHlink");

  return {
    ...(dark1 ? { dark1 } : {}),
    ...(light1 ? { light1 } : {}),
    ...(dark2 ? { dark2 } : {}),
    ...(light2 ? { light2 } : {}),
    accent1,
    ...(accent2 ? { accent2 } : {}),
    ...(accent3 ? { accent3 } : {}),
    ...(accent4 ? { accent4 } : {}),
    ...(accent5 ? { accent5 } : {}),
    ...(accent6 ? { accent6 } : {}),
    ...(hyperlink ? { hyperlink } : {}),
    ...(followedHyperlink ? { followedHyperlink } : {}),
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
  const pagination = parseParagraphPagination(properties);
  const frame = parseParagraphFrame(properties.framePr);
  const tabs = parseParagraphTabs(properties.tabs);
  const parsed = {
    ...(typeof alignment.val === "string" ? { alignment: alignment.val as ParagraphAlignment } : {}),
    ...(spacing ? { spacing } : {}),
    ...(indent ? { indent } : {}),
    ...(shading ? { shading } : {}),
    ...(borders ? { borders } : {}),
    ...(pagination ? { pagination } : {}),
    ...(frame ? { frame } : {}),
    ...(tabs ? { tabs } : {}),
  };

  return Object.keys(parsed).length > 0 ? parsed : undefined;
}

function parseParagraphPagination(properties: XmlNode): NonNullable<ParagraphNode["pagination"]> | undefined {
  const keepNext = parsePaginationToggle(properties.keepNext);
  const keepLines = parsePaginationToggle(properties.keepLines);
  const pageBreakBefore = parsePaginationToggle(properties.pageBreakBefore);
  const widowControl = parsePaginationToggle(properties.widowControl);
  const suppressLineNumbers = parsePaginationToggle(properties.suppressLineNumbers);
  const suppressAutoHyphens = parsePaginationToggle(properties.suppressAutoHyphens);
  const contextualSpacing = parsePaginationToggle(properties.contextualSpacing);
  const mirrorIndents = parsePaginationToggle(properties.mirrorIndents);
  const overflowPunct = parsePaginationToggle(properties.overflowPunct);
  const topLinePunct = parsePaginationToggle(properties.topLinePunct);
  const textAlignment = asObject(properties.textAlignment);
  const textDirection = asObject(properties.textDirection);
  const adjustRightInd = parsePaginationToggle(properties.adjustRightInd);
  const autoSpaceDE = parsePaginationToggle(properties.autoSpaceDE);
  const autoSpaceDN = parsePaginationToggle(properties.autoSpaceDN);
  const pagination = {
    ...(keepNext !== undefined ? { keepNext } : {}),
    ...(keepLines !== undefined ? { keepLines } : {}),
    ...(pageBreakBefore !== undefined ? { pageBreakBefore } : {}),
    ...(widowControl !== undefined ? { widowControl } : {}),
    ...(suppressLineNumbers !== undefined ? { suppressLineNumbers } : {}),
    ...(suppressAutoHyphens !== undefined ? { suppressAutoHyphens } : {}),
    ...(contextualSpacing !== undefined ? { contextualSpacing } : {}),
    ...(mirrorIndents !== undefined ? { mirrorIndents } : {}),
    ...(overflowPunct !== undefined ? { overflowPunct } : {}),
    ...(topLinePunct !== undefined ? { topLinePunct } : {}),
    ...(typeof textAlignment.val === "string" ? { textAlignment: textAlignment.val as NonNullable<ParagraphNode["pagination"]>["textAlignment"] } : {}),
    ...(typeof textDirection.val === "string" ? { textDirection: textDirection.val as NonNullable<ParagraphNode["pagination"]>["textDirection"] } : {}),
    ...(adjustRightInd !== undefined ? { adjustRightInd } : {}),
    ...(autoSpaceDE !== undefined ? { autoSpaceDE } : {}),
    ...(autoSpaceDN !== undefined ? { autoSpaceDN } : {}),
  };

  return Object.keys(pagination).length > 0 ? pagination : undefined;
}

function parsePaginationToggle(value: unknown): boolean | undefined {
  if (value === undefined) return undefined;

  const rawValue = asObject(value).val;
  return rawValue === "0" || rawValue === false || rawValue === "false" || rawValue === "off" ? false : true;
}

function parseParagraphFrame(value: unknown): ParagraphNode["frame"] | undefined {
  const frame = asObject(value);
  const parsed = {
    ...(frame.w !== undefined ? { width: parseNumber(frame.w) } : {}),
    ...(frame.h !== undefined ? { height: parseNumber(frame.h) } : {}),
    ...(frame.x !== undefined ? { x: parseNumber(frame.x) } : {}),
    ...(frame.y !== undefined ? { y: parseNumber(frame.y) } : {}),
    ...(typeof frame.hAnchor === "string" ? { horizontalAnchor: frame.hAnchor as NonNullable<ParagraphNode["frame"]>["horizontalAnchor"] } : {}),
    ...(typeof frame.vAnchor === "string" ? { verticalAnchor: frame.vAnchor as NonNullable<ParagraphNode["frame"]>["verticalAnchor"] } : {}),
    ...(typeof frame.xAlign === "string" ? { xAlign: frame.xAlign as NonNullable<ParagraphNode["frame"]>["xAlign"] } : {}),
    ...(typeof frame.yAlign === "string" ? { yAlign: frame.yAlign as NonNullable<ParagraphNode["frame"]>["yAlign"] } : {}),
    ...(typeof frame.wrap === "string" ? { wrap: frame.wrap as NonNullable<ParagraphNode["frame"]>["wrap"] } : {}),
    ...(typeof frame.dropCap === "string" ? { dropCap: frame.dropCap as NonNullable<ParagraphNode["frame"]>["dropCap"] } : {}),
    ...(frame.lines !== undefined ? { lines: parseNumber(frame.lines) } : {}),
    ...(frame.anchorLock !== undefined ? { anchorLock: parseOnOff(frame.anchorLock) } : {}),
    ...(typeof frame.hRule === "string" ? { heightRule: frame.hRule as NonNullable<ParagraphNode["frame"]>["heightRule"] } : {}),
  };

  return Object.keys(parsed).length > 0 ? parsed : undefined;
}

function parseParagraphTabs(value: unknown): ParagraphNode["tabs"] | undefined {
  const tabs = asArray(asObject(value).tab)
    .map((tabValue) => {
      const tab = asObject(tabValue);
      return typeof tab.val === "string" && tab.pos !== undefined
        ? {
          value: tab.val as NonNullable<ParagraphNode["tabs"]>[number]["value"],
          position: parseNumber(tab.pos),
          ...(typeof tab.leader === "string" ? { leader: tab.leader as NonNullable<ParagraphNode["tabs"]>[number]["leader"] } : {}),
        }
        : undefined;
    })
    .filter((tab): tab is NonNullable<ParagraphNode["tabs"]>[number] => tab !== undefined);

  return tabs.length > 0 ? tabs : undefined;
}

function parseStyleRunProperties(value: unknown): StyleRunProperties | undefined {
  const properties = asObject(value);
  const fonts = asObject(properties.rFonts);
  const size = asObject(properties.sz);
  const complexScriptSize = asObject(properties.szCs);
  const color = asObject(properties.color);
  const highlight = asObject(properties.highlight);
  const language = parseRunLanguage(properties.lang);
  const characterPosition = asObject(properties.position);
  const kerning = asObject(properties.kern);
  const verticalAlign = asObject(properties.vertAlign);
  const characterSpacing = asObject(properties.spacing);
  const scale = asObject(properties.w);
  const fitText = asObject(properties.fitText);
  const emphasis = asObject(properties.em);
  const border = parseBorder(properties.bdr);
  const parsed = {
    ...parseStyleOnOffRunProperty(properties.b, "bold"),
    ...parseStyleOnOffRunProperty(properties.i, "italic"),
    ...parseUnderline(properties.u),
    ...(typeof fonts.ascii === "string" ? { fontFamily: fonts.ascii } : {}),
    ...(typeof fonts.eastAsia === "string" ? { eastAsiaFontFamily: fonts.eastAsia } : {}),
    ...(typeof fonts.cs === "string" ? { complexScriptFontFamily: fonts.cs } : {}),
    ...(typeof fonts.asciiTheme === "string" ? { fontTheme: fonts.asciiTheme } : {}),
    ...(typeof fonts.eastAsiaTheme === "string" ? { eastAsiaFontTheme: fonts.eastAsiaTheme } : {}),
    ...(typeof fonts.cstheme === "string" ? { complexScriptFontTheme: fonts.cstheme } : {}),
    ...(typeof fonts.hint === "string" ? { fontHint: fonts.hint as NonNullable<StyleRunProperties["fontHint"]> } : {}),
    ...(typeof size.val === "number" ? { fontSize: size.val / 2 } : {}),
    ...(typeof size.val === "string" ? { fontSize: Number.parseInt(size.val, 10) / 2 } : {}),
    ...(typeof complexScriptSize.val === "number" ? { complexScriptFontSize: complexScriptSize.val / 2 } : {}),
    ...(typeof complexScriptSize.val === "string" ? { complexScriptFontSize: Number.parseInt(complexScriptSize.val, 10) / 2 } : {}),
    ...(typeof color.val === "string" ? { color: color.val } : {}),
    ...(typeof highlight.val === "string" ? { highlight: highlight.val as NonNullable<StyleRunProperties["highlight"]> } : {}),
    ...parseStyleOnOffRunProperty(properties.strike, "strike"),
    ...parseStyleOnOffRunProperty(properties.dstrike, "doubleStrike"),
    ...parseStyleOnOffRunProperty(properties.smallCaps, "smallCaps"),
    ...parseStyleOnOffRunProperty(properties.caps, "allCaps"),
    ...parseStyleOnOffRunProperty(properties.shadow, "shadow"),
    ...parseStyleOnOffRunProperty(properties.outline, "outline"),
    ...parseStyleOnOffRunProperty(properties.emboss, "emboss"),
    ...parseStyleOnOffRunProperty(properties.imprint, "imprint"),
    ...parseStyleOnOffRunProperty(properties.rtl, "rtl"),
    ...parseStyleOnOffRunProperty(properties.cs, "complexScript"),
    ...parseStyleOnOffRunProperty(properties.specVanish, "specVanish"),
    ...parseStyleOnOffRunProperty(properties.vanish, "hidden"),
    ...parseStyleOnOffRunProperty(properties.webHidden, "webHidden"),
    ...parseStyleOnOffRunProperty(properties.snapToGrid, "snapToGrid"),
    ...parseStyleOnOffRunProperty(properties.noProof, "noProof"),
    ...parseStyleOnOffRunProperty(properties.oMath, "officeMath"),
    ...(language ? { language } : {}),
    ...(characterPosition.val !== undefined ? { characterPosition: parseNumber(characterPosition.val) } : {}),
    ...(kerning.val !== undefined ? { kerning: parseNumber(kerning.val) } : {}),
    ...(typeof verticalAlign.val === "string" ? { verticalAlign: verticalAlign.val as NonNullable<StyleRunProperties["verticalAlign"]> } : {}),
    ...(characterSpacing.val !== undefined ? { characterSpacing: parseNumber(characterSpacing.val) } : {}),
    ...(scale.val !== undefined ? { scale: parseNumber(scale.val) } : {}),
    ...(fitText.val !== undefined
      ? { fitText: { width: parseNumber(fitText.val), ...(fitText.id !== undefined ? { id: parseNumber(fitText.id) } : {}) } }
      : {}),
    ...(typeof emphasis.val === "string" ? { emphasis: emphasis.val as NonNullable<StyleRunProperties["emphasis"]> } : {}),
    ...(border ? { border } : {}),
  };

  return Object.keys(parsed).length > 0 ? parsed : undefined;
}

function parseStyleOnOffRunProperty<K extends keyof StyleRunProperties>(value: unknown, key: K): Partial<Pick<StyleRunProperties, K>> {
  if (value === undefined) {
    return {};
  }
  const val = asObject(value).val;
  return { [key]: !(val === "0" || val === false || val === "false") } as Partial<Pick<StyleRunProperties, K>>;
}

function parseStyleTableProperties(value: unknown): TableStyleDefinition["table"] | undefined {
  const properties = asObject(value);
  const borders = parseTableBorders(properties.tblBorders);

  return borders ? { borders } : undefined;
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
    const pageNumbering = parsePageNumbering(body.sectPr);
    const lineNumbering = parseLineNumbering(body.sectPr);
    const footnoteProperties = parseNoteProperties(body.sectPr, "footnotePr");
    const endnoteProperties = parseNoteProperties(body.sectPr, "endnotePr");
    const noEndnote = parseNoEndnote(body.sectPr);
    const documentGrid = parseDocumentGrid(body.sectPr);
    const verticalAlignment = parseSectionVerticalAlignment(body.sectPr);
    const textDirection = parseSectionTextDirection(body.sectPr);
    const bidi = parseSectionBidi(body.sectPr);
    const rtlGutter = parseRtlGutter(body.sectPr);
    const mirrorMargins = parseMirrorMargins(body.sectPr);
    return [{
      ...(page ? { page } : {}),
      ...(pageNumbering ? { pageNumbering } : {}),
      ...(lineNumbering ? { lineNumbering } : {}),
      ...(footnoteProperties ? { footnoteProperties } : {}),
      ...(endnoteProperties ? { endnoteProperties } : {}),
      ...(noEndnote ? { noEndnote } : {}),
      ...(documentGrid ? { documentGrid } : {}),
      ...(verticalAlignment ? { verticalAlignment } : {}),
      ...(textDirection ? { textDirection } : {}),
      ...(bidi ? { bidi } : {}),
      ...(rtlGutter ? { rtlGutter } : {}),
      ...(mirrorMargins ? { mirrorMargins } : {}),
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
    const pageNumbering = parsePageNumbering(sectPr);
    const lineNumbering = parseLineNumbering(sectPr);
    const footnoteProperties = parseNoteProperties(sectPr, "footnotePr");
    const endnoteProperties = parseNoteProperties(sectPr, "endnotePr");
    const noEndnote = parseNoEndnote(sectPr);
    const documentGrid = parseDocumentGrid(sectPr);
    const verticalAlignment = parseSectionVerticalAlignment(sectPr);
    const textDirection = parseSectionTextDirection(sectPr);
    const bidi = parseSectionBidi(sectPr);
    const rtlGutter = parseRtlGutter(sectPr);
    const mirrorMargins = parseMirrorMargins(sectPr);

    return {
      ...(breakType ? { breakType } : {}),
      ...(page ? { page } : {}),
      ...(pageNumbering ? { pageNumbering } : {}),
      ...(lineNumbering ? { lineNumbering } : {}),
      ...(footnoteProperties ? { footnoteProperties } : {}),
      ...(endnoteProperties ? { endnoteProperties } : {}),
      ...(noEndnote ? { noEndnote } : {}),
      ...(documentGrid ? { documentGrid } : {}),
      ...(verticalAlignment ? { verticalAlignment } : {}),
      ...(textDirection ? { textDirection } : {}),
      ...(bidi ? { bidi } : {}),
      ...(rtlGutter ? { rtlGutter } : {}),
      ...(mirrorMargins ? { mirrorMargins } : {}),
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
): Promise<Pick<SectionNode, "headers" | "footers" | "titlePage" | "watermark">> {
  const sectionProperties = asObject(sectionPropertiesValue);
  const headers = await parseHeaderFooterReferences(zip, sectionProperties.headerReference, relationships, comments);
  const footers = await parseHeaderFooterReferences(zip, sectionProperties.footerReference, relationships, comments);
  const watermark = await parseWatermarkFromHeaderReferences(zip, sectionProperties.headerReference, relationships);

  return {
    ...(sectionProperties.titlePg !== undefined ? { titlePage: true } : {}),
    ...(watermark ? { watermark } : {}),
    ...(headers ? { headers } : {}),
    ...(footers ? { footers } : {}),
  };
}

async function parseHeaderFooterReferences(
  zip: JSZip,
  value: unknown,
  relationships: RelationshipMap,
  comments: CommentMap,
): Promise<SectionNode["headers"] | undefined> {
  const entries = await Promise.all(asArray(value).map(async (referenceValue) => {
    const reference = asObject(referenceValue);
    const blocks = await parseHeaderFooterReference(zip, reference, relationships, comments);
    const type = headerFooterReferenceType(reference.type);

    return blocks && blocks.length > 0 ? { type, blocks } : undefined;
  }));
  const parsed = entries.reduce<NonNullable<SectionNode["headers"]>>((accumulator, entry) => {
    if (entry) {
      accumulator[entry.type] = entry.blocks;
    }

    return accumulator;
  }, {});

  return Object.keys(parsed).length > 0 ? parsed : undefined;
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

  return asArray(root.p)
    .filter((paragraph) => !isWatermarkParagraph(paragraph))
    .map((paragraph) => parseParagraph(paragraph, relationships, comments));
}

async function parseWatermarkFromHeaderReferences(
  zip: JSZip,
  value: unknown,
  relationships: RelationshipMap,
): Promise<SectionNode["watermark"] | undefined> {
  for (const referenceValue of asArray(value)) {
    const reference = asObject(referenceValue);
    if (headerFooterReferenceType(reference.type) !== "default" || typeof reference.id !== "string") {
      continue;
    }

    const target = relationships[reference.id];
    const file = target ? zip.file(`word/${target}`) : undefined;
    if (!file) {
      continue;
    }

    const watermark = parseWatermarkXml(await file.async("string"));
    if (watermark) {
      return watermark;
    }
  }

  return undefined;
}

function parseWatermarkXml(xml: string): SectionNode["watermark"] | undefined {
  const text = xml.match(/<v:textpath\b[^>]*\bstring="([^"]*)"/)?.[1];
  if (!text) {
    return undefined;
  }

  const shape = xml.match(/<v:shape\b[^>]*>/)?.[0] ?? "";
  const fill = xml.match(/<v:fill\b[^>]*\bopacity="([^"]*)"/)?.[1];
  const fontFamily = xml.match(/font-family:&quot;([^&]*)&quot;/)?.[1];
  const color = shape.match(/\bfillcolor="#?([^"\s]*)"/)?.[1];
  const rotation = shape.match(/rotation:([^;"]+)/)?.[1];

  return {
    text: unescapeXml(text),
    ...(color ? { color } : {}),
    ...(fill !== undefined ? { opacity: Number.parseFloat(fill) } : {}),
    ...(rotation !== undefined ? { rotation: Number.parseFloat(rotation) } : {}),
    ...(fontFamily ? { fontFamily: unescapeXml(fontFamily) } : {}),
  };
}

function isWatermarkParagraph(value: unknown): boolean {
  return asArray(asObject(value).r).some((run) => asObject(run).pict !== undefined);
}

function headerFooterReferenceType(value: unknown): keyof NonNullable<SectionNode["headers"]> {
  return value === "first" || value === "even" ? value : "default";
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

function parsePageNumbering(sectionPropertiesValue: unknown): SectionNode["pageNumbering"] | undefined {
  const pageNumbering = asObject(asObject(sectionPropertiesValue).pgNumType);
  const parsed = {
    ...(pageNumbering.start !== undefined ? { start: parseNumber(pageNumbering.start) } : {}),
    ...(typeof pageNumbering.fmt === "string" ? { format: pageNumbering.fmt as NonNullable<SectionNode["pageNumbering"]>["format"] } : {}),
    ...(pageNumbering.chapStyle !== undefined ? { chapterStyle: parseNumber(pageNumbering.chapStyle) } : {}),
    ...(typeof pageNumbering.chapSep === "string" ? { chapterSeparator: pageNumbering.chapSep as NonNullable<SectionNode["pageNumbering"]>["chapterSeparator"] } : {}),
  };

  return Object.keys(parsed).length > 0 ? parsed : undefined;
}

function parseLineNumbering(sectionPropertiesValue: unknown): SectionNode["lineNumbering"] | undefined {
  const lineNumbering = asObject(asObject(sectionPropertiesValue).lnNumType);
  const parsed = {
    ...(lineNumbering.start !== undefined ? { start: parseNumber(lineNumbering.start) } : {}),
    ...(lineNumbering.countBy !== undefined ? { countBy: parseNumber(lineNumbering.countBy) } : {}),
    ...(lineNumbering.distance !== undefined ? { distance: parseNumber(lineNumbering.distance) } : {}),
    ...(typeof lineNumbering.restart === "string" ? { restart: lineNumbering.restart as NonNullable<SectionNode["lineNumbering"]>["restart"] } : {}),
  };

  return Object.keys(parsed).length > 0 ? parsed : undefined;
}

function parseNoteProperties(sectionPropertiesValue: unknown, key: "footnotePr" | "endnotePr"): SectionNode["footnoteProperties"] | undefined {
  const properties = asObject(asObject(sectionPropertiesValue)[key]);
  const position = asObject(properties.pos);
  const format = asObject(properties.numFmt);
  const start = asObject(properties.numStart);
  const restart = asObject(properties.numRestart);
  const numbering = {
    ...(typeof format.val === "string" ? { format: format.val as NonNullable<NonNullable<SectionNode["footnoteProperties"]>["numbering"]>["format"] } : {}),
    ...(start.val !== undefined ? { start: parseNumber(start.val) } : {}),
    ...(typeof restart.val === "string" ? { restart: restart.val as NonNullable<NonNullable<SectionNode["footnoteProperties"]>["numbering"]>["restart"] } : {}),
  };
  const parsed = {
    ...(typeof position.val === "string" ? { position: position.val as NonNullable<SectionNode["footnoteProperties"]>["position"] } : {}),
    ...(Object.keys(numbering).length > 0 ? { numbering } : {}),
  };

  return Object.keys(parsed).length > 0 ? parsed : undefined;
}

function parseNoEndnote(sectionPropertiesValue: unknown): boolean | undefined {
  return asObject(sectionPropertiesValue).noEndnote !== undefined ? true : undefined;
}

function parseDocumentGrid(sectionPropertiesValue: unknown): SectionNode["documentGrid"] | undefined {
  const documentGrid = asObject(asObject(sectionPropertiesValue).docGrid);
  const parsed = {
    ...(typeof documentGrid.type === "string" ? { type: documentGrid.type as NonNullable<SectionNode["documentGrid"]>["type"] } : {}),
    ...(documentGrid.linePitch !== undefined ? { linePitch: parseNumber(documentGrid.linePitch) } : {}),
    ...(documentGrid.charSpace !== undefined ? { charSpace: parseNumber(documentGrid.charSpace) } : {}),
  };

  return Object.keys(parsed).length > 0 ? parsed : undefined;
}

function parseSectionVerticalAlignment(sectionPropertiesValue: unknown): SectionNode["verticalAlignment"] | undefined {
  const verticalAlignment = asObject(asObject(sectionPropertiesValue).vAlign);

  return typeof verticalAlignment.val === "string"
    ? verticalAlignment.val as NonNullable<SectionNode["verticalAlignment"]>
    : undefined;
}

function parseSectionTextDirection(sectionPropertiesValue: unknown): SectionNode["textDirection"] | undefined {
  const textDirection = asObject(asObject(sectionPropertiesValue).textDirection);

  return typeof textDirection.val === "string"
    ? textDirection.val as NonNullable<SectionNode["textDirection"]>
    : undefined;
}

function parseSectionBidi(sectionPropertiesValue: unknown): boolean | undefined {
  return asObject(sectionPropertiesValue).bidi !== undefined ? true : undefined;
}

function parseRtlGutter(sectionPropertiesValue: unknown): boolean | undefined {
  return asObject(sectionPropertiesValue).rtlGutter !== undefined ? true : undefined;
}

function parseMirrorMargins(sectionPropertiesValue: unknown): boolean | undefined {
  return asObject(sectionPropertiesValue).mirrorMargins !== undefined ? true : undefined;
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
  const frame = parseParagraphFrame(properties.framePr);
  const tabs = parseParagraphTabs(properties.tabs);
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
    ...(tabs ? { tabs } : {}),
    ...(numbering ? { list: numbering } : {}),
    ...(pagination ? { pagination } : {}),
    ...(frame ? { frame } : {}),
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

function parseTableBorders(value: unknown): TableNode["borders"] | undefined {
  const borders = asObject(value);
  const parsed = {
    ...parseTableBorderSide(borders.top, "top"),
    ...parseTableBorderSide(borders.left, "left"),
    ...parseTableBorderSide(borders.bottom, "bottom"),
    ...parseTableBorderSide(borders.right, "right"),
    ...parseTableBorderSide(borders.insideH, "insideH"),
    ...parseTableBorderSide(borders.insideV, "insideV"),
  };

  if (Object.keys(parsed).length === 0) {
    return undefined;
  }

  return isDefaultSingleTableBorders(parsed) ? "single" : parsed;
}

function parseTableBorderSide(value: unknown, side: keyof Exclude<NonNullable<TableNode["borders"]>, string>): Partial<Exclude<NonNullable<TableNode["borders"]>, string>> {
  const border = parseBorder(value);

  return border ? { [side]: border } : {};
}

function isDefaultSingleTableBorders(borders: Exclude<NonNullable<TableNode["borders"]>, string>): boolean {
  const sides: (keyof typeof borders)[] = ["top", "left", "bottom", "right", "insideH", "insideV"];
  return sides.every((side) => {
    const border = borders[side];
    return border?.style === "single" &&
      border.size === 4 &&
      border.space === 0 &&
      border.color === "auto";
  });
}

function parseParagraphBorderSide(value: unknown, side: keyof NonNullable<ParagraphNode["borders"]>): Partial<NonNullable<ParagraphNode["borders"]>> {
  const border = parseBorder(value);

  return border ? { [side]: border } : {};
}

function parseBorder(value: unknown): NonNullable<TextRun["border"]> | undefined {
  const border = asObject(value);
  const style = parseBorderStyle(border.val);

  if (!style) {
    return undefined;
  }

  return {
    style,
    ...(border.sz !== undefined ? { size: parseNumber(border.sz) } : {}),
    ...(typeof border.color === "string" ? { color: border.color } : {}),
    ...(border.space !== undefined ? { space: parseNumber(border.space) } : {}),
  };
}

function parseBorderStyle(value: unknown): NonNullable<TextRun["border"]>["style"] | undefined {
  return value === "single" ||
    value === "double" ||
    value === "dashed" ||
    value === "dotted" ||
    value === "nil" ||
    value === "none"
    ? value
    : undefined;
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
  const keepNext = parsePaginationToggle(properties.keepNext);
  const keepLines = parsePaginationToggle(properties.keepLines);
  const pageBreakBefore = parsePaginationToggle(properties.pageBreakBefore);
  const widowControl = parsePaginationToggle(properties.widowControl);
  const suppressLineNumbers = parsePaginationToggle(properties.suppressLineNumbers);
  const suppressAutoHyphens = parsePaginationToggle(properties.suppressAutoHyphens);
  const contextualSpacing = parsePaginationToggle(properties.contextualSpacing);
  const mirrorIndents = parsePaginationToggle(properties.mirrorIndents);
  const overflowPunct = parsePaginationToggle(properties.overflowPunct);
  const topLinePunct = parsePaginationToggle(properties.topLinePunct);
  const textAlignment = asObject(properties.textAlignment);
  const textDirection = asObject(properties.textDirection);
  const adjustRightInd = parsePaginationToggle(properties.adjustRightInd);
  const autoSpaceDE = parsePaginationToggle(properties.autoSpaceDE);
  const autoSpaceDN = parsePaginationToggle(properties.autoSpaceDN);
  const pagination = {
    ...(keepNext !== undefined ? { keepNext } : {}),
    ...(keepLines !== undefined ? { keepLines } : {}),
    ...(pageBreakBefore !== undefined ? { pageBreakBefore } : {}),
    ...(widowControl !== undefined ? { widowControl } : {}),
    ...(suppressLineNumbers !== undefined ? { suppressLineNumbers } : {}),
    ...(suppressAutoHyphens !== undefined ? { suppressAutoHyphens } : {}),
    ...(contextualSpacing !== undefined ? { contextualSpacing } : {}),
    ...(mirrorIndents !== undefined ? { mirrorIndents } : {}),
    ...(overflowPunct !== undefined ? { overflowPunct } : {}),
    ...(topLinePunct !== undefined ? { topLinePunct } : {}),
    ...(typeof textAlignment.val === "string" ? { textAlignment: textAlignment.val as NonNullable<ParagraphNode["pagination"]>["textAlignment"] } : {}),
    ...(typeof textDirection.val === "string" ? { textDirection: textDirection.val as NonNullable<ParagraphNode["pagination"]>["textDirection"] } : {}),
    ...(adjustRightInd !== undefined ? { adjustRightInd } : {}),
    ...(autoSpaceDE !== undefined ? { autoSpaceDE } : {}),
    ...(autoSpaceDN !== undefined ? { autoSpaceDN } : {}),
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
  if (paragraphXml?.includes("<m:oMath")) {
    return parseParagraphRunsWithMath(paragraphXml, footnotes, endnotes);
  }

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

function parseParagraphRunsWithMath(paragraphXml: string, footnotes: NoteMap, endnotes: NoteMap): TextRun[] {
  const runs: TextRun[] = [];
  const tokenPattern = /<w:r\b[\s\S]*?<\/w:r>|<m:oMath\b[\s\S]*?<\/m:oMath>/g;
  let match: RegExpExecArray | null;

  while ((match = tokenPattern.exec(paragraphXml)) !== null) {
    const token = match[0];
    if (token.startsWith("<m:oMath")) {
      const mathRun = parseMathRunXml(token);
      if (mathRun) {
        runs.push(mathRun);
      }
      continue;
    }

    const parsed = parser.parse(
      `<root xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${token}</root>`,
    ) as XmlNode;
    const run = withMatchingNotes(parseRun(asObject(parsed.root).r), footnotes, endnotes);
    if (run.text !== "" || run.break !== undefined || run.field !== undefined || run.footnote !== undefined || run.endnote !== undefined) {
      runs.push(run);
    }
  }

  return runs;
}

function parseMathRunXml(xml: string): TextRun | undefined {
  const parsed = parser.parse(
    `<root xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">${xml}</root>`,
  ) as XmlNode;
  const math = asObject(asObject(parsed.root).oMath);
  const nodes = parseMathNodes(math);

  if (nodes.some((node) => node.type !== "text")) {
    return { text: "", math: { nodes } };
  }

  const text = asArray(math.r)
    .map((run) => parseText(asObject(run).t))
    .join("");

  return text ? { text: "", math: { text } } : undefined;
}

function parseMathNodes(container: XmlNode): NonNullable<NonNullable<TextRun["math"]>["nodes"]> {
  return [
    ...asArray(container.r)
      .map((run) => {
        const text = parseText(asObject(run).t);
        return text ? { type: "text" as const, text } : undefined;
      })
      .filter((node): node is { type: "text"; text: string } => node !== undefined),
    ...asArray(container.f)
      .map((fraction) => {
        const fractionNode = asObject(fraction);
        const fractionProperties = asObject(fractionNode.fPr);
        const controlProperties = parseMathControlProperties(asObject(fractionProperties.ctrlPr).rPr);
        const fractionType = fractionTypeValue(asObject(fractionProperties.type).val);
        return {
          type: "fraction" as const,
          ...(fractionType ? { fractionType } : {}),
          ...(controlProperties ? { controlProperties } : {}),
          numerator: parseMathNodes(asObject(fractionNode.num)),
          denominator: parseMathNodes(asObject(fractionNode.den)),
        };
      }),
    ...asArray(container.sSup)
      .map((superscript) => {
        const superscriptNode = asObject(superscript);
        const controlProperties = parseMathControlProperties(asObject(asObject(superscriptNode.sSupPr).ctrlPr).rPr);
        return {
          type: "superscript" as const,
          ...(controlProperties ? { controlProperties } : {}),
          base: parseMathNodes(asObject(superscriptNode.e)),
          superscript: parseMathNodes(asObject(superscriptNode.sup)),
        };
      }),
    ...asArray(container.sSub)
      .map((subscript) => {
        const subscriptNode = asObject(subscript);
        const controlProperties = parseMathControlProperties(asObject(asObject(subscriptNode.sSubPr).ctrlPr).rPr);
        return {
          type: "subscript" as const,
          ...(controlProperties ? { controlProperties } : {}),
          base: parseMathNodes(asObject(subscriptNode.e)),
          subscript: parseMathNodes(asObject(subscriptNode.sub)),
        };
      }),
    ...asArray(container.sSubSup)
      .map((subSup) => {
        const subSupNode = asObject(subSup);
        const controlProperties = parseMathControlProperties(asObject(asObject(subSupNode.sSubSupPr).ctrlPr).rPr);
        return {
          type: "subSup" as const,
          ...(controlProperties ? { controlProperties } : {}),
          base: parseMathNodes(asObject(subSupNode.e)),
          subscript: parseMathNodes(asObject(subSupNode.sub)),
          superscript: parseMathNodes(asObject(subSupNode.sup)),
        };
      }),
    ...asArray(container.sPre)
      .map((sPre) => {
        const sPreNode = asObject(sPre);
        const controlProperties = parseMathControlProperties(asObject(asObject(sPreNode.sPrePr).ctrlPr).rPr);
        return {
          type: "sPre" as const,
          ...(controlProperties ? { controlProperties } : {}),
          base: parseMathNodes(asObject(sPreNode.e)),
          subscript: parseMathNodes(asObject(sPreNode.sub)),
          superscript: parseMathNodes(asObject(sPreNode.sup)),
        };
      }),
    ...asArray(container.preSubSup)
      .map((preSubSup) => {
        const preSubSupNode = asObject(preSubSup);
        const controlProperties = parseMathControlProperties(asObject(asObject(preSubSupNode.preSubSupPr).ctrlPr).rPr);
        return {
          type: "preSubSup" as const,
          ...(controlProperties ? { controlProperties } : {}),
          base: parseMathNodes(asObject(preSubSupNode.e)),
          subscript: parseMathNodes(asObject(preSubSupNode.sub)),
          superscript: parseMathNodes(asObject(preSubSupNode.sup)),
        };
      }),
    ...asArray(container.rad)
      .map((radical) => {
        const radicalNode = asObject(radical);
        const radicalProperties = asObject(radicalNode.radPr);
        const controlProperties = parseMathControlProperties(asObject(radicalProperties.ctrlPr).rPr);
        const degree = parseMathNodes(asObject(radicalNode.deg));
        return {
          type: "radical" as const,
          ...(controlProperties ? { controlProperties } : {}),
          ...mathOptionalBooleanProperty(radicalProperties.degHide, "hideDegree"),
          ...(degree.length > 0 ? { degree } : {}),
          content: parseMathNodes(asObject(radicalNode.e)),
        };
      }),
    ...asArray(container.nary)
      .map((nary) => {
        const naryNode = asObject(nary);
        const lowerLimit = parseMathNodes(asObject(naryNode.sub));
        const upperLimit = parseMathNodes(asObject(naryNode.sup));
        const naryProperties = asObject(naryNode.naryPr);
        const controlProperties = parseMathControlProperties(asObject(naryProperties.ctrlPr).rPr);
        const limitLocation = naryLimitLocationValue(asObject(naryProperties.limLoc).val);
        const operatorCharacter = asObject(naryProperties.chr).val;
        const operator = naryOperatorValue(operatorCharacter);
        return {
          type: "nary" as const,
          operator,
          ...(typeof operatorCharacter === "string" && operatorCharacter !== naryOperatorCharacter(operator) ? { operatorCharacter } : {}),
          ...(controlProperties ? { controlProperties } : {}),
          ...(limitLocation ? { limitLocation } : {}),
          ...mathOptionalBooleanProperty(naryProperties.grow, "grow"),
          ...mathOptionalBooleanProperty(naryProperties.subHide, "hideLowerLimit"),
          ...mathOptionalBooleanProperty(naryProperties.supHide, "hideUpperLimit"),
          ...(lowerLimit.length > 0 ? { lowerLimit } : {}),
          ...(upperLimit.length > 0 ? { upperLimit } : {}),
          body: parseMathNodes(asObject(naryNode.e)),
        };
      }),
    ...asArray(container.m)
      .map((matrix) => {
        const matrixNode = asObject(matrix);
        const matrixProperties = asObject(matrixNode.mPr);
        const baseJustification = matrixBaseJustificationValue(asObject(matrixProperties.baseJc).val);
        const rowSpacing = asObject(matrixProperties.rSp).val;
        const rowSpacingRule = matrixSpacingRuleValue(asObject(matrixProperties.rSpRule).val);
        const columnSpacing = asObject(matrixProperties.cSp).val;
        const columnSpacingRule = matrixSpacingRuleValue(asObject(matrixProperties.cSpRule).val);
        const columnJustifications = matrixColumnJustificationValues(matrixProperties.mcs);
        const columnCounts = matrixColumnCountValues(matrixProperties.mcs);
        const controlProperties = parseMathControlProperties(asObject(matrixProperties.ctrlPr).rPr);
        return {
          type: "matrix" as const,
          ...(baseJustification ? { baseJustification } : {}),
          ...(rowSpacing !== undefined ? { rowSpacing: parseNumber(rowSpacing) } : {}),
          ...(rowSpacingRule ? { rowSpacingRule } : {}),
          ...(columnSpacing !== undefined ? { columnSpacing: parseNumber(columnSpacing) } : {}),
          ...(columnSpacingRule ? { columnSpacingRule } : {}),
          ...(columnJustifications.length > 0 ? { columnJustifications } : {}),
          ...(columnCounts.length > 0 ? { columnCounts } : {}),
          ...(controlProperties ? { controlProperties } : {}),
          rows: asArray(matrixNode.mr).map((row) =>
            asArray(asObject(row).e).map((cell) => parseMathNodes(asObject(cell))),
          ),
        };
      }),
    ...asArray(container.d)
      .map((delimiter) => {
        const delimiterNode = asObject(delimiter);
        const delimiterProperties = asObject(delimiterNode.dPr);
        const begin = asObject(delimiterProperties.begChr).val;
        const end = asObject(delimiterProperties.endChr).val;
        const separator = asObject(delimiterProperties.sepChr).val;
        const controlProperties = parseMathControlProperties(asObject(delimiterProperties.ctrlPr).rPr);
        return {
          type: "delimiter" as const,
          ...(typeof begin === "string" ? { begin } : {}),
          ...(typeof end === "string" ? { end } : {}),
          ...mathOptionalBooleanProperty(delimiterProperties.grow, "grow"),
          ...(typeof separator === "string" ? { separator } : {}),
          ...(controlProperties ? { controlProperties } : {}),
          content: parseMathNodes(asObject(delimiterNode.e)),
        };
      }),
    ...asArray(container.acc)
      .map((accent) => {
        const accentNode = asObject(accent);
        const accentProperties = asObject(accentNode.accPr);
        const mark = asObject(accentProperties.chr).val;
        const controlProperties = parseMathControlProperties(asObject(accentProperties.ctrlPr).rPr);
        return {
          type: "accent" as const,
          mark: typeof mark === "string" ? mark : "",
          ...(controlProperties ? { controlProperties } : {}),
          content: parseMathNodes(asObject(accentNode.e)),
        };
      }),
    ...asArray(container.bar)
      .map((bar) => {
        const barNode = asObject(bar);
        const barProperties = asObject(barNode.barPr);
        const controlProperties = parseMathControlProperties(asObject(barProperties.ctrlPr).rPr);
        return {
          type: "bar" as const,
          position: barPositionValue(asObject(barProperties.pos).val),
          ...(controlProperties ? { controlProperties } : {}),
          content: parseMathNodes(asObject(barNode.e)),
        };
      }),
    ...asArray(container.func)
      .map((func) => {
        const funcNode = asObject(func);
        const controlProperties = parseMathControlProperties(asObject(asObject(funcNode.funcPr).ctrlPr).rPr);
        return {
          type: "function" as const,
          ...(controlProperties ? { controlProperties } : {}),
          name: parseMathNodes(asObject(funcNode.fName)),
          argument: parseMathNodes(asObject(funcNode.e)),
        };
      }),
    ...asArray(container.limLow)
      .map((limitLower) => {
        const limitNode = asObject(limitLower);
        const controlProperties = parseMathControlProperties(asObject(asObject(limitNode.limLowPr).ctrlPr).rPr);
        return {
          type: "limitLower" as const,
          ...(controlProperties ? { controlProperties } : {}),
          base: parseMathNodes(asObject(limitNode.e)),
          limit: parseMathNodes(asObject(limitNode.lim)),
        };
      }),
    ...asArray(container.limUpp)
      .map((limitUpper) => {
        const limitNode = asObject(limitUpper);
        const controlProperties = parseMathControlProperties(asObject(asObject(limitNode.limUppPr).ctrlPr).rPr);
        return {
          type: "limitUpper" as const,
          ...(controlProperties ? { controlProperties } : {}),
          base: parseMathNodes(asObject(limitNode.e)),
          limit: parseMathNodes(asObject(limitNode.lim)),
        };
      }),
    ...asArray(container.eqArr)
      .map((equationArray) => {
        const equationArrayNode = asObject(equationArray);
        const equationArrayProperties = asObject(equationArrayNode.eqArrPr);
        const baseJustification = matrixBaseJustificationValue(asObject(equationArrayProperties.baseJc).val);
        const verticalJustification = groupCharacterVerticalJustificationValue(asObject(equationArrayProperties.vertJc).val);
        const rowSpacing = asObject(equationArrayProperties.rSp).val;
        const rowSpacingRule = matrixSpacingRuleValue(asObject(equationArrayProperties.rSpRule).val);
        const controlProperties = parseMathControlProperties(asObject(equationArrayProperties.ctrlPr).rPr);
        return {
          type: "equationArray" as const,
          ...(baseJustification ? { baseJustification } : {}),
          ...(verticalJustification ? { verticalJustification } : {}),
          ...mathOptionalBooleanProperty(equationArrayProperties.aln, "alignment"),
          ...(rowSpacing !== undefined ? { rowSpacing: parseNumber(rowSpacing) } : {}),
          ...(rowSpacingRule ? { rowSpacingRule } : {}),
          ...mathOptionalBooleanProperty(equationArrayProperties.objDist, "objectDistribution"),
          ...mathOptionalBooleanProperty(equationArrayProperties.maxDist, "maxDistribution"),
          ...(controlProperties ? { controlProperties } : {}),
          rows: asArray(equationArrayNode.e).map((row) => parseMathNodes(asObject(row))),
        };
      }),
    ...asArray(container.box)
      .map((box) => {
        const boxNode = asObject(box);
        const boxProperties = asObject(boxNode.boxPr);
        const controlProperties = parseMathControlProperties(asObject(boxProperties.ctrlPr).rPr);
        return {
          type: "box" as const,
          ...mathOptionalBooleanProperty(boxProperties.hideTop, "hideTop"),
          ...mathOptionalBooleanProperty(boxProperties.hideBot, "hideBottom"),
          ...mathOptionalBooleanProperty(boxProperties.hideLeft, "hideLeft"),
          ...mathOptionalBooleanProperty(boxProperties.hideRight, "hideRight"),
          ...(controlProperties ? { controlProperties } : {}),
          content: parseMathNodes(asObject(boxNode.e)),
        };
      }),
    ...asArray(container.borderBox)
      .map((borderBox) => {
        const borderBoxNode = asObject(borderBox);
        const borderBoxProperties = asObject(borderBoxNode.borderBoxPr);
        const controlProperties = parseMathControlProperties(asObject(borderBoxProperties.ctrlPr).rPr);
        return {
          type: "borderBox" as const,
          ...mathOptionalBooleanProperty(borderBoxProperties.hideTop, "hideTop"),
          ...mathOptionalBooleanProperty(borderBoxProperties.hideBot, "hideBottom"),
          ...mathOptionalBooleanProperty(borderBoxProperties.hideLeft, "hideLeft"),
          ...mathOptionalBooleanProperty(borderBoxProperties.hideRight, "hideRight"),
          ...(controlProperties ? { controlProperties } : {}),
          content: parseMathNodes(asObject(borderBoxNode.e)),
        };
      }),
    ...asArray(container.phant)
      .map((phantom) => {
        const phantomNode = asObject(phantom);
        const phantomProperties = asObject(phantomNode.phantPr);
        const controlProperties = parseMathControlProperties(asObject(phantomProperties.ctrlPr).rPr);
        return {
          type: "phantom" as const,
          ...mathOptionalBooleanProperty(phantomProperties.show, "show"),
          ...mathOptionalBooleanProperty(phantomProperties.zeroWid, "zeroWidth"),
          ...mathOptionalBooleanProperty(phantomProperties.zeroAsc, "zeroAscent"),
          ...mathOptionalBooleanProperty(phantomProperties.zeroDesc, "zeroDescent"),
          ...mathOptionalBooleanProperty(phantomProperties.transp, "transparent"),
          ...(controlProperties ? { controlProperties } : {}),
          content: parseMathNodes(asObject(phantomNode.e)),
        };
      }),
    ...asArray(container.groupChr)
      .map((groupCharacter) => {
        const groupCharacterNode = asObject(groupCharacter);
        const groupCharacterProperties = asObject(groupCharacterNode.groupChrPr);
        const character = asObject(groupCharacterProperties.chr).val;
        const position = groupCharacterPositionValue(asObject(groupCharacterProperties.pos).val);
        const verticalJustification = groupCharacterVerticalJustificationValue(asObject(groupCharacterProperties.vertJc).val);
        const controlProperties = parseMathControlProperties(asObject(groupCharacterProperties.ctrlPr).rPr);
        return {
          type: "groupCharacter" as const,
          ...(typeof character === "string" ? { character } : {}),
          ...(position ? { position } : {}),
          ...(verticalJustification ? { verticalJustification } : {}),
          ...(controlProperties ? { controlProperties } : {}),
          content: parseMathNodes(asObject(groupCharacterNode.e)),
        };
      }),
  ];
}

function parseMathControlProperties(value: unknown): MathControlProperties | undefined {
  const properties = asObject(value);
  const parsed: MathControlProperties = {
    ...parseRunStyle(properties),
    ...(properties.b !== undefined ? { bold: true } : {}),
    ...(properties.i !== undefined ? { italic: true } : {}),
    ...(properties.u !== undefined ? { underline: true } : {}),
    ...parseRunFont(properties),
  };

  return Object.keys(parsed).length > 0 ? parsed : undefined;
}

function fractionTypeValue(value: unknown): Extract<MathNode, { type: "fraction" }>["fractionType"] | undefined {
  if (value === "skw") {
    return "skewed";
  }
  if (value === "lin") {
    return "linear";
  }
  return value === "bar" || value === "noBar" ? value : undefined;
}

function groupCharacterPositionValue(value: unknown): "top" | "bottom" | undefined {
  return value === "top" || value === "bottom" ? value : undefined;
}

function groupCharacterVerticalJustificationValue(value: unknown): "top" | "bottom" | undefined {
  return value === "top" || value === "bottom" ? value : undefined;
}

function mathOptionalBooleanProperty<K extends string>(node: unknown, key: K): Partial<Record<K, boolean>> {
  if (node === undefined) {
    return {};
  }
  return { [key]: mathBooleanValue(node) } as Partial<Record<K, boolean>>;
}

function mathBooleanProperty<K extends string>(node: unknown, key: K): Partial<Record<K, true>> {
  if (mathBooleanValue(node)) {
    return { [key]: true } as Partial<Record<K, true>>;
  }
  return {};
}

function mathBooleanValue(node: unknown): boolean {
  if (node === undefined) {
    return false;
  }
  const value = asObject(node).val;
  return value === undefined || value === "1" || value === true;
}

function barPositionValue(value: unknown): "top" | "bottom" {
  return value === "bottom" ? "bottom" : "top";
}

function naryOperatorValue(value: unknown): Extract<MathNode, { type: "nary" }>["operator"] {
  const values = {
    "∑": "sum",
    "∫": "integral",
    "∏": "product",
    "∐": "coproduct",
    "⋂": "intersection",
    "⋃": "union",
  } as const;
  return typeof value === "string" && value in values ? values[value as keyof typeof values] : "sum";
}

function naryOperatorCharacter(operator: Extract<MathNode, { type: "nary" }>["operator"]): string {
  const values = {
    sum: "∑",
    integral: "∫",
    product: "∏",
    coproduct: "∐",
    intersection: "⋂",
    union: "⋃",
  } satisfies Record<Extract<MathNode, { type: "nary" }>["operator"], string>;
  return values[operator];
}

function naryLimitLocationValue(value: unknown): Extract<MathNode, { type: "nary" }>["limitLocation"] | undefined {
  if (value === "undOvr") {
    return "underOver";
  }
  return value === "subSup" ? "subSup" : undefined;
}

function matrixBaseJustificationValue(value: unknown): Extract<MathNode, { type: "matrix" }>["baseJustification"] | undefined {
  if (value === "top" || value === "center") {
    return value;
  }
  if (value === "bot") {
    return "bottom";
  }
  return undefined;
}

function matrixColumnJustificationValues(node: unknown): NonNullable<Extract<MathNode, { type: "matrix" }>["columnJustifications"]> {
  return asArray(asObject(node).mc)
    .map((column) => matrixColumnJustificationValue(asObject(asObject(column).mcPr).mcJc))
    .filter((value): value is "left" | "center" | "right" => value !== undefined);
}

function matrixColumnJustificationValue(node: unknown): "left" | "center" | "right" | undefined {
  const value = asObject(node).val;
  return value === "left" || value === "center" || value === "right" ? value : undefined;
}

function matrixColumnCountValues(node: unknown): number[] {
  return asArray(asObject(node).mc)
    .map((column) => {
      const count = asObject(asObject(column).mcPr).count;
      const value = asObject(count).val;
      return value !== undefined ? parseNumber(value) : undefined;
    })
    .filter((value): value is number => value !== undefined);
}

function matrixSpacingRuleValue(value: unknown): "single" | "oneAndHalf" | "double" | "exactly" | "multiple" | undefined {
  return value === "single" || value === "oneAndHalf" || value === "double" || value === "exactly" || value === "multiple"
    ? value
    : undefined;
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
    ...parseOnOffRunProperty(properties.b, "bold"),
    ...parseOnOffRunProperty(properties.i, "italic"),
    ...parseUnderline(properties.u),
    ...parseRunFont(properties),
  };
}

function parseRunStyle(properties: XmlNode): Partial<TextRun> {
  const style = asObject(properties.rStyle);

  return typeof style.val === "string" ? { styleId: style.val } : {};
}

function parseUnderline(value: unknown): Partial<Pick<TextRun, "underline">> {
  if (value === undefined) {
    return {};
  }
  const val = asObject(value).val;
  return { underline: !(val === "none" || val === "0" || val === false) };
}

function parseOnOffRunProperty<K extends keyof TextRun>(value: unknown, key: K): Partial<Pick<TextRun, K>> {
  if (value === undefined) {
    return {};
  }
  const val = asObject(value).val;
  return { [key]: !(val === "0" || val === false || val === "false") } as Partial<Pick<TextRun, K>>;
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
  const complexScriptSize = asObject(properties.szCs);
  const color = asObject(properties.color);
  const highlight = asObject(properties.highlight);
  const verticalAlign = asObject(properties.vertAlign);
  const characterPosition = asObject(properties.position);
  const kerning = asObject(properties.kern);
  const characterSpacing = asObject(properties.spacing);
  const scale = asObject(properties.w);
  const fitText = asObject(properties.fitText);
  const emphasis = asObject(properties.em);
  const language = parseRunLanguage(properties.lang);
  const border = parseBorder(properties.bdr);

  return {
    ...(typeof fonts.ascii === "string" ? { fontFamily: fonts.ascii } : {}),
    ...(typeof fonts.eastAsia === "string" ? { eastAsiaFontFamily: fonts.eastAsia } : {}),
    ...(typeof fonts.cs === "string" ? { complexScriptFontFamily: fonts.cs } : {}),
    ...(typeof fonts.asciiTheme === "string" ? { fontTheme: fonts.asciiTheme } : {}),
    ...(typeof fonts.eastAsiaTheme === "string" ? { eastAsiaFontTheme: fonts.eastAsiaTheme } : {}),
    ...(typeof fonts.cstheme === "string" ? { complexScriptFontTheme: fonts.cstheme } : {}),
    ...(typeof fonts.hint === "string" ? { fontHint: fonts.hint as NonNullable<TextRun["fontHint"]> } : {}),
    ...(typeof size.val === "number" ? { fontSize: size.val / 2 } : {}),
    ...(typeof size.val === "string" ? { fontSize: Number.parseInt(size.val, 10) / 2 } : {}),
    ...(typeof complexScriptSize.val === "number" ? { complexScriptFontSize: complexScriptSize.val / 2 } : {}),
    ...(typeof complexScriptSize.val === "string" ? { complexScriptFontSize: Number.parseInt(complexScriptSize.val, 10) / 2 } : {}),
    ...(typeof color.val === "string" ? { color: color.val } : {}),
    ...(typeof highlight.val === "string" ? { highlight: highlight.val as NonNullable<TextRun["highlight"]> } : {}),
    ...parseOnOffRunProperty(properties.strike, "strike"),
    ...parseOnOffRunProperty(properties.dstrike, "doubleStrike"),
    ...parseOnOffRunProperty(properties.smallCaps, "smallCaps"),
    ...parseOnOffRunProperty(properties.caps, "allCaps"),
    ...parseOnOffRunProperty(properties.shadow, "shadow"),
    ...parseOnOffRunProperty(properties.outline, "outline"),
    ...parseOnOffRunProperty(properties.emboss, "emboss"),
    ...parseOnOffRunProperty(properties.imprint, "imprint"),
    ...parseOnOffRunProperty(properties.rtl, "rtl"),
    ...parseOnOffRunProperty(properties.cs, "complexScript"),
    ...parseOnOffRunProperty(properties.specVanish, "specVanish"),
    ...parseOnOffRunProperty(properties.vanish, "hidden"),
    ...parseOnOffRunProperty(properties.webHidden, "webHidden"),
    ...parseOnOffRunProperty(properties.snapToGrid, "snapToGrid"),
    ...parseOnOffRunProperty(properties.noProof, "noProof"),
    ...parseOnOffRunProperty(properties.oMath, "officeMath"),
    ...(language ? { language } : {}),
    ...(characterPosition.val !== undefined ? { characterPosition: parseNumber(characterPosition.val) } : {}),
    ...(kerning.val !== undefined ? { kerning: parseNumber(kerning.val) } : {}),
    ...(typeof verticalAlign.val === "string" ? { verticalAlign: verticalAlign.val as NonNullable<TextRun["verticalAlign"]> } : {}),
    ...(characterSpacing.val !== undefined ? { characterSpacing: parseNumber(characterSpacing.val) } : {}),
    ...(scale.val !== undefined ? { scale: parseNumber(scale.val) } : {}),
    ...(fitText.val !== undefined
      ? { fitText: { width: parseNumber(fitText.val), ...(fitText.id !== undefined ? { id: parseNumber(fitText.id) } : {}) } }
      : {}),
    ...(typeof emphasis.val === "string" ? { emphasis: emphasis.val as NonNullable<TextRun["emphasis"]> } : {}),
    ...(border ? { border } : {}),
  };
}

function parseRunLanguage(value: unknown): RunLanguage | undefined {
  const language = asObject(value);
  const parsed = {
    ...(typeof language.val === "string" ? { value: language.val } : {}),
    ...(typeof language.eastAsia === "string" ? { eastAsia: language.eastAsia } : {}),
    ...(typeof language.bidi === "string" ? { bidi: language.bidi } : {}),
  };

  return Object.keys(parsed).length > 0 ? parsed : undefined;
}

function parseTable(value: unknown, relationships: RelationshipMap, comments: CommentMap, footnotes: NoteMap, endnotes: NoteMap, numberingContext: NumberingContext): TableNode {
  const table = asObject(value);
  const properties = asObject(table.tblPr);
  const style = asObject(properties.tblStyle);
  const position = parseTablePosition(properties.tblpPr);
  const overlap = parseTableOverlap(properties.tblOverlap);
  const caption = asObject(properties.tblCaption);
  const description = asObject(properties.tblDescription);
  const width = asObject(properties.tblW);
  const borders = parseTableBorders(properties.tblBorders);
  const alignment = asObject(properties.jc);
  const cellSpacing = asObject(properties.tblCellSpacing);
  const indent = parseTableIndent(properties.tblInd);
  const layout = parseTableLayout(properties.tblLayout);
  const look = parseTableLook(properties.tblLook);
  const widthType = parseTableWidthType(width.type);
  const propertyRevision = parsePropertyRevision(properties.tblPrChange);
  const grid = parseTableGrid(table.tblGrid);

  return {
    type: "table",
    ...(typeof style.val === "string" ? { styleId: style.val } : {}),
    ...(position ? { position } : {}),
    ...(overlap ? { overlap } : {}),
    ...(typeof caption.val === "string" ? { caption: caption.val } : {}),
    ...(typeof description.val === "string" ? { description: description.val } : {}),
    ...(grid ? { grid } : {}),
    ...(width.w !== undefined ? { width: parseNumber(width.w) } : {}),
    ...(widthType && widthType !== "dxa" ? { widthType } : {}),
    ...(borders ? { borders } : {}),
    ...(typeof alignment.val === "string" ? { alignment: alignment.val as NonNullable<TableNode["alignment"]> } : {}),
    ...(cellSpacing.w !== undefined ? { cellSpacing: parseNumber(cellSpacing.w) } : {}),
    ...(indent ? { indent } : {}),
    ...(layout ? { layout } : {}),
    ...(look ? { look } : {}),
    ...(propertyRevision ? { propertyRevision } : {}),
    rows: asArray(table.tr).map((rowValue) => {
      const row = asObject(rowValue);
      const rowProperties = asObject(row.trPr);
      const propertyExceptions = parseTablePropertyExceptions(rowProperties.tblPrEx);
      const revision = parseTableRowRevision(rowProperties);
      const height = parseTableRowHeight(rowProperties);
      const repeatHeader = parseTableRowRepeatHeader(rowProperties);
      const cantSplit = parseTableRowCantSplit(rowProperties);

      return {
        ...(propertyExceptions ? { propertyExceptions } : {}),
        ...(revision ? { revision } : {}),
        ...(height ? { height } : {}),
        ...(repeatHeader ? { repeatHeader } : {}),
        ...(cantSplit ? { cantSplit } : {}),
        cells: asArray(row.tc).map((cell) => parseTableCell(cell, relationships, comments, footnotes, endnotes, numberingContext)),
      };
    }),
  };
}

function parseTablePropertyExceptions(value: unknown): TableNode["rows"][number]["propertyExceptions"] | undefined {
  const properties = asObject(value);
  const width = asObject(properties.tblW);
  const cellSpacing = asObject(properties.tblCellSpacing);
  const indent = parseTableIndent(properties.tblInd);
  const layout = parseTableLayout(properties.tblLayout);
  const look = parseTableLook(properties.tblLook);
  const widthType = parseTableWidthType(width.type);
  const parsed = {
    ...(width.w !== undefined ? { width: parseNumber(width.w) } : {}),
    ...(widthType && widthType !== "dxa" ? { widthType } : {}),
    ...(cellSpacing.w !== undefined ? { cellSpacing: parseNumber(cellSpacing.w) } : {}),
    ...(indent ? { indent } : {}),
    ...(layout ? { layout } : {}),
    ...(look ? { look } : {}),
  };

  return Object.keys(parsed).length > 0 ? parsed : undefined;
}

function parseTableOverlap(value: unknown): TableNode["overlap"] | undefined {
  const overlap = asObject(value).val;
  return overlap === "never" || overlap === "overlap" ? overlap : undefined;
}

function parseTablePosition(value: unknown): TableNode["position"] | undefined {
  const position = asObject(value);
  const parsed = {
    ...(typeof position.horzAnchor === "string" ? { horizontalAnchor: position.horzAnchor as NonNullable<TableNode["position"]>["horizontalAnchor"] } : {}),
    ...(typeof position.vertAnchor === "string" ? { verticalAnchor: position.vertAnchor as NonNullable<TableNode["position"]>["verticalAnchor"] } : {}),
    ...(position.tblpX !== undefined ? { x: parseNumber(position.tblpX) } : {}),
    ...(position.tblpY !== undefined ? { y: parseNumber(position.tblpY) } : {}),
    ...(typeof position.tblpXSpec === "string" ? { xAlign: position.tblpXSpec as NonNullable<TableNode["position"]>["xAlign"] } : {}),
    ...(typeof position.tblpYSpec === "string" ? { yAlign: position.tblpYSpec as NonNullable<TableNode["position"]>["yAlign"] } : {}),
    ...(position.leftFromText !== undefined ? { leftFromText: parseNumber(position.leftFromText) } : {}),
    ...(position.rightFromText !== undefined ? { rightFromText: parseNumber(position.rightFromText) } : {}),
    ...(position.topFromText !== undefined ? { topFromText: parseNumber(position.topFromText) } : {}),
    ...(position.bottomFromText !== undefined ? { bottomFromText: parseNumber(position.bottomFromText) } : {}),
  };

  return Object.keys(parsed).length > 0 ? parsed : undefined;
}

function parseTableIndent(value: unknown): TableNode["indent"] | undefined {
  const indent = asObject(value);
  if (indent.w === undefined) {
    return undefined;
  }

  const type = indent.type === "dxa" || indent.type === "nil" || indent.type === "pct" ? indent.type : undefined;
  return {
    width: parseNumber(indent.w),
    ...(type && type !== "dxa" ? { type } : {}),
  };
}

function parseTableWidthType(value: unknown): TableNode["widthType"] | undefined {
  return value === "auto" || value === "dxa" || value === "nil" || value === "pct" ? value : undefined;
}

function parseTableLayout(value: unknown): TableNode["layout"] | undefined {
  const type = asObject(value).type;
  return type === "autofit" || type === "fixed" ? type : undefined;
}

function parseTableLook(value: unknown): TableNode["look"] | undefined {
  const look = asObject(value);
  const parsed = {
    ...parseTableLookValue(look.val),
    ...parseTableLookFlag(look.firstRow, "firstRow"),
    ...parseTableLookFlag(look.lastRow, "lastRow"),
    ...parseTableLookFlag(look.firstColumn, "firstColumn"),
    ...parseTableLookFlag(look.lastColumn, "lastColumn"),
    ...parseInvertedTableLookFlag(look.noHBand, "bandedRows"),
    ...parseInvertedTableLookFlag(look.noVBand, "bandedColumns"),
  };

  return Object.keys(parsed).length > 0 ? parsed : undefined;
}

function parseTableLookValue(value: unknown): TableNode["look"] | undefined {
  if (typeof value !== "string" && typeof value !== "number") {
    return undefined;
  }

  const parsed = Number.parseInt(String(value), 16);
  if (Number.isNaN(parsed)) {
    return undefined;
  }

  return {
    firstRow: (parsed & 0x0020) !== 0,
    lastRow: (parsed & 0x0040) !== 0,
    firstColumn: (parsed & 0x0080) !== 0,
    lastColumn: (parsed & 0x0100) !== 0,
    bandedRows: (parsed & 0x0200) === 0,
    bandedColumns: (parsed & 0x0400) === 0,
  };
}

function parseTableLookFlag<K extends keyof NonNullable<TableNode["look"]>>(value: unknown, key: K): Partial<Pick<NonNullable<TableNode["look"]>, K>> {
  const parsed = parseTableLookOnOff(value);
  return parsed === undefined ? {} : { [key]: parsed } as Partial<Pick<NonNullable<TableNode["look"]>, K>>;
}

function parseInvertedTableLookFlag<K extends keyof NonNullable<TableNode["look"]>>(value: unknown, key: K): Partial<Pick<NonNullable<TableNode["look"]>, K>> {
  const parsed = parseTableLookOnOff(value);
  return parsed === undefined ? {} : { [key]: !parsed } as Partial<Pick<NonNullable<TableNode["look"]>, K>>;
}

function parseTableLookOnOff(value: unknown): boolean | undefined {
  if (value === undefined) {
    return undefined;
  }

  return !(value === "0" || value === false || value === "false" || value === "off");
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

function parseTableRowRepeatHeader(value: unknown): boolean | undefined {
  const tableHeader = asObject(value).tblHeader;
  if (tableHeader === undefined) {
    return undefined;
  }

  const val = asObject(tableHeader).val;
  return val === "0" || val === false || val === "false" ? undefined : true;
}

function parseTableRowCantSplit(value: unknown): boolean | undefined {
  const cantSplit = asObject(value).cantSplit;
  if (cantSplit === undefined) {
    return undefined;
  }

  const val = asObject(cantSplit).val;
  return val === undefined || parseOnOff(val) ? true : undefined;
}

function parseTableCell(value: unknown, relationships: RelationshipMap, comments: CommentMap, footnotes: NoteMap, endnotes: NoteMap, numberingContext: NumberingContext): TableCellNode {
  const cell = asObject(value);
  const properties = asObject(cell.tcPr);
  const width = asObject(properties.tcW);
  const widthType = parseTableWidthType(width.type);
  const gridSpan = asObject(properties.gridSpan);
  const verticalMerge = asObject(properties.vMerge);
  const verticalAlignment = asObject(properties.vAlign);
  const shading = parseTableCellShading(properties.shd);
  const borders = parseParagraphBorders(properties.tcBorders);
  const noWrap = parseTableCellOnOff(properties.noWrap);
  const textDirection = asObject(properties.textDirection);
  const fitText = parseTableCellOnOff(properties.tcFitText);
  const margins = parseTableCellMargins(properties.tcMar);
  const propertyRevision = parsePropertyRevision(properties.tcPrChange);

  return {
    ...(width.w !== undefined ? { width: parseNumber(width.w) } : {}),
    ...(widthType && widthType !== "dxa" ? { widthType } : {}),
    ...(gridSpan.val !== undefined ? { colSpan: parseNumber(gridSpan.val) } : {}),
    ...(typeof verticalMerge.val === "string" ? { verticalMerge: verticalMerge.val as NonNullable<TableCellNode["verticalMerge"]> } : {}),
    ...(typeof verticalAlignment.val === "string" ? { verticalAlignment: verticalAlignment.val as NonNullable<TableCellNode["verticalAlignment"]> } : {}),
    ...(shading ? { shading } : {}),
    ...(borders ? { borders } : {}),
    ...(noWrap ? { noWrap } : {}),
    ...(typeof textDirection.val === "string" ? { textDirection: textDirection.val as NonNullable<TableCellNode["textDirection"]> } : {}),
    ...(fitText ? { fitText } : {}),
    ...(margins ? { margins } : {}),
    ...(propertyRevision ? { propertyRevision } : {}),
    blocks: asArray(cell.p).map((paragraph) => parseParagraph(paragraph, relationships, comments, footnotes, endnotes, numberingContext)),
  };
}

function parseTableCellOnOff(value: unknown): boolean | undefined {
  if (value === undefined) {
    return undefined;
  }

  const val = asObject(value).val;
  return val === undefined || parseOnOff(val) ? true : undefined;
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

  if (margin.w === undefined) {
    return {};
  }

  const width = parseNumber(margin.w);
  const type = parseTableWidthType(margin.type);

  return type && type !== "dxa"
    ? { [side]: { width, type } }
    : { [side]: width };
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

function unescapeXml(value: string): string {
  return value
    .replaceAll("&quot;", '"')
    .replaceAll("&apos;", "'")
    .replaceAll("&lt;", "<")
    .replaceAll("&gt;", ">")
    .replaceAll("&amp;", "&");
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
