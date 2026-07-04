import JSZip from "jszip";
import type {
  BorderDefinition,
  DocumentBlock,
  DocumentJson,
  ImageNode,
  MathControlProperties,
  MathNode,
  PageSettings,
  ParagraphNode,
  RunLanguage,
  RunRevision,
  SectionNode,
  StyleParagraphProperties,
  StyleRunProperties,
  TableCellNode,
  TableNode,
  TextRun,
} from "./schema.js";

type WriterContext = {
  hyperlinks: HyperlinkRelationship[];
  comments: CommentEntry[];
  images: ImageRelationship[];
  headers: HeaderFooterRelationship[];
  footers: HeaderFooterRelationship[];
  footnotes: NoteEntry[];
  endnotes: NoteEntry[];
  theme: DocumentJson["theme"];
  settings: DocumentJson["settings"];
  properties: DocumentJson["properties"];
  fonts: DocumentJson["fonts"];
  customXmlParts: NonNullable<DocumentJson["customXmlParts"]>;
  bookmarkId: number;
};

type HyperlinkRelationship = {
  id: string;
  url: string;
};

type CommentEntry = {
  id: number;
  author: string;
  initials?: string;
  date?: string;
  text: string;
};

type ImageRelationship = {
  id: string;
  filename: string;
  contentType: ImageNode["contentType"];
  data: string;
};

type HeaderFooterRelationship = {
  id: string;
  filename: string;
  blocks: ParagraphNode[];
  watermark?: SectionNode["watermark"];
};

type NoteEntry = {
  id: number;
  blocks: ParagraphNode[];
};

export async function buildDocx(document: DocumentJson): Promise<Buffer> {
  const zip = new JSZip();
  const context: WriterContext = { hyperlinks: [], comments: [], images: [], headers: [], footers: [], footnotes: [], endnotes: [], theme: document.theme, settings: document.settings, properties: document.properties, fonts: document.fonts, customXmlParts: document.customXmlParts ?? [], bookmarkId: 0 };

  zip.folder("_rels")!.file(".rels", packageRelsXml(context));
  zip.folder("word")!.file("document.xml", documentXml(document, context));
  if (document.properties?.core) {
    zip.folder("docProps")!.file("core.xml", corePropertiesXml(document.properties.core));
  }
  if (document.properties?.app) {
    zip.folder("docProps")!.file("app.xml", appPropertiesXml(document.properties.app));
  }
  if (document.properties?.custom && document.properties.custom.length > 0) {
    zip.folder("docProps")!.file("custom.xml", customPropertiesXml(document.properties.custom));
  }
  context.customXmlParts.forEach((part, index) => {
    zip.file(part.path, part.xml);

    if (part.properties) {
      const propertiesPath = part.properties.path ?? defaultCustomXmlPropertiesPath(index);
      zip.file(propertiesPath, customXmlPropertiesXml(part.properties));
      zip.file(customXmlRelationshipPath(part.path), customXmlItemRelsXml(propertiesPath));
    }
  });
  for (const header of context.headers) {
    zip.folder("word")!.file(header.filename, headerFooterXml("hdr", header.blocks, context, header.watermark));
  }
  for (const footer of context.footers) {
    zip.folder("word")!.file(footer.filename, headerFooterXml("ftr", footer.blocks, context));
  }
  for (const image of context.images) {
    zip.folder("word")!.folder("media")!.file(image.filename, Buffer.from(image.data, "base64"));
  }
  zip.folder("word")!.file("styles.xml", stylesXml(document));
  zip.folder("word")!.file("numbering.xml", numberingXml(document));
  if (document.theme) {
    zip.folder("word")!.folder("theme")!.file("theme1.xml", themeXml(document.theme));
  }
  if (document.settings) {
    zip.folder("word")!.file("settings.xml", settingsXml(document.settings));
  }
  if (document.settings?.web) {
    zip.folder("word")!.file("webSettings.xml", webSettingsXml(document.settings.web));
  }
  if (document.fonts && document.fonts.length > 0) {
    zip.folder("word")!.file("fontTable.xml", fontTableXml(document.fonts));
  }
  if (context.comments.length > 0) {
    zip.folder("word")!.file("comments.xml", commentsXml(context));
  }
  if (context.footnotes.length > 0) {
    zip.folder("word")!.file("footnotes.xml", notesXml("footnotes", "footnote", context.footnotes, context));
  }
  if (context.endnotes.length > 0) {
    zip.folder("word")!.file("endnotes.xml", notesXml("endnotes", "endnote", context.endnotes, context));
  }
  zip.file("[Content_Types].xml", contentTypesXml(context));
  zip.folder("word")!.folder("_rels")!.file("document.xml.rels", documentRelsXml(context));

  return zip.generateAsync({ type: "nodebuffer" });
}

function settingsXml(settings: NonNullable<DocumentJson["settings"]>): string {
  const mathNamespace = settings.math ? ` xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"` : "";

  return xmlDeclaration(
    `<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"${mathNamespace}>` +
      viewSettingsXml(settings.view) +
      proofingSettingsXml(settings.proofing) +
      (settings.trackRevisions ? "<w:trackRevisions/>" : "") +
      (settings.defaultTabStop !== undefined ? `<w:defaultTabStop w:val="${settings.defaultTabStop}"/>` : "") +
      (settings.evenAndOddHeaders ? "<w:evenAndOddHeaders/>" : "") +
      (settings.updateFields ? "<w:updateFields/>" : "") +
      compatibilitySettingsXml(settings.compatibility) +
      documentProtectionXml(settings.protection) +
      mailMergeSettingsXml(settings.mailMerge) +
      writeProtectionXml(settings.writeProtection) +
      mathSettingsXml(settings.math) +
      `</w:settings>`,
  );
}

function mathSettingsXml(math: NonNullable<DocumentJson["settings"]>["math"]): string {
  if (!math) {
    return "";
  }

  const children = [
    math.mathFont ? `<m:mathFont m:val="${escapeAttribute(math.mathFont)}"/>` : "",
    math.breakBinary ? `<m:brkBin m:val="${escapeAttribute(math.breakBinary)}"/>` : "",
    math.smallFraction !== undefined ? `<m:smallFrac m:val="${math.smallFraction ? "1" : "0"}"/>` : "",
    math.displayDefaults ? "<m:dispDef/>" : "",
  ].join("");

  return children ? `<m:mathPr>${children}</m:mathPr>` : "";
}

function writeProtectionXml(writeProtection: NonNullable<DocumentJson["settings"]>["writeProtection"]): string {
  if (!writeProtection) {
    return "";
  }

  const attributes = [
    writeProtection.recommended !== undefined ? `w:recommended="${writeProtection.recommended ? "1" : "0"}"` : "",
    writeProtection.cryptProviderType ? `w:cryptProviderType="${escapeAttribute(writeProtection.cryptProviderType)}"` : "",
    writeProtection.cryptAlgorithmClass ? `w:cryptAlgorithmClass="${escapeAttribute(writeProtection.cryptAlgorithmClass)}"` : "",
    writeProtection.cryptAlgorithmType ? `w:cryptAlgorithmType="${escapeAttribute(writeProtection.cryptAlgorithmType)}"` : "",
    writeProtection.cryptAlgorithmSid !== undefined ? `w:cryptAlgorithmSid="${writeProtection.cryptAlgorithmSid}"` : "",
    writeProtection.cryptSpinCount !== undefined ? `w:cryptSpinCount="${writeProtection.cryptSpinCount}"` : "",
    writeProtection.hash ? `w:hash="${escapeAttribute(writeProtection.hash)}"` : "",
    writeProtection.salt ? `w:salt="${escapeAttribute(writeProtection.salt)}"` : "",
  ].filter(Boolean).join(" ");

  return attributes ? `<w:writeProtection ${attributes}/>` : "";
}

function mailMergeSettingsXml(mailMerge: NonNullable<DocumentJson["settings"]>["mailMerge"]): string {
  if (!mailMerge) {
    return "";
  }

  const children = [
    mailMerge.mainDocumentType ? `<w:mainDocumentType w:val="${escapeAttribute(mailMerge.mainDocumentType)}"/>` : "",
    mailMerge.dataType ? `<w:dataType w:val="${escapeAttribute(mailMerge.dataType)}"/>` : "",
    mailMerge.connectString ? `<w:connectString w:val="${escapeAttribute(mailMerge.connectString)}"/>` : "",
    mailMerge.query ? `<w:query w:val="${escapeAttribute(mailMerge.query)}"/>` : "",
    mailMerge.viewMergedData ? "<w:viewMergedData/>" : "",
    mailMerge.activeRecord !== undefined ? `<w:activeRecord w:val="${mailMerge.activeRecord}"/>` : "",
    mailMerge.checkErrors !== undefined ? `<w:checkErrors w:val="${mailMerge.checkErrors}"/>` : "",
  ].join("");

  return children ? `<w:mailMerge>${children}</w:mailMerge>` : "";
}

function documentProtectionXml(protection: NonNullable<DocumentJson["settings"]>["protection"]): string {
  if (!protection) {
    return "";
  }

  const attributes = [
    protection.edit ? `w:edit="${escapeAttribute(protection.edit)}"` : "",
    protection.enforcement !== undefined ? `w:enforcement="${protection.enforcement ? "1" : "0"}"` : "",
    protection.cryptProviderType ? `w:cryptProviderType="${escapeAttribute(protection.cryptProviderType)}"` : "",
    protection.cryptAlgorithmClass ? `w:cryptAlgorithmClass="${escapeAttribute(protection.cryptAlgorithmClass)}"` : "",
    protection.cryptAlgorithmType ? `w:cryptAlgorithmType="${escapeAttribute(protection.cryptAlgorithmType)}"` : "",
    protection.cryptAlgorithmSid !== undefined ? `w:cryptAlgorithmSid="${protection.cryptAlgorithmSid}"` : "",
    protection.cryptSpinCount !== undefined ? `w:cryptSpinCount="${protection.cryptSpinCount}"` : "",
    protection.hash ? `w:hash="${escapeAttribute(protection.hash)}"` : "",
    protection.salt ? `w:salt="${escapeAttribute(protection.salt)}"` : "",
  ].filter(Boolean).join(" ");

  return attributes ? `<w:documentProtection ${attributes}/>` : "";
}

function viewSettingsXml(view: NonNullable<DocumentJson["settings"]>["view"]): string {
  if (!view) {
    return "";
  }

  const viewXml = view.mode ? `<w:view w:val="${escapeAttribute(view.mode)}"/>` : "";
  const zoomAttributes = [
    view.zoom?.preset ? `w:val="${escapeAttribute(view.zoom.preset)}"` : "",
    view.zoom?.percent !== undefined ? `w:percent="${view.zoom.percent}"` : "",
  ].filter(Boolean).join(" ");
  const zoomXml = zoomAttributes ? `<w:zoom ${zoomAttributes}/>` : "";

  return viewXml + zoomXml;
}

function proofingSettingsXml(proofing: NonNullable<DocumentJson["settings"]>["proofing"]): string {
  if (!proofing) {
    return "";
  }

  const proofStateAttributes = [
    proofing.spelling ? `w:spelling="${escapeAttribute(proofing.spelling)}"` : "",
    proofing.grammar ? `w:grammar="${escapeAttribute(proofing.grammar)}"` : "",
  ].filter(Boolean).join(" ");
  const proofState = proofStateAttributes ? `<w:proofState ${proofStateAttributes}/>` : "";
  const doNotHyphenateCaps = proofing.doNotHyphenateCaps ? "<w:doNotHyphenateCaps/>" : "";
  const hyphenationZone = proofing.hyphenationZone !== undefined ? `<w:hyphenationZone w:val="${proofing.hyphenationZone}"/>` : "";

  return proofState + doNotHyphenateCaps + hyphenationZone;
}

function compatibilitySettingsXml(compatibility: NonNullable<DocumentJson["settings"]>["compatibility"]): string {
  if (!compatibility?.compatMode && !compatibility?.settings?.length) {
    return "";
  }

  const compatMode = compatibility.compatMode
    ? `<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="${escapeAttribute(compatibility.compatMode)}"/>`
    : "";
  const settings = (compatibility.settings ?? [])
    .map((setting) => `<w:compatSetting w:name="${escapeAttribute(setting.name)}" w:uri="${escapeAttribute(setting.uri)}" w:val="${escapeAttribute(setting.value)}"/>`)
    .join("");

  return `<w:compat>${compatMode}${settings}</w:compat>`;
}

function webSettingsXml(settings: NonNullable<NonNullable<DocumentJson["settings"]>["web"]>): string {
  return xmlDeclaration(
    `<w:webSettings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
      (settings.optimizeForBrowser ? "<w:optimizeForBrowser/>" : "") +
      (settings.allowPng ? "<w:allowPNG/>" : "") +
      (settings.doNotSaveAsSingleFile ? "<w:doNotSaveAsSingleFile/>" : "") +
      (settings.pixelsPerInch !== undefined ? `<w:pixelsPerInch w:val="${settings.pixelsPerInch}"/>` : "") +
      `</w:webSettings>`,
  );
}

function fontTableXml(fonts: NonNullable<DocumentJson["fonts"]>): string {
  const fontEntries = fonts.map((font) =>
    `<w:font w:name="${escapeAttribute(font.name)}">` +
      (font.panose1 ? `<w:panose1 w:val="${escapeAttribute(font.panose1)}"/>` : "") +
      (font.charset ? `<w:charset w:val="${escapeAttribute(font.charset)}"/>` : "") +
      (font.family ? `<w:family w:val="${font.family}"/>` : "") +
      (font.pitch ? `<w:pitch w:val="${font.pitch}"/>` : "") +
      `</w:font>`,
  ).join("");

  return xmlDeclaration(
    `<w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
      fontEntries +
      `</w:fonts>`,
  );
}

function documentXml(document: DocumentJson, context: WriterContext): string {
  const sections = document.sections.length > 0 ? document.sections : [{ blocks: [] }];
  const body = sections
    .map((section, index) => {
      const blocks = section.blocks.map((block) => blockXml(block, context)).join("");
      const isLast = index === sections.length - 1;

      return isLast
        ? blocks
        : `${blocks}<w:p><w:pPr>${sectionPropertiesXml(section, context)}</w:pPr></w:p>`;
    })
    .join("");
  const finalSection = sections[sections.length - 1];

  return xmlDeclaration(
    `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">` +
      `<w:body>${body}${sectionPropertiesXml(finalSection, context)}</w:body>` +
      `</w:document>`,
  );
}

function blockXml(block: DocumentBlock, context: WriterContext): string {
  if (block.type === "paragraph") {
    return paragraphXml(block, context);
  }

  if (block.type === "table") {
    return tableXml(block, context);
  }

  return imageXml(block, context);
}

function sectionPropertiesXml(section: SectionNode, context: WriterContext): string {
  const page = section.page ?? defaultPageSettings();
  const orientation = page.orientation && page.orientation !== "portrait"
    ? ` w:orient="${page.orientation}"`
    : "";
  const headerReference = headerFooterReferencesXml(sectionHeadersWithWatermark(section), "header", context, section.watermark);
  const footerReference = headerFooterReferencesXml(section.footers, "footer", context);
  const titlePage = section.titlePage ? "<w:titlePg/>" : "";
  const breakType = section.breakType
    ? `<w:type w:val="${sectionBreakValue(section.breakType)}"/>`
    : "";
  const pageNumbering = pageNumberingXml(section.pageNumbering);
  const lineNumbering = lineNumberingXml(section.lineNumbering);
  const footnoteProperties = notePropertiesXml("footnotePr", section.footnoteProperties);
  const endnoteProperties = notePropertiesXml("endnotePr", section.endnoteProperties);
  const noEndnote = section.noEndnote ? "<w:noEndnote/>" : "";
  const documentGrid = documentGridXml(section.documentGrid);
  const verticalAlignment = section.verticalAlignment
    ? `<w:vAlign w:val="${section.verticalAlignment}"/>`
    : "";
  const textDirection = section.textDirection
    ? `<w:textDirection w:val="${section.textDirection}"/>`
    : "";
  const bidi = section.bidi ? "<w:bidi/>" : "";
  const rtlGutter = section.rtlGutter ? "<w:rtlGutter/>" : "";
  const mirrorMargins = section.mirrorMargins ? "<w:mirrorMargins/>" : "";
  const columns = section.columns
    ? `<w:cols w:num="${section.columns.count}"${section.columns.space ? ` w:space="${section.columns.space}"` : ""}/>`
    : "";

  return `<w:sectPr>` +
    headerReference +
    footerReference +
    footnoteProperties +
    endnoteProperties +
    breakType +
    `<w:pgSz w:w="${page.width}" w:h="${page.height}"${orientation}/>` +
    `<w:pgMar w:top="${page.margins.top}" w:right="${page.margins.right}" w:bottom="${page.margins.bottom}" w:left="${page.margins.left}" w:header="${page.margins.header}" w:footer="${page.margins.footer}" w:gutter="${page.margins.gutter}"/>` +
    lineNumbering +
    pageNumbering +
    columns +
    verticalAlignment +
    noEndnote +
    titlePage +
    textDirection +
    bidi +
    rtlGutter +
    mirrorMargins +
    documentGrid +
    `</w:sectPr>`;
}

function pageNumberingXml(pageNumbering?: SectionNode["pageNumbering"]): string {
  if (!pageNumbering) {
    return "";
  }

  const attributes = [
    pageNumbering.start !== undefined ? ` w:start="${pageNumbering.start}"` : "",
    pageNumbering.format ? ` w:fmt="${pageNumbering.format}"` : "",
    pageNumbering.chapterStyle !== undefined ? ` w:chapStyle="${pageNumbering.chapterStyle}"` : "",
    pageNumbering.chapterSeparator ? ` w:chapSep="${pageNumbering.chapterSeparator}"` : "",
  ].join("");

  return attributes ? `<w:pgNumType${attributes}/>` : "";
}

function lineNumberingXml(lineNumbering?: SectionNode["lineNumbering"]): string {
  if (!lineNumbering) {
    return "";
  }

  const attributes = [
    lineNumbering.start !== undefined ? ` w:start="${lineNumbering.start}"` : "",
    lineNumbering.countBy !== undefined ? ` w:countBy="${lineNumbering.countBy}"` : "",
    lineNumbering.distance !== undefined ? ` w:distance="${lineNumbering.distance}"` : "",
    lineNumbering.restart ? ` w:restart="${lineNumbering.restart}"` : "",
  ].join("");

  return attributes ? `<w:lnNumType${attributes}/>` : "";
}

function notePropertiesXml(root: "footnotePr" | "endnotePr", properties?: SectionNode["footnoteProperties"]): string {
  if (!properties) {
    return "";
  }

  const children = [
    properties.position ? `<w:pos w:val="${properties.position}"/>` : "",
    properties.numbering?.format ? `<w:numFmt w:val="${properties.numbering.format}"/>` : "",
    properties.numbering?.start !== undefined ? `<w:numStart w:val="${properties.numbering.start}"/>` : "",
    properties.numbering?.restart ? `<w:numRestart w:val="${properties.numbering.restart}"/>` : "",
  ].join("");

  return children ? `<w:${root}>${children}</w:${root}>` : "";
}

function documentGridXml(documentGrid?: SectionNode["documentGrid"]): string {
  if (!documentGrid) {
    return "";
  }

  const attributes = [
    documentGrid.type ? ` w:type="${documentGrid.type}"` : "",
    documentGrid.linePitch !== undefined ? ` w:linePitch="${documentGrid.linePitch}"` : "",
    documentGrid.charSpace !== undefined ? ` w:charSpace="${documentGrid.charSpace}"` : "",
  ].join("");

  return attributes ? `<w:docGrid${attributes}/>` : "";
}

function sectionBreakValue(value: SectionNode["breakType"]): string {
  if (value === "evenPage") {
    return "evenPage";
  }

  if (value === "oddPage") {
    return "oddPage";
  }

  if (value === "continuous") {
    return "continuous";
  }

  return "nextPage";
}

function sectionHeadersWithWatermark(section: SectionNode): SectionNode["headers"] {
  if (!section.watermark) {
    return section.headers;
  }

  return {
    ...section.headers,
    default: section.headers?.default ?? [],
  };
}

function headerFooterReferencesXml(content: SectionNode["headers"], root: "header" | "footer", context: WriterContext, watermark?: SectionNode["watermark"]): string {
  if (!content) {
    return "";
  }

  return (["default", "first", "even"] as const)
    .map((type) => {
      const blocks = content[type];
      if (!blocks) {
        return "";
      }

      return root === "header"
        ? createHeaderReference(type, blocks, context, type === "default" ? watermark : undefined)
        : createFooterReference(type, blocks, context);
    })
    .join("");
}

function createHeaderReference(type: keyof NonNullable<SectionNode["headers"]>, blocks: ParagraphNode[], context: WriterContext, watermark?: SectionNode["watermark"]): string {
  const index = context.headers.length + 1;
  const id = `rIdHeader${index}`;
  context.headers.push({ id, filename: `header${index}.xml`, blocks, watermark });
  return `<w:headerReference w:type="${type}" r:id="${id}"/>`;
}

function createFooterReference(type: keyof NonNullable<SectionNode["footers"]>, blocks: ParagraphNode[], context: WriterContext): string {
  const index = context.footers.length + 1;
  const id = `rIdFooter${index}`;
  context.footers.push({ id, filename: `footer${index}.xml`, blocks });
  return `<w:footerReference w:type="${type}" r:id="${id}"/>`;
}

function defaultPageSettings(): PageSettings {
  return {
    width: 12240,
    height: 15840,
    margins: {
      top: 1440,
      right: 1440,
      bottom: 1440,
      left: 1440,
      header: 720,
      footer: 720,
      gutter: 0,
    },
  };
}

function paragraphXml(paragraph: ParagraphNode, context: WriterContext): string {
  const properties = paragraphPropertiesXml(paragraph);
  const runs = paragraphRunsXml(paragraph.runs, context);
  const plainParagraph = `<w:p>${properties}${runs}</w:p>`;
  const startComment = paragraph.commentRangeStart
    ? paragraphCommentRangeStartXml(paragraph.commentRangeStart, context)
    : "";
  const endComment = paragraph.commentRangeEnd
    ? paragraphCommentRangeEndXml(paragraph.commentRangeEnd)
    : "";
  const controlledParagraph = paragraph.contentControl
    ? contentControlXml(paragraph.contentControl, plainParagraph)
    : plainParagraph;

  if (!paragraph.revision) {
    return `${startComment}${controlledParagraph}${endComment}`;
  }

  const wrapper = revisionElement(paragraph.revision.type);
  return `${startComment}<w:${wrapper}${revisionAttributes(paragraph.revision)}>${controlledParagraph}</w:${wrapper}>${endComment}`;
}

function contentControlXml(contentControl: NonNullable<ParagraphNode["contentControl"]>, content: string): string {
  const properties = [
    contentControl.alias ? `<w:alias w:val="${escapeAttribute(contentControl.alias)}"/>` : "",
    contentControl.tag ? `<w:tag w:val="${escapeAttribute(contentControl.tag)}"/>` : "",
    contentControl.lock ? `<w:lock w:val="${contentControl.lock}"/>` : "",
    contentControl.appearance ? `<w:appearance w:val="${contentControl.appearance}"/>` : "",
    contentControl.color ? `<w:color w:val="${escapeAttribute(contentControl.color)}"/>` : "",
    contentControl.showingPlaceholder ? `<w:showingPlcHdr/>` : "",
    contentControl.dataBinding ? dataBindingContentControlXml(contentControl.dataBinding) : "",
    contentControl.placeholder ? placeholderContentControlXml(contentControl.placeholder) : "",
    contentControl.checkbox ? checkboxContentControlXml(contentControl.checkbox) : "",
    contentControl.dropdown ? dropdownContentControlXml(contentControl.dropdown) : "",
    contentControl.comboBox ? comboBoxContentControlXml(contentControl.comboBox) : "",
    contentControl.date ? dateContentControlXml(contentControl.date) : "",
    contentControl.repeatingSection ? repeatingSectionContentControlXml(contentControl.repeatingSection) : "",
    contentControl.repeatingSectionItem ? repeatingSectionItemContentControlXml(contentControl.repeatingSectionItem) : "",
  ].join("");

  return `<w:sdt><w:sdtPr>${properties}</w:sdtPr><w:sdtContent>${content}</w:sdtContent></w:sdt>`;
}

function dataBindingContentControlXml(dataBinding: NonNullable<NonNullable<ParagraphNode["contentControl"]>["dataBinding"]>): string {
  const attributes = [
    dataBinding.storeItemId ? `w:storeItemID="${escapeAttribute(dataBinding.storeItemId)}"` : "",
    dataBinding.xpath ? `w:xpath="${escapeAttribute(dataBinding.xpath)}"` : "",
    dataBinding.prefixMappings ? `w:prefixMappings="${escapeAttribute(dataBinding.prefixMappings)}"` : "",
  ].filter(Boolean).join(" ");

  return `<w:dataBinding${attributes ? ` ${attributes}` : ""}/>`;
}

function placeholderContentControlXml(placeholder: NonNullable<NonNullable<ParagraphNode["contentControl"]>["placeholder"]>): string {
  return `<w:placeholder><w:docPart w:val="${escapeAttribute(placeholder.docPart)}"/></w:placeholder>`;
}

function checkboxContentControlXml(checkbox: NonNullable<NonNullable<ParagraphNode["contentControl"]>["checkbox"]>): string {
  return `<w:checkBox>` +
    `<w:checked w:val="${checkbox.checked ? 1 : 0}"/>` +
    (checkbox.checkedSymbol ? `<w:checkedState w:val="${escapeAttribute(checkbox.checkedSymbol)}"/>` : "") +
    (checkbox.uncheckedSymbol ? `<w:uncheckedState w:val="${escapeAttribute(checkbox.uncheckedSymbol)}"/>` : "") +
    `</w:checkBox>`;
}

function dropdownContentControlXml(dropdown: NonNullable<NonNullable<ParagraphNode["contentControl"]>["dropdown"]>): string {
  const items = dropdown.items
    .map((item) => `<w:listItem w:displayText="${escapeAttribute(item.displayText)}" w:value="${escapeAttribute(item.value)}"/>`)
    .join("");

  return `<w:dropDownList>${items}</w:dropDownList>`;
}

function comboBoxContentControlXml(comboBox: NonNullable<NonNullable<ParagraphNode["contentControl"]>["comboBox"]>): string {
  const items = comboBox.items
    .map((item) => `<w:listItem w:displayText="${escapeAttribute(item.displayText)}" w:value="${escapeAttribute(item.value)}"/>`)
    .join("");

  return `<w:comboBox>${items}</w:comboBox>`;
}

function dateContentControlXml(date: NonNullable<NonNullable<ParagraphNode["contentControl"]>["date"]>): string {
  return `<w:date>` +
    (date.fullDate ? `<w:fullDate w:val="${escapeAttribute(date.fullDate)}"/>` : "") +
    (date.format ? `<w:dateFormat w:val="${escapeAttribute(date.format)}"/>` : "") +
    `</w:date>`;
}

function repeatingSectionContentControlXml(repeatingSection: NonNullable<NonNullable<ParagraphNode["contentControl"]>["repeatingSection"]>): string {
  return `<w:repeatingSection>` +
    (repeatingSection.sectionTitle ? `<w:sectionTitle w:val="${escapeAttribute(repeatingSection.sectionTitle)}"/>` : "") +
    (repeatingSection.doNotAllowInsertDeleteSection ? `<w:doNotAllowInsertDeleteSection/>` : "") +
    `</w:repeatingSection>`;
}

function repeatingSectionItemContentControlXml(repeatingSectionItem: NonNullable<NonNullable<ParagraphNode["contentControl"]>["repeatingSectionItem"]>): string {
  return `<w:repeatingSectionItem>` +
    (repeatingSectionItem.id ? `<w:id w:val="${escapeAttribute(repeatingSectionItem.id)}"/>` : "") +
    `</w:repeatingSectionItem>`;
}

function paragraphCommentRangeStartXml(comment: NonNullable<ParagraphNode["commentRangeStart"]>, context: WriterContext): string {
  const id = comment.id ?? nextCommentId(context);
  ensureComment(comment, id, context);
  return `<w:commentRangeStart w:id="${id}"/>`;
}

function paragraphCommentRangeEndXml(commentRangeEnd: NonNullable<ParagraphNode["commentRangeEnd"]>): string {
  return `<w:commentRangeEnd w:id="${commentRangeEnd.id}"/>` +
    `<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="${commentRangeEnd.id}"/></w:r>`;
}

function paragraphRunsXml(runs: TextRun[], context: WriterContext): string {
  const chunks: string[] = [];
  let index = 0;

  while (index < runs.length) {
    const run = runs[index];
    const commentId = explicitCommentId(run);

    if (commentId === undefined) {
      chunks.push(runXml(run, context));
      index += 1;
      continue;
    }

    let endIndex = index + 1;
    while (endIndex < runs.length && explicitCommentId(runs[endIndex]) === commentId) {
      endIndex += 1;
    }

    if (endIndex === index + 1) {
      chunks.push(runXml(run, context));
      index += 1;
      continue;
    }

    ensureCommentEntry(run, commentId, context);
    const groupedRuns = runs
      .slice(index, endIndex)
      .map((groupedRun) => runXml(groupedRun, context, { skipComment: true }))
      .join("");
    chunks.push(
      `<w:commentRangeStart w:id="${commentId}"/>` +
      groupedRuns +
      `<w:commentRangeEnd w:id="${commentId}"/>` +
      `<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="${commentId}"/></w:r>`,
    );
    index = endIndex;
  }

  return chunks.join("");
}

function explicitCommentId(run: TextRun): number | undefined {
  return run.comment?.id;
}

function paragraphPropertiesXml(paragraph: ParagraphNode): string {
  const styleId = paragraph.styleId ?? (paragraph.style && paragraph.style !== "normal"
    ? paragraphStyleId(paragraph.style)
    : undefined);
  const style = styleId
    ? `<w:pStyle w:val="${escapeAttribute(styleId)}"/>`
    : "";
  const alignment = paragraph.alignment
    ? `<w:jc w:val="${paragraph.alignment}"/>`
    : "";
  const spacing = paragraph.spacing ? paragraphSpacingXml(paragraph.spacing) : "";
  const indent = paragraph.indent ? paragraphIndentXml(paragraph.indent) : "";
  const shading = paragraph.shading ? shadingXml(paragraph.shading) : "";
  const borders = paragraph.borders ? paragraphBordersXml(paragraph.borders) : "";
  const tabs = paragraphTabsXml(paragraph.tabs);
  const list = paragraph.list
    ? `<w:numPr><w:ilvl w:val="${paragraph.list.level}"/><w:numId w:val="${paragraph.list.numberingId ?? (paragraph.list.type === "bullet" ? 1 : 2)}"/></w:numPr>`
    : "";
  const pagination = paragraph.pagination;
  const frame = paragraphFrameXml(paragraph.frame);
  const propertyRevision = paragraph.propertyRevision
    ? propertyRevisionXml("pPr", "pPrChange", paragraph.propertyRevision)
    : "";
  const properties = [
    style,
    frame,
    paragraphPaginationToggleXml("keepNext", pagination?.keepNext),
    paragraphPaginationToggleXml("keepLines", pagination?.keepLines),
    paragraphPaginationToggleXml("pageBreakBefore", pagination?.pageBreakBefore),
    paragraphPaginationToggleXml("widowControl", pagination?.widowControl),
    list,
    paragraphPaginationToggleXml("suppressLineNumbers", pagination?.suppressLineNumbers),
    borders,
    shading,
    tabs,
    paragraphPaginationToggleXml("suppressAutoHyphens", pagination?.suppressAutoHyphens),
    paragraphPaginationToggleXml("overflowPunct", pagination?.overflowPunct),
    paragraphPaginationToggleXml("topLinePunct", pagination?.topLinePunct),
    paragraphPaginationToggleXml("autoSpaceDE", pagination?.autoSpaceDE),
    paragraphPaginationToggleXml("autoSpaceDN", pagination?.autoSpaceDN),
    paragraphPaginationToggleXml("adjustRightInd", pagination?.adjustRightInd),
    spacing,
    indent,
    paragraphPaginationToggleXml("contextualSpacing", pagination?.contextualSpacing),
    paragraphPaginationToggleXml("mirrorIndents", pagination?.mirrorIndents),
    alignment,
    pagination?.textDirection ? `<w:textDirection w:val="${pagination.textDirection}"/>` : "",
    pagination?.textAlignment ? `<w:textAlignment w:val="${pagination.textAlignment}"/>` : "",
    propertyRevision,
  ].join("");

  return properties ? `<w:pPr>${properties}</w:pPr>` : "";
}

function shadingXml(shading: { fill: string; value?: string; color?: string; themeFill?: string; themeFillTint?: string; themeFillShade?: string; themeColor?: string; themeTint?: string; themeShade?: string }): string {
  return `<w:shd w:val="${escapeAttribute(shading.value ?? "clear")}"` +
    (shading.color ? ` w:color="${escapeAttribute(shading.color)}"` : "") +
    ` w:fill="${escapeAttribute(shading.fill)}"` +
    (shading.themeFill ? ` w:themeFill="${escapeAttribute(shading.themeFill)}"` : "") +
    (shading.themeFillTint ? ` w:themeFillTint="${escapeAttribute(shading.themeFillTint)}"` : "") +
    (shading.themeFillShade ? ` w:themeFillShade="${escapeAttribute(shading.themeFillShade)}"` : "") +
    (shading.themeColor ? ` w:themeColor="${escapeAttribute(shading.themeColor)}"` : "") +
    (shading.themeTint ? ` w:themeTint="${escapeAttribute(shading.themeTint)}"` : "") +
    (shading.themeShade ? ` w:themeShade="${escapeAttribute(shading.themeShade)}"` : "") +
    `/>`;
}

function paragraphBordersXml(borders: NonNullable<ParagraphNode["borders"]>): string {
  const sides = [
    borders.top ? borderSideXml("top", borders.top) : "",
    borders.left ? borderSideXml("left", borders.left) : "",
    borders.bottom ? borderSideXml("bottom", borders.bottom) : "",
    borders.right ? borderSideXml("right", borders.right) : "",
    borders.between ? borderSideXml("between", borders.between) : "",
    borders.bar ? borderSideXml("bar", borders.bar) : "",
  ].join("");

  return sides ? `<w:pBdr>${sides}</w:pBdr>` : "";
}

function borderSideXml(side: "top" | "left" | "bottom" | "right" | "between" | "bar" | "insideH" | "insideV" | "bdr", border: BorderDefinition): string {
  return `<w:${side} w:val="${border.style}"` +
    (border.size !== undefined ? ` w:sz="${border.size}"` : "") +
    (border.space !== undefined ? ` w:space="${border.space}"` : "") +
    (border.color ? ` w:color="${escapeAttribute(border.color)}"` : "") +
    (border.themeColor ? ` w:themeColor="${escapeAttribute(border.themeColor)}"` : "") +
    (border.themeTint ? ` w:themeTint="${escapeAttribute(border.themeTint)}"` : "") +
    (border.themeShade ? ` w:themeShade="${escapeAttribute(border.themeShade)}"` : "") +
    `/>`;
}

function paragraphSpacingXml(spacing: NonNullable<ParagraphNode["spacing"]>): string {
  return `<w:spacing` +
    (spacing.before !== undefined ? ` w:before="${spacing.before}"` : "") +
    (spacing.after !== undefined ? ` w:after="${spacing.after}"` : "") +
    (spacing.line !== undefined ? ` w:line="${spacing.line}"` : "") +
    (spacing.lineRule ? ` w:lineRule="${spacing.lineRule}"` : "") +
    `/>`;
}

function paragraphIndentXml(indent: NonNullable<ParagraphNode["indent"]>): string {
  return `<w:ind` +
    (indent.left !== undefined ? ` w:left="${indent.left}"` : "") +
    (indent.right !== undefined ? ` w:right="${indent.right}"` : "") +
    (indent.firstLine !== undefined ? ` w:firstLine="${indent.firstLine}"` : "") +
    (indent.hanging !== undefined ? ` w:hanging="${indent.hanging}"` : "") +
    `/>`;
}

function paragraphFrameXml(frame?: ParagraphNode["frame"]): string {
  if (!frame) return "";

  return `<w:framePr` +
    (frame.width !== undefined ? ` w:w="${frame.width}"` : "") +
    (frame.height !== undefined ? ` w:h="${frame.height}"` : "") +
    (frame.x !== undefined ? ` w:x="${frame.x}"` : "") +
    (frame.y !== undefined ? ` w:y="${frame.y}"` : "") +
    (frame.horizontalAnchor ? ` w:hAnchor="${frame.horizontalAnchor}"` : "") +
    (frame.verticalAnchor ? ` w:vAnchor="${frame.verticalAnchor}"` : "") +
    (frame.xAlign ? ` w:xAlign="${frame.xAlign}"` : "") +
    (frame.yAlign ? ` w:yAlign="${frame.yAlign}"` : "") +
    (frame.wrap ? ` w:wrap="${frame.wrap}"` : "") +
    (frame.dropCap ? ` w:dropCap="${frame.dropCap}"` : "") +
    (frame.lines !== undefined ? ` w:lines="${frame.lines}"` : "") +
    (frame.anchorLock !== undefined ? ` w:anchorLock="${frame.anchorLock ? 1 : 0}"` : "") +
    (frame.heightRule ? ` w:hRule="${frame.heightRule}"` : "") +
    `/>`;
}

function paragraphTabsXml(tabs?: ParagraphNode["tabs"]): string {
  if (!tabs?.length) return "";

  return `<w:tabs>${tabs.map((tab) =>
    `<w:tab w:val="${tab.value}" w:pos="${tab.position}"${tab.leader ? ` w:leader="${tab.leader}"` : ""}/>`
  ).join("")}</w:tabs>`;
}

function paragraphPaginationToggleXml(name: string, value: boolean | undefined): string {
  if (value === undefined) return "";
  return value ? `<w:${name}/>` : `<w:${name} w:val="0"/>`;
}

function paragraphPaginationXml(pagination?: ParagraphNode["pagination"]): string {
  return pagination
    ? [
      paragraphPaginationToggleXml("keepNext", pagination.keepNext),
      paragraphPaginationToggleXml("keepLines", pagination.keepLines),
      paragraphPaginationToggleXml("pageBreakBefore", pagination.pageBreakBefore),
      paragraphPaginationToggleXml("widowControl", pagination.widowControl),
      paragraphPaginationToggleXml("suppressLineNumbers", pagination.suppressLineNumbers),
      paragraphPaginationToggleXml("suppressAutoHyphens", pagination.suppressAutoHyphens),
      paragraphPaginationToggleXml("contextualSpacing", pagination.contextualSpacing),
      paragraphPaginationToggleXml("mirrorIndents", pagination.mirrorIndents),
      paragraphPaginationToggleXml("overflowPunct", pagination.overflowPunct),
      paragraphPaginationToggleXml("topLinePunct", pagination.topLinePunct),
      pagination.textAlignment ? `<w:textAlignment w:val="${pagination.textAlignment}"/>` : "",
      pagination.textDirection ? `<w:textDirection w:val="${pagination.textDirection}"/>` : "",
      paragraphPaginationToggleXml("adjustRightInd", pagination.adjustRightInd),
      paragraphPaginationToggleXml("autoSpaceDE", pagination.autoSpaceDE),
      paragraphPaginationToggleXml("autoSpaceDN", pagination.autoSpaceDN),
    ].join("")
    : "";
}

function runXml(run: TextRun, context: WriterContext, options: { skipComment?: boolean } = {}): string {
  if (run.footnote) {
    const id = context.footnotes.length + 1;
    context.footnotes.push({ id, blocks: run.footnote.blocks });
    return `<w:r><w:footnoteReference w:id="${id}"/></w:r>`;
  }

  if (run.endnote) {
    const id = context.endnotes.length + 1;
    context.endnotes.push({ id, blocks: run.endnote.blocks });
    return `<w:r><w:endnoteReference w:id="${id}"/></w:r>`;
  }

  if (run.math) {
    return mathRunXml(run.math);
  }

  if (run.field) {
    return fieldRunXml(run.field);
  }

  if (run.break) {
    return run.break === "page"
      ? '<w:r><w:br w:type="page"/></w:r>'
      : "<w:r><w:br/></w:r>";
  }

  const properties = runPropertiesXml(run);
  const textSpace = /^\s|\s$/.test(run.text) ? ' xml:space="preserve"' : "";
  const plainRun = `<w:r>${properties}<w:t${textSpace}>${escapeXml(run.text)}</w:t></w:r>`;
  const controlledRun = run.contentControl ? contentControlXml(run.contentControl, plainRun) : plainRun;
  const revisedRun = wrapRevisionIfNeeded(run, controlledRun);
  const bookmarkedRun = wrapBookmarkIfNeeded(run, revisedRun, context);
  const runContent = options.skipComment ? bookmarkedRun : wrapCommentIfNeeded(run, bookmarkedRun, context);

  if (!run.link) {
    return runContent;
  }

  if ("anchor" in run.link) {
    return `<w:hyperlink w:anchor="${escapeAttribute(run.link.anchor)}">${runContent}</w:hyperlink>`;
  }

  const relationshipId = `rIdHyperlink${context.hyperlinks.length + 1}`;
  context.hyperlinks.push({ id: relationshipId, url: run.link.url });
  return `<w:hyperlink r:id="${relationshipId}">${runContent}</w:hyperlink>`;
}

function mathRunXml(math: NonNullable<TextRun["math"]>): string {
  const nodes = math.nodes ?? (math.text !== undefined ? [{ type: "text" as const, text: math.text }] : []);

  return `<m:oMath>${nodes.map((node) => mathNodeXml(node)).join("")}</m:oMath>`;
}

function mathNodeXml(node: MathNode): string {
  if (node.type === "text") {
    return `<m:r><m:t>${escapeXml(node.text)}</m:t></m:r>`;
  }

  if (node.type === "fraction") {
    const properties = mathControlPropertiesXml(node.controlProperties);
    const fractionProperties = [
      fractionTypeXml(node.fractionType),
      properties ? `<m:ctrlPr>${properties}</m:ctrlPr>` : "",
    ].join("");
    return `<m:f>${fractionProperties ? `<m:fPr>${fractionProperties}</m:fPr>` : ""}<m:num>${node.numerator.map((child) => mathNodeXml(child)).join("")}</m:num><m:den>${node.denominator.map((child) => mathNodeXml(child)).join("")}</m:den></m:f>`;
  }

  if (node.type === "superscript") {
    const properties = mathControlPropertiesXml(node.controlProperties);
    return `<m:sSup>${properties ? `<m:sSupPr><m:ctrlPr>${properties}</m:ctrlPr></m:sSupPr>` : ""}<m:e>${node.base.map((child) => mathNodeXml(child)).join("")}</m:e><m:sup>${node.superscript.map((child) => mathNodeXml(child)).join("")}</m:sup></m:sSup>`;
  }

  if (node.type === "subscript") {
    const properties = mathControlPropertiesXml(node.controlProperties);
    return `<m:sSub>${properties ? `<m:sSubPr><m:ctrlPr>${properties}</m:ctrlPr></m:sSubPr>` : ""}<m:e>${node.base.map((child) => mathNodeXml(child)).join("")}</m:e><m:sub>${node.subscript.map((child) => mathNodeXml(child)).join("")}</m:sub></m:sSub>`;
  }

  if (node.type === "subSup") {
    const properties = mathControlPropertiesXml(node.controlProperties);
    return `<m:sSubSup>${properties ? `<m:sSubSupPr><m:ctrlPr>${properties}</m:ctrlPr></m:sSubSupPr>` : ""}<m:e>${node.base.map((child) => mathNodeXml(child)).join("")}</m:e><m:sub>${node.subscript.map((child) => mathNodeXml(child)).join("")}</m:sub><m:sup>${node.superscript.map((child) => mathNodeXml(child)).join("")}</m:sup></m:sSubSup>`;
  }

  if (node.type === "sPre") {
    const properties = mathControlPropertiesXml(node.controlProperties);
    return `<m:sPre>${properties ? `<m:sPrePr><m:ctrlPr>${properties}</m:ctrlPr></m:sPrePr>` : ""}<m:sub>${node.subscript.map((child) => mathNodeXml(child)).join("")}</m:sub><m:sup>${node.superscript.map((child) => mathNodeXml(child)).join("")}</m:sup><m:e>${node.base.map((child) => mathNodeXml(child)).join("")}</m:e></m:sPre>`;
  }

  if (node.type === "preSubSup") {
    const properties = mathControlPropertiesXml(node.controlProperties);
    return `<m:preSubSup>${properties ? `<m:preSubSupPr><m:ctrlPr>${properties}</m:ctrlPr></m:preSubSupPr>` : ""}<m:e>${node.base.map((child) => mathNodeXml(child)).join("")}</m:e><m:sub>${node.subscript.map((child) => mathNodeXml(child)).join("")}</m:sub><m:sup>${node.superscript.map((child) => mathNodeXml(child)).join("")}</m:sup></m:preSubSup>`;
  }

  if (node.type === "radical") {
    const properties = mathControlPropertiesXml(node.controlProperties);
    const radicalProperties = [
      node.hideDegree !== undefined ? `<m:degHide m:val="${node.hideDegree ? "1" : "0"}"/>` : "",
      properties ? `<m:ctrlPr>${properties}</m:ctrlPr>` : "",
    ].join("");
    const degree = node.degree ? `<m:deg>${node.degree.map((child) => mathNodeXml(child)).join("")}</m:deg>` : "";
    return `<m:rad>${radicalProperties ? `<m:radPr>${radicalProperties}</m:radPr>` : ""}${degree}<m:e>${node.content.map((child) => mathNodeXml(child)).join("")}</m:e></m:rad>`;
  }

  if (node.type === "matrix") {
    const controlProperties = mathControlPropertiesXml(node.controlProperties);
    const columnProperties = matrixColumnPropertiesXml(node.columnJustifications, node.columnCounts);
    const properties = [
      node.baseJustification !== undefined ? `<m:baseJc m:val="${matrixBaseJustificationXml(node.baseJustification)}"/>` : "",
      node.rowSpacing !== undefined ? `<m:rSp m:val="${node.rowSpacing}"/>` : "",
      node.rowSpacingRule !== undefined ? `<m:rSpRule m:val="${node.rowSpacingRule}"/>` : "",
      node.columnSpacing !== undefined ? `<m:cSp m:val="${node.columnSpacing}"/>` : "",
      node.columnSpacingRule !== undefined ? `<m:cSpRule m:val="${node.columnSpacingRule}"/>` : "",
      columnProperties,
      controlProperties ? `<m:ctrlPr>${controlProperties}</m:ctrlPr>` : "",
    ].join("");
    return `<m:m>${properties ? `<m:mPr>${properties}</m:mPr>` : ""}${node.rows.map((row) => `<m:mr>${row.map((cell) => `<m:e>${cell.map((child) => mathNodeXml(child)).join("")}</m:e>`).join("")}</m:mr>`).join("")}</m:m>`;
  }

  if (node.type === "delimiter") {
    const controlProperties = mathControlPropertiesXml(node.controlProperties);
    const delimiterProperties = [
      node.begin !== undefined ? `<m:begChr m:val="${escapeAttribute(node.begin)}"/>` : "",
      node.end !== undefined ? `<m:endChr m:val="${escapeAttribute(node.end)}"/>` : "",
      node.grow !== undefined ? `<m:grow m:val="${node.grow ? "1" : "0"}"/>` : "",
      node.separator !== undefined ? `<m:sepChr m:val="${escapeAttribute(node.separator)}"/>` : "",
      controlProperties ? `<m:ctrlPr>${controlProperties}</m:ctrlPr>` : "",
    ].join("");
    return `<m:d>${delimiterProperties ? `<m:dPr>${delimiterProperties}</m:dPr>` : ""}<m:e>${node.content.map((child) => mathNodeXml(child)).join("")}</m:e></m:d>`;
  }

  if (node.type === "accent") {
    const properties = mathControlPropertiesXml(node.controlProperties);
    return `<m:acc><m:accPr><m:chr m:val="${escapeAttribute(node.mark)}"/>${properties ? `<m:ctrlPr>${properties}</m:ctrlPr>` : ""}</m:accPr><m:e>${node.content.map((child) => mathNodeXml(child)).join("")}</m:e></m:acc>`;
  }

  if (node.type === "bar") {
    const properties = mathControlPropertiesXml(node.controlProperties);
    return `<m:bar><m:barPr><m:pos m:val="${node.position}"/>${properties ? `<m:ctrlPr>${properties}</m:ctrlPr>` : ""}</m:barPr><m:e>${node.content.map((child) => mathNodeXml(child)).join("")}</m:e></m:bar>`;
  }

  if (node.type === "function") {
    const properties = mathControlPropertiesXml(node.controlProperties);
    return `<m:func>${properties ? `<m:funcPr><m:ctrlPr>${properties}</m:ctrlPr></m:funcPr>` : ""}<m:fName>${node.name.map((child) => mathNodeXml(child)).join("")}</m:fName><m:e>${node.argument.map((child) => mathNodeXml(child)).join("")}</m:e></m:func>`;
  }

  if (node.type === "limitLower") {
    const properties = mathControlPropertiesXml(node.controlProperties);
    return `<m:limLow>${properties ? `<m:limLowPr><m:ctrlPr>${properties}</m:ctrlPr></m:limLowPr>` : ""}<m:e>${node.base.map((child) => mathNodeXml(child)).join("")}</m:e><m:lim>${node.limit.map((child) => mathNodeXml(child)).join("")}</m:lim></m:limLow>`;
  }

  if (node.type === "limitUpper") {
    const properties = mathControlPropertiesXml(node.controlProperties);
    return `<m:limUpp>${properties ? `<m:limUppPr><m:ctrlPr>${properties}</m:ctrlPr></m:limUppPr>` : ""}<m:e>${node.base.map((child) => mathNodeXml(child)).join("")}</m:e><m:lim>${node.limit.map((child) => mathNodeXml(child)).join("")}</m:lim></m:limUpp>`;
  }

  if (node.type === "equationArray") {
    const properties = mathControlPropertiesXml(node.controlProperties);
    const equationArrayProperties = [
      node.baseJustification !== undefined ? `<m:baseJc m:val="${matrixBaseJustificationXml(node.baseJustification)}"/>` : "",
      node.verticalJustification !== undefined ? `<m:vertJc m:val="${node.verticalJustification}"/>` : "",
      node.alignment !== undefined ? `<m:aln m:val="${node.alignment ? "1" : "0"}"/>` : "",
      node.rowSpacing !== undefined ? `<m:rSp m:val="${node.rowSpacing}"/>` : "",
      node.rowSpacingRule !== undefined ? `<m:rSpRule m:val="${node.rowSpacingRule}"/>` : "",
      node.objectDistribution !== undefined ? `<m:objDist m:val="${node.objectDistribution ? "1" : "0"}"/>` : "",
      node.maxDistribution !== undefined ? `<m:maxDist m:val="${node.maxDistribution ? "1" : "0"}"/>` : "",
      properties ? `<m:ctrlPr>${properties}</m:ctrlPr>` : "",
    ].join("");
    return `<m:eqArr>${equationArrayProperties ? `<m:eqArrPr>${equationArrayProperties}</m:eqArrPr>` : ""}${node.rows.map((row) => `<m:e>${row.map((child) => mathNodeXml(child)).join("")}</m:e>`).join("")}</m:eqArr>`;
  }

  if (node.type === "box") {
    const controlProperties = mathControlPropertiesXml(node.controlProperties);
    const properties = [
      node.hideTop !== undefined ? `<m:hideTop m:val="${node.hideTop ? "1" : "0"}"/>` : "",
      node.hideBottom !== undefined ? `<m:hideBot m:val="${node.hideBottom ? "1" : "0"}"/>` : "",
      node.hideLeft !== undefined ? `<m:hideLeft m:val="${node.hideLeft ? "1" : "0"}"/>` : "",
      node.hideRight !== undefined ? `<m:hideRight m:val="${node.hideRight ? "1" : "0"}"/>` : "",
      controlProperties ? `<m:ctrlPr>${controlProperties}</m:ctrlPr>` : "",
    ].join("");
    return `<m:box>${properties ? `<m:boxPr>${properties}</m:boxPr>` : ""}<m:e>${node.content.map((child) => mathNodeXml(child)).join("")}</m:e></m:box>`;
  }

  if (node.type === "borderBox") {
    const controlProperties = mathControlPropertiesXml(node.controlProperties);
    const properties = [
      node.hideTop !== undefined ? `<m:hideTop m:val="${node.hideTop ? "1" : "0"}"/>` : "",
      node.hideBottom !== undefined ? `<m:hideBot m:val="${node.hideBottom ? "1" : "0"}"/>` : "",
      node.hideLeft !== undefined ? `<m:hideLeft m:val="${node.hideLeft ? "1" : "0"}"/>` : "",
      node.hideRight !== undefined ? `<m:hideRight m:val="${node.hideRight ? "1" : "0"}"/>` : "",
      controlProperties ? `<m:ctrlPr>${controlProperties}</m:ctrlPr>` : "",
    ].join("");
    return `<m:borderBox>${properties ? `<m:borderBoxPr>${properties}</m:borderBoxPr>` : ""}<m:e>${node.content.map((child) => mathNodeXml(child)).join("")}</m:e></m:borderBox>`;
  }

  if (node.type === "phantom") {
    const controlProperties = mathControlPropertiesXml(node.controlProperties);
    const properties = [
      node.show !== undefined ? `<m:show m:val="${node.show ? "1" : "0"}"/>` : "",
      node.zeroWidth !== undefined ? `<m:zeroWid m:val="${node.zeroWidth ? "1" : "0"}"/>` : "",
      node.zeroAscent !== undefined ? `<m:zeroAsc m:val="${node.zeroAscent ? "1" : "0"}"/>` : "",
      node.zeroDescent !== undefined ? `<m:zeroDesc m:val="${node.zeroDescent ? "1" : "0"}"/>` : "",
      node.transparent !== undefined ? `<m:transp m:val="${node.transparent ? "1" : "0"}"/>` : "",
      controlProperties ? `<m:ctrlPr>${controlProperties}</m:ctrlPr>` : "",
    ].join("");
    return `<m:phant>${properties ? `<m:phantPr>${properties}</m:phantPr>` : ""}<m:e>${node.content.map((child) => mathNodeXml(child)).join("")}</m:e></m:phant>`;
  }

  if (node.type === "groupCharacter") {
    const controlProperties = mathControlPropertiesXml(node.controlProperties);
    const properties = [
      node.character !== undefined ? `<m:chr m:val="${escapeAttribute(node.character)}"/>` : "",
      node.position !== undefined ? `<m:pos m:val="${node.position}"/>` : "",
      node.verticalJustification !== undefined ? `<m:vertJc m:val="${node.verticalJustification}"/>` : "",
      controlProperties ? `<m:ctrlPr>${controlProperties}</m:ctrlPr>` : "",
    ].join("");
    return `<m:groupChr>${properties ? `<m:groupChrPr>${properties}</m:groupChrPr>` : ""}<m:e>${node.content.map((child) => mathNodeXml(child)).join("")}</m:e></m:groupChr>`;
  }

  const properties = mathControlPropertiesXml(node.controlProperties);
  const naryProperties = [
    naryOperatorXml(node.operator, node.operatorCharacter),
    naryLimitLocationXml(node.limitLocation),
    node.grow !== undefined ? `<m:grow m:val="${node.grow ? "1" : "0"}"/>` : "",
    node.hideLowerLimit !== undefined ? `<m:subHide m:val="${node.hideLowerLimit ? "1" : "0"}"/>` : "",
    node.hideUpperLimit !== undefined ? `<m:supHide m:val="${node.hideUpperLimit ? "1" : "0"}"/>` : "",
    properties ? `<m:ctrlPr>${properties}</m:ctrlPr>` : "",
  ].join("");
  return `<m:nary><m:naryPr>${naryProperties}</m:naryPr>` +
    (node.lowerLimit ? `<m:sub>${node.lowerLimit.map((child) => mathNodeXml(child)).join("")}</m:sub>` : "") +
    (node.upperLimit ? `<m:sup>${node.upperLimit.map((child) => mathNodeXml(child)).join("")}</m:sup>` : "") +
    `<m:e>${node.body.map((child) => mathNodeXml(child)).join("")}</m:e></m:nary>`;
}

function mathControlPropertiesXml(properties: MathControlProperties | undefined): string {
  return properties ? runPropertiesXml({ text: "", ...properties }) : "";
}

function fractionTypeXml(type: Extract<MathNode, { type: "fraction" }>["fractionType"]): string {
  if (type === undefined) {
    return "";
  }
  const value = type === "skewed" ? "skw" : type === "linear" ? "lin" : type;
  return `<m:type m:val="${value}"/>`;
}

function naryOperatorXml(operator: Extract<MathNode, { type: "nary" }>["operator"], character?: string): string {
  return `<m:chr m:val="${escapeAttribute(character ?? naryOperatorCharacter(operator))}"/>`;
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

function naryLimitLocationXml(location: Extract<MathNode, { type: "nary" }>["limitLocation"]): string {
  if (location === undefined) {
    return "";
  }
  return `<m:limLoc m:val="${location === "underOver" ? "undOvr" : "subSup"}"/>`;
}

function matrixBaseJustificationXml(value: "top" | "center" | "bottom" | undefined): string {
  return value === "bottom" ? "bot" : value ?? "center";
}

function matrixColumnPropertiesXml(
  justifications: Extract<MathNode, { type: "matrix" }>["columnJustifications"],
  counts: Extract<MathNode, { type: "matrix" }>["columnCounts"],
): string {
  const columnCount = Math.max(justifications?.length ?? 0, counts?.length ?? 0);
  if (columnCount === 0) {
    return "";
  }
  return `<m:mcs>${Array.from({ length: columnCount }, (_, index) => {
    const count = counts?.[index];
    const justification = justifications?.[index];
    const properties = [
      count !== undefined ? `<m:count m:val="${count}"/>` : "",
      justification !== undefined ? `<m:mcJc m:val="${justification}"/>` : "",
    ].join("");
    return `<m:mc><m:mcPr>${properties}</m:mcPr></m:mc>`;
  }).join("")}</m:mcs>`;
}

function wrapRevisionIfNeeded(run: TextRun, runContent: string): string {
  if (!run.revision) {
    return runContent;
  }

  if (run.revision.type === "insert") {
    return `<w:ins${revisionAttributes(run.revision)}>${runContent}</w:ins>`;
  }

  if (run.revision.type === "moveTo") {
    return `<w:moveTo${revisionAttributes(run.revision)}>${runContent}</w:moveTo>`;
  }

  const textSpace = /^\s|\s$/.test(run.text) ? ' xml:space="preserve"' : "";
  const wrapper = run.revision.type === "moveFrom" ? "moveFrom" : "del";
  return `<w:${wrapper}${revisionAttributes(run.revision)}><w:r>${runPropertiesXml(run)}<w:delText${textSpace}>${escapeXml(run.text)}</w:delText></w:r></w:${wrapper}>`;
}

function revisionAttributes(revision: RunRevision): string {
  return ` w:id="${revision.id}" w:author="${escapeAttribute(revision.author)}"` +
    (revision.date ? ` w:date="${escapeAttribute(revision.date)}"` : "");
}

function revisionElement(type: RunRevision["type"]): "ins" | "del" | "moveFrom" | "moveTo" {
  if (type === "insert") {
    return "ins";
  }

  if (type === "delete") {
    return "del";
  }

  return type;
}

function fieldRunXml(field: TextRun["field"]): string {
  const instruction = fieldInstruction(field);
  const result = typeof field === "object" && "result" in field && field.result !== undefined
    ? field.result
    : "";

  return `<w:r><w:fldChar w:fldCharType="begin"/></w:r>` +
    `<w:r><w:instrText xml:space="preserve">${instruction}</w:instrText></w:r>` +
    `<w:r><w:fldChar w:fldCharType="separate"/></w:r>` +
    `<w:r><w:t>${escapeXml(result)}</w:t></w:r>` +
    `<w:r><w:fldChar w:fldCharType="end"/></w:r>`;
}

function fieldInstruction(field: TextRun["field"]): string {
  if (typeof field === "object" && field.type === "toc") {
    return `TOC${field.switches ? ` ${tocSwitchesInstruction(field.switches)}` : ""}`;
  }

  if (typeof field === "object" && field.type === "page") {
    return "PAGE";
  }

  if (typeof field === "object" && field.type === "numPages") {
    return "NUMPAGES";
  }

  if (field === "page") {
    return "PAGE";
  }

  if (field === "numPages") {
    return "NUMPAGES";
  }

  if (field?.type === "ref") {
    return `REF ${field.target}`;
  }

  if (field?.type === "pageRef") {
    return `PAGEREF ${field.target}`;
  }

  return "PAGE";
}

function tocSwitchesInstruction(switches: string): string {
  return switches
    .split(/ (?=[a-z]\b)/i)
    .map((value) => `\\${value}`)
    .join(" ");
}

function wrapBookmarkIfNeeded(run: TextRun, runContent: string, context: WriterContext): string {
  if (!run.bookmark) {
    return runContent;
  }

  const id = context.bookmarkId;
  context.bookmarkId += 1;

  return `<w:bookmarkStart w:id="${id}" w:name="${escapeAttribute(run.bookmark.name)}"/>` +
    runContent +
    `<w:bookmarkEnd w:id="${id}"/>`;
}

function wrapCommentIfNeeded(run: TextRun, runContent: string, context: WriterContext): string {
  if (!run.comment) {
    return runContent;
  }

  const id = run.comment.id ?? nextCommentId(context);
  ensureCommentEntry(run, id, context);

  return `<w:commentRangeStart w:id="${id}"/>` +
    runContent +
    `<w:commentRangeEnd w:id="${id}"/>` +
    `<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="${id}"/></w:r>`;
}

function ensureCommentEntry(run: TextRun, id: number, context: WriterContext): void {
  if (!run.comment) {
    return;
  }

  ensureComment(run.comment, id, context);
}

function ensureComment(comment: TextRun["comment"], id: number, context: WriterContext): void {
  if (!comment || context.comments.some((entry) => entry.id === id)) {
    return;
  }

  context.comments.push({ id, ...comment });
}

function nextCommentId(context: WriterContext): number {
  const usedIds = new Set(context.comments.map((comment) => comment.id));
  let id = 0;

  while (usedIds.has(id)) {
    id += 1;
  }

  return id;
}

function runPropertiesXml(run: TextRun): string {
  const properties = [
    run.styleId ? `<w:rStyle w:val="${escapeAttribute(run.styleId)}"/>` : "",
    runFontsXml(run),
    run.bold !== undefined ? (run.bold ? "<w:b/>" : '<w:b w:val="0"/>') : "",
    run.italic !== undefined ? (run.italic ? "<w:i/>" : '<w:i w:val="0"/>') : "",
    run.allCaps !== undefined ? (run.allCaps ? "<w:caps/>" : '<w:caps w:val="0"/>') : "",
    run.smallCaps !== undefined ? (run.smallCaps ? "<w:smallCaps/>" : '<w:smallCaps w:val="0"/>') : "",
    run.strike !== undefined ? (run.strike ? "<w:strike/>" : '<w:strike w:val="0"/>') : "",
    run.doubleStrike !== undefined ? (run.doubleStrike ? "<w:dstrike/>" : '<w:dstrike w:val="0"/>') : "",
    run.outline !== undefined ? (run.outline ? "<w:outline/>" : '<w:outline w:val="0"/>') : "",
    run.shadow !== undefined ? (run.shadow ? "<w:shadow/>" : '<w:shadow w:val="0"/>') : "",
    run.emboss !== undefined ? (run.emboss ? "<w:emboss/>" : '<w:emboss w:val="0"/>') : "",
    run.imprint !== undefined ? (run.imprint ? "<w:imprint/>" : '<w:imprint w:val="0"/>') : "",
    run.noProof !== undefined ? (run.noProof ? "<w:noProof/>" : '<w:noProof w:val="0"/>') : "",
    run.snapToGrid !== undefined ? (run.snapToGrid ? "<w:snapToGrid/>" : '<w:snapToGrid w:val="0"/>') : "",
    run.hidden !== undefined ? (run.hidden ? "<w:vanish/>" : '<w:vanish w:val="0"/>') : "",
    run.webHidden !== undefined ? (run.webHidden ? "<w:webHidden/>" : '<w:webHidden w:val="0"/>') : "",
    run.color ? `<w:color w:val="${escapeAttribute(run.color)}"/>` : "",
    run.characterSpacing !== undefined ? `<w:spacing w:val="${run.characterSpacing}"/>` : "",
    run.scale !== undefined ? `<w:w w:val="${run.scale}"/>` : "",
    run.kerning !== undefined ? `<w:kern w:val="${run.kerning}"/>` : "",
    run.characterPosition !== undefined ? `<w:position w:val="${run.characterPosition}"/>` : "",
    run.fontSize ? `<w:sz w:val="${run.fontSize * 2}"/>` : "",
    run.complexScriptFontSize ? `<w:szCs w:val="${run.complexScriptFontSize * 2}"/>` : "",
    run.highlight ? `<w:highlight w:val="${run.highlight}"/>` : "",
    run.underline !== undefined ? `<w:u w:val="${run.underline ? "single" : "none"}"/>` : "",
    run.border ? borderSideXml("bdr", run.border) : "",
    run.fitText ? `<w:fitText w:val="${run.fitText.width}"${run.fitText.id !== undefined ? ` w:id="${run.fitText.id}"` : ""}/>` : "",
    run.verticalAlign ? `<w:vertAlign w:val="${run.verticalAlign}"/>` : "",
    run.rtl !== undefined ? (run.rtl ? "<w:rtl/>" : '<w:rtl w:val="0"/>') : "",
    run.complexScript !== undefined ? (run.complexScript ? "<w:cs/>" : '<w:cs w:val="0"/>') : "",
    run.emphasis ? `<w:em w:val="${run.emphasis}"/>` : "",
    run.language ? runLanguageXml(run.language) : "",
    run.specVanish !== undefined ? (run.specVanish ? "<w:specVanish/>" : '<w:specVanish w:val="0"/>') : "",
    run.officeMath !== undefined ? (run.officeMath ? "<w:oMath/>" : '<w:oMath w:val="0"/>') : "",
  ].join("");

  return properties ? `<w:rPr>${properties}</w:rPr>` : "";
}

function runFontsXml(run: TextRun): string {
  const attributes = [
    run.fontFamily ? ` w:ascii="${escapeAttribute(run.fontFamily)}" w:hAnsi="${escapeAttribute(run.fontFamily)}"` : "",
    run.eastAsiaFontFamily ? ` w:eastAsia="${escapeAttribute(run.eastAsiaFontFamily)}"` : "",
    run.complexScriptFontFamily ? ` w:cs="${escapeAttribute(run.complexScriptFontFamily)}"` : "",
    run.fontTheme ? ` w:asciiTheme="${escapeAttribute(run.fontTheme)}" w:hAnsiTheme="${escapeAttribute(run.fontTheme)}"` : "",
    run.eastAsiaFontTheme ? ` w:eastAsiaTheme="${escapeAttribute(run.eastAsiaFontTheme)}"` : "",
    run.complexScriptFontTheme ? ` w:cstheme="${escapeAttribute(run.complexScriptFontTheme)}"` : "",
    run.fontHint ? ` w:hint="${run.fontHint}"` : "",
  ].join("");

  return attributes ? `<w:rFonts${attributes}/>` : "";
}

function runLanguageXml(language: RunLanguage): string {
  const attributes = [
    language.value ? ` w:val="${escapeAttribute(language.value)}"` : "",
    language.eastAsia ? ` w:eastAsia="${escapeAttribute(language.eastAsia)}"` : "",
    language.bidi ? ` w:bidi="${escapeAttribute(language.bidi)}"` : "",
  ].join("");

  return attributes ? `<w:lang${attributes}/>` : "";
}

function tableXml(table: TableNode, context: WriterContext): string {
  const properties = [
    table.styleId ? `<w:tblStyle w:val="${escapeAttribute(table.styleId)}"/>` : "",
    table.position ? tablePositionXml(table.position) : "",
    table.overlap ? `<w:tblOverlap w:val="${table.overlap}"/>` : "",
    table.caption !== undefined ? `<w:tblCaption w:val="${escapeAttribute(table.caption)}"/>` : "",
    table.description !== undefined ? `<w:tblDescription w:val="${escapeAttribute(table.description)}"/>` : "",
    table.width !== undefined ? `<w:tblW w:w="${table.width}" w:type="${table.widthType ?? "dxa"}"/>` : "",
    table.borders ? tableBordersXml(table.borders) : "",
    table.alignment ? `<w:jc w:val="${table.alignment}"/>` : "",
    table.cellSpacing !== undefined ? `<w:tblCellSpacing w:w="${table.cellSpacing}" w:type="dxa"/>` : "",
    table.indent ? `<w:tblInd w:w="${table.indent.width}" w:type="${table.indent.type ?? "dxa"}"/>` : "",
    table.layout ? `<w:tblLayout w:type="${table.layout}"/>` : "",
    table.look ? tableLookXml(table.look) : "",
    table.propertyRevision ? propertyRevisionXml("tblPr", "tblPrChange", table.propertyRevision) : "",
  ].join("");
  const rows = table.rows
    .map((row) => {
      const propertyExceptions = row.propertyExceptions ? tablePropertyExceptionsXml(row.propertyExceptions) : "";
      const rowRevision = row.revision
        ? `<w:${revisionElement(row.revision.type)}${revisionAttributes(row.revision)}/>`
        : "";
      const rowHeight = row.height
        ? `<w:trHeight w:val="${row.height.value}"${row.height.rule ? ` w:hRule="${row.height.rule}"` : ""}/>`
        : "";
      const repeatHeader = row.repeatHeader ? "<w:tblHeader/>" : "";
      const cantSplit = row.cantSplit ? "<w:cantSplit/>" : "";
      const rowProperties = propertyExceptions || rowRevision || repeatHeader || cantSplit || rowHeight ? `<w:trPr>${propertyExceptions}${rowRevision}${repeatHeader}${cantSplit}${rowHeight}</w:trPr>` : "";

      return `<w:tr>${rowProperties}${row.cells.map((cell) => tableCellXml(cell, context)).join("")}</w:tr>`;
    })
    .join("");
  const grid = table.grid
    ? `<w:tblGrid>${table.grid.map((width) => `<w:gridCol w:w="${width}"/>`).join("")}</w:tblGrid>`
    : "";

  return `<w:tbl>${properties ? `<w:tblPr>${properties}</w:tblPr>` : ""}${grid}${rows}</w:tbl>`;
}

function tablePropertyExceptionsXml(exceptions: NonNullable<TableNode["rows"][number]["propertyExceptions"]>): string {
  const properties = [
    exceptions.width !== undefined ? `<w:tblW w:w="${exceptions.width}" w:type="${exceptions.widthType ?? "dxa"}"/>` : "",
    exceptions.cellSpacing !== undefined ? `<w:tblCellSpacing w:w="${exceptions.cellSpacing}" w:type="dxa"/>` : "",
    exceptions.indent ? `<w:tblInd w:w="${exceptions.indent.width}" w:type="${exceptions.indent.type ?? "dxa"}"/>` : "",
    exceptions.layout ? `<w:tblLayout w:type="${exceptions.layout}"/>` : "",
    exceptions.look ? tableLookXml(exceptions.look) : "",
  ].join("");

  return properties ? `<w:tblPrEx>${properties}</w:tblPrEx>` : "";
}

function tablePositionXml(position: NonNullable<TableNode["position"]>): string {
  const attributes = [
    position.leftFromText !== undefined ? ` w:leftFromText="${position.leftFromText}"` : "",
    position.rightFromText !== undefined ? ` w:rightFromText="${position.rightFromText}"` : "",
    position.topFromText !== undefined ? ` w:topFromText="${position.topFromText}"` : "",
    position.bottomFromText !== undefined ? ` w:bottomFromText="${position.bottomFromText}"` : "",
    position.horizontalAnchor ? ` w:horzAnchor="${position.horizontalAnchor}"` : "",
    position.verticalAnchor ? ` w:vertAnchor="${position.verticalAnchor}"` : "",
    position.x !== undefined ? ` w:tblpX="${position.x}"` : "",
    position.y !== undefined ? ` w:tblpY="${position.y}"` : "",
    position.xAlign ? ` w:tblpXSpec="${position.xAlign}"` : "",
    position.yAlign ? ` w:tblpYSpec="${position.yAlign}"` : "",
  ].join("");

  return attributes ? `<w:tblpPr${attributes}/>` : "";
}

function tableLookXml(look: NonNullable<TableNode["look"]>): string {
  const hasFlags = Object.values(look).some((value) => value !== undefined);
  if (!hasFlags) {
    return "";
  }

  let value = 0;
  if (look.firstRow) value |= 0x0020;
  if (look.lastRow) value |= 0x0040;
  if (look.firstColumn) value |= 0x0080;
  if (look.lastColumn) value |= 0x0100;
  if (look.bandedRows === false) value |= 0x0200;
  if (look.bandedColumns === false) value |= 0x0400;

  return `<w:tblLook w:val="${value.toString(16).toUpperCase().padStart(4, "0")}"/>`;
}

function tableCellXml(cell: TableCellNode, context: WriterContext): string {
  const properties = [
    cell.width !== undefined ? `<w:tcW w:w="${cell.width}" w:type="${cell.widthType ?? "dxa"}"/>` : "",
    cell.colSpan ? `<w:gridSpan w:val="${cell.colSpan}"/>` : "",
    cell.verticalMerge ? `<w:vMerge w:val="${cell.verticalMerge}"/>` : "",
    cell.borders ? tableCellBordersXml(cell.borders) : "",
    cell.shading ? shadingXml(cell.shading) : "",
    cell.noWrap ? "<w:noWrap/>" : "",
    cell.margins ? tableCellMarginsXml(cell.margins) : "",
    cell.textDirection ? `<w:textDirection w:val="${cell.textDirection}"/>` : "",
    cell.fitText ? "<w:tcFitText/>" : "",
    cell.verticalAlignment ? `<w:vAlign w:val="${cell.verticalAlignment}"/>` : "",
    cell.propertyRevision ? propertyRevisionXml("tcPr", "tcPrChange", cell.propertyRevision) : "",
  ].join("");
  const blocks = cell.blocks.map((block) => paragraphXml(block, context)).join("");

  return `<w:tc>${properties ? `<w:tcPr>${properties}</w:tcPr>` : ""}${blocks}</w:tc>`;
}

function propertyRevisionXml(
  propertyElement: "pPr" | "tblPr" | "tcPr",
  changeElement: "pPrChange" | "tblPrChange" | "tcPrChange",
  revision: NonNullable<ParagraphNode["propertyRevision"]>,
): string {
  return `<w:${changeElement} w:id="${revision.id}" w:author="${escapeAttribute(revision.author)}"` +
    (revision.date ? ` w:date="${escapeAttribute(revision.date)}"` : "") +
    `><w:${propertyElement}/></w:${changeElement}>`;
}

function tableCellBordersXml(borders: NonNullable<TableCellNode["borders"]>): string {
  const sides = [
    borders.top ? borderSideXml("top", borders.top) : "",
    borders.left ? borderSideXml("left", borders.left) : "",
    borders.bottom ? borderSideXml("bottom", borders.bottom) : "",
    borders.right ? borderSideXml("right", borders.right) : "",
    borders.insideH ? borderSideXml("insideH", borders.insideH) : "",
    borders.insideV ? borderSideXml("insideV", borders.insideV) : "",
  ].join("");

  return sides ? `<w:tcBorders>${sides}</w:tcBorders>` : "";
}

function tableCellMarginsXml(margins: NonNullable<TableCellNode["margins"]>): string {
  const sides = [
    tableCellMarginSideXml("top", margins.top),
    tableCellMarginSideXml("right", margins.right),
    tableCellMarginSideXml("bottom", margins.bottom),
    tableCellMarginSideXml("left", margins.left),
  ].join("");

  return sides ? `<w:tcMar>${sides}</w:tcMar>` : "";
}

function tableCellMarginSideXml(side: "top" | "right" | "bottom" | "left", margin: NonNullable<TableCellNode["margins"]>[typeof side] | undefined): string {
  if (margin === undefined) {
    return "";
  }

  const width = typeof margin === "number" ? margin : margin.width;
  const type = typeof margin === "number" ? "dxa" : margin.type ?? "dxa";
  return `<w:${side} w:w="${width}" w:type="${type}"/>`;
}

function tableBordersXml(border: NonNullable<TableNode["borders"]>): string {
  if (border === "single") {
    const defaultBorder = { style: border, size: 4, space: 0, color: "auto" } satisfies BorderDefinition;
    return `<w:tblBorders>` +
      borderSideXml("top", defaultBorder) +
      borderSideXml("left", defaultBorder) +
      borderSideXml("bottom", defaultBorder) +
      borderSideXml("right", defaultBorder) +
      borderSideXml("insideH", defaultBorder) +
      borderSideXml("insideV", defaultBorder) +
      `</w:tblBorders>`;
  }

  const sides = [
    border.top ? borderSideXml("top", border.top) : "",
    border.left ? borderSideXml("left", border.left) : "",
    border.bottom ? borderSideXml("bottom", border.bottom) : "",
    border.right ? borderSideXml("right", border.right) : "",
    border.insideH ? borderSideXml("insideH", border.insideH) : "",
    border.insideV ? borderSideXml("insideV", border.insideV) : "",
  ].join("");

  return sides ? `<w:tblBorders>${sides}</w:tblBorders>` : "";
}

function imageXml(image: ImageNode, context: WriterContext): string {
  const id = context.images.length + 1;
  const relationshipId = `rIdImage${id}`;
  const filename = `image${id}.${image.contentType === "image/png" ? "png" : "jpg"}`;
  const widthEmu = image.width * 9525;
  const heightEmu = image.height * 9525;
  const rotation = image.rotation !== undefined ? ` rot="${image.rotation * 60000}"` : "";
  const crop = image.crop ? `<a:srcRect` +
    (image.crop.left !== undefined ? ` l="${image.crop.left}"` : "") +
    (image.crop.top !== undefined ? ` t="${image.crop.top}"` : "") +
    (image.crop.right !== undefined ? ` r="${image.crop.right}"` : "") +
    (image.crop.bottom !== undefined ? ` b="${image.crop.bottom}"` : "") +
    `/>` : "";

  context.images.push({
    id: relationshipId,
    filename,
    contentType: image.contentType,
    data: image.data,
  });

  const graphic = `<wp:extent cx="${widthEmu}" cy="${heightEmu}"/>` +
    `<wp:docPr id="${id}" name="Image ${id}"${image.altText ? ` descr="${escapeAttribute(image.altText)}"` : ""}/>` +
    `<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">` +
    `<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
    `<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
    `<pic:nvPicPr><pic:cNvPr id="${id}" name="Image ${id}"${image.altText ? ` descr="${escapeAttribute(image.altText)}"` : ""}/><pic:cNvPicPr/></pic:nvPicPr>` +
    `<pic:blipFill><a:blip r:embed="${relationshipId}"/>${crop}</pic:blipFill>` +
    `<pic:spPr><a:xfrm${rotation}><a:ext cx="${widthEmu}" cy="${heightEmu}"/></a:xfrm></pic:spPr>` +
    `</pic:pic></a:graphicData></a:graphic>`;
  const drawing = image.floating
    ? `<wp:anchor xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" simplePos="${image.floating.simplePosition ? "1" : "0"}" relativeHeight="${image.floating.relativeHeight ?? 0}"` +
      (image.floating.distanceTop !== undefined ? ` distT="${image.floating.distanceTop}"` : "") +
      (image.floating.distanceBottom !== undefined ? ` distB="${image.floating.distanceBottom}"` : "") +
      (image.floating.distanceLeft !== undefined ? ` distL="${image.floating.distanceLeft}"` : "") +
      (image.floating.distanceRight !== undefined ? ` distR="${image.floating.distanceRight}"` : "") +
      ` behindDoc="${image.floating.behindDoc ? "1" : "0"}" locked="${image.floating.locked ? "1" : "0"}" layoutInCell="${image.floating.layoutInCell === false ? "0" : "1"}" allowOverlap="${image.floating.allowOverlap === false ? "0" : "1"}">` +
      imageSimplePositionXml(image.floating) +
      imagePositionHXml(image.floating) +
      imagePositionVXml(image.floating) +
      imageEffectExtentXml(image.floating) +
      imageWrapXml(image.floating) +
      graphic +
      `</wp:anchor>`
    : `<wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">${graphic}</wp:inline>`;

  return `<w:p><w:r><w:drawing>` +
    drawing +
    `</w:drawing></w:r></w:p>`;
}

function imageEffectExtentXml(floating: NonNullable<ImageNode["floating"]>): string {
  return floating.effectExtent
    ? `<wp:effectExtent t="${floating.effectExtent.top}" b="${floating.effectExtent.bottom}" l="${floating.effectExtent.left}" r="${floating.effectExtent.right}"/>`
    : "";
}

function imageSimplePositionXml(floating: NonNullable<ImageNode["floating"]>): string {
  return floating.simplePosition
    ? `<wp:simplePos x="${floating.simplePosition.x}" y="${floating.simplePosition.y}"/>`
    : "";
}

function imagePositionHXml(floating: NonNullable<ImageNode["floating"]>): string {
  const position = floating.horizontalAlign
    ? `<wp:align>${floating.horizontalAlign}</wp:align>`
    : `<wp:posOffset>${floating.horizontalOffset}</wp:posOffset>`;

  return `<wp:positionH relativeFrom="${floating.horizontalRelativeFrom ?? "page"}">${position}</wp:positionH>`;
}

function imagePositionVXml(floating: NonNullable<ImageNode["floating"]>): string {
  const position = floating.verticalAlign
    ? `<wp:align>${floating.verticalAlign}</wp:align>`
    : `<wp:posOffset>${floating.verticalOffset}</wp:posOffset>`;

  return `<wp:positionV relativeFrom="${floating.verticalRelativeFrom ?? "page"}">${position}</wp:positionV>`;
}

function imageWrapXml(floating: NonNullable<ImageNode["floating"]>): string {
  const wrap = floating.wrap;

  if (wrap === "none") {
    return "<wp:wrapNone/>";
  }

  if (wrap === "topAndBottom") {
    return "<wp:wrapTopAndBottom/>";
  }

  if (wrap === "tight") {
    return imagePolygonWrapXml("wrapTight", floating);
  }

  if (wrap === "through") {
    return imagePolygonWrapXml("wrapThrough", floating);
  }

  return floating.wrapText
    ? `<wp:wrapSquare wrapText="${floating.wrapText}"/>`
    : "<wp:wrapSquare/>";
}

function imagePolygonWrapXml(tag: "wrapTight" | "wrapThrough", floating: NonNullable<ImageNode["floating"]>): string {
  const polygon = floating.wrapPolygon ?? { edited: false, start: { x: 0, y: 0 }, points: [{ x: 0, y: 0 }] };
  const points = polygon.points
    .map((point) => `<wp:lineTo x="${point.x}" y="${point.y}"/>`)
    .join("");

  return `<wp:${tag} wrapText="${floating.wrapText ?? "bothSides"}">` +
    `<wp:wrapPolygon edited="${polygon.edited ? "1" : "0"}">` +
    `<wp:start x="${polygon.start.x}" y="${polygon.start.y}"/>` +
    points +
    `</wp:wrapPolygon></wp:${tag}>`;
}

function contentTypesXml(context: WriterContext): string {
  const commentsOverride = context.comments.length > 0
    ? `<Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>`
    : "";
  const footnotesOverride = context.footnotes.length > 0
    ? `<Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>`
    : "";
  const endnotesOverride = context.endnotes.length > 0
    ? `<Override PartName="/word/endnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"/>`
    : "";
  const themeOverride = context.theme
    ? `<Override PartName="/word/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>`
    : "";
  const settingsOverride = context.settings
    ? `<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>`
    : "";
  const webSettingsOverride = context.settings?.web
    ? `<Override PartName="/word/webSettings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml"/>`
    : "";
  const fontTableOverride = context.fonts && context.fonts.length > 0
    ? `<Override PartName="/word/fontTable.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"/>`
    : "";
  const corePropertiesOverride = context.properties?.core
    ? `<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>`
    : "";
  const appPropertiesOverride = context.properties?.app
    ? `<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>`
    : "";
  const customPropertiesOverride = context.properties?.custom && context.properties.custom.length > 0
    ? `<Override PartName="/docProps/custom.xml" ContentType="application/vnd.openxmlformats-officedocument.custom-properties+xml"/>`
    : "";

  const imageDefaults = [
    context.images.some((image) => image.contentType === "image/png")
      ? `<Default Extension="png" ContentType="image/png"/>`
      : "",
    context.images.some((image) => image.contentType === "image/jpeg")
      ? `<Default Extension="jpg" ContentType="image/jpeg"/>`
      : "",
  ].join("");
  const headerOverrides = context.headers
    .map((header) => `<Override PartName="/word/${header.filename}" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>`)
    .join("");
  const footerOverrides = context.footers
    .map((footer) => `<Override PartName="/word/${footer.filename}" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>`)
    .join("");
  const customXmlPropertiesOverrides = context.customXmlParts
    .map((part, index) => part.properties ? (part.properties.path ?? defaultCustomXmlPropertiesPath(index)) : undefined)
    .filter((path): path is string => path !== undefined)
    .map((path) => `<Override PartName="/${escapeAttribute(path)}" ContentType="application/vnd.openxmlformats-officedocument.customXmlProperties+xml"/>`)
    .join("");

  return xmlDeclaration(
    `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
      `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
      `<Default Extension="xml" ContentType="application/xml"/>` +
      imageDefaults +
      `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>` +
      `<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>` +
      `<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>` +
      headerOverrides +
      footerOverrides +
      commentsOverride +
      footnotesOverride +
      endnotesOverride +
      themeOverride +
      settingsOverride +
      webSettingsOverride +
      fontTableOverride +
      corePropertiesOverride +
      appPropertiesOverride +
      customPropertiesOverride +
      customXmlPropertiesOverrides +
      `</Types>`,
  );
}

function packageRelsXml(context: WriterContext): string {
  const customXmlRelationships = context.customXmlParts
    .map((part, index) => `<Relationship Id="rIdCustomXml${index + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml" Target="${escapeAttribute(part.path)}"/>`)
    .join("");
  const corePropertiesRelationship = context.properties?.core
    ? `<Relationship Id="rIdCoreProperties" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>`
    : "";
  const appPropertiesRelationship = context.properties?.app
    ? `<Relationship Id="rIdAppProperties" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>`
    : "";
  const customPropertiesRelationship = context.properties?.custom && context.properties.custom.length > 0
    ? `<Relationship Id="rIdCustomProperties" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties" Target="docProps/custom.xml"/>`
    : "";

  return xmlDeclaration(
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>` +
      corePropertiesRelationship +
      appPropertiesRelationship +
      customPropertiesRelationship +
      customXmlRelationships +
      `</Relationships>`,
  );
}

function corePropertiesXml(core: NonNullable<NonNullable<DocumentJson["properties"]>["core"]>): string {
  return xmlDeclaration(
    `<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">` +
      (core.title ? `<dc:title>${escapeXml(core.title)}</dc:title>` : "") +
      (core.subject ? `<dc:subject>${escapeXml(core.subject)}</dc:subject>` : "") +
      (core.creator ? `<dc:creator>${escapeXml(core.creator)}</dc:creator>` : "") +
      (core.keywords ? `<cp:keywords>${escapeXml(core.keywords)}</cp:keywords>` : "") +
      (core.description ? `<dc:description>${escapeXml(core.description)}</dc:description>` : "") +
      (core.lastModifiedBy ? `<cp:lastModifiedBy>${escapeXml(core.lastModifiedBy)}</cp:lastModifiedBy>` : "") +
      (core.created ? `<dcterms:created xsi:type="dcterms:W3CDTF">${escapeXml(core.created)}</dcterms:created>` : "") +
      (core.modified ? `<dcterms:modified xsi:type="dcterms:W3CDTF">${escapeXml(core.modified)}</dcterms:modified>` : "") +
      `</cp:coreProperties>`,
  );
}

function appPropertiesXml(app: NonNullable<NonNullable<DocumentJson["properties"]>["app"]>): string {
  return xmlDeclaration(
    `<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">` +
      (app.application ? `<Application>${escapeXml(app.application)}</Application>` : "") +
      (app.company ? `<Company>${escapeXml(app.company)}</Company>` : "") +
      (app.manager ? `<Manager>${escapeXml(app.manager)}</Manager>` : "") +
      (app.pages !== undefined ? `<Pages>${app.pages}</Pages>` : "") +
      (app.words !== undefined ? `<Words>${app.words}</Words>` : "") +
      (app.characters !== undefined ? `<Characters>${app.characters}</Characters>` : "") +
      (app.lines !== undefined ? `<Lines>${app.lines}</Lines>` : "") +
      (app.paragraphs !== undefined ? `<Paragraphs>${app.paragraphs}</Paragraphs>` : "") +
      `</Properties>`,
  );
}

function customPropertiesXml(custom: NonNullable<NonNullable<DocumentJson["properties"]>["custom"]>): string {
  const properties = custom
    .map((property, index) => `<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="${index + 2}" name="${escapeAttribute(property.name)}">${customPropertyValueXml(property)}</property>`)
    .join("");

  return xmlDeclaration(
    `<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">` +
      properties +
      `</Properties>`,
  );
}

function customPropertyValueXml(property: NonNullable<NonNullable<DocumentJson["properties"]>["custom"]>[number]): string {
  if (property.type === "number") {
    return `<vt:i4>${property.value}</vt:i4>`;
  }

  if (property.type === "boolean") {
    return `<vt:bool>${property.value ? "true" : "false"}</vt:bool>`;
  }

  if (property.type === "date") {
    return `<vt:filetime>${escapeXml(String(property.value))}</vt:filetime>`;
  }

  return `<vt:lpwstr>${escapeXml(String(property.value))}</vt:lpwstr>`;
}

function customXmlPropertiesXml(properties: NonNullable<NonNullable<DocumentJson["customXmlParts"]>[number]["properties"]>): string {
  const schemaRefs = (properties.schemaRefs ?? [])
    .map((schemaRef) => `<ds:schemaRef ds:uri="${escapeAttribute(schemaRef)}"/>`)
    .join("");

  return xmlDeclaration(
    `<ds:datastoreItem ds:itemID="${escapeAttribute(properties.storeItemId ?? "")}" xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml">` +
      `<ds:schemaRefs>${schemaRefs}</ds:schemaRefs>` +
      `</ds:datastoreItem>`,
  );
}

function customXmlItemRelsXml(propertiesPath: string): string {
  return xmlDeclaration(
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      `<Relationship Id="rIdCustomXmlProps1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps" Target="${escapeAttribute(pathBasename(propertiesPath))}"/>` +
      `</Relationships>`,
  );
}

function customXmlRelationshipPath(partPath: string): string {
  return `${pathDirname(partPath)}/_rels/${pathBasename(partPath)}.rels`;
}

function defaultCustomXmlPropertiesPath(index: number): string {
  return `customXml/itemProps${index + 1}.xml`;
}

function documentRelsXml(context: WriterContext): string {
  const hyperlinkRelationships = context.hyperlinks
    .map((relationship) => `<Relationship Id="${relationship.id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="${escapeAttribute(relationship.url)}" TargetMode="External"/>`)
    .join("");
  const imageRelationships = context.images
    .map((image) => `<Relationship Id="${image.id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/${image.filename}"/>`)
    .join("");
  const headerRelationships = context.headers
    .map((header) => `<Relationship Id="${header.id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="${header.filename}"/>`)
    .join("");
  const footerRelationships = context.footers
    .map((footer) => `<Relationship Id="${footer.id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="${footer.filename}"/>`)
    .join("");

  return xmlDeclaration(
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      `<Relationship Id="rIdNumbering" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>` +
      (context.comments.length > 0 ? `<Relationship Id="rIdComments" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/>` : "") +
      (context.footnotes.length > 0 ? `<Relationship Id="rIdFootnotes" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/>` : "") +
      (context.endnotes.length > 0 ? `<Relationship Id="rIdEndnotes" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes" Target="endnotes.xml"/>` : "") +
      (context.theme ? `<Relationship Id="rIdTheme" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>` : "") +
      (context.settings ? `<Relationship Id="rIdSettings" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>` : "") +
      (context.settings?.web ? `<Relationship Id="rIdWebSettings" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings" Target="webSettings.xml"/>` : "") +
      (context.fonts && context.fonts.length > 0 ? `<Relationship Id="rIdFontTable" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable" Target="fontTable.xml"/>` : "") +
      headerRelationships +
      footerRelationships +
      imageRelationships +
      hyperlinkRelationships +
      `</Relationships>`,
  );
}

function themeXml(theme: NonNullable<DocumentJson["theme"]>): string {
  const colorScheme = [
    themeColorXml("dk1", theme.colors.dark1 ?? "000000"),
    themeColorXml("lt1", theme.colors.light1 ?? "FFFFFF"),
    themeColorXml("dk2", theme.colors.dark2 ?? "1F497D"),
    themeColorXml("lt2", theme.colors.light2 ?? "EEECE1"),
    themeColorXml("accent1", theme.colors.accent1),
    themeColorXml("accent2", theme.colors.accent2 ?? "C0504D"),
    themeColorXml("accent3", theme.colors.accent3 ?? "9BBB59"),
    themeColorXml("accent4", theme.colors.accent4 ?? "8064A2"),
    themeColorXml("accent5", theme.colors.accent5 ?? "4BACC6"),
    themeColorXml("accent6", theme.colors.accent6 ?? "F79646"),
    themeColorXml("hlink", theme.colors.hyperlink ?? "0000FF"),
    themeColorXml("folHlink", theme.colors.followedHyperlink ?? "800080"),
  ].join("");

  return xmlDeclaration(
    `<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="${escapeAttribute(theme.name)}">` +
      `<a:themeElements>` +
      `<a:clrScheme name="${escapeAttribute(theme.name)}">` +
      colorScheme +
      `</a:clrScheme>` +
      `<a:fontScheme name="${escapeAttribute(theme.name)}">` +
      themeFontXml("majorFont", theme.fonts.major, theme.fonts.majorEastAsia, theme.fonts.majorComplexScript, theme.fonts.supplemental?.filter((font) => font.group === "major") ?? []) +
      themeFontXml("minorFont", theme.fonts.minor, theme.fonts.minorEastAsia, theme.fonts.minorComplexScript, theme.fonts.supplemental?.filter((font) => font.group === "minor") ?? []) +
      `</a:fontScheme>` +
      (theme.formatScheme ? themeFormatSchemeXml(theme.formatScheme) : "") +
      `</a:themeElements>` +
      `</a:theme>`,
  );
}

function themeFontXml(tag: "majorFont" | "minorFont", latin: string, eastAsia: string | undefined, complexScript: string | undefined, supplemental: NonNullable<DocumentJson["theme"]>["fonts"]["supplemental"]): string {
  return `<a:${tag}>` +
    `<a:latin typeface="${escapeAttribute(latin)}"/>` +
    `<a:ea typeface="${escapeAttribute(eastAsia ?? "")}"/>` +
    `<a:cs typeface="${escapeAttribute(complexScript ?? "")}"/>` +
    (supplemental ?? []).map((font) => `<a:font script="${escapeAttribute(font.script)}" typeface="${escapeAttribute(font.typeface)}"/>`).join("") +
    `</a:${tag}>`;
}

function themeColorXml(tag: string, color: string | undefined): string {
  return color ? `<a:${tag}><a:srgbClr val="${escapeAttribute(color)}"/></a:${tag}>` : "";
}

function themeFormatSchemeXml(formatScheme: NonNullable<NonNullable<DocumentJson["theme"]>["formatScheme"]>): string {
  const fillStyles = padThemeColors(formatScheme.fillStyleColors, ["FFFFFF", "EEECE1", "D9EAF7"])
    .map((color) => `<a:solidFill><a:srgbClr val="${escapeAttribute(color)}"/></a:solidFill>`)
    .join("");
  const lineStyles = padThemeColors(formatScheme.lineStyleColors, ["000000", "808080", "A6A6A6"])
    .map((color) => `<a:ln w="9525"><a:solidFill><a:srgbClr val="${escapeAttribute(color)}"/></a:solidFill></a:ln>`)
    .join("");
  const effectStyles = padThemeColors(formatScheme.effectStyleColors, ["000000", "808080", "A6A6A6"])
    .map((color) => `<a:effectStyle><a:effectLst><a:outerShdw><a:srgbClr val="${escapeAttribute(color)}"/></a:outerShdw></a:effectLst></a:effectStyle>`)
    .join("");
  const backgroundFillStyles = padThemeColors(formatScheme.backgroundFillStyleColors, ["FFFFFF", "EEECE1", "D9EAF7"])
    .map((color) => `<a:solidFill><a:srgbClr val="${escapeAttribute(color)}"/></a:solidFill>`)
    .join("");

  return `<a:fmtScheme name="${escapeAttribute(formatScheme.name)}">` +
    `<a:fillStyleLst>${fillStyles}</a:fillStyleLst>` +
    `<a:lnStyleLst>${lineStyles}</a:lnStyleLst>` +
    `<a:effectStyleLst>${effectStyles}</a:effectStyleLst>` +
    `<a:bgFillStyleLst>${backgroundFillStyles}</a:bgFillStyleLst>` +
    `</a:fmtScheme>`;
}

function padThemeColors(colors: string[] | undefined, defaults: string[]): string[] {
  const values = [...(colors ?? [])];
  for (const color of defaults) {
    if (values.length >= defaults.length) break;
    values.push(color);
  }
  return values;
}

function stylesXml(document: DocumentJson): string {
  const defaults = document.styles?.defaults ? docDefaultsXml(document.styles.defaults) : "";
  const paragraphStyles = (document.styles?.paragraph ?? [])
    .map((style) => `<w:style w:type="paragraph" w:styleId="${escapeAttribute(style.id)}">` +
      `<w:name w:val="${escapeAttribute(style.name)}"/>` +
      (style.basedOn ? `<w:basedOn w:val="${escapeAttribute(style.basedOn)}"/>` : "") +
      (style.next ? `<w:next w:val="${escapeAttribute(style.next)}"/>` : "") +
      paragraphStylePropertiesXml(style.paragraph) +
      styleRunPropertiesXml(style.run) +
      `</w:style>`)
    .join("");
  const characterStyles = (document.styles?.character ?? [])
    .map((style) => styleXml("character", style.id, style.name, style.basedOn, style.run))
    .join("");
  const tableStyles = (document.styles?.table ?? [])
    .map((style) => styleXml("table", style.id, style.name, style.basedOn, style.run, tableStylePropertiesXml(style.table)))
    .join("");

  return xmlDeclaration(
    `<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
      `<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>` +
      `<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:qFormat/></w:style>` +
      `<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:qFormat/></w:style>` +
      `<w:style w:type="paragraph" w:styleId="Heading3"><w:name w:val="heading 3"/><w:basedOn w:val="Normal"/><w:qFormat/></w:style>` +
      defaults +
      paragraphStyles +
      characterStyles +
      tableStyles +
      `</w:styles>`,
  );
}

function docDefaultsXml(defaults: NonNullable<NonNullable<DocumentJson["styles"]>["defaults"]>): string {
  const runDefaults = defaults.run
    ? `<w:rPrDefault>${styleRunPropertiesXml(defaults.run)}</w:rPrDefault>`
    : "";
  const paragraphDefaults = defaults.paragraph
    ? `<w:pPrDefault>${paragraphStylePropertiesXml(defaults.paragraph)}</w:pPrDefault>`
    : "";

  return runDefaults || paragraphDefaults ? `<w:docDefaults>${runDefaults}${paragraphDefaults}</w:docDefaults>` : "";
}

function styleXml(type: "character" | "table", id: string, name: string, basedOn?: string, run?: StyleRunProperties, properties = ""): string {
  return `<w:style w:type="${type}" w:styleId="${escapeAttribute(id)}">` +
    `<w:name w:val="${escapeAttribute(name)}"/>` +
    (basedOn ? `<w:basedOn w:val="${escapeAttribute(basedOn)}"/>` : "") +
    styleRunPropertiesXml(run) +
    properties +
    `</w:style>`;
}

function paragraphStylePropertiesXml(properties?: StyleParagraphProperties): string {
  if (!properties) {
    return "";
  }

  const paragraphProperties = [
    properties.alignment ? `<w:jc w:val="${properties.alignment}"/>` : "",
    properties.spacing ? paragraphSpacingXml(properties.spacing) : "",
    properties.indent ? paragraphIndentXml(properties.indent) : "",
    properties.shading ? shadingXml(properties.shading) : "",
    properties.borders ? paragraphBordersXml(properties.borders) : "",
    paragraphTabsXml(properties.tabs),
    paragraphPaginationXml(properties.pagination),
    paragraphFrameXml(properties.frame),
  ].join("");

  return paragraphProperties ? `<w:pPr>${paragraphProperties}</w:pPr>` : "";
}

function styleRunPropertiesXml(run?: StyleRunProperties): string {
  if (!run) {
    return "";
  }

  const properties = [
    styleOnOffXml("b", run.bold),
    styleOnOffXml("i", run.italic),
    run.underline !== undefined ? `<w:u w:val="${run.underline ? "single" : "none"}"/>` : "",
    styleRunFontsXml(run),
    run.fontSize ? `<w:sz w:val="${run.fontSize * 2}"/>` : "",
    run.complexScriptFontSize ? `<w:szCs w:val="${run.complexScriptFontSize * 2}"/>` : "",
    run.color ? `<w:color w:val="${escapeAttribute(run.color)}"/>` : "",
    run.highlight ? `<w:highlight w:val="${run.highlight}"/>` : "",
    styleOnOffXml("strike", run.strike),
    styleOnOffXml("dstrike", run.doubleStrike),
    styleOnOffXml("smallCaps", run.smallCaps),
    styleOnOffXml("caps", run.allCaps),
    styleOnOffXml("shadow", run.shadow),
    styleOnOffXml("outline", run.outline),
    styleOnOffXml("emboss", run.emboss),
    styleOnOffXml("imprint", run.imprint),
    styleOnOffXml("rtl", run.rtl),
    styleOnOffXml("cs", run.complexScript),
    styleOnOffXml("specVanish", run.specVanish),
    styleOnOffXml("vanish", run.hidden),
    styleOnOffXml("webHidden", run.webHidden),
    styleOnOffXml("snapToGrid", run.snapToGrid),
    styleOnOffXml("noProof", run.noProof),
    styleOnOffXml("oMath", run.officeMath),
    run.language ? runLanguageXml(run.language) : "",
    run.characterPosition !== undefined ? `<w:position w:val="${run.characterPosition}"/>` : "",
    run.kerning !== undefined ? `<w:kern w:val="${run.kerning}"/>` : "",
    run.verticalAlign ? `<w:vertAlign w:val="${run.verticalAlign}"/>` : "",
    run.characterSpacing !== undefined ? `<w:spacing w:val="${run.characterSpacing}"/>` : "",
    run.scale !== undefined ? `<w:w w:val="${run.scale}"/>` : "",
    run.fitText ? `<w:fitText w:val="${run.fitText.width}"${run.fitText.id !== undefined ? ` w:id="${run.fitText.id}"` : ""}/>` : "",
    run.emphasis ? `<w:em w:val="${run.emphasis}"/>` : "",
    run.border ? borderSideXml("bdr", run.border) : "",
  ].join("");

  return properties ? `<w:rPr>${properties}</w:rPr>` : "";
}

function styleOnOffXml(tag: string, value: boolean | undefined): string {
  return value !== undefined ? (value ? `<w:${tag}/>` : `<w:${tag} w:val="0"/>`) : "";
}

function styleRunFontsXml(run: StyleRunProperties): string {
  const attributes = [
    run.fontFamily ? ` w:ascii="${escapeAttribute(run.fontFamily)}" w:hAnsi="${escapeAttribute(run.fontFamily)}"` : "",
    run.eastAsiaFontFamily ? ` w:eastAsia="${escapeAttribute(run.eastAsiaFontFamily)}"` : "",
    run.complexScriptFontFamily ? ` w:cs="${escapeAttribute(run.complexScriptFontFamily)}"` : "",
    run.fontTheme ? ` w:asciiTheme="${escapeAttribute(run.fontTheme)}" w:hAnsiTheme="${escapeAttribute(run.fontTheme)}"` : "",
    run.eastAsiaFontTheme ? ` w:eastAsiaTheme="${escapeAttribute(run.eastAsiaFontTheme)}"` : "",
    run.complexScriptFontTheme ? ` w:cstheme="${escapeAttribute(run.complexScriptFontTheme)}"` : "",
    run.fontHint ? ` w:hint="${run.fontHint}"` : "",
  ].join("");

  return attributes ? `<w:rFonts${attributes}/>` : "";
}

function tableStylePropertiesXml(properties: NonNullable<NonNullable<DocumentJson["styles"]>["table"]>[number]["table"]): string {
  if (!properties?.borders) {
    return "";
  }

  return `<w:tblPr>${tableBordersXml(properties.borders)}</w:tblPr>`;
}

function numberingXml(document: DocumentJson): string {
  const customAbstractNums = (document.numbering?.abstractNums ?? [])
    .map((abstractNum) => abstractNumberingXml(abstractNum))
    .join("");
  const customNums = (document.numbering?.nums ?? [])
    .map((num) => numberingInstanceXml(num))
    .join("");

  return xmlDeclaration(
    `<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
      `<w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="•"/></w:lvl><w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="◦"/></w:lvl></w:abstractNum>` +
      `<w:abstractNum w:abstractNumId="2"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl><w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%2."/></w:lvl></w:abstractNum>` +
      customAbstractNums +
      `<w:num w:numId="1"><w:abstractNumId w:val="1"/></w:num>` +
      `<w:num w:numId="2"><w:abstractNumId w:val="2"/></w:num>` +
      customNums +
      `</w:numbering>`,
  );
}

function numberingInstanceXml(num: NonNullable<DocumentJson["numbering"]>["nums"][number]): string {
  const overrides = (num.overrides ?? [])
    .map((override) => numberingLevelOverrideXml(override))
    .join("");

  return `<w:num w:numId="${num.id}"><w:abstractNumId w:val="${num.abstractId}"/>${overrides}</w:num>`;
}

function numberingLevelOverrideXml(override: NonNullable<NonNullable<DocumentJson["numbering"]>["nums"][number]["overrides"]>[number]): string {
  const start = override.start !== undefined ? `<w:startOverride w:val="${override.start}"/>` : "";
  const definition = override.definition ? numberingLevelXml(override.definition) : "";

  return `<w:lvlOverride w:ilvl="${override.level}">${start}${definition}</w:lvlOverride>`;
}

function abstractNumberingXml(abstractNum: NonNullable<DocumentJson["numbering"]>["abstractNums"][number]): string {
  const metadata = [
    abstractNum.nsid ? `<w:nsid w:val="${escapeAttribute(abstractNum.nsid)}"/>` : "",
    abstractNum.multiLevelType ? `<w:multiLevelType w:val="${abstractNum.multiLevelType}"/>` : "",
    abstractNum.templateCode ? `<w:tmpl w:val="${escapeAttribute(abstractNum.templateCode)}"/>` : "",
    abstractNum.styleLink ? `<w:styleLink w:val="${escapeAttribute(abstractNum.styleLink)}"/>` : "",
    abstractNum.numberingStyleLink ? `<w:numStyleLink w:val="${escapeAttribute(abstractNum.numberingStyleLink)}"/>` : "",
  ].join("");

  return `<w:abstractNum w:abstractNumId="${abstractNum.id}">` +
    metadata +
    abstractNum.levels.map((level) => numberingLevelXml(level)).join("") +
    `</w:abstractNum>`;
}

function numberingLevelXml(level: NonNullable<DocumentJson["numbering"]>["abstractNums"][number]["levels"][number]): string {
  const indentation = level.left !== undefined || level.hanging !== undefined
    ? `<w:pPr><w:ind${level.left !== undefined ? ` w:left="${level.left}"` : ""}${level.hanging !== undefined ? ` w:hanging="${level.hanging}"` : ""}/></w:pPr>`
    : "";
  const style = level.styleId ? `<w:pStyle w:val="${escapeAttribute(level.styleId)}"/>` : "";
  const suffix = level.suffix ? `<w:suff w:val="${level.suffix}"/>` : "";
  const restart = level.restart !== undefined ? `<w:lvlRestart w:val="${level.restart}"/>` : "";
  const legal = level.legal !== undefined ? (level.legal ? "<w:isLgl/>" : '<w:isLgl w:val="0"/>') : "";
  const alignment = level.alignment ? `<w:lvlJc w:val="${level.alignment}"/>` : "";
  const runProperties = styleRunPropertiesXml(level.run);

  return `<w:lvl w:ilvl="${level.level}">` +
    `<w:start w:val="${level.start ?? 1}"/>` +
    `<w:numFmt w:val="${level.format}"/>` +
    restart +
    style +
    legal +
    suffix +
    `<w:lvlText w:val="${escapeAttribute(level.text)}"/>` +
    alignment +
    indentation +
    runProperties +
    `</w:lvl>`;
}

function commentsXml(context: WriterContext): string {
  const comments = context.comments
    .map((comment) => `<w:comment w:id="${comment.id}" w:author="${escapeAttribute(comment.author)}"${comment.initials ? ` w:initials="${escapeAttribute(comment.initials)}"` : ""}${comment.date ? ` w:date="${escapeAttribute(comment.date)}"` : ""}><w:p><w:r><w:t>${escapeXml(comment.text)}</w:t></w:r></w:p></w:comment>`)
    .join("");

  return xmlDeclaration(
    `<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${comments}</w:comments>`,
  );
}

function headerFooterXml(root: "hdr" | "ftr", blocks: ParagraphNode[], context: WriterContext, watermark?: SectionNode["watermark"]): string {
  const watermarkXml = root === "hdr" && watermark ? watermarkParagraphXml(watermark) : "";

  return xmlDeclaration(
    `<w:${root} xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office">` +
      watermarkXml +
      blocks.map((block) => paragraphXml(block, context)).join("") +
      `</w:${root}>`,
  );
}

function watermarkParagraphXml(watermark: NonNullable<SectionNode["watermark"]>): string {
  const color = watermark.color ?? "C0C0C0";
  const opacity = watermark.opacity ?? 0.15;
  const rotation = watermark.rotation ?? 315;
  const fontFamily = watermark.fontFamily ?? "Calibri";

  return `<w:p><w:r><w:pict>` +
    `<v:shape id="Word2JsonWatermark" o:spid="_x0000_s1025" type="#_x0000_t136" ` +
    `style="position:absolute;margin-left:0;margin-top:0;width:468pt;height:468pt;rotation:${rotation};z-index:-251654144;mso-position-horizontal:center;mso-position-vertical:center;mso-wrap-edited:f;" ` +
    `fillcolor="#${escapeAttribute(color)}" stroked="f">` +
    `<v:fill opacity="${opacity}"/>` +
    `<v:textpath style="font-family:&quot;${escapeAttribute(fontFamily)}&quot;;font-size:1pt" string="${escapeAttribute(watermark.text)}"/>` +
    `<v:path textpathok="t"/>` +
    `</v:shape>` +
    `</w:pict></w:r></w:p>`;
}

function notesXml(root: "footnotes" | "endnotes", item: "footnote" | "endnote", notes: NoteEntry[], context: WriterContext): string {
  const separatorType = item === "footnote" ? "footnote" : "endnote";
  const entries = [
    `<w:${item} w:id="-1" w:type="separator"><w:p><w:r><w:separator/></w:r></w:p></w:${item}>`,
    `<w:${item} w:id="0" w:type="continuationSeparator"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:${item}>`,
    ...notes.map((note) => `<w:${item} w:id="${note.id}">${note.blocks.map((block) => paragraphXml(block, context)).join("")}</w:${item}>`),
  ].join("");

  return xmlDeclaration(
    `<w:${root} xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${entries}</w:${root}>`,
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

function escapeAttribute(value: string): string {
  return escapeXml(value);
}

function pathBasename(path: string): string {
  return path.split("/").pop() ?? path;
}

function pathDirname(path: string): string {
  const parts = path.split("/");
  parts.pop();

  return parts.join("/");
}

function paragraphStyleId(style: string): string {
  if (style === "heading1") {
    return "Heading1";
  }

  if (style === "heading2") {
    return "Heading2";
  }

  if (style === "heading3") {
    return "Heading3";
  }

  return "Normal";
}
