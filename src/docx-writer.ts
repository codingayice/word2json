import JSZip from "jszip";
import type {
  BorderDefinition,
  DocumentBlock,
  DocumentJson,
  ImageNode,
  PageSettings,
  ParagraphNode,
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
};

type NoteEntry = {
  id: number;
  blocks: ParagraphNode[];
};

export async function buildDocx(document: DocumentJson): Promise<Buffer> {
  const zip = new JSZip();
  const context: WriterContext = { hyperlinks: [], comments: [], images: [], headers: [], footers: [], footnotes: [], endnotes: [], theme: document.theme, settings: document.settings, bookmarkId: 0 };

  zip.folder("_rels")!.file(".rels", packageRelsXml());
  zip.folder("word")!.file("document.xml", documentXml(document, context));
  for (const header of context.headers) {
    zip.folder("word")!.file(header.filename, headerFooterXml("hdr", header.blocks, context));
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
  return xmlDeclaration(
    `<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
      (settings.defaultTabStop !== undefined ? `<w:defaultTabStop w:val="${settings.defaultTabStop}"/>` : "") +
      (settings.evenAndOddHeaders ? "<w:evenAndOddHeaders/>" : "") +
      (settings.updateFields ? "<w:updateFields/>" : "") +
      `</w:settings>`,
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
    `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
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
  const headerReference = section.headers?.default
    ? createHeaderReference(section.headers.default, context)
    : "";
  const footerReference = section.footers?.default
    ? createFooterReference(section.footers.default, context)
    : "";
  const breakType = section.breakType
    ? `<w:type w:val="${sectionBreakValue(section.breakType)}"/>`
    : "";
  const columns = section.columns
    ? `<w:cols w:num="${section.columns.count}"${section.columns.space ? ` w:space="${section.columns.space}"` : ""}/>`
    : "";

  return `<w:sectPr>` +
    headerReference +
    footerReference +
    breakType +
    `<w:pgSz w:w="${page.width}" w:h="${page.height}"${orientation}/>` +
    `<w:pgMar w:top="${page.margins.top}" w:right="${page.margins.right}" w:bottom="${page.margins.bottom}" w:left="${page.margins.left}" w:header="${page.margins.header}" w:footer="${page.margins.footer}" w:gutter="${page.margins.gutter}"/>` +
    columns +
    `</w:sectPr>`;
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

function createHeaderReference(blocks: ParagraphNode[], context: WriterContext): string {
  const index = context.headers.length + 1;
  const id = `rIdHeader${index}`;
  context.headers.push({ id, filename: `header${index}.xml`, blocks });
  return `<w:headerReference w:type="default" r:id="${id}"/>`;
}

function createFooterReference(blocks: ParagraphNode[], context: WriterContext): string {
  const index = context.footers.length + 1;
  const id = `rIdFooter${index}`;
  context.footers.push({ id, filename: `footer${index}.xml`, blocks });
  return `<w:footerReference w:type="default" r:id="${id}"/>`;
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
  const runs = paragraph.runs.map((run) => runXml(run, context)).join("");

  return `<w:p>${properties}${runs}</w:p>`;
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
  const list = paragraph.list
    ? `<w:numPr><w:ilvl w:val="${paragraph.list.level}"/><w:numId w:val="${paragraph.list.numberingId ?? (paragraph.list.type === "bullet" ? 1 : 2)}"/></w:numPr>`
    : "";
  const pagination = paragraph.pagination
    ? [
      paragraph.pagination.keepNext ? "<w:keepNext/>" : "",
      paragraph.pagination.keepLines ? "<w:keepLines/>" : "",
      paragraph.pagination.pageBreakBefore ? "<w:pageBreakBefore/>" : "",
    ].join("")
    : "";
  const properties = `${style}${alignment}${spacing}${indent}${shading}${borders}${list}${pagination}`;

  return properties ? `<w:pPr>${properties}</w:pPr>` : "";
}

function shadingXml(shading: { fill: string }): string {
  return `<w:shd w:fill="${escapeAttribute(shading.fill)}"/>`;
}

function paragraphBordersXml(borders: NonNullable<ParagraphNode["borders"]>): string {
  const sides = [
    borders.top ? borderSideXml("top", borders.top) : "",
    borders.left ? borderSideXml("left", borders.left) : "",
    borders.bottom ? borderSideXml("bottom", borders.bottom) : "",
    borders.right ? borderSideXml("right", borders.right) : "",
  ].join("");

  return sides ? `<w:pBdr>${sides}</w:pBdr>` : "";
}

function borderSideXml(side: "top" | "left" | "bottom" | "right" | "bdr", border: BorderDefinition): string {
  return `<w:${side} w:val="${border.style}"` +
    (border.size !== undefined ? ` w:sz="${border.size}"` : "") +
    (border.space !== undefined ? ` w:space="${border.space}"` : "") +
    (border.color ? ` w:color="${escapeAttribute(border.color)}"` : "") +
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

function runXml(run: TextRun, context: WriterContext): string {
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
  const bookmarkedRun = wrapBookmarkIfNeeded(run, plainRun, context);
  const runContent = wrapCommentIfNeeded(run, bookmarkedRun, context);

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

  const id = context.comments.length;
  context.comments.push({ id, ...run.comment });

  return `<w:commentRangeStart w:id="${id}"/>` +
    runContent +
    `<w:commentRangeEnd w:id="${id}"/>` +
    `<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="${id}"/></w:r>`;
}

function runPropertiesXml(run: TextRun): string {
  const properties = [
    run.styleId ? `<w:rStyle w:val="${escapeAttribute(run.styleId)}"/>` : "",
    run.bold ? "<w:b/>" : "",
    run.italic ? "<w:i/>" : "",
    run.underline ? '<w:u w:val="single"/>' : "",
    run.fontFamily ? `<w:rFonts w:ascii="${escapeAttribute(run.fontFamily)}" w:hAnsi="${escapeAttribute(run.fontFamily)}"/>` : "",
    run.fontSize ? `<w:sz w:val="${run.fontSize * 2}"/>` : "",
    run.color ? `<w:color w:val="${escapeAttribute(run.color)}"/>` : "",
    run.highlight ? `<w:highlight w:val="${run.highlight}"/>` : "",
    run.strike ? "<w:strike/>" : "",
    run.doubleStrike ? "<w:dstrike/>" : "",
    run.smallCaps ? "<w:smallCaps/>" : "",
    run.allCaps ? "<w:caps/>" : "",
    run.verticalAlign ? `<w:vertAlign w:val="${run.verticalAlign}"/>` : "",
    run.characterSpacing !== undefined ? `<w:spacing w:val="${run.characterSpacing}"/>` : "",
    run.scale !== undefined ? `<w:w w:val="${run.scale}"/>` : "",
    run.border ? borderSideXml("bdr", run.border) : "",
  ].join("");

  return properties ? `<w:rPr>${properties}</w:rPr>` : "";
}

function tableXml(table: TableNode, context: WriterContext): string {
  const properties = [
    table.styleId ? `<w:tblStyle w:val="${escapeAttribute(table.styleId)}"/>` : "",
    table.width ? `<w:tblW w:w="${table.width}" w:type="dxa"/>` : "",
    table.borders ? tableBordersXml(table.borders) : "",
    table.alignment ? `<w:jc w:val="${table.alignment}"/>` : "",
    table.cellSpacing !== undefined ? `<w:tblCellSpacing w:w="${table.cellSpacing}" w:type="dxa"/>` : "",
  ].join("");
  const rows = table.rows
    .map((row) => {
      const rowProperties = row.height
        ? `<w:trPr><w:trHeight w:val="${row.height.value}"${row.height.rule ? ` w:hRule="${row.height.rule}"` : ""}/></w:trPr>`
        : "";

      return `<w:tr>${rowProperties}${row.cells.map((cell) => tableCellXml(cell, context)).join("")}</w:tr>`;
    })
    .join("");
  const grid = table.grid
    ? `<w:tblGrid>${table.grid.map((width) => `<w:gridCol w:w="${width}"/>`).join("")}</w:tblGrid>`
    : "";

  return `<w:tbl>${properties ? `<w:tblPr>${properties}</w:tblPr>` : ""}${grid}${rows}</w:tbl>`;
}

function tableCellXml(cell: TableCellNode, context: WriterContext): string {
  const properties = [
    cell.width ? `<w:tcW w:w="${cell.width}" w:type="dxa"/>` : "",
    cell.colSpan ? `<w:gridSpan w:val="${cell.colSpan}"/>` : "",
    cell.verticalMerge ? `<w:vMerge w:val="${cell.verticalMerge}"/>` : "",
    cell.verticalAlignment ? `<w:vAlign w:val="${cell.verticalAlignment}"/>` : "",
    cell.shading ? `<w:shd w:fill="${escapeAttribute(cell.shading.fill)}"/>` : "",
    cell.borders ? tableCellBordersXml(cell.borders) : "",
    cell.textDirection ? `<w:textDirection w:val="${cell.textDirection}"/>` : "",
    cell.margins ? tableCellMarginsXml(cell.margins) : "",
  ].join("");
  const blocks = cell.blocks.map((block) => paragraphXml(block, context)).join("");

  return `<w:tc>${properties ? `<w:tcPr>${properties}</w:tcPr>` : ""}${blocks}</w:tc>`;
}

function tableCellBordersXml(borders: NonNullable<TableCellNode["borders"]>): string {
  const sides = [
    borders.top ? borderSideXml("top", borders.top) : "",
    borders.left ? borderSideXml("left", borders.left) : "",
    borders.bottom ? borderSideXml("bottom", borders.bottom) : "",
    borders.right ? borderSideXml("right", borders.right) : "",
  ].join("");

  return sides ? `<w:tcBorders>${sides}</w:tcBorders>` : "";
}

function tableCellMarginsXml(margins: NonNullable<TableCellNode["margins"]>): string {
  const sides = [
    margins.top !== undefined ? `<w:top w:w="${margins.top}" w:type="dxa"/>` : "",
    margins.right !== undefined ? `<w:right w:w="${margins.right}" w:type="dxa"/>` : "",
    margins.bottom !== undefined ? `<w:bottom w:w="${margins.bottom}" w:type="dxa"/>` : "",
    margins.left !== undefined ? `<w:left w:w="${margins.left}" w:type="dxa"/>` : "",
  ].join("");

  return sides ? `<w:tcMar>${sides}</w:tcMar>` : "";
}

function tableBordersXml(border: "single"): string {
  return `<w:tblBorders>` +
    `<w:top w:val="${border}" w:sz="4" w:space="0" w:color="auto"/>` +
    `<w:left w:val="${border}" w:sz="4" w:space="0" w:color="auto"/>` +
    `<w:bottom w:val="${border}" w:sz="4" w:space="0" w:color="auto"/>` +
    `<w:right w:val="${border}" w:sz="4" w:space="0" w:color="auto"/>` +
    `<w:insideH w:val="${border}" w:sz="4" w:space="0" w:color="auto"/>` +
    `<w:insideV w:val="${border}" w:sz="4" w:space="0" w:color="auto"/>` +
    `</w:tblBorders>`;
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
    `<pic:blipFill><a:blip r:embed="${relationshipId}"/>${crop}</pic:blipFill>` +
    `<pic:spPr><a:xfrm${rotation}><a:ext cx="${widthEmu}" cy="${heightEmu}"/></a:xfrm></pic:spPr>` +
    `</pic:pic></a:graphicData></a:graphic>`;
  const drawing = image.floating
    ? `<wp:anchor xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" simplePos="0" relativeHeight="0" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">` +
      `<wp:positionH relativeFrom="page"><wp:posOffset>${image.floating.horizontalOffset}</wp:posOffset></wp:positionH>` +
      `<wp:positionV relativeFrom="page"><wp:posOffset>${image.floating.verticalOffset}</wp:posOffset></wp:positionV>` +
      `<wp:wrapSquare/>` +
      graphic +
      `</wp:anchor>`
    : `<wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">${graphic}</wp:inline>`;

  return `<w:p><w:r><w:drawing>` +
    drawing +
    `</w:drawing></w:r></w:p>`;
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
      headerRelationships +
      footerRelationships +
      imageRelationships +
      hyperlinkRelationships +
      `</Relationships>`,
  );
}

function themeXml(theme: NonNullable<DocumentJson["theme"]>): string {
  return xmlDeclaration(
    `<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="${escapeAttribute(theme.name)}">` +
      `<a:themeElements>` +
      `<a:clrScheme name="${escapeAttribute(theme.name)}">` +
      `<a:accent1><a:srgbClr val="${escapeAttribute(theme.colors.accent1)}"/></a:accent1>` +
      `</a:clrScheme>` +
      `<a:fontScheme name="${escapeAttribute(theme.name)}">` +
      `<a:majorFont><a:latin typeface="${escapeAttribute(theme.fonts.major)}"/></a:majorFont>` +
      `<a:minorFont><a:latin typeface="${escapeAttribute(theme.fonts.minor)}"/></a:minorFont>` +
      `</a:fontScheme>` +
      `</a:themeElements>` +
      `</a:theme>`,
  );
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
  ].join("");

  return paragraphProperties ? `<w:pPr>${paragraphProperties}</w:pPr>` : "";
}

function styleRunPropertiesXml(run?: StyleRunProperties): string {
  if (!run) {
    return "";
  }

  const properties = [
    run.bold ? "<w:b/>" : "",
    run.italic ? "<w:i/>" : "",
    run.underline ? '<w:u w:val="single"/>' : "",
    run.fontFamily ? `<w:rFonts w:ascii="${escapeAttribute(run.fontFamily)}" w:hAnsi="${escapeAttribute(run.fontFamily)}"/>` : "",
    run.fontSize ? `<w:sz w:val="${run.fontSize * 2}"/>` : "",
    run.color ? `<w:color w:val="${escapeAttribute(run.color)}"/>` : "",
    run.highlight ? `<w:highlight w:val="${run.highlight}"/>` : "",
    run.strike ? "<w:strike/>" : "",
    run.doubleStrike ? "<w:dstrike/>" : "",
    run.smallCaps ? "<w:smallCaps/>" : "",
    run.allCaps ? "<w:caps/>" : "",
    run.verticalAlign ? `<w:vertAlign w:val="${run.verticalAlign}"/>` : "",
    run.characterSpacing !== undefined ? `<w:spacing w:val="${run.characterSpacing}"/>` : "",
    run.scale !== undefined ? `<w:w w:val="${run.scale}"/>` : "",
    run.border ? borderSideXml("bdr", run.border) : "",
  ].join("");

  return properties ? `<w:rPr>${properties}</w:rPr>` : "";
}

function tableStylePropertiesXml(properties: NonNullable<NonNullable<DocumentJson["styles"]>["table"]>[number]["table"]): string {
  if (!properties?.borders) {
    return "";
  }

  return `<w:tblPr>${tableBordersXml(properties.borders)}</w:tblPr>`;
}

function numberingXml(document: DocumentJson): string {
  const customAbstractNums = (document.numbering?.abstractNums ?? [])
    .map((abstractNum) => `<w:abstractNum w:abstractNumId="${abstractNum.id}">` +
      abstractNum.levels.map((level) => numberingLevelXml(level)).join("") +
      `</w:abstractNum>`)
    .join("");
  const customNums = (document.numbering?.nums ?? [])
    .map((num) => `<w:num w:numId="${num.id}"><w:abstractNumId w:val="${num.abstractId}"/></w:num>`)
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

function numberingLevelXml(level: NonNullable<DocumentJson["numbering"]>["abstractNums"][number]["levels"][number]): string {
  const indentation = level.left !== undefined || level.hanging !== undefined
    ? `<w:pPr><w:ind${level.left !== undefined ? ` w:left="${level.left}"` : ""}${level.hanging !== undefined ? ` w:hanging="${level.hanging}"` : ""}/></w:pPr>`
    : "";

  return `<w:lvl w:ilvl="${level.level}">` +
    `<w:start w:val="${level.start ?? 1}"/>` +
    `<w:numFmt w:val="${level.format}"/>` +
    `<w:lvlText w:val="${escapeAttribute(level.text)}"/>` +
    indentation +
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

function headerFooterXml(root: "hdr" | "ftr", blocks: ParagraphNode[], context: WriterContext): string {
  return xmlDeclaration(
    `<w:${root} xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
      blocks.map((block) => paragraphXml(block, context)).join("") +
      `</w:${root}>`,
  );
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
