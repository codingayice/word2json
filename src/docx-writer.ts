import JSZip from "jszip";
import type {
  DocumentBlock,
  DocumentJson,
  PageSettings,
  ParagraphNode,
  SectionNode,
  TableCellNode,
  TableNode,
  TextRun,
} from "./schema.js";

type WriterContext = {
  hyperlinks: HyperlinkRelationship[];
  comments: CommentEntry[];
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

export async function buildDocx(document: DocumentJson): Promise<Buffer> {
  const zip = new JSZip();
  const context: WriterContext = { hyperlinks: [], comments: [], bookmarkId: 0 };

  zip.folder("_rels")!.file(".rels", packageRelsXml());
  zip.folder("word")!.file("document.xml", documentXml(document, context));
  zip.folder("word")!.file("styles.xml", stylesXml());
  zip.folder("word")!.file("numbering.xml", numberingXml());
  if (context.comments.length > 0) {
    zip.folder("word")!.file("comments.xml", commentsXml(context));
  }
  zip.file("[Content_Types].xml", contentTypesXml(context));
  zip.folder("word")!.folder("_rels")!.file("document.xml.rels", documentRelsXml(context));

  return zip.generateAsync({ type: "nodebuffer" });
}

function documentXml(document: DocumentJson, context: WriterContext): string {
  const section = document.sections[0] ?? { blocks: [] };
  const body = document.sections
    .flatMap((currentSection) => currentSection.blocks)
    .map((block) => blockXml(block, context))
    .join("");

  return xmlDeclaration(
    `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
      `<w:body>${body}${sectionPropertiesXml(section)}</w:body>` +
      `</w:document>`,
  );
}

function blockXml(block: DocumentBlock, context: WriterContext): string {
  if (block.type === "paragraph") {
    return paragraphXml(block, context);
  }

  return tableXml(block, context);
}

function sectionPropertiesXml(section: SectionNode): string {
  const page = section.page ?? defaultPageSettings();
  const orientation = page.orientation && page.orientation !== "portrait"
    ? ` w:orient="${page.orientation}"`
    : "";

  return `<w:sectPr>` +
    `<w:pgSz w:w="${page.width}" w:h="${page.height}"${orientation}/>` +
    `<w:pgMar w:top="${page.margins.top}" w:right="${page.margins.right}" w:bottom="${page.margins.bottom}" w:left="${page.margins.left}" w:header="${page.margins.header}" w:footer="${page.margins.footer}" w:gutter="${page.margins.gutter}"/>` +
    `</w:sectPr>`;
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
  const style = paragraph.style && paragraph.style !== "normal"
    ? `<w:pStyle w:val="${paragraphStyleId(paragraph.style)}"/>`
    : "";
  const alignment = paragraph.alignment
    ? `<w:jc w:val="${paragraph.alignment}"/>`
    : "";
  const list = paragraph.list
    ? `<w:numPr><w:ilvl w:val="${paragraph.list.level}"/><w:numId w:val="${paragraph.list.type === "bullet" ? 1 : 2}"/></w:numPr>`
    : "";
  const properties = `${style}${alignment}${list}`;

  return properties ? `<w:pPr>${properties}</w:pPr>` : "";
}

function runXml(run: TextRun, context: WriterContext): string {
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

  const relationshipId = `rIdHyperlink${context.hyperlinks.length + 1}`;
  context.hyperlinks.push({ id: relationshipId, url: run.link.url });
  return `<w:hyperlink r:id="${relationshipId}">${runContent}</w:hyperlink>`;
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
    run.bold ? "<w:b/>" : "",
    run.italic ? "<w:i/>" : "",
    run.underline ? '<w:u w:val="single"/>' : "",
    run.fontFamily ? `<w:rFonts w:ascii="${escapeAttribute(run.fontFamily)}" w:hAnsi="${escapeAttribute(run.fontFamily)}"/>` : "",
    run.fontSize ? `<w:sz w:val="${run.fontSize * 2}"/>` : "",
    run.color ? `<w:color w:val="${escapeAttribute(run.color)}"/>` : "",
  ].join("");

  return properties ? `<w:rPr>${properties}</w:rPr>` : "";
}

function tableXml(table: TableNode, context: WriterContext): string {
  const properties = [
    table.width ? `<w:tblW w:w="${table.width}" w:type="dxa"/>` : "",
    table.borders ? tableBordersXml(table.borders) : "",
  ].join("");
  const rows = table.rows
    .map((row) => `<w:tr>${row.cells.map((cell) => tableCellXml(cell, context)).join("")}</w:tr>`)
    .join("");

  return `<w:tbl>${properties ? `<w:tblPr>${properties}</w:tblPr>` : ""}${rows}</w:tbl>`;
}

function tableCellXml(cell: TableCellNode, context: WriterContext): string {
  const properties = [
    cell.width ? `<w:tcW w:w="${cell.width}" w:type="dxa"/>` : "",
    cell.colSpan ? `<w:gridSpan w:val="${cell.colSpan}"/>` : "",
  ].join("");
  const blocks = cell.blocks.map((block) => paragraphXml(block, context)).join("");

  return `<w:tc>${properties ? `<w:tcPr>${properties}</w:tcPr>` : ""}${blocks}</w:tc>`;
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

function contentTypesXml(context: WriterContext): string {
  const commentsOverride = context.comments.length > 0
    ? `<Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>`
    : "";

  return xmlDeclaration(
    `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
      `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
      `<Default Extension="xml" ContentType="application/xml"/>` +
      `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>` +
      `<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>` +
      `<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>` +
      commentsOverride +
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

  return xmlDeclaration(
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      `<Relationship Id="rIdNumbering" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>` +
      (context.comments.length > 0 ? `<Relationship Id="rIdComments" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/>` : "") +
      hyperlinkRelationships +
      `</Relationships>`,
  );
}

function stylesXml(): string {
  return xmlDeclaration(
    `<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
      `<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>` +
      `<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:qFormat/></w:style>` +
      `<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:qFormat/></w:style>` +
      `<w:style w:type="paragraph" w:styleId="Heading3"><w:name w:val="heading 3"/><w:basedOn w:val="Normal"/><w:qFormat/></w:style>` +
      `</w:styles>`,
  );
}

function numberingXml(): string {
  return xmlDeclaration(
    `<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
      `<w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="•"/></w:lvl><w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="◦"/></w:lvl></w:abstractNum>` +
      `<w:abstractNum w:abstractNumId="2"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl><w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%2."/></w:lvl></w:abstractNum>` +
      `<w:num w:numId="1"><w:abstractNumId w:val="1"/></w:num>` +
      `<w:num w:numId="2"><w:abstractNumId w:val="2"/></w:num>` +
      `</w:numbering>`,
  );
}

function commentsXml(context: WriterContext): string {
  const comments = context.comments
    .map((comment) => `<w:comment w:id="${comment.id}" w:author="${escapeAttribute(comment.author)}"${comment.initials ? ` w:initials="${escapeAttribute(comment.initials)}"` : ""}${comment.date ? ` w:date="${escapeAttribute(comment.date)}"` : ""}><w:p><w:r><w:t>${escapeXml(comment.text)}</w:t></w:r></w:p></w:comment>`)
    .join("");

  return xmlDeclaration(
    `<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${comments}</w:comments>`,
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
