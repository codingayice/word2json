export type DocumentJson = {
  version: "1.0";
  sections: SectionNode[];
};

export type SectionNode = {
  page?: PageSettings;
  blocks: DocumentBlock[];
};

export type PageSettings = {
  width: number;
  height: number;
  orientation?: "portrait" | "landscape";
  margins: PageMargins;
};

export type PageMargins = {
  top: number;
  right: number;
  bottom: number;
  left: number;
  header: number;
  footer: number;
  gutter: number;
};

export type DocumentBlock = ParagraphNode | TableNode;

export type ParagraphNode = {
  type: "paragraph";
  style?: ParagraphStyle;
  alignment?: ParagraphAlignment;
  list?: ListSettings;
  runs: TextRun[];
};

export type ListSettings = {
  type: "bullet" | "ordered";
  level: number;
};

export type ParagraphStyle = "normal" | "heading1" | "heading2" | "heading3";

export type ParagraphAlignment = "left" | "center" | "right" | "both";

export type TextRun = {
  text: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  fontFamily?: string;
  fontSize?: number;
  color?: string;
  link?: Hyperlink;
  comment?: Comment;
  bookmark?: Bookmark;
  break?: BreakKind;
};

export type Hyperlink = {
  url: string;
};

export type Comment = {
  author: string;
  initials?: string;
  date?: string;
  text: string;
};

export type Bookmark = {
  name: string;
};

export type BreakKind = "line" | "page";

export type TableNode = {
  type: "table";
  width?: number;
  borders?: "single";
  rows: TableRowNode[];
};

export type TableRowNode = {
  cells: TableCellNode[];
};

export type TableCellNode = {
  width?: number;
  colSpan?: number;
  blocks: ParagraphNode[];
};

export function createDocumentJson(blocks: DocumentBlock[]): DocumentJson {
  return {
    version: "1.0",
    sections: [{ blocks }],
  };
}
