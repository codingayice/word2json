export type DocumentJson = {
  version: "1.0";
  theme?: DocumentTheme;
  styles?: DocumentStyles;
  sections: SectionNode[];
};

export type DocumentTheme = {
  name: string;
  fonts: ThemeFonts;
  colors: ThemeColors;
};

export type ThemeFonts = {
  major: string;
  minor: string;
};

export type ThemeColors = {
  accent1: string;
};

export type DocumentStyles = {
  paragraph?: ParagraphStyleDefinition[];
  character?: StyleDefinition[];
  table?: StyleDefinition[];
};

export type StyleDefinition = {
  id: string;
  name: string;
  basedOn?: string;
};

export type ParagraphStyleDefinition = StyleDefinition & {
  next?: string;
};

export type SectionNode = {
  breakType?: SectionBreakType;
  page?: PageSettings;
  headers?: HeaderFooterContent;
  footers?: HeaderFooterContent;
  columns?: ColumnSettings;
  blocks: DocumentBlock[];
};

export type SectionBreakType = "nextPage" | "continuous" | "evenPage" | "oddPage";

export type ColumnSettings = {
  count: number;
  space?: number;
};

export type HeaderFooterContent = {
  default?: ParagraphNode[];
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

export type DocumentBlock = ParagraphNode | TableNode | ImageNode;

export type ParagraphNode = {
  type: "paragraph";
  style?: ParagraphStyle;
  styleId?: string;
  alignment?: ParagraphAlignment;
  list?: ListSettings;
  pagination?: ParagraphPagination;
  runs: TextRun[];
};

export type ParagraphPagination = {
  keepNext?: boolean;
  keepLines?: boolean;
  pageBreakBefore?: boolean;
};

export type ListSettings = {
  type: "bullet" | "ordered";
  level: number;
};

export type ParagraphStyle = "normal" | "heading1" | "heading2" | "heading3";

export type ParagraphAlignment = "left" | "center" | "right" | "both";

export type TextRun = {
  text: string;
  styleId?: string;
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
  field?: FieldKind | ReferenceField;
  footnote?: NoteContent;
  endnote?: NoteContent;
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

export type FieldKind = "page" | "numPages";

export type ReferenceField = {
  type: "ref" | "pageRef";
  target: string;
};

export type NoteContent = {
  blocks: ParagraphNode[];
};

export type TableNode = {
  type: "table";
  styleId?: string;
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

export type ImageNode = {
  type: "image";
  data: string;
  contentType: "image/png" | "image/jpeg";
  width: number;
  height: number;
  altText?: string;
};

export function createDocumentJson(blocks: DocumentBlock[]): DocumentJson {
  return {
    version: "1.0",
    sections: [{ blocks }],
  };
}
