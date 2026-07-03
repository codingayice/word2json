export type DocumentJson = {
  version: "1.0";
  settings?: DocumentSettings;
  theme?: DocumentTheme;
  styles?: DocumentStyles;
  numbering?: DocumentNumbering;
  sections: SectionNode[];
};

export type DocumentSettings = {
  defaultTabStop?: number;
  evenAndOddHeaders?: boolean;
  updateFields?: boolean;
  trackRevisions?: boolean;
};

export type DocumentNumbering = {
  abstractNums: AbstractNumberingDefinition[];
  nums: NumberingInstance[];
};

export type AbstractNumberingDefinition = {
  id: number;
  levels: NumberingLevelDefinition[];
};

export type NumberingInstance = {
  id: number;
  abstractId: number;
};

export type NumberingLevelDefinition = {
  level: number;
  format: NumberingFormat;
  text: string;
  start?: number;
  left?: number;
  hanging?: number;
};

export type NumberingFormat = "bullet" | "decimal" | "lowerLetter" | "upperLetter" | "lowerRoman" | "upperRoman";

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
  defaults?: DocumentStyleDefaults;
  paragraph?: ParagraphStyleDefinition[];
  character?: StyleDefinition[];
  table?: TableStyleDefinition[];
};

export type DocumentStyleDefaults = {
  run?: StyleRunProperties;
  paragraph?: StyleParagraphProperties;
};

export type StyleDefinition = {
  id: string;
  name: string;
  basedOn?: string;
  run?: StyleRunProperties;
};

export type ParagraphStyleDefinition = StyleDefinition & {
  next?: string;
  paragraph?: StyleParagraphProperties;
};

export type StyleRunProperties = {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  fontFamily?: string;
  fontSize?: number;
  color?: string;
  highlight?: HighlightColor;
  strike?: boolean;
  doubleStrike?: boolean;
  smallCaps?: boolean;
  allCaps?: boolean;
  verticalAlign?: RunVerticalAlign;
  characterSpacing?: number;
  scale?: number;
  border?: BorderDefinition;
};

export type HighlightColor = "yellow" | "green" | "cyan" | "magenta" | "blue" | "red" | "darkBlue" | "darkCyan" | "darkGreen" | "darkMagenta" | "darkRed" | "darkYellow" | "darkGray" | "lightGray" | "black";

export type RunVerticalAlign = "superscript" | "subscript" | "baseline";

export type StyleParagraphProperties = {
  alignment?: ParagraphAlignment;
  spacing?: ParagraphSpacing;
  indent?: ParagraphIndent;
  shading?: ShadingDefinition;
  borders?: ParagraphBorders;
};

export type ShadingDefinition = {
  fill: string;
};

export type BorderDefinition = {
  style: "single";
  size?: number;
  color?: string;
  space?: number;
};

export type ParagraphBorders = {
  top?: BorderDefinition;
  left?: BorderDefinition;
  bottom?: BorderDefinition;
  right?: BorderDefinition;
};

export type TableStyleDefinition = StyleDefinition & {
  table?: StyleTableProperties;
};

export type StyleTableProperties = {
  borders?: "single";
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
  revision?: RunRevision;
  propertyRevision?: ParagraphPropertyRevision;
  alignment?: ParagraphAlignment;
  spacing?: ParagraphSpacing;
  indent?: ParagraphIndent;
  shading?: ShadingDefinition;
  borders?: ParagraphBorders;
  list?: ListSettings;
  pagination?: ParagraphPagination;
  runs: TextRun[];
};

export type ParagraphSpacing = {
  before?: number;
  after?: number;
  line?: number;
  lineRule?: "auto" | "exact" | "atLeast";
};

export type ParagraphIndent = {
  left?: number;
  right?: number;
  firstLine?: number;
  hanging?: number;
};

export type ParagraphPagination = {
  keepNext?: boolean;
  keepLines?: boolean;
  pageBreakBefore?: boolean;
};

export type ListSettings = {
  type: "bullet" | "ordered";
  level: number;
  numberingId?: number;
};

export type ParagraphStyle = "normal" | "heading1" | "heading2" | "heading3";

export type ParagraphAlignment = "left" | "center" | "right" | "both";

export type TextRun = StyleRunProperties & {
  text: string;
  styleId?: string;
  revision?: RunRevision;
  link?: Hyperlink;
  comment?: Comment;
  bookmark?: Bookmark;
  break?: BreakKind;
  field?: FieldKind | FieldWithResult | ReferenceField | TocField;
  footnote?: NoteContent;
  endnote?: NoteContent;
};

export type RunRevision = {
  type: "insert" | "delete" | "moveFrom" | "moveTo";
  id: number;
  author: string;
  date?: string;
};

export type ParagraphPropertyRevision = {
  id: number;
  author: string;
  date?: string;
};

export type Hyperlink = {
  url: string;
} | {
  anchor: string;
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

export type FieldWithResult = {
  type: FieldKind;
  result?: string;
};

export type ReferenceField = {
  type: "ref" | "pageRef";
  target: string;
  result?: string;
};

export type TocField = {
  type: "toc";
  switches?: string;
  result?: string;
};

export type NoteContent = {
  blocks: ParagraphNode[];
};

export type TableNode = {
  type: "table";
  styleId?: string;
  grid?: number[];
  width?: number;
  borders?: "single";
  alignment?: ParagraphAlignment;
  cellSpacing?: number;
  propertyRevision?: ParagraphPropertyRevision;
  rows: TableRowNode[];
};

export type TableRowNode = {
  revision?: RunRevision;
  height?: TableRowHeight;
  cells: TableCellNode[];
};

export type TableRowHeight = {
  value: number;
  rule?: "auto" | "atLeast" | "exact";
};

export type TableCellNode = {
  width?: number;
  colSpan?: number;
  verticalMerge?: "restart" | "continue";
  verticalAlignment?: "top" | "center" | "bottom";
  shading?: TableCellShading;
  borders?: ParagraphBorders;
  textDirection?: "lrTb" | "tbRl" | "btLr";
  margins?: TableCellMargins;
  propertyRevision?: ParagraphPropertyRevision;
  blocks: ParagraphNode[];
};

export type TableCellShading = {
  fill: string;
};

export type TableCellMargins = {
  top?: number;
  right?: number;
  bottom?: number;
  left?: number;
};

export type ImageNode = {
  type: "image";
  data: string;
  contentType: "image/png" | "image/jpeg";
  width: number;
  height: number;
  altText?: string;
  crop?: ImageCrop;
  rotation?: number;
  floating?: ImageFloatingLayout;
};

export type ImageCrop = {
  left?: number;
  top?: number;
  right?: number;
  bottom?: number;
};

export type ImageFloatingLayout = {
  wrap: "square";
  horizontalOffset: number;
  verticalOffset: number;
};

export function createDocumentJson(blocks: DocumentBlock[]): DocumentJson {
  return {
    version: "1.0",
    sections: [{ blocks }],
  };
}
