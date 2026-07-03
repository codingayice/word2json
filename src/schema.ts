export type DocumentJson = {
  version: "1.0";
  settings?: DocumentSettings;
  properties?: DocumentProperties;
  theme?: DocumentTheme;
  styles?: DocumentStyles;
  numbering?: DocumentNumbering;
  customXmlParts?: CustomXmlPart[];
  fonts?: DocumentFont[];
  sections: SectionNode[];
};

export type DocumentFont = {
  name: string;
  family?: "decorative" | "modern" | "roman" | "script" | "swiss" | "system";
  pitch?: "fixed" | "variable" | "default";
  charset?: string;
  panose1?: string;
};

export type DocumentProperties = {
  core?: DocumentCoreProperties;
  app?: DocumentAppProperties;
  custom?: CustomDocumentProperty[];
};

export type DocumentCoreProperties = {
  title?: string;
  subject?: string;
  creator?: string;
  keywords?: string;
  description?: string;
  lastModifiedBy?: string;
  created?: string;
  modified?: string;
};

export type DocumentAppProperties = {
  application?: string;
  company?: string;
  manager?: string;
  pages?: number;
  words?: number;
  characters?: number;
  lines?: number;
  paragraphs?: number;
};

export type CustomDocumentProperty = {
  name: string;
  type: "string" | "number" | "boolean" | "date";
  value: string | number | boolean;
};

export type CustomXmlPart = {
  path: string;
  xml: string;
  properties?: CustomXmlPartProperties;
};

export type CustomXmlPartProperties = {
  path?: string;
  storeItemId?: string;
  schemaRefs?: string[];
};

export type DocumentSettings = {
  defaultTabStop?: number;
  evenAndOddHeaders?: boolean;
  updateFields?: boolean;
  trackRevisions?: boolean;
  compatibility?: DocumentCompatibilitySettings;
  proofing?: DocumentProofingSettings;
  protection?: DocumentProtectionSettings;
  mailMerge?: DocumentMailMergeSettings;
  writeProtection?: DocumentWriteProtectionSettings;
  math?: DocumentMathSettings;
  view?: DocumentViewSettings;
  web?: DocumentWebSettings;
};

export type DocumentMathSettings = {
  mathFont?: string;
  breakBinary?: "before" | "after" | "repeat";
  smallFraction?: boolean;
  displayDefaults?: boolean;
};

export type DocumentWriteProtectionSettings = {
  recommended?: boolean;
  cryptProviderType?: string;
  cryptAlgorithmClass?: string;
  cryptAlgorithmType?: string;
  cryptAlgorithmSid?: number;
  cryptSpinCount?: number;
  hash?: string;
  salt?: string;
};

export type DocumentMailMergeSettings = {
  mainDocumentType?: string;
  dataType?: string;
  connectString?: string;
  query?: string;
  viewMergedData?: boolean;
  activeRecord?: number;
  checkErrors?: number;
};

export type DocumentProtectionSettings = {
  edit?: "none" | "readOnly" | "comments" | "trackedChanges" | "forms";
  enforcement?: boolean;
  cryptProviderType?: string;
  cryptAlgorithmClass?: string;
  cryptAlgorithmType?: string;
  cryptAlgorithmSid?: number;
  cryptSpinCount?: number;
  hash?: string;
  salt?: string;
};

export type DocumentViewSettings = {
  mode?: "none" | "print" | "outline" | "masterPages" | "normal" | "web";
  zoom?: {
    preset?: "none" | "fullPage" | "bestFit" | "textFit";
    percent?: number;
  };
};

export type DocumentProofingSettings = {
  spelling?: "clean" | "dirty";
  grammar?: "clean" | "dirty";
  doNotHyphenateCaps?: boolean;
  hyphenationZone?: number;
};

export type DocumentCompatibilitySettings = {
  compatMode?: string;
  settings?: DocumentCompatSetting[];
};

export type DocumentCompatSetting = {
  name: string;
  uri: string;
  value: string;
};

export type DocumentWebSettings = {
  optimizeForBrowser?: boolean;
  allowPng?: boolean;
  doNotSaveAsSingleFile?: boolean;
  pixelsPerInch?: number;
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
  formatScheme?: ThemeFormatScheme;
};

export type ThemeFonts = {
  major: string;
  minor: string;
  majorEastAsia?: string;
  majorComplexScript?: string;
  minorEastAsia?: string;
  minorComplexScript?: string;
  supplemental?: ThemeSupplementalFont[];
};

export type ThemeSupplementalFont = {
  group: "major" | "minor";
  script: string;
  typeface: string;
};

export type ThemeColors = {
  dark1?: string;
  light1?: string;
  dark2?: string;
  light2?: string;
  accent1: string;
  accent2?: string;
  accent3?: string;
  accent4?: string;
  accent5?: string;
  accent6?: string;
  hyperlink?: string;
  followedHyperlink?: string;
};

export type ThemeFormatScheme = {
  name: string;
  fillStyleColors?: string[];
  lineStyleColors?: string[];
  effectStyleColors?: string[];
  backgroundFillStyleColors?: string[];
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
  contentControl?: ContentControl;
  alignment?: ParagraphAlignment;
  spacing?: ParagraphSpacing;
  indent?: ParagraphIndent;
  shading?: ShadingDefinition;
  borders?: ParagraphBorders;
  list?: ListSettings;
  pagination?: ParagraphPagination;
  commentRangeStart?: Comment;
  commentRangeEnd?: CommentRangeEnd;
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
  contentControl?: ContentControl;
  revision?: RunRevision;
  link?: Hyperlink;
  comment?: Comment;
  bookmark?: Bookmark;
  break?: BreakKind;
  field?: FieldKind | FieldWithResult | ReferenceField | TocField;
  footnote?: NoteContent;
  endnote?: NoteContent;
  math?: MathRun;
};

export type MathRun = {
  text?: string;
  nodes?: MathNode[];
};

export type MathNode =
  | { type: "text"; text: string }
  | { type: "fraction"; numerator: MathNode[]; denominator: MathNode[] }
  | { type: "superscript"; base: MathNode[]; superscript: MathNode[] }
  | { type: "subscript"; base: MathNode[]; subscript: MathNode[] }
  | { type: "radical"; degree?: MathNode[]; content: MathNode[] }
  | { type: "nary"; operator: "sum"; lowerLimit?: MathNode[]; upperLimit?: MathNode[]; body: MathNode[] }
  | { type: "matrix"; rows: MathNode[][][] }
  | { type: "delimiter"; begin?: string; end?: string; content: MathNode[] }
  | { type: "accent"; mark: string; content: MathNode[] }
  | { type: "bar"; position: "top" | "bottom"; content: MathNode[] }
  | { type: "function"; name: MathNode[]; argument: MathNode[] }
  | { type: "limitLower"; base: MathNode[]; limit: MathNode[] }
  | { type: "limitUpper"; base: MathNode[]; limit: MathNode[] };

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

export type ContentControl = {
  alias?: string;
  tag?: string;
  lock?: "sdtLocked" | "contentLocked" | "sdtContentLocked" | "unlocked";
  appearance?: "boundingBox" | "tags" | "hidden";
  color?: string;
  showingPlaceholder?: boolean;
  dataBinding?: ContentControlDataBinding;
  placeholder?: ContentControlPlaceholder;
  checkbox?: CheckboxContentControl;
  dropdown?: DropdownContentControl;
  comboBox?: DropdownContentControl;
  date?: DateContentControl;
  repeatingSection?: RepeatingSectionContentControl;
  repeatingSectionItem?: RepeatingSectionItemContentControl;
};

export type ContentControlDataBinding = {
  storeItemId?: string;
  xpath?: string;
  prefixMappings?: string;
};

export type ContentControlPlaceholder = {
  docPart: string;
};

export type CheckboxContentControl = {
  checked: boolean;
  checkedSymbol?: string;
  uncheckedSymbol?: string;
};

export type DropdownContentControl = {
  items: DropdownItem[];
};

export type DropdownItem = {
  displayText: string;
  value: string;
};

export type DateContentControl = {
  fullDate?: string;
  format?: string;
};

export type RepeatingSectionContentControl = {
  sectionTitle?: string;
  doNotAllowInsertDeleteSection?: boolean;
};

export type RepeatingSectionItemContentControl = {
  id?: string;
};

export type Hyperlink = {
  url: string;
} | {
  anchor: string;
};

export type Comment = {
  id?: number;
  author: string;
  initials?: string;
  date?: string;
  text: string;
};

export type CommentRangeEnd = {
  id: number;
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
