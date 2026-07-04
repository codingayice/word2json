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
  nsid?: string;
  multiLevelType?: "singleLevel" | "multilevel" | "hybridMultilevel";
  templateCode?: string;
  styleLink?: string;
  numberingStyleLink?: string;
  levels: NumberingLevelDefinition[];
};

export type NumberingInstance = {
  id: number;
  abstractId: number;
  overrides?: NumberingLevelOverride[];
};

export type NumberingLevelOverride = {
  level: number;
  start?: number;
  definition?: NumberingLevelDefinition;
};

export type NumberingLevelDefinition = {
  level: number;
  format: NumberingFormat;
  text: string;
  start?: number;
  styleId?: string;
  suffix?: "nothing" | "space" | "tab";
  restart?: number;
  legal?: boolean;
  alignment?: ParagraphAlignment;
  run?: StyleRunProperties;
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
  eastAsiaFontFamily?: string;
  complexScriptFontFamily?: string;
  fontTheme?: string;
  eastAsiaFontTheme?: string;
  complexScriptFontTheme?: string;
  fontHint?: RunFontHint;
  fontSize?: number;
  complexScriptFontSize?: number;
  color?: string;
  highlight?: HighlightColor;
  strike?: boolean;
  doubleStrike?: boolean;
  smallCaps?: boolean;
  allCaps?: boolean;
  shadow?: boolean;
  outline?: boolean;
  emboss?: boolean;
  imprint?: boolean;
  rtl?: boolean;
  complexScript?: boolean;
  specVanish?: boolean;
  hidden?: boolean;
  webHidden?: boolean;
  snapToGrid?: boolean;
  noProof?: boolean;
  officeMath?: boolean;
  language?: RunLanguage;
  characterPosition?: number;
  kerning?: number;
  verticalAlign?: RunVerticalAlign;
  characterSpacing?: number;
  scale?: number;
  fitText?: RunFitText;
  emphasis?: RunEmphasis;
  border?: BorderDefinition;
};

export type HighlightColor = "yellow" | "green" | "cyan" | "magenta" | "blue" | "red" | "darkBlue" | "darkCyan" | "darkGreen" | "darkMagenta" | "darkRed" | "darkYellow" | "darkGray" | "lightGray" | "black";

export type RunVerticalAlign = "superscript" | "subscript" | "baseline";

export type RunFontHint = "default" | "eastAsia" | "cs";

export type RunFitText = {
  width: number;
  id?: number;
};

export type RunEmphasis = "dot" | "comma" | "circle" | "underDot" | "none";

export type RunLanguage = {
  value?: string;
  eastAsia?: string;
  bidi?: string;
};

export type StyleParagraphProperties = {
  alignment?: ParagraphAlignment;
  spacing?: ParagraphSpacing;
  indent?: ParagraphIndent;
  shading?: ShadingDefinition;
  borders?: ParagraphBorders;
  pagination?: ParagraphPagination;
  frame?: ParagraphFrameProperties;
  tabs?: ParagraphTabStop[];
};

export type ShadingDefinition = {
  fill: string;
  value?: string;
  color?: string;
  themeFill?: string;
  themeFillTint?: string;
  themeFillShade?: string;
  themeColor?: string;
  themeTint?: string;
  themeShade?: string;
};

export type BorderStyle = "single" | "double" | "dashed" | "dotted" | "nil" | "none" | "thick" | "dotDash" | "dotDotDash" | "triple" | "wave";

export type BorderDefinition = {
  style: BorderStyle;
  size?: number;
  color?: string;
  themeColor?: string;
  themeTint?: string;
  themeShade?: string;
  space?: number;
};

export type ParagraphBorders = {
  top?: BorderDefinition;
  left?: BorderDefinition;
  bottom?: BorderDefinition;
  right?: BorderDefinition;
  between?: BorderDefinition;
  bar?: BorderDefinition;
};

export type TableCellBorders = ParagraphBorders & {
  insideH?: BorderDefinition;
  insideV?: BorderDefinition;
};

export type TableStyleDefinition = StyleDefinition & {
  table?: StyleTableProperties;
};

export type StyleTableProperties = {
  borders?: TableBorders;
};

export type SectionNode = {
  breakType?: SectionBreakType;
  titlePage?: boolean;
  page?: PageSettings;
  pageNumbering?: SectionPageNumbering;
  lineNumbering?: SectionLineNumbering;
  footnoteProperties?: SectionNoteProperties;
  endnoteProperties?: SectionNoteProperties;
  noEndnote?: boolean;
  documentGrid?: SectionDocumentGrid;
  verticalAlignment?: SectionVerticalAlignment;
  textDirection?: SectionTextDirection;
  bidi?: boolean;
  rtlGutter?: boolean;
  mirrorMargins?: boolean;
  watermark?: TextWatermark;
  headers?: HeaderFooterContent;
  footers?: HeaderFooterContent;
  columns?: ColumnSettings;
  blocks: DocumentBlock[];
};

export type TextWatermark = {
  text: string;
  color?: string;
  opacity?: number;
  rotation?: number;
  fontFamily?: string;
};

export type SectionBreakType = "nextPage" | "continuous" | "evenPage" | "oddPage";

export type ColumnSettings = {
  count: number;
  space?: number;
};

export type SectionNoteProperties = {
  position?: "pageBottom" | "beneathText" | "sectEnd" | "docEnd";
  numbering?: {
    format?: NumberingFormat;
    start?: number;
    restart?: "continuous" | "eachSect" | "eachPage";
  };
};

export type SectionDocumentGrid = {
  type?: "default" | "lines" | "linesAndChars" | "snapToChars";
  linePitch?: number;
  charSpace?: number;
};

export type SectionVerticalAlignment = "top" | "center" | "both" | "bottom";

export type SectionTextDirection = "lrTb" | "tbRl" | "btLr";

export type HeaderFooterContent = {
  default?: ParagraphNode[];
  first?: ParagraphNode[];
  even?: ParagraphNode[];
};

export type SectionPageNumbering = {
  start?: number;
  format?: NumberingFormat;
  chapterStyle?: number;
  chapterSeparator?: "colon" | "emDash" | "enDash" | "hyphen" | "period";
};

export type SectionLineNumbering = {
  start?: number;
  countBy?: number;
  distance?: number;
  restart?: "continuous" | "newPage" | "newSection";
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
  frame?: ParagraphFrameProperties;
  tabs?: ParagraphTabStop[];
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
  widowControl?: boolean;
  suppressLineNumbers?: boolean;
  suppressAutoHyphens?: boolean;
  contextualSpacing?: boolean;
  mirrorIndents?: boolean;
  overflowPunct?: boolean;
  topLinePunct?: boolean;
  textAlignment?: "auto" | "baseline" | "bottom" | "center" | "top";
  textDirection?: "lrTb" | "tbRl" | "btLr";
  adjustRightInd?: boolean;
  autoSpaceDE?: boolean;
  autoSpaceDN?: boolean;
};

export type ParagraphFrameProperties = {
  width?: number;
  height?: number;
  x?: number;
  y?: number;
  horizontalAnchor?: "text" | "margin" | "page";
  verticalAnchor?: "text" | "margin" | "page";
  xAlign?: "left" | "center" | "right" | "inside" | "outside";
  yAlign?: "top" | "center" | "bottom" | "inside" | "outside";
  wrap?: "around" | "auto" | "none" | "notBeside" | "through" | "tight";
  dropCap?: "drop" | "margin" | "none";
  lines?: number;
  anchorLock?: boolean;
  heightRule?: "auto" | "atLeast" | "exact";
};

export type ParagraphTabStop = {
  value: "bar" | "center" | "clear" | "decimal" | "end" | "left" | "num" | "right" | "start";
  position: number;
  leader?: "dot" | "heavy" | "hyphen" | "middleDot" | "none" | "underscore";
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

export type MathControlProperties = Pick<StyleRunProperties, "bold" | "italic" | "underline" | "fontFamily" | "fontSize" | "color" | "highlight"> & {
  styleId?: string;
};

export type MathNode =
  | { type: "text"; text: string }
  | { type: "fraction"; fractionType?: "bar" | "skewed" | "linear" | "noBar"; controlProperties?: MathControlProperties; numerator: MathNode[]; denominator: MathNode[] }
  | { type: "superscript"; controlProperties?: MathControlProperties; base: MathNode[]; superscript: MathNode[] }
  | { type: "subscript"; controlProperties?: MathControlProperties; base: MathNode[]; subscript: MathNode[] }
  | { type: "subSup"; controlProperties?: MathControlProperties; base: MathNode[]; subscript: MathNode[]; superscript: MathNode[] }
  | { type: "sPre"; controlProperties?: MathControlProperties; base: MathNode[]; subscript: MathNode[]; superscript: MathNode[] }
  | { type: "preSubSup"; controlProperties?: MathControlProperties; base: MathNode[]; subscript: MathNode[]; superscript: MathNode[] }
  | { type: "radical"; controlProperties?: MathControlProperties; hideDegree?: boolean; degree?: MathNode[]; content: MathNode[] }
  | { type: "nary"; controlProperties?: MathControlProperties; operator: "sum" | "integral" | "product" | "coproduct" | "intersection" | "union"; operatorCharacter?: string; limitLocation?: "underOver" | "subSup"; grow?: boolean; hideLowerLimit?: boolean; hideUpperLimit?: boolean; lowerLimit?: MathNode[]; upperLimit?: MathNode[]; body: MathNode[] }
  | { type: "matrix"; controlProperties?: MathControlProperties; baseJustification?: "top" | "center" | "bottom"; rowSpacing?: number; rowSpacingRule?: "single" | "oneAndHalf" | "double" | "exactly" | "multiple"; columnSpacing?: number; columnSpacingRule?: "single" | "oneAndHalf" | "double" | "exactly" | "multiple"; columnJustifications?: ("left" | "center" | "right")[]; columnCounts?: number[]; rows: MathNode[][][] }
  | { type: "delimiter"; controlProperties?: MathControlProperties; begin?: string; end?: string; grow?: boolean; separator?: string; content: MathNode[] }
  | { type: "accent"; controlProperties?: MathControlProperties; mark: string; content: MathNode[] }
  | { type: "bar"; controlProperties?: MathControlProperties; position: "top" | "bottom"; content: MathNode[] }
  | { type: "function"; controlProperties?: MathControlProperties; name: MathNode[]; argument: MathNode[] }
  | { type: "limitLower"; controlProperties?: MathControlProperties; base: MathNode[]; limit: MathNode[] }
  | { type: "limitUpper"; controlProperties?: MathControlProperties; base: MathNode[]; limit: MathNode[] }
  | { type: "equationArray"; controlProperties?: MathControlProperties; baseJustification?: "top" | "center" | "bottom"; verticalJustification?: "top" | "bottom"; alignment?: boolean; rowSpacing?: number; rowSpacingRule?: "single" | "oneAndHalf" | "double" | "exactly" | "multiple"; objectDistribution?: boolean; maxDistribution?: boolean; rows: MathNode[][] }
  | { type: "box"; controlProperties?: MathControlProperties; hideTop?: boolean; hideBottom?: boolean; hideLeft?: boolean; hideRight?: boolean; content: MathNode[] }
  | { type: "borderBox"; controlProperties?: MathControlProperties; hideTop?: boolean; hideBottom?: boolean; hideLeft?: boolean; hideRight?: boolean; content: MathNode[] }
  | { type: "phantom"; controlProperties?: MathControlProperties; show?: boolean; zeroWidth?: boolean; zeroAscent?: boolean; zeroDescent?: boolean; transparent?: boolean; content: MathNode[] }
  | { type: "groupCharacter"; controlProperties?: MathControlProperties; character?: string; position?: "top" | "bottom"; verticalJustification?: "top" | "bottom"; content: MathNode[] };

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
  position?: TablePosition;
  overlap?: "never" | "overlap";
  caption?: string;
  description?: string;
  look?: TableLook;
  layout?: "autofit" | "fixed";
  grid?: number[];
  width?: number;
  widthType?: "auto" | "dxa" | "nil" | "pct";
  borders?: TableBorders;
  alignment?: ParagraphAlignment;
  cellSpacing?: number;
  indent?: TableIndent;
  propertyRevision?: ParagraphPropertyRevision;
  rows: TableRowNode[];
};

export type TablePosition = {
  horizontalAnchor?: "margin" | "page" | "text";
  verticalAnchor?: "margin" | "page" | "text";
  x?: number;
  y?: number;
  xAlign?: "center" | "inside" | "left" | "outside" | "right";
  yAlign?: "bottom" | "center" | "inside" | "outside" | "top";
  leftFromText?: number;
  rightFromText?: number;
  topFromText?: number;
  bottomFromText?: number;
};

export type TableIndent = {
  width: number;
  type?: "dxa" | "nil" | "pct";
};

export type TableLook = {
  firstRow?: boolean;
  lastRow?: boolean;
  firstColumn?: boolean;
  lastColumn?: boolean;
  bandedRows?: boolean;
  bandedColumns?: boolean;
};

export type TableBorders = "single" | {
  top?: BorderDefinition;
  left?: BorderDefinition;
  bottom?: BorderDefinition;
  right?: BorderDefinition;
  insideH?: BorderDefinition;
  insideV?: BorderDefinition;
};

export type TableRowNode = {
  propertyExceptions?: TablePropertyExceptions;
  revision?: RunRevision;
  height?: TableRowHeight;
  repeatHeader?: boolean;
  cantSplit?: boolean;
  cells: TableCellNode[];
};

export type TablePropertyExceptions = {
  width?: number;
  widthType?: "auto" | "dxa" | "nil" | "pct";
  cellSpacing?: number;
  indent?: TableIndent;
  layout?: "autofit" | "fixed";
  look?: TableLook;
};

export type TableRowHeight = {
  value: number;
  rule?: "auto" | "atLeast" | "exact";
};

export type TableCellNode = {
  width?: number;
  widthType?: "auto" | "dxa" | "nil" | "pct";
  colSpan?: number;
  verticalMerge?: "restart" | "continue";
  verticalAlignment?: "top" | "center" | "bottom";
  shading?: TableCellShading;
  borders?: TableCellBorders;
  noWrap?: boolean;
  textDirection?: "lrTb" | "tbRl" | "btLr";
  fitText?: boolean;
  margins?: TableCellMargins;
  propertyRevision?: ParagraphPropertyRevision;
  blocks: ParagraphNode[];
};

export type TableCellShading = ShadingDefinition;

export type TableCellMargins = {
  top?: TableCellMarginSide;
  right?: TableCellMarginSide;
  bottom?: TableCellMarginSide;
  left?: TableCellMarginSide;
};

export type TableCellMarginSide = number | {
  width: number;
  type?: "auto" | "dxa" | "nil" | "pct";
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
  wrap: "none" | "square" | "topAndBottom";
  horizontalOffset: number;
  verticalOffset: number;
  horizontalRelativeFrom?: "page" | "margin" | "column" | "character";
  verticalRelativeFrom?: "page" | "margin" | "paragraph" | "line";
  horizontalAlign?: "left" | "center" | "right" | "inside" | "outside";
  verticalAlign?: "top" | "center" | "bottom" | "inside" | "outside";
  simplePosition?: { x: number; y: number };
  relativeHeight?: number;
  locked?: boolean;
  distanceTop?: number;
  distanceBottom?: number;
  distanceLeft?: number;
  distanceRight?: number;
  behindDoc?: boolean;
  allowOverlap?: boolean;
  layoutInCell?: boolean;
};

export function createDocumentJson(blocks: DocumentBlock[]): DocumentJson {
  return {
    version: "1.0",
    sections: [{ blocks }],
  };
}
