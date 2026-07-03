export type DocumentJson = {
  version: "1.0";
  sections: SectionNode[];
};

export type SectionNode = {
  blocks: DocumentBlock[];
};

export type DocumentBlock = ParagraphNode;

export type ParagraphNode = {
  type: "paragraph";
  alignment?: ParagraphAlignment;
  runs: TextRun[];
};

export type ParagraphAlignment = "left" | "center" | "right" | "both";

export type TextRun = {
  text: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
};

export function createDocumentJson(blocks: DocumentBlock[]): DocumentJson {
  return {
    version: "1.0",
    sections: [{ blocks }],
  };
}
