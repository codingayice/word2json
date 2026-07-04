# Phase 152 Paragraph Frame Properties Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Preserve paragraph frame properties used by Word for floating or text-box-like paragraph layout.

**Architecture:** Add a focused `ParagraphFrameProperties` object on `ParagraphNode` and `StyleParagraphProperties`. The writer emits `<w:framePr>` inside paragraph properties for document paragraphs and styles, and the reader parses the same attributes back into JSON.

**Tech Stack:** TypeScript, JSZip-based DOCX package tests, existing DOCX reader/writer XML helpers, Vitest.

---

### Task 1: Add RED Tests For Paragraph Frame Properties

**Files:**
- Modify: `tests/docx-core.test.ts`

- [x] **Step 1: Write paragraph writer test before implementation**

Add this test near the existing paragraph writer property tests:

```ts
  it("writes paragraph frame properties", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        frame: {
          width: 2880,
          height: 1440,
          x: 720,
          y: 360,
          horizontalAnchor: "margin",
          verticalAnchor: "page",
          xAlign: "center",
          yAlign: "top",
          wrap: "around",
          dropCap: "drop",
          lines: 3,
          anchorLock: true,
          heightRule: "exact",
        },
        runs: [{ text: "Framed paragraph" }],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<w:framePr w:w="2880" w:h="1440" w:x="720" w:y="360" w:hAnchor="margin" w:vAnchor="page" w:xAlign="center" w:yAlign="top" w:wrap="around" w:dropCap="drop" w:lines="3" w:anchorLock="1" w:hRule="exact"/>');
  });
```

- [x] **Step 2: Write paragraph style writer test before implementation**

Add this test near the existing paragraph style property tests:

```ts
  it("writes paragraph style frame properties", async () => {
    const document = {
      version: "1.0" as const,
      styles: {
        paragraph: [{
          id: "Sidebar",
          name: "Sidebar",
          paragraph: {
            frame: {
              width: 2160,
              horizontalAnchor: "page",
              verticalAnchor: "margin",
              xAlign: "right",
              yAlign: "bottom",
              wrap: "notBeside",
              anchorLock: false,
            },
          },
        }],
      },
      sections: [{ blocks: [{ type: "paragraph" as const, styleId: "Sidebar", runs: [{ text: "Sidebar" }] }] }],
    };

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const styles = await zip.file("word/styles.xml")!.async("string");

    expect(styles).toContain('<w:style w:type="paragraph" w:styleId="Sidebar"><w:name w:val="Sidebar"/><w:pPr><w:framePr w:w="2160" w:hAnchor="page" w:vAnchor="margin" w:xAlign="right" w:yAlign="bottom" w:wrap="notBeside" w:anchorLock="0"/></w:pPr></w:style>');
  });
```

- [x] **Step 3: Write reader round-trip test before implementation**

Add this test near the existing paragraph round-trip tests:

```ts
  it("round-trips paragraph frame properties", async () => {
    const source = {
      version: "1.0" as const,
      styles: {
        paragraph: [{
          id: "Sidebar",
          name: "Sidebar",
          paragraph: {
            frame: {
              width: 2160,
              horizontalAnchor: "page",
              verticalAnchor: "margin",
              xAlign: "right",
              yAlign: "bottom",
              wrap: "notBeside",
              anchorLock: false,
            },
          },
        }],
      },
      sections: [{
        blocks: [{
          type: "paragraph" as const,
          styleId: "Sidebar",
          frame: {
            width: 2880,
            height: 1440,
            x: 720,
            y: 360,
            horizontalAnchor: "margin",
            verticalAnchor: "page",
            xAlign: "center",
            yAlign: "top",
            wrap: "around",
            dropCap: "drop",
            lines: 3,
            anchorLock: true,
            heightRule: "exact",
          },
          runs: [{ text: "Framed paragraph" }],
        }],
      }],
    };

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });
```

- [x] **Step 4: Run RED tests**

Run: `npm test -- tests/docx-core.test.ts -t "paragraph frame properties"`

Expected: FAIL because paragraph frame properties are not yet written or parsed.

### Task 2: Implement Paragraph Frame Properties

**Files:**
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`
- Modify: `src/docx-reader.ts`

- [x] **Step 1: Extend schema**

Add a `ParagraphFrameProperties` type and reference it from `ParagraphNode` and `StyleParagraphProperties`:

```ts
  frame?: ParagraphFrameProperties;
```

```ts
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
```

- [x] **Step 2: Emit framePr XML**

Add a `paragraphFrameXml(frame?: ParagraphFrameProperties): string` helper and call it from both `paragraphPropertiesXml` and `paragraphStylePropertiesXml`:

```ts
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
```

- [x] **Step 3: Parse framePr XML**

Add a `parseParagraphFrame(value: unknown): ParagraphNode["frame"] | undefined` helper and call it from both style paragraph parsing and document paragraph parsing:

```ts
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
```

- [x] **Step 4: Run targeted GREEN tests**

Run: `npm test -- tests/docx-core.test.ts -t "paragraph frame properties"`

Expected: PASS with three targeted tests passing.

### Task 3: Verify, Commit, Push

**Files:**
- Modify: `docs/superpowers/plans/2026-07-04-phase-152-paragraph-frame-properties.md`

- [x] **Step 1: Run full verification**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "paragraph frame properties"
npm test
npm run build
git diff --check
```

Expected: Targeted tests, full tests, and build pass. `git diff --check` may print LF/CRLF warnings but must exit 0.

- [ ] **Step 2: Commit implementation**

```bash
git add src/schema.ts src/docx-writer.ts src/docx-reader.ts tests/docx-core.test.ts docs/superpowers/plans/2026-07-04-phase-152-paragraph-frame-properties.md
git commit -m "feat: add phase 152 paragraph frame properties"
```

- [ ] **Step 3: Push branch**

```bash
git push -u origin phase-152-paragraph-frame-properties
```

- [ ] **Step 4: Record push metadata**

Append the pushed commit hash and remote branch to this plan, then commit the plan update:

```bash
git add docs/superpowers/plans/2026-07-04-phase-152-paragraph-frame-properties.md
git commit -m "docs: mark phase 152 pushed"
git push
```

## Push Record

- Pending.
