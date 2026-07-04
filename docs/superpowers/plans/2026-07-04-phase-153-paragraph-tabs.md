# Phase 153 Paragraph Tabs Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Preserve paragraph tab stops for document paragraphs and paragraph style definitions.

**Architecture:** Add `tabs?: ParagraphTabStop[]` to paragraph nodes and paragraph style properties. The writer emits `<w:tabs><w:tab .../></w:tabs>` inside `<w:pPr>`, and the reader parses single or multiple `w:tab` nodes back into ordered JSON tab stop arrays.

**Tech Stack:** TypeScript, JSZip-based DOCX package tests, existing DOCX reader/writer XML helpers, Vitest.

---

### Task 1: Add RED Tests For Paragraph Tabs

**Files:**
- Modify: `tests/docx-core.test.ts`

- [x] **Step 1: Write paragraph writer test before implementation**

Add this test near the existing paragraph writer property tests:

```ts
  it("writes paragraph tab stops", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        tabs: [
          { value: "left", position: 720 },
          { value: "right", position: 4320, leader: "dot" },
        ],
        runs: [{ text: "Label\tValue" }],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<w:tabs><w:tab w:val="left" w:pos="720"/><w:tab w:val="right" w:pos="4320" w:leader="dot"/></w:tabs>');
  });
```

- [x] **Step 2: Write paragraph style writer test before implementation**

Add this test near the existing paragraph style property tests:

```ts
  it("writes paragraph style tab stops", async () => {
    const document = {
      version: "1.0" as const,
      styles: {
        paragraph: [{
          id: "Tabular",
          name: "Tabular",
          paragraph: {
            tabs: [
              { value: "center", position: 2160, leader: "hyphen" },
              { value: "clear", position: 3600 },
            ],
          },
        }],
      },
      sections: [{ blocks: [{ type: "paragraph" as const, styleId: "Tabular", runs: [{ text: "A\tB" }] }] }],
    };

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const styles = await zip.file("word/styles.xml")!.async("string");

    expect(styles).toContain('<w:style w:type="paragraph" w:styleId="Tabular"><w:name w:val="Tabular"/><w:pPr><w:tabs><w:tab w:val="center" w:pos="2160" w:leader="hyphen"/><w:tab w:val="clear" w:pos="3600"/></w:tabs></w:pPr></w:style>');
  });
```

- [x] **Step 3: Write reader round-trip test before implementation**

Add this test near the existing paragraph round-trip tests:

```ts
  it("round-trips paragraph tab stops", async () => {
    const source = {
      version: "1.0" as const,
      styles: {
        paragraph: [{
          id: "Tabular",
          name: "Tabular",
          paragraph: {
            tabs: [
              { value: "center", position: 2160, leader: "hyphen" },
              { value: "clear", position: 3600 },
            ],
          },
        }],
      },
      sections: [{
        blocks: [{
          type: "paragraph" as const,
          styleId: "Tabular",
          tabs: [
            { value: "left", position: 720 },
            { value: "right", position: 4320, leader: "dot" },
          ],
          runs: [{ text: "Label\tValue" }],
        }],
      }],
    };

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });
```

- [x] **Step 4: Run RED tests**

Run: `npm test -- tests/docx-core.test.ts -t "paragraph tab stops"`

Expected: FAIL because paragraph tab stops are not yet written or parsed.

### Task 2: Implement Paragraph Tabs

**Files:**
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`
- Modify: `src/docx-reader.ts`

- [x] **Step 1: Extend schema**

Add `tabs?: ParagraphTabStop[]` to `ParagraphNode` and `StyleParagraphProperties`, then define:

```ts
export type ParagraphTabStop = {
  value: "bar" | "center" | "clear" | "decimal" | "end" | "left" | "num" | "right" | "start";
  position: number;
  leader?: "dot" | "heavy" | "hyphen" | "middleDot" | "none" | "underscore";
};
```

- [x] **Step 2: Emit tabs XML**

Add `paragraphTabsXml(tabs?: ParagraphTabStop[]): string` and call it from both paragraph and style paragraph property writers:

```ts
function paragraphTabsXml(tabs?: ParagraphNode["tabs"]): string {
  if (!tabs?.length) return "";

  return `<w:tabs>${tabs.map((tab) =>
    `<w:tab w:val="${tab.value}" w:pos="${tab.position}"${tab.leader ? ` w:leader="${tab.leader}"` : ""}/>`
  ).join("")}</w:tabs>`;
}
```

- [x] **Step 3: Parse tabs XML**

Add `parseParagraphTabs(value: unknown): ParagraphNode["tabs"] | undefined` and call it from both paragraph and style paragraph property parsers:

```ts
function parseParagraphTabs(value: unknown): ParagraphNode["tabs"] | undefined {
  const tabs = asArray(asObject(value).tab)
    .map((tabValue) => {
      const tab = asObject(tabValue);
      return typeof tab.val === "string" && tab.pos !== undefined
        ? {
          value: tab.val as NonNullable<ParagraphNode["tabs"]>[number]["value"],
          position: parseNumber(tab.pos),
          ...(typeof tab.leader === "string" ? { leader: tab.leader as NonNullable<ParagraphNode["tabs"]>[number]["leader"] } : {}),
        }
        : undefined;
    })
    .filter((tab): tab is NonNullable<ParagraphNode["tabs"]>[number] => tab !== undefined);

  return tabs.length > 0 ? tabs : undefined;
}
```

- [x] **Step 4: Run targeted GREEN tests**

Run: `npm test -- tests/docx-core.test.ts -t "paragraph tab stops"`

Expected: PASS with three targeted tests passing.

### Task 3: Verify, Commit, Push

**Files:**
- Modify: `docs/superpowers/plans/2026-07-04-phase-153-paragraph-tabs.md`

- [x] **Step 1: Run full verification**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "paragraph tab stops"
npm test
npm run build
git diff --check
```

Expected: Targeted tests, full tests, and build pass. `git diff --check` may print LF/CRLF warnings but must exit 0.

- [ ] **Step 2: Commit implementation**

```bash
git add src/schema.ts src/docx-writer.ts src/docx-reader.ts tests/docx-core.test.ts docs/superpowers/plans/2026-07-04-phase-153-paragraph-tabs.md
git commit -m "feat: add phase 153 paragraph tabs"
```

- [ ] **Step 3: Push branch**

```bash
git push -u origin phase-153-paragraph-tabs
```

- [ ] **Step 4: Record push metadata**

Append the pushed commit hash and remote branch to this plan, then commit the plan update:

```bash
git add docs/superpowers/plans/2026-07-04-phase-153-paragraph-tabs.md
git commit -m "docs: mark phase 153 pushed"
git push
```

## Push Record

- Pending.
