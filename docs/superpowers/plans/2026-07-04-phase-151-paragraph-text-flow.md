# Phase 151 Paragraph Text Flow Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Preserve paragraph-level text flow properties for document paragraphs and paragraph style definitions.

**Architecture:** Extend the existing paragraph `<w:pPr>` property bucket with value properties `textAlignment` and `textDirection`, plus on/off controls `adjustRightInd`, `autoSpaceDE`, and `autoSpaceDN`. Reuse the existing paragraph on/off helper for boolean XML, while adding small value helpers for `<w:textAlignment w:val="..."/>` and `<w:textDirection w:val="..."/>`.

**Tech Stack:** TypeScript, JSZip-based DOCX package tests, existing DOCX reader/writer XML helpers, Vitest.

---

### Task 1: Add RED Tests For Paragraph Text Flow

**Files:**
- Modify: `tests/docx-core.test.ts`

- [x] **Step 1: Write paragraph writer test before implementation**

Add this test near the existing paragraph pagination/layout writer tests:

```ts
  it("writes paragraph text flow properties", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        pagination: {
          textAlignment: "center",
          textDirection: "tbRl",
          adjustRightInd: true,
          autoSpaceDE: false,
          autoSpaceDN: true,
        },
        runs: [{ text: "Vertical paragraph" }],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<w:textAlignment w:val="center"/>');
    expect(xml).toContain('<w:textDirection w:val="tbRl"/>');
    expect(xml).toContain("<w:adjustRightInd/>");
    expect(xml).toContain('<w:autoSpaceDE w:val="0"/>');
    expect(xml).toContain("<w:autoSpaceDN/>");
  });
```

- [x] **Step 2: Write paragraph style writer test before implementation**

Add this test near the existing paragraph style pagination/layout writer tests:

```ts
  it("writes paragraph style text flow properties", async () => {
    const document = {
      version: "1.0" as const,
      styles: {
        paragraph: [{
          id: "VerticalBody",
          name: "Vertical Body",
          paragraph: {
            pagination: {
              textAlignment: "baseline",
              textDirection: "btLr",
              adjustRightInd: false,
              autoSpaceDE: true,
              autoSpaceDN: false,
            },
          },
        }],
      },
      sections: [{ blocks: [{ type: "paragraph" as const, styleId: "VerticalBody", runs: [{ text: "Body" }] }] }],
    };

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const styles = await zip.file("word/styles.xml")!.async("string");

    expect(styles).toContain('<w:style w:type="paragraph" w:styleId="VerticalBody"><w:name w:val="Vertical Body"/><w:pPr><w:textAlignment w:val="baseline"/><w:textDirection w:val="btLr"/><w:adjustRightInd w:val="0"/><w:autoSpaceDE/><w:autoSpaceDN w:val="0"/></w:pPr></w:style>');
  });
```

- [x] **Step 3: Write reader round-trip test before implementation**

Add this test near the existing paragraph round-trip tests:

```ts
  it("round-trips paragraph text flow properties", async () => {
    const source = {
      version: "1.0" as const,
      styles: {
        paragraph: [{
          id: "VerticalBody",
          name: "Vertical Body",
          paragraph: {
            pagination: {
              textAlignment: "baseline",
              textDirection: "btLr",
              adjustRightInd: false,
              autoSpaceDE: true,
              autoSpaceDN: false,
            },
          },
        }],
      },
      sections: [{
        blocks: [{
          type: "paragraph" as const,
          styleId: "VerticalBody",
          pagination: {
            textAlignment: "center",
            textDirection: "tbRl",
            adjustRightInd: true,
            autoSpaceDE: false,
            autoSpaceDN: true,
          },
          runs: [{ text: "Vertical paragraph" }],
        }],
      }],
    };

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });
```

- [x] **Step 4: Run RED tests**

Run: `npm test -- tests/docx-core.test.ts -t "paragraph text flow"`

Expected: FAIL because paragraph text flow properties are not written or parsed yet.

### Task 2: Implement Paragraph Text Flow

**Files:**
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`
- Modify: `src/docx-reader.ts`

- [x] **Step 1: Extend schema**

Add these fields to `ParagraphPagination`:

```ts
  textAlignment?: "auto" | "baseline" | "bottom" | "center" | "top";
  textDirection?: "lrTb" | "tbRl" | "btLr";
  adjustRightInd?: boolean;
  autoSpaceDE?: boolean;
  autoSpaceDN?: boolean;
```

- [x] **Step 2: Emit text flow XML**

Add these entries to `paragraphPaginationXml`:

```ts
      pagination.textAlignment ? `<w:textAlignment w:val="${pagination.textAlignment}"/>` : "",
      pagination.textDirection ? `<w:textDirection w:val="${pagination.textDirection}"/>` : "",
      paragraphPaginationToggleXml("adjustRightInd", pagination.adjustRightInd),
      paragraphPaginationToggleXml("autoSpaceDE", pagination.autoSpaceDE),
      paragraphPaginationToggleXml("autoSpaceDN", pagination.autoSpaceDN),
```

- [x] **Step 3: Parse text flow XML**

Add parsing in both `parseParagraphPagination` and `parsePagination`:

```ts
  const textAlignment = asObject(properties.textAlignment);
  const textDirection = asObject(properties.textDirection);
  const adjustRightInd = parsePaginationToggle(properties.adjustRightInd);
  const autoSpaceDE = parsePaginationToggle(properties.autoSpaceDE);
  const autoSpaceDN = parsePaginationToggle(properties.autoSpaceDN);
```

Add fields to the returned object:

```ts
    ...(typeof textAlignment.val === "string" ? { textAlignment: textAlignment.val as NonNullable<ParagraphNode["pagination"]>["textAlignment"] } : {}),
    ...(typeof textDirection.val === "string" ? { textDirection: textDirection.val as NonNullable<ParagraphNode["pagination"]>["textDirection"] } : {}),
    ...(adjustRightInd !== undefined ? { adjustRightInd } : {}),
    ...(autoSpaceDE !== undefined ? { autoSpaceDE } : {}),
    ...(autoSpaceDN !== undefined ? { autoSpaceDN } : {}),
```

- [x] **Step 4: Run targeted GREEN tests**

Run: `npm test -- tests/docx-core.test.ts -t "paragraph text flow"`

Expected: PASS with three targeted tests passing.

### Task 3: Verify, Commit, Push

**Files:**
- Modify: `docs/superpowers/plans/2026-07-04-phase-151-paragraph-text-flow.md`

- [x] **Step 1: Run full verification**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "paragraph text flow"
npm test
npm run build
git diff --check
```

Expected: Targeted tests, full tests, and build pass. `git diff --check` may print LF/CRLF warnings but must exit 0.

- [ ] **Step 2: Commit implementation**

```bash
git add src/schema.ts src/docx-writer.ts src/docx-reader.ts tests/docx-core.test.ts docs/superpowers/plans/2026-07-04-phase-151-paragraph-text-flow.md
git commit -m "feat: add phase 151 paragraph text flow"
```

- [ ] **Step 3: Push branch**

```bash
git push -u origin phase-151-paragraph-text-flow
```

- [ ] **Step 4: Record push metadata**

Append the pushed commit hash and remote branch to this plan, then commit the plan update:

```bash
git add docs/superpowers/plans/2026-07-04-phase-151-paragraph-text-flow.md
git commit -m "docs: mark phase 151 pushed"
git push
```

## Push Record

- Pending.
