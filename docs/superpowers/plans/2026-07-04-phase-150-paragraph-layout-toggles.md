# Phase 150 Paragraph Layout Toggles Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Preserve additional paragraph layout on/off controls for paragraph nodes and paragraph style definitions.

**Architecture:** Extend the existing `ParagraphPagination` structure with four more `<w:pPr>` on/off controls: `contextualSpacing`, `mirrorIndents`, `overflowPunct`, and `topLinePunct`. Reuse the existing paragraph pagination on/off writer and reader helpers so both `true` and explicit `false` round-trip for document paragraphs and paragraph styles.

**Tech Stack:** TypeScript, JSZip-based DOCX package tests, existing DOCX reader/writer XML helpers, Vitest.

---

### Task 1: Add RED Tests For Paragraph Layout Toggles

**Files:**
- Modify: `tests/docx-core.test.ts`

- [x] **Step 1: Write paragraph writer test before implementation**

Add this test near the existing paragraph pagination writer tests:

```ts
  it("writes paragraph layout toggles", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        pagination: { contextualSpacing: true, mirrorIndents: false, overflowPunct: true, topLinePunct: false },
        runs: [{ text: "Layout-sensitive paragraph" }],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain("<w:contextualSpacing/>");
    expect(xml).toContain('<w:mirrorIndents w:val="0"/>');
    expect(xml).toContain("<w:overflowPunct/>");
    expect(xml).toContain('<w:topLinePunct w:val="0"/>');
  });
```

- [x] **Step 2: Write paragraph style writer test before implementation**

Add this test near the existing paragraph style pagination writer tests:

```ts
  it("writes paragraph style layout toggles", async () => {
    const document = {
      version: "1.0" as const,
      styles: {
        paragraph: [{
          id: "LayoutBody",
          name: "Layout Body",
          paragraph: {
            pagination: { contextualSpacing: false, mirrorIndents: true, overflowPunct: false, topLinePunct: true },
          },
        }],
      },
      sections: [{ blocks: [{ type: "paragraph" as const, styleId: "LayoutBody", runs: [{ text: "Body" }] }] }],
    };

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const styles = await zip.file("word/styles.xml")!.async("string");

    expect(styles).toContain('<w:style w:type="paragraph" w:styleId="LayoutBody"><w:name w:val="Layout Body"/><w:pPr><w:contextualSpacing w:val="0"/><w:mirrorIndents/><w:overflowPunct w:val="0"/><w:topLinePunct/></w:pPr></w:style>');
  });
```

- [x] **Step 3: Write reader round-trip test before implementation**

Add this test near the existing pagination round-trip tests:

```ts
  it("round-trips paragraph layout toggles", async () => {
    const source = {
      version: "1.0" as const,
      styles: {
        paragraph: [{
          id: "LayoutBody",
          name: "Layout Body",
          paragraph: {
            pagination: { contextualSpacing: false, mirrorIndents: true, overflowPunct: false, topLinePunct: true },
          },
        }],
      },
      sections: [{
        blocks: [{
          type: "paragraph" as const,
          styleId: "LayoutBody",
          pagination: { contextualSpacing: true, mirrorIndents: false, overflowPunct: true, topLinePunct: false },
          runs: [{ text: "Layout-sensitive paragraph" }],
        }],
      }],
    };

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });
```

- [x] **Step 4: Run RED tests**

Run: `npm test -- tests/docx-core.test.ts -t "paragraph layout toggles"`

Expected: FAIL because `contextualSpacing`, `mirrorIndents`, `overflowPunct`, and `topLinePunct` are not yet written or parsed.

### Task 2: Implement Paragraph Layout Toggles

**Files:**
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`
- Modify: `src/docx-reader.ts`

- [x] **Step 1: Extend schema**

Add these fields to `ParagraphPagination`:

```ts
  contextualSpacing?: boolean;
  mirrorIndents?: boolean;
  overflowPunct?: boolean;
  topLinePunct?: boolean;
```

- [x] **Step 2: Emit layout toggle XML**

Add these calls to `paragraphPaginationXml`:

```ts
      paragraphPaginationToggleXml("contextualSpacing", pagination.contextualSpacing),
      paragraphPaginationToggleXml("mirrorIndents", pagination.mirrorIndents),
      paragraphPaginationToggleXml("overflowPunct", pagination.overflowPunct),
      paragraphPaginationToggleXml("topLinePunct", pagination.topLinePunct),
```

- [x] **Step 3: Parse layout toggle XML**

Add these values to both `parseParagraphPagination` and `parsePagination`:

```ts
  const contextualSpacing = parsePaginationToggle(properties.contextualSpacing);
  const mirrorIndents = parsePaginationToggle(properties.mirrorIndents);
  const overflowPunct = parsePaginationToggle(properties.overflowPunct);
  const topLinePunct = parsePaginationToggle(properties.topLinePunct);
```

Add them to the returned object:

```ts
    ...(contextualSpacing !== undefined ? { contextualSpacing } : {}),
    ...(mirrorIndents !== undefined ? { mirrorIndents } : {}),
    ...(overflowPunct !== undefined ? { overflowPunct } : {}),
    ...(topLinePunct !== undefined ? { topLinePunct } : {}),
```

- [x] **Step 4: Run targeted GREEN tests**

Run: `npm test -- tests/docx-core.test.ts -t "paragraph layout toggles"`

Expected: PASS with three targeted tests passing.

### Task 3: Verify, Commit, Push

**Files:**
- Modify: `docs/superpowers/plans/2026-07-04-phase-150-paragraph-layout-toggles.md`

- [x] **Step 1: Run full verification**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "paragraph layout toggles"
npm test
npm run build
git diff --check
```

Expected: Targeted tests, full tests, and build pass. `git diff --check` may print LF/CRLF warnings but must exit 0.

- [x] **Step 2: Commit implementation**

```bash
git add src/schema.ts src/docx-writer.ts src/docx-reader.ts tests/docx-core.test.ts docs/superpowers/plans/2026-07-04-phase-150-paragraph-layout-toggles.md
git commit -m "feat: add phase 150 paragraph layout toggles"
```

- [x] **Step 3: Push branch**

```bash
git push -u origin phase-150-paragraph-layout-toggles
```

- [x] **Step 4: Record push metadata**

Append the pushed commit hash and remote branch to this plan, then commit the plan update:

```bash
git add docs/superpowers/plans/2026-07-04-phase-150-paragraph-layout-toggles.md
git commit -m "docs: mark phase 150 pushed"
git push
```

## Push Record

- Branch: `phase-150-paragraph-layout-toggles`
- Remote: `origin`
- Repository: `https://github.com/codingayice/word2json.git`
- Implementation commit: `7c76e9a5e1fce039712ea94baa6cc068b12e754b`
- Pull request URL: `https://github.com/codingayice/word2json/pull/new/phase-150-paragraph-layout-toggles`
