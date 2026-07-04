# Phase 149 Paragraph Pagination Extra Toggles Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Preserve additional paragraph pagination and line-layout on/off controls for paragraph nodes and paragraph style definitions.

**Architecture:** Extend the existing `ParagraphPagination` shape with `widowControl`, `suppressLineNumbers`, and `suppressAutoHyphens`. Reuse the Phase 148 on/off XML helper so `true` emits empty OOXML elements and explicit `false` emits `w:val="0"` for both document paragraphs and paragraph styles.

**Tech Stack:** TypeScript, JSZip-based DOCX package tests, existing DOCX reader/writer XML helpers, Vitest.

---

### Task 1: Add RED Tests For Extra Paragraph Pagination Toggles

**Files:**
- Modify: `tests/docx-core.test.ts`

- [x] **Step 1: Write paragraph writer test before implementation**

Add this test near the existing paragraph pagination writer tests:

```ts
  it("writes extra paragraph pagination toggles", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        pagination: { widowControl: true, suppressLineNumbers: true, suppressAutoHyphens: false },
        runs: [{ text: "Controlled typography" }],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain("<w:widowControl/>");
    expect(xml).toContain("<w:suppressLineNumbers/>");
    expect(xml).toContain('<w:suppressAutoHyphens w:val="0"/>');
  });
```

- [x] **Step 2: Write paragraph style writer test before implementation**

Add this test near the existing paragraph style pagination writer tests:

```ts
  it("writes extra paragraph style pagination toggles", async () => {
    const document = {
      version: "1.0" as const,
      styles: {
        paragraph: [{
          id: "ControlledBody",
          name: "Controlled Body",
          paragraph: {
            pagination: { widowControl: false, suppressLineNumbers: true, suppressAutoHyphens: true },
          },
        }],
      },
      sections: [{ blocks: [{ type: "paragraph" as const, styleId: "ControlledBody", runs: [{ text: "Body" }] }] }],
    };

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const styles = await zip.file("word/styles.xml")!.async("string");

    expect(styles).toContain('<w:style w:type="paragraph" w:styleId="ControlledBody"><w:name w:val="Controlled Body"/><w:pPr><w:widowControl w:val="0"/><w:suppressLineNumbers/><w:suppressAutoHyphens/></w:pPr></w:style>');
  });
```

- [x] **Step 3: Write reader round-trip test before implementation**

Add this test near the existing pagination round-trip tests:

```ts
  it("round-trips extra paragraph pagination toggles", async () => {
    const source = {
      version: "1.0" as const,
      styles: {
        paragraph: [{
          id: "ControlledBody",
          name: "Controlled Body",
          paragraph: {
            pagination: { widowControl: false, suppressLineNumbers: true, suppressAutoHyphens: true },
          },
        }],
      },
      sections: [{
        blocks: [{
          type: "paragraph" as const,
          styleId: "ControlledBody",
          pagination: { widowControl: true, suppressLineNumbers: true, suppressAutoHyphens: false },
          runs: [{ text: "Controlled typography" }],
        }],
      }],
    };

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });
```

- [x] **Step 4: Run RED tests**

Run: `npm test -- tests/docx-core.test.ts -t "extra paragraph pagination toggles"`

Expected: FAIL because the schema and XML helpers do not yet preserve `widowControl`, `suppressLineNumbers`, or `suppressAutoHyphens`.

### Task 2: Implement Extra Paragraph Pagination Toggles

**Files:**
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`
- Modify: `src/docx-reader.ts`

- [x] **Step 1: Extend schema**

Add these fields to `ParagraphPagination`:

```ts
  widowControl?: boolean;
  suppressLineNumbers?: boolean;
  suppressAutoHyphens?: boolean;
```

- [x] **Step 2: Emit extra toggle XML**

Add these calls to `paragraphPaginationXml` after `pageBreakBefore`:

```ts
      paragraphPaginationToggleXml("widowControl", pagination.widowControl),
      paragraphPaginationToggleXml("suppressLineNumbers", pagination.suppressLineNumbers),
      paragraphPaginationToggleXml("suppressAutoHyphens", pagination.suppressAutoHyphens),
```

- [x] **Step 3: Parse extra toggle XML**

Add these parsed values to both `parseParagraphPagination` and `parsePagination`:

```ts
  const widowControl = parsePaginationToggle(properties.widowControl);
  const suppressLineNumbers = parsePaginationToggle(properties.suppressLineNumbers);
  const suppressAutoHyphens = parsePaginationToggle(properties.suppressAutoHyphens);
```

Add them to the returned object:

```ts
    ...(widowControl !== undefined ? { widowControl } : {}),
    ...(suppressLineNumbers !== undefined ? { suppressLineNumbers } : {}),
    ...(suppressAutoHyphens !== undefined ? { suppressAutoHyphens } : {}),
```

- [x] **Step 4: Run targeted GREEN tests**

Run: `npm test -- tests/docx-core.test.ts -t "extra paragraph pagination toggles"`

Expected: PASS with three targeted tests passing.

### Task 3: Verify, Commit, Push

**Files:**
- Modify: `docs/superpowers/plans/2026-07-04-phase-149-paragraph-pagination-extra-toggles.md`

- [x] **Step 1: Run full verification**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "extra paragraph pagination toggles"
npm test
npm run build
git diff --check
```

Expected: Targeted tests, full tests, and build pass. `git diff --check` may print LF/CRLF warnings but must exit 0.

- [x] **Step 2: Commit implementation**

```bash
git add src/schema.ts src/docx-writer.ts src/docx-reader.ts tests/docx-core.test.ts docs/superpowers/plans/2026-07-04-phase-149-paragraph-pagination-extra-toggles.md
git commit -m "feat: add phase 149 paragraph pagination toggles"
```

- [x] **Step 3: Push branch**

```bash
git push -u origin phase-149-paragraph-pagination-extra-toggles
```

- [x] **Step 4: Record push metadata**

Append the pushed commit hash and remote branch to this plan, then commit the plan update:

```bash
git add docs/superpowers/plans/2026-07-04-phase-149-paragraph-pagination-extra-toggles.md
git commit -m "docs: mark phase 149 pushed"
git push
```

## Push Record

- Branch: `phase-149-paragraph-pagination-extra-toggles`
- Remote: `origin`
- Repository: `https://github.com/codingayice/word2json.git`
- Implementation commit: `537d1d54762327ea951b54df6d334b1b276e3b32`
- Pull request URL: `https://github.com/codingayice/word2json/pull/new/phase-149-paragraph-pagination-extra-toggles`
