# Phase 147 Paragraph Style Pagination Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Preserve paragraph style pagination controls in Word style definitions.

**Architecture:** Paragraph nodes already support `pagination` for `keepNext`, `keepLines`, and `pageBreakBefore`. This phase adds the same shape to `StyleParagraphProperties`, writes it into style `<w:pPr>`, and parses it back from styles and defaults.

**Tech Stack:** TypeScript, JSZip-based DOCX package tests, existing DOCX reader/writer XML helpers, Vitest.

---

### Task 1: Add RED Tests For Paragraph Style Pagination

**Files:**
- Modify: `tests/docx-core.test.ts`

- [x] **Step 1: Write writer test before implementation**

Add a writer test near the existing paragraph style writer tests:

```ts
  it("writes paragraph style pagination properties", async () => {
    const document = {
      version: "1.0" as const,
      styles: {
        paragraph: [{
          id: "KeepHeading",
          name: "Keep Heading",
          paragraph: {
            pagination: { keepNext: true, keepLines: true, pageBreakBefore: true },
          },
        }],
      },
      sections: [{ blocks: [{ type: "paragraph" as const, styleId: "KeepHeading", runs: [{ text: "Heading" }] }] }],
    };

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const styles = await zip.file("word/styles.xml")!.async("string");

    expect(styles).toContain('<w:style w:type="paragraph" w:styleId="KeepHeading"><w:name w:val="Keep Heading"/><w:pPr><w:keepNext/><w:keepLines/><w:pageBreakBefore/></w:pPr></w:style>');
  });
```

- [x] **Step 2: Write reader round-trip test before implementation**

Add a reader test near the existing paragraph style round-trip tests:

```ts
  it("round-trips paragraph style pagination properties", async () => {
    const source = {
      version: "1.0" as const,
      styles: {
        paragraph: [{
          id: "KeepHeading",
          name: "Keep Heading",
          paragraph: {
            pagination: { keepNext: true, keepLines: true, pageBreakBefore: true },
          },
        }],
      },
      sections: [{ blocks: [{ type: "paragraph" as const, styleId: "KeepHeading", runs: [{ text: "Heading" }] }] }],
    };

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });
```

- [x] **Step 3: Run writer RED test**

Run: `npm test -- tests/docx-core.test.ts -t "writes paragraph style pagination"`

Expected: FAIL because style `<w:pPr>` does not emit pagination controls yet.

- [x] **Step 4: Run reader RED test**

Run: `npm test -- tests/docx-core.test.ts -t "round-trips paragraph style pagination"`

Expected: FAIL because style parsing drops paragraph pagination controls.

### Task 2: Implement Paragraph Style Pagination

**Files:**
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`
- Modify: `src/docx-reader.ts`

- [x] **Step 1: Add schema field**

Add to `StyleParagraphProperties`:

```ts
  pagination?: ParagraphPagination;
```

- [x] **Step 2: Emit paragraph style pagination XML**

Add a shared writer helper:

```ts
function paragraphPaginationXml(pagination?: ParagraphPagination): string {
  return pagination
    ? [
      pagination.keepNext ? "<w:keepNext/>" : "",
      pagination.keepLines ? "<w:keepLines/>" : "",
      pagination.pageBreakBefore ? "<w:pageBreakBefore/>" : "",
    ].join("")
    : "";
}
```

Use it from `paragraphPropertiesXml` and `paragraphStylePropertiesXml`.

- [x] **Step 3: Parse paragraph style pagination XML**

Add a shared parser helper:

```ts
function parseParagraphPagination(properties: XmlNode): ParagraphPagination | undefined {
  const pagination = {
    ...(properties.keepNext !== undefined ? { keepNext: true } : {}),
    ...(properties.keepLines !== undefined ? { keepLines: true } : {}),
    ...(properties.pageBreakBefore !== undefined ? { pageBreakBefore: true } : {}),
  };

  return Object.keys(pagination).length > 0 ? pagination : undefined;
}
```

Use it from `parseStyleParagraphProperties`.

- [x] **Step 4: Run targeted GREEN tests**

Run: `npm test -- tests/docx-core.test.ts -t "paragraph style pagination"`

Expected: PASS with both targeted tests passing.

### Task 3: Verify, Commit, Push

**Files:**
- Modify: `docs/superpowers/plans/2026-07-04-phase-147-paragraph-style-pagination.md`

- [x] **Step 1: Run full verification**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "paragraph style pagination"
npm test
npm run build
git diff --check
```

Expected: Targeted tests, full tests, and build pass. `git diff --check` may print existing LF/CRLF warnings but must exit 0.

- [ ] **Step 2: Commit implementation**

```bash
git add src/schema.ts src/docx-writer.ts src/docx-reader.ts tests/docx-core.test.ts docs/superpowers/plans/2026-07-04-phase-147-paragraph-style-pagination.md
git commit -m "feat: add phase 147 paragraph style pagination"
```

- [ ] **Step 3: Push branch**

```bash
git push -u origin phase-147-paragraph-style-pagination
```

- [ ] **Step 4: Record push metadata**

Append the pushed commit hash and remote branch to this plan, then commit the plan update:

```bash
git add docs/superpowers/plans/2026-07-04-phase-147-paragraph-style-pagination.md
git commit -m "docs: mark phase 147 pushed"
git push
```

## Push Record

- Pending.
