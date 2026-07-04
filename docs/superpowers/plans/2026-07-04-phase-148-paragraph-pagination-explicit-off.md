# Phase 148 Paragraph Pagination Explicit Off Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Preserve explicit off values for paragraph pagination controls in both paragraph nodes and paragraph style definitions.

**Architecture:** `ParagraphPagination` already stores optional booleans, which can represent `false` without a schema change. This phase updates the writer to emit `w:val="0"` for explicit false values and updates the reader to parse OOXML on/off values back to booleans for paragraph and style pagination.

**Tech Stack:** TypeScript, JSZip-based DOCX package tests, existing DOCX reader/writer XML helpers, Vitest.

---

### Task 1: Add RED Tests For Explicit Off Pagination

**Files:**
- Modify: `tests/docx-core.test.ts`

- [x] **Step 1: Write writer test before implementation**

Add a writer test near the existing paragraph pagination writer tests:

```ts
  it("writes explicit off paragraph pagination controls", async () => {
    const document = {
      version: "1.0" as const,
      sections: [{
        blocks: [{
          type: "paragraph" as const,
          pagination: { keepNext: false, keepLines: false, pageBreakBefore: false },
          runs: [{ text: "Loose paragraph" }],
        }],
      }],
    };

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<w:pPr><w:keepNext w:val="0"/><w:keepLines w:val="0"/><w:pageBreakBefore w:val="0"/></w:pPr>');
  });
```

- [x] **Step 2: Write style writer test before implementation**

Add a writer test near the existing paragraph style pagination writer test:

```ts
  it("writes explicit off paragraph style pagination properties", async () => {
    const document = {
      version: "1.0" as const,
      styles: {
        paragraph: [{
          id: "LooseHeading",
          name: "Loose Heading",
          paragraph: {
            pagination: { keepNext: false, keepLines: false, pageBreakBefore: false },
          },
        }],
      },
      sections: [{ blocks: [{ type: "paragraph" as const, styleId: "LooseHeading", runs: [{ text: "Heading" }] }] }],
    };

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const styles = await zip.file("word/styles.xml")!.async("string");

    expect(styles).toContain('<w:style w:type="paragraph" w:styleId="LooseHeading"><w:name w:val="Loose Heading"/><w:pPr><w:keepNext w:val="0"/><w:keepLines w:val="0"/><w:pageBreakBefore w:val="0"/></w:pPr></w:style>');
  });
```

- [x] **Step 3: Write reader round-trip test before implementation**

Add a reader test near the existing pagination round-trip tests:

```ts
  it("round-trips explicit off paragraph pagination controls", async () => {
    const source = {
      version: "1.0" as const,
      styles: {
        paragraph: [{
          id: "LooseHeading",
          name: "Loose Heading",
          paragraph: {
            pagination: { keepNext: false, keepLines: false, pageBreakBefore: false },
          },
        }],
      },
      sections: [{
        blocks: [{
          type: "paragraph" as const,
          styleId: "LooseHeading",
          pagination: { keepNext: false, keepLines: false, pageBreakBefore: false },
          runs: [{ text: "Loose paragraph" }],
        }],
      }],
    };

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });
```

- [x] **Step 4: Run RED tests**

Run: `npm test -- tests/docx-core.test.ts -t "explicit off paragraph pagination"`

Expected: FAIL because false values are currently omitted instead of written as `w:val="0"`, and the reader does not parse off values back to `false`.

### Task 2: Implement Explicit Off Pagination

**Files:**
- Modify: `src/docx-writer.ts`
- Modify: `src/docx-reader.ts`

- [x] **Step 1: Emit explicit false pagination XML**

Replace `paragraphPaginationXml` with an implementation that distinguishes `undefined` from `false`:

```ts
function paragraphPaginationToggleXml(name: string, value: boolean | undefined): string {
  if (value === undefined) return "";
  return value ? `<w:${name}/>` : `<w:${name} w:val="0"/>`;
}

function paragraphPaginationXml(pagination?: ParagraphNode["pagination"]): string {
  return pagination
    ? [
      paragraphPaginationToggleXml("keepNext", pagination.keepNext),
      paragraphPaginationToggleXml("keepLines", pagination.keepLines),
      paragraphPaginationToggleXml("pageBreakBefore", pagination.pageBreakBefore),
    ].join("")
    : "";
}
```

- [x] **Step 2: Parse explicit false pagination XML**

Add a helper to parse OOXML on/off attributes for pagination elements:

```ts
function parsePaginationToggle(value: unknown): boolean | undefined {
  if (value === undefined) return undefined;

  const node = asObject(value);
  const rawValue = node.val;
  return rawValue === "0" || rawValue === "false" || rawValue === "off" ? false : true;
}
```

Use it from both `parseParagraphPagination` and `parsePagination`:

```ts
function parseParagraphPagination(properties: XmlNode): NonNullable<ParagraphNode["pagination"]> | undefined {
  const keepNext = parsePaginationToggle(properties.keepNext);
  const keepLines = parsePaginationToggle(properties.keepLines);
  const pageBreakBefore = parsePaginationToggle(properties.pageBreakBefore);
  const pagination = {
    ...(keepNext !== undefined ? { keepNext } : {}),
    ...(keepLines !== undefined ? { keepLines } : {}),
    ...(pageBreakBefore !== undefined ? { pageBreakBefore } : {}),
  };

  return Object.keys(pagination).length > 0 ? pagination : undefined;
}
```

- [x] **Step 3: Run targeted GREEN tests**

Run: `npm test -- tests/docx-core.test.ts -t "explicit off paragraph pagination"`

Expected: PASS with three targeted tests passing.

### Task 3: Verify, Commit, Push

**Files:**
- Modify: `docs/superpowers/plans/2026-07-04-phase-148-paragraph-pagination-explicit-off.md`

- [x] **Step 1: Run full verification**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "explicit off paragraph pagination"
npm test
npm run build
git diff --check
```

Expected: Targeted tests, full tests, and build pass. `git diff --check` may print LF/CRLF warnings but must exit 0.

- [ ] **Step 2: Commit implementation**

```bash
git add src/docx-writer.ts src/docx-reader.ts tests/docx-core.test.ts docs/superpowers/plans/2026-07-04-phase-148-paragraph-pagination-explicit-off.md
git commit -m "feat: add phase 148 paragraph pagination explicit off"
```

- [ ] **Step 3: Push branch**

```bash
git push -u origin phase-148-paragraph-pagination-explicit-off
```

- [ ] **Step 4: Record push metadata**

Append the pushed commit hash and remote branch to this plan, then commit the plan update:

```bash
git add docs/superpowers/plans/2026-07-04-phase-148-paragraph-pagination-explicit-off.md
git commit -m "docs: mark phase 148 pushed"
git push
```

## Push Record

- Pending.
