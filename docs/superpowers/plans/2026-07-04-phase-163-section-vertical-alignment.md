# Phase 163 Section Vertical Alignment Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add high-fidelity JSON and DOCX round-trip support for section-level vertical page alignment.

**Architecture:** Extend `SectionNode` with `verticalAlignment`, serialize it as `<w:vAlign>` inside section properties, and parse it back from `document.xml`. This follows the existing section page-numbering, line-numbering, note-property, document-grid, and columns patterns.

**Tech Stack:** TypeScript, JSZip, Vitest, OOXML WordprocessingML.

---

### Task 1: Add Failing Section Vertical Alignment Tests

**Files:**
- Modify: `tests/docx-core.test.ts`

- [x] **Step 1: Write the failing writer test**

Add a test near the existing section document grid and note property writer tests:

```ts
it("writes section vertical alignment", async () => {
  const document = {
    version: "1.0" as const,
    sections: [
      {
        verticalAlignment: "center" as const,
        blocks: [{ type: "paragraph" as const, runs: [{ text: "Centered" }] }],
      },
    ],
  };

  const buffer = await buildDocx(document);
  const zip = await JSZip.loadAsync(buffer);
  const xml = await zip.file("word/document.xml")!.async("string");

  expect(xml).toContain('<w:vAlign w:val="center"/>');
});
```

- [x] **Step 2: Write the failing round-trip test**

Add a test near the existing section document grid and note property round-trip tests:

```ts
it("round-trips section vertical alignment", async () => {
  const source = {
    version: "1.0" as const,
    sections: [
      {
        verticalAlignment: "both" as const,
        blocks: [{ type: "paragraph" as const, runs: [{ text: "Distributed" }] }],
      },
    ],
  };

  const docx = await buildDocx(source);
  const parsed = await parseDocx(docx);

  expect(parsed).toEqual(source);
});
```

- [x] **Step 3: Run RED verification**

Run: `npm test -- tests/docx-core.test.ts -t "section vertical alignment"`

Expected: the writer and round-trip tests fail because section vertical alignment is not serialized or parsed yet.

### Task 2: Implement Section Vertical Alignment Schema

**Files:**
- Modify: `src/schema.ts`

- [x] **Step 1: Add the section field**

Add to `SectionNode`:

```ts
verticalAlignment?: SectionVerticalAlignment;
```

- [x] **Step 2: Add the alignment type**

Add near the existing section types:

```ts
export type SectionVerticalAlignment = "top" | "center" | "both" | "bottom";
```

### Task 3: Serialize Section Vertical Alignment

**Files:**
- Modify: `src/docx-writer.ts`

- [x] **Step 1: Emit vertical alignment from section properties**

In `sectionPropertiesXml`, compute:

```ts
const verticalAlignment = section.verticalAlignment
  ? `<w:vAlign w:val="${section.verticalAlignment}"/>`
  : "";
```

Append it inside `<w:sectPr>` near the other section layout properties.

### Task 4: Parse Section Vertical Alignment

**Files:**
- Modify: `src/docx-reader.ts`

- [x] **Step 1: Parse vertical alignment in both section paths**

In the no-explicit-section path and per-section path, compute:

```ts
const verticalAlignment = parseSectionVerticalAlignment(sectPrValue);
```

Include it in the returned section object only when present.

- [x] **Step 2: Add parser helper**

Add near `parseDocumentGrid`:

```ts
function parseSectionVerticalAlignment(sectionPropertiesValue: unknown): SectionNode["verticalAlignment"] | undefined {
  const verticalAlignment = asObject(asObject(sectionPropertiesValue).vAlign);

  return typeof verticalAlignment.val === "string"
    ? verticalAlignment.val as NonNullable<SectionNode["verticalAlignment"]>
    : undefined;
}
```

### Task 5: Verify, Commit, Push

**Files:**
- Modify: `docs/superpowers/plans/2026-07-04-phase-163-section-vertical-alignment.md`

- [x] **Step 1: Run targeted GREEN verification**

Run: `npm test -- tests/docx-core.test.ts -t "section vertical alignment"`

Expected: both tests pass.

- [x] **Step 2: Run full verification**

Run:

```bash
npm test
npm run build
git diff --check
```

Expected: tests and build exit 0; `git diff --check` exits 0. CRLF warnings are acceptable only when the command still exits 0.

- [ ] **Step 3: Commit feature**

Run:

```bash
git add src/schema.ts src/docx-writer.ts src/docx-reader.ts tests/docx-core.test.ts docs/superpowers/plans/2026-07-04-phase-163-section-vertical-alignment.md
git commit -m "feat: add phase 163 section vertical alignment"
```

- [ ] **Step 4: Push feature branch**

Run:

```bash
git push -u origin phase-163-section-vertical-alignment
```

- [ ] **Step 5: Record push and commit docs**

Add a Push Record section with the branch, commits, verification commands, and PR URL:

```text
https://github.com/codingayice/word2json/pull/new/phase-163-section-vertical-alignment
```

Then run:

```bash
git add docs/superpowers/plans/2026-07-04-phase-163-section-vertical-alignment.md
git commit -m "docs: mark phase 163 pushed"
git push
```
