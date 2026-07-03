# Phase 2 Daily Structures Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Expand the supported high-fidelity subset from plain paragraphs to the daily document structures used in reports, contracts, notices, and business forms.

**Architecture:** Continue the explicit OOXML writer/reader approach from Phase 1. Extend the public JSON schema first, then teach the writer to emit the relevant OOXML parts and the reader to parse the same supported subset back into stable JSON.

**Tech Stack:** TypeScript, Vitest, JSZip, fast-xml-parser, tsx.

---

## File Structure

- `src/schema.ts`: extend block, section, paragraph, run, table, list, and page setting types.
- `src/docx-writer.ts`: emit styles, numbering, paragraphs, tables, list paragraphs, and section page settings.
- `src/docx-reader.ts`: parse the emitted OOXML subset back into the JSON schema.
- `tests/docx-core.test.ts`: add TDD coverage for Phase 2 round-trip and XML-level assertions.

## Task 1: Text Styles and Headings

**Files:**
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`
- Modify: `src/docx-reader.ts`
- Test: `tests/docx-core.test.ts`

- [ ] **Step 1: Write failing tests**

Add tests proving:
- heading paragraphs write `<w:pStyle w:val="Heading1"/>`
- runs write font family, half-point font size, and color
- `json -> docx -> json` preserves `style: "heading1"`, `fontFamily`, `fontSize`, and `color`

- [ ] **Step 2: Run the targeted tests**

Run: `npm test -- tests/docx-core.test.ts -t "text styles and headings"`

Expected: FAIL because schema, writer, and reader do not support these properties yet.

- [ ] **Step 3: Implement minimal support**

Add `ParagraphStyle = "normal" | "heading1" | "heading2" | "heading3"`, run fields `fontFamily`, `fontSize`, and `color`, write the corresponding `w:pStyle`, `w:rFonts`, `w:sz`, and `w:color`, then parse them back.

- [ ] **Step 4: Run the targeted tests**

Run: `npm test -- tests/docx-core.test.ts -t "text styles and headings"`

Expected: PASS.

## Task 2: Page Settings

**Files:**
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`
- Modify: `src/docx-reader.ts`
- Test: `tests/docx-core.test.ts`

- [ ] **Step 1: Write failing tests**

Add tests proving section page settings preserve paper size, orientation, and margins through generated XML and round-trip parsing.

- [ ] **Step 2: Run the targeted tests**

Run: `npm test -- tests/docx-core.test.ts -t "page settings"`

Expected: FAIL because sections have no page settings yet.

- [ ] **Step 3: Implement minimal support**

Add `SectionNode.page` with `width`, `height`, `orientation`, and margin fields. Write these into `w:sectPr/w:pgSz` and `w:pgMar`, then parse them back.

- [ ] **Step 4: Run the targeted tests**

Run: `npm test -- tests/docx-core.test.ts -t "page settings"`

Expected: PASS.

## Task 3: Tables

**Files:**
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`
- Modify: `src/docx-reader.ts`
- Test: `tests/docx-core.test.ts`

- [ ] **Step 1: Write failing tests**

Add tests proving table rows, cells, width, borders, and grid span are emitted in OOXML and parsed back.

- [ ] **Step 2: Run the targeted tests**

Run: `npm test -- tests/docx-core.test.ts -t "tables"`

Expected: FAIL because tables are not supported yet.

- [ ] **Step 3: Implement minimal support**

Add `TableNode`, `TableRowNode`, and `TableCellNode`. Write `w:tbl`, `w:tr`, `w:tc`, `w:tcW`, `w:gridSpan`, and basic single-line cell paragraphs. Parse the same subset back.

- [ ] **Step 4: Run the targeted tests**

Run: `npm test -- tests/docx-core.test.ts -t "tables"`

Expected: PASS.

## Task 4: Lists

**Files:**
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`
- Modify: `src/docx-reader.ts`
- Test: `tests/docx-core.test.ts`

- [ ] **Step 1: Write failing tests**

Add tests proving bullet and ordered list paragraphs emit `word/numbering.xml`, paragraph `w:numPr`, and round-trip list metadata.

- [ ] **Step 2: Run the targeted tests**

Run: `npm test -- tests/docx-core.test.ts -t "lists"`

Expected: FAIL because numbering is not supported yet.

- [ ] **Step 3: Implement minimal support**

Add paragraph `list` metadata with `type` and `level`. Emit `word/numbering.xml`, relationships/content-types entries, and paragraph `w:numPr`. Parse `ilvl` and `numId` back to bullet or ordered list metadata.

- [ ] **Step 4: Run the targeted tests**

Run: `npm test -- tests/docx-core.test.ts -t "lists"`

Expected: PASS.

## Task 5: Full Verification and Push

**Files:**
- Modify: all Phase 2 files

- [ ] **Step 1: Run full verification**

Run: `npm test`

Expected: all tests pass.

Run: `npm run build`

Expected: TypeScript build exits with code 0.

- [ ] **Step 2: Commit**

```bash
git add .
git commit -m "feat: add phase 2 daily document structures"
```

- [ ] **Step 3: Push**

```bash
git push -u origin phase-2-daily-structures
```

Expected: branch `phase-2-daily-structures` is pushed to `https://github.com/codingayice/word2json.git`.

## Self-Review

- Spec coverage: The plan extends the high-fidelity subset with the daily structures named after Phase 1: headings, styling, page setup, tables, and lists.
- Placeholder scan: No unresolved placeholder language remains.
- Type consistency: New schema names are used consistently across tests, writer, and reader.
