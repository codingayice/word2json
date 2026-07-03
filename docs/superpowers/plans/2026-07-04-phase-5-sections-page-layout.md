# Phase 5 Sections Page Layout Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Support multi-section DOCX layout features: section breaks, per-section page settings, columns, and paragraph pagination controls.

**Architecture:** Keep the explicit OOXML writer/reader and extend the supported writer-shaped subset. The writer will emit intermediate section properties in paragraph properties and final section properties in the document body; the reader will parse that structure back into ordered JSON sections.

**Tech Stack:** TypeScript, Vitest, JSZip, fast-xml-parser, tsx.

---

## File Structure

- `src/schema.ts`: add section break metadata, column settings, and paragraph pagination controls.
- `src/docx-writer.ts`: write multiple section properties, `w:type`, `w:cols`, and paragraph pagination properties.
- `src/docx-reader.ts`: parse writer-shaped multiple sections and layout controls back into JSON.
- `tests/docx-core.test.ts`: add XML and round-trip tests.

## Task 1: Multiple Sections and Section Breaks

**Files:**
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`
- Modify: `src/docx-reader.ts`
- Test: `tests/docx-core.test.ts`

- [ ] **Step 1: Write failing tests**

Add tests proving two sections write two `w:sectPr` blocks with independent page settings and section break type, then round-trip as two sections.

- [ ] **Step 2: Run targeted tests**

Run: `npm test -- tests/docx-core.test.ts -t "multiple sections"`

Expected: FAIL because the writer currently flattens all sections into one body.

- [ ] **Step 3: Implement minimal support**

Add `SectionNode.breakType`, write intermediate section properties inside an empty paragraph `w:pPr/w:sectPr`, keep the last section property in body-level `w:sectPr`, and parse the writer-shaped structure back into ordered sections.

- [ ] **Step 4: Run targeted tests**

Run: `npm test -- tests/docx-core.test.ts -t "multiple sections"`

Expected: PASS.

## Task 2: Columns

**Files:**
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`
- Modify: `src/docx-reader.ts`
- Test: `tests/docx-core.test.ts`

- [ ] **Step 1: Write failing tests**

Add tests proving section `columns` write `w:cols` and round-trip.

- [ ] **Step 2: Run targeted tests**

Run: `npm test -- tests/docx-core.test.ts -t "columns"`

Expected: FAIL because column settings are not supported yet.

- [ ] **Step 3: Implement minimal support**

Add `SectionNode.columns`, emit `w:cols w:num w:space`, and parse the values back.

- [ ] **Step 4: Run targeted tests**

Run: `npm test -- tests/docx-core.test.ts -t "columns"`

Expected: PASS.

## Task 3: Paragraph Pagination Controls

**Files:**
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`
- Modify: `src/docx-reader.ts`
- Test: `tests/docx-core.test.ts`

- [ ] **Step 1: Write failing tests**

Add tests proving `keepNext`, `keepLines`, and `pageBreakBefore` emit paragraph properties and round-trip.

- [ ] **Step 2: Run targeted tests**

Run: `npm test -- tests/docx-core.test.ts -t "pagination controls"`

Expected: FAIL because paragraph pagination controls are not supported yet.

- [ ] **Step 3: Implement minimal support**

Add `ParagraphNode.pagination`, emit `w:keepNext`, `w:keepLines`, `w:pageBreakBefore`, and parse them back.

- [ ] **Step 4: Run targeted tests**

Run: `npm test -- tests/docx-core.test.ts -t "pagination controls"`

Expected: PASS.

## Task 4: Full Verification and Push

**Files:**
- Modify: all Phase 5 files

- [ ] **Step 1: Run full verification**

Run: `npm test`

Expected: all tests pass.

Run: `npm run build`

Expected: TypeScript build exits with code 0.

- [ ] **Step 2: Commit**

```bash
git add .
git commit -m "feat: add phase 5 sections page layout"
```

- [ ] **Step 3: Push**

```bash
git push -u origin phase-5-sections-page-layout
```

Expected: branch `phase-5-sections-page-layout` is pushed to `https://github.com/codingayice/word2json.git`.

## Self-Review

- Spec coverage: This plan covers multiple sections, section break types, columns, and paragraph pagination controls.
- Placeholder scan: No unresolved placeholders remain.
- Type consistency: New public schema names are `SectionNode.breakType`, `SectionNode.columns`, and `ParagraphNode.pagination`.
