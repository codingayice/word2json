# Phase 4 Media Headers Footers Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add media and repeating page-region features that are common in real DOCX documents: inline images, headers, footers, and page number fields.

**Architecture:** Continue the explicit OOXML package writer/reader. The schema will model image blocks plus section headers/footers, while the writer creates media parts, relationships, header/footer parts, and field-code runs; the reader parses the same supported subset back into stable JSON.

**Tech Stack:** TypeScript, Vitest, JSZip, fast-xml-parser, tsx.

---

## File Structure

- `src/schema.ts`: add image blocks and section header/footer models.
- `src/docx-writer.ts`: emit media parts, drawing nodes, header/footer parts, relationships, and page field runs.
- `src/docx-reader.ts`: parse media relationships, drawing nodes, header/footer parts, and page fields back into JSON.
- `tests/docx-core.test.ts`: add XML-level and round-trip tests for Phase 4 capabilities.

## Task 1: Inline Images

**Files:**
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`
- Modify: `src/docx-reader.ts`
- Test: `tests/docx-core.test.ts`

- [ ] **Step 1: Write failing tests**

Add tests proving an image block writes `word/media/image1.png`, `a:blip r:embed`, dimensions, alt text, and an image relationship, then round-trips back to JSON.

- [ ] **Step 2: Run targeted tests**

Run: `npm test -- tests/docx-core.test.ts -t "images"`

Expected: FAIL because images are not supported yet.

- [ ] **Step 3: Implement minimal support**

Add `ImageNode`, store base64 image data in JSON, emit an inline drawing with EMU dimensions, create media file and relationship, then parse the supported writer shape back to JSON.

- [ ] **Step 4: Run targeted tests**

Run: `npm test -- tests/docx-core.test.ts -t "images"`

Expected: PASS.

## Task 2: Headers and Footers

**Files:**
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`
- Modify: `src/docx-reader.ts`
- Test: `tests/docx-core.test.ts`

- [ ] **Step 1: Write failing tests**

Add tests proving section headers and footers write `word/header1.xml`, `word/footer1.xml`, section references, relationships, and round-trip paragraph content.

- [ ] **Step 2: Run targeted tests**

Run: `npm test -- tests/docx-core.test.ts -t "headers and footers"`

Expected: FAIL because header/footer parts are not supported yet.

- [ ] **Step 3: Implement minimal support**

Add `SectionNode.headers.default` and `SectionNode.footers.default`, emit header/footer relationship entries and section references, then parse parts back into section JSON.

- [ ] **Step 4: Run targeted tests**

Run: `npm test -- tests/docx-core.test.ts -t "headers and footers"`

Expected: PASS.

## Task 3: Page Number Fields

**Files:**
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`
- Modify: `src/docx-reader.ts`
- Test: `tests/docx-core.test.ts`

- [ ] **Step 1: Write failing tests**

Add tests proving footer runs can represent `PAGE` and `NUMPAGES` fields and round-trip as structured JSON.

- [ ] **Step 2: Run targeted tests**

Run: `npm test -- tests/docx-core.test.ts -t "page fields"`

Expected: FAIL because field runs are not supported yet.

- [ ] **Step 3: Implement minimal support**

Add `TextRun.field` for `page` and `numPages`, emit simple field character/instruction runs, and parse those writer-shaped fields back.

- [ ] **Step 4: Run targeted tests**

Run: `npm test -- tests/docx-core.test.ts -t "page fields"`

Expected: PASS.

## Task 4: Full Verification and Push

**Files:**
- Modify: all Phase 4 files

- [ ] **Step 1: Run full verification**

Run: `npm test`

Expected: all tests pass.

Run: `npm run build`

Expected: TypeScript build exits with code 0.

- [ ] **Step 2: Commit**

```bash
git add .
git commit -m "feat: add phase 4 media headers footers"
```

- [ ] **Step 3: Push**

```bash
git push -u origin phase-4-media-headers-footers
```

Expected: branch `phase-4-media-headers-footers` is pushed to `https://github.com/codingayice/word2json.git`.

## Self-Review

- Spec coverage: This plan covers inline images, headers, footers, and page-number fields.
- Placeholder scan: No unresolved placeholders remain.
- Type consistency: New public schema names are `ImageNode`, `SectionNode.headers`, `SectionNode.footers`, and `TextRun.field`.
