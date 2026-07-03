# Phase 7 Styles Theme Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add reusable style definitions and a basic theme part so generated DOCX files can preserve document/template styling more faithfully.

**Architecture:** Extend `DocumentJson` with `styles` and `theme`. The writer emits `word/styles.xml` from JSON definitions and `word/theme/theme1.xml` when theme data exists; the reader parses the supported writer-shaped style/theme subset back into JSON.

**Tech Stack:** TypeScript, Vitest, JSZip, fast-xml-parser, tsx.

---

## File Structure

- `src/schema.ts`: add document styles, run style ids, table style ids, and theme model.
- `src/docx-writer.ts`: emit custom paragraph/character/table styles and theme relationship/content type.
- `src/docx-reader.ts`: parse styles and theme back into JSON and read style ids from content nodes.
- `tests/docx-core.test.ts`: add XML-level and round-trip tests.

## Task 1: Paragraph Style Definitions

**Files:**
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`
- Modify: `src/docx-reader.ts`
- Test: `tests/docx-core.test.ts`

- [x] **Step 1: Write failing tests**

Add tests proving custom paragraph styles write `w:style`, `w:basedOn`, `w:next`, paragraph `w:pStyle`, and round-trip.

- [x] **Step 2: Run targeted tests**

Run: `npm test -- tests/docx-core.test.ts -t "paragraph style definitions"`

Expected: FAIL because custom styles are not supported yet.

- [x] **Step 3: Implement minimal support**

Add `DocumentJson.styles.paragraph`, map `ParagraphNode.styleId` to `w:pStyle`, emit style definitions, and parse definitions plus paragraph style ids back.

- [x] **Step 4: Run targeted tests**

Run: `npm test -- tests/docx-core.test.ts -t "paragraph style definitions"`

Expected: PASS.

## Task 2: Character and Table Styles

**Files:**
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`
- Modify: `src/docx-reader.ts`
- Test: `tests/docx-core.test.ts`

- [x] **Step 1: Write failing tests**

Add tests proving character style definitions, run `w:rStyle`, table style definitions, table `w:tblStyle`, and round-trip.

- [x] **Step 2: Run targeted tests**

Run: `npm test -- tests/docx-core.test.ts -t "character and table styles"`

Expected: FAIL because character/table styles are not supported yet.

- [x] **Step 3: Implement minimal support**

Add `DocumentJson.styles.character`, `DocumentJson.styles.table`, `TextRun.styleId`, and `TableNode.styleId`, write and parse the supported style subset.

- [x] **Step 4: Run targeted tests**

Run: `npm test -- tests/docx-core.test.ts -t "character and table styles"`

Expected: PASS.

## Task 3: Theme Font and Color Part

**Files:**
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`
- Modify: `src/docx-reader.ts`
- Test: `tests/docx-core.test.ts`

- [x] **Step 1: Write failing tests**

Add tests proving `word/theme/theme1.xml`, theme relationship/content type, and round-trip of major/minor fonts and accent color.

- [x] **Step 2: Run targeted tests**

Run: `npm test -- tests/docx-core.test.ts -t "theme part"`

Expected: FAIL because theme parts are not supported yet.

- [x] **Step 3: Implement minimal support**

Add `DocumentJson.theme`, emit `theme1.xml`, add document relationship and content type override, and parse the supported theme subset back.

- [x] **Step 4: Run targeted tests**

Run: `npm test -- tests/docx-core.test.ts -t "theme part"`

Expected: PASS.

## Task 4: Full Verification and Push

**Files:**
- Modify: all Phase 7 files

- [x] **Step 1: Run full verification**

Run: `npm test`

Expected: all tests pass.

Run: `npm run build`

Expected: TypeScript build exits with code 0.

- [ ] **Step 2: Commit**

```bash
git add .
git commit -m "feat: add phase 7 styles theme"
```

- [ ] **Step 3: Push**

```bash
git push -u origin phase-7-styles-theme
```

Expected: branch `phase-7-styles-theme` is pushed to `https://github.com/codingayice/word2json.git`.

## Self-Review

- Spec coverage: This plan covers paragraph styles, character styles, table styles, and a basic theme part.
- Placeholder scan: No unresolved placeholders remain.
- Type consistency: New schema names are `DocumentJson.styles`, `ParagraphNode.styleId`, `TextRun.styleId`, `TableNode.styleId`, and `DocumentJson.theme`.
