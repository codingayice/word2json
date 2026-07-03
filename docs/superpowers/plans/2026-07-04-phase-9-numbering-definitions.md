# Phase 9 Numbering Definitions Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Preserve custom Word numbering definitions so ordered and bullet lists can round-trip with explicit `abstractNum`, `num`, level formatting, text patterns, and indentation.

**Architecture:** Extend `DocumentJson` with a `numbering` model and let paragraphs reference custom list instances by `list.numberingId`. The writer emits custom definitions into `word/numbering.xml` alongside the built-in defaults, and the reader parses non-default definitions plus paragraph num IDs back into JSON.

**Tech Stack:** TypeScript, Vitest, JSZip, fast-xml-parser.

---

## File Structure

- `src/schema.ts`: add numbering definition types and `ListSettings.numberingId`.
- `src/docx-writer.ts`: write custom abstract numbering definitions and num instances.
- `src/docx-reader.ts`: parse custom numbering definitions and preserve paragraph list `numberingId`.
- `tests/docx-core.test.ts`: add XML and round-trip tests.

## Task 1: Custom Numbering Writer

**Files:**
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`
- Test: `tests/docx-core.test.ts`

- [x] **Step 1: Write failing tests**

Add a writer test with:

```ts
numbering: {
  abstractNums: [{
    id: 10,
    levels: [
      { level: 0, format: "decimal", text: "%1.", start: 1, left: 720, hanging: 360 },
      { level: 1, format: "lowerLetter", text: "%2)", start: 1, left: 1440, hanging: 360 },
    ],
  }],
  nums: [{ id: 10, abstractId: 10 }],
}
```

The paragraph uses `list: { type: "ordered", level: 1, numberingId: 10 }`.

- [x] **Step 2: Run targeted tests**

Run: `npm test -- tests/docx-core.test.ts -t "custom numbering definitions"`

Expected: FAIL because custom numbering definitions are not supported yet.

- [x] **Step 3: Implement minimal support**

Add numbering schema types, emit custom `w:abstractNum`, `w:lvl`, `w:numFmt`, `w:lvlText`, `w:pPr/w:ind`, and `w:num` entries. Use `list.numberingId` when writing paragraph `w:numId`.

- [x] **Step 4: Run targeted tests**

Run: `npm test -- tests/docx-core.test.ts -t "custom numbering definitions"`

Expected: PASS for writer XML assertions.

## Task 2: Custom Numbering Reader

**Files:**
- Modify: `src/docx-reader.ts`
- Test: `tests/docx-core.test.ts`

- [x] **Step 1: Write failing tests**

Add a reader round-trip test for the same JSON from Task 1.

- [x] **Step 2: Run targeted tests**

Run: `npm test -- tests/docx-core.test.ts -t "round-trips custom numbering definitions"`

Expected: FAIL because custom numbering definitions and paragraph `numberingId` are not parsed yet.

- [x] **Step 3: Implement minimal support**

Parse `word/numbering.xml`, ignore built-in defaults `abstractNumId` 1/2 and `numId` 1/2, return `DocumentJson.numbering`, and parse custom paragraph `w:numId` into `list.numberingId`.

- [x] **Step 4: Run targeted tests**

Run: `npm test -- tests/docx-core.test.ts -t "round-trips custom numbering definitions"`

Expected: PASS.

## Task 3: Full Verification and Push

**Files:**
- Modify: all Phase 9 files

- [x] **Step 1: Run full verification**

Run: `npm test`

Expected: all tests pass.

Run: `npm run build`

Expected: TypeScript build exits with code 0.

- [ ] **Step 2: Commit**

```bash
git add .
git commit -m "feat: add phase 9 numbering definitions"
```

- [ ] **Step 3: Push**

```bash
git push -u origin phase-9-numbering-definitions
```

Expected: branch `phase-9-numbering-definitions` is pushed to `https://github.com/codingayice/word2json.git`.

## Self-Review

- Spec coverage: This plan adds custom numbering definitions, num instances, level formatting, indentation, paragraph references, and round-trip parsing.
- Placeholder scan: No unresolved placeholders remain.
- Type consistency: New schema names are `DocumentNumbering`, `AbstractNumberingDefinition`, `NumberingInstance`, `NumberingLevelDefinition`, and `ListSettings.numberingId`.
