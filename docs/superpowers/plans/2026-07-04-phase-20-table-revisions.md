# Phase 20 Table Revisions Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Preserve table-level, row-level, and cell-level tracked change metadata for reviewed Word documents.

**Architecture:** Extend table, row, and cell JSON nodes with revision metadata that mirrors WordprocessingML revision markers. The writer emits `w:tblPrChange`, row `w:ins/w:del` markers inside `w:trPr`, and `w:tcPrChange`; the reader parses those markers back into JSON without changing existing table layout behavior.

**Tech Stack:** TypeScript, Vitest, JSZip, fast-xml-parser.

---

## File Structure

- `src/schema.ts`: add table property, row, and cell property revision metadata types.
- `src/docx-writer.ts`: emit `w:tblPrChange`, row-level `w:ins/w:del`, and `w:tcPrChange`.
- `src/docx-reader.ts`: parse table and cell property change metadata plus row revision metadata.
- `tests/docx-core.test.ts`: add writer and round-trip coverage for table revision metadata.

## Task 1: Table Property Change Metadata

**Files:**
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`
- Modify: `src/docx-reader.ts`
- Test: `tests/docx-core.test.ts`

- [x] **Step 1: Write failing tests**

Add writer and round-trip tests for:

```ts
{
  type: "table",
  width: 7200,
  propertyRevision: { id: 20, author: "Ada", date: "2026-07-04T05:00:00.000Z" },
  rows: [
    { cells: [{ blocks: [{ type: "paragraph", runs: [{ text: "Table changed" }] }] }] },
  ],
}
```

- [x] **Step 2: Run targeted tests**

Run: `npm test -- tests/docx-core.test.ts -t "table property change"`

Expected: FAIL because `w:tblPrChange` is not supported yet.

- [x] **Step 3: Implement minimal support**

Add `propertyRevision` to `TableNode`, emit `<w:tblPrChange ...><w:tblPr/></w:tblPrChange>` inside `w:tblPr`, and parse it back.

- [x] **Step 4: Run targeted tests**

Run: `npm test -- tests/docx-core.test.ts -t "table property change"`

Expected: PASS.

## Task 2: Table Row Insert and Delete Metadata

**Files:**
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`
- Modify: `src/docx-reader.ts`
- Test: `tests/docx-core.test.ts`

- [x] **Step 1: Write failing tests**

Add writer and round-trip tests for inserted and deleted rows:

```ts
{
  revision: { type: "insert", id: 21, author: "Lin", date: "2026-07-04T06:00:00.000Z" },
  cells: [{ blocks: [{ type: "paragraph", runs: [{ text: "Inserted row" }] }] }],
}
```

```ts
{
  revision: { type: "delete", id: 22, author: "Mira", date: "2026-07-04T07:00:00.000Z" },
  cells: [{ blocks: [{ type: "paragraph", runs: [{ text: "Deleted row" }] }] }],
}
```

- [x] **Step 2: Run targeted tests**

Run: `npm test -- tests/docx-core.test.ts -t "table row revisions"`

Expected: FAIL because row revision markers are not supported yet.

- [x] **Step 3: Implement minimal support**

Add `revision` to `TableRowNode`, emit row revision markers in `w:trPr`, and parse `w:trPr/w:ins` and `w:trPr/w:del` back into row metadata.

- [x] **Step 4: Run targeted tests**

Run: `npm test -- tests/docx-core.test.ts -t "table row revisions"`

Expected: PASS.

## Task 3: Table Cell Property Change Metadata

**Files:**
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`
- Modify: `src/docx-reader.ts`
- Test: `tests/docx-core.test.ts`

- [x] **Step 1: Write failing tests**

Add writer and round-trip tests for:

```ts
{
  shading: { fill: "D9EAF7" },
  propertyRevision: { id: 23, author: "Noor", date: "2026-07-04T08:00:00.000Z" },
  blocks: [{ type: "paragraph", runs: [{ text: "Cell changed" }] }],
}
```

- [x] **Step 2: Run targeted tests**

Run: `npm test -- tests/docx-core.test.ts -t "cell property change"`

Expected: FAIL because `w:tcPrChange` is not supported yet.

- [x] **Step 3: Implement minimal support**

Add `propertyRevision` to `TableCellNode`, emit `<w:tcPrChange ...><w:tcPr/></w:tcPrChange>` inside `w:tcPr`, and parse it back.

- [x] **Step 4: Run targeted tests**

Run: `npm test -- tests/docx-core.test.ts -t "cell property change"`

Expected: PASS.

## Task 4: Full Verification and Push

**Files:**
- Modify: all Phase 20 files

- [x] **Step 1: Run full verification**

Run: `npm test`

Expected: all tests pass.

Run: `npm run build`

Expected: TypeScript build exits with code 0.

- [x] **Step 2: Commit**

```bash
git add .
git commit -m "feat: add phase 20 table revisions"
```

- [x] **Step 3: Push**

```bash
git push -u origin phase-20-table-revisions
```

Expected: branch `phase-20-table-revisions` is pushed to `https://github.com/codingayice/word2json.git`.

## Self-Review

- Spec coverage: This plan covers table property revision metadata, row insert/delete revision metadata, and cell property revision metadata.
- Placeholder scan: No unresolved placeholders remain.
- Type consistency: row revisions reuse `RunRevision`; table and cell property revisions reuse paragraph property revision shape.
