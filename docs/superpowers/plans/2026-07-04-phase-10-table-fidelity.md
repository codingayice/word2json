# Phase 10 Table Fidelity Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Preserve common high-fidelity Word table layout properties including explicit grids, row heights, vertical merges, cell vertical alignment, shading, and cell margins.

**Architecture:** Extend table, row, and cell JSON nodes with a small set of OpenXML-shaped layout properties. The writer emits these properties into `w:tblGrid`, `w:trPr`, and `w:tcPr`; the reader parses the supported subset back into JSON.

**Tech Stack:** TypeScript, Vitest, JSZip, fast-xml-parser.

---

## File Structure

- `src/schema.ts`: add `TableNode.grid`, row height, cell vertical merge/alignment/shading/margins.
- `src/docx-writer.ts`: emit `w:tblGrid`, `w:trPr`, `w:vMerge`, `w:vAlign`, `w:shd`, and `w:tcMar`.
- `src/docx-reader.ts`: parse the same table subset back into JSON.
- `tests/docx-core.test.ts`: add XML and round-trip tests.

## Task 1: Table Grid and Row Height

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
  grid: [2400, 3600],
  rows: [{
    height: { value: 480, rule: "exact" },
    cells: [
      { blocks: [{ type: "paragraph", runs: [{ text: "A" }] }] },
      { blocks: [{ type: "paragraph", runs: [{ text: "B" }] }] },
    ],
  }],
}
```

- [x] **Step 2: Run targeted tests**

Run: `npm test -- tests/docx-core.test.ts -t "table grid and row height"`

Expected: FAIL because grid and row height are not supported yet.

- [x] **Step 3: Implement minimal support**

Add `TableNode.grid`, `TableRowNode.height`, writer XML for `w:tblGrid/w:gridCol` and `w:trPr/w:trHeight`, and reader parsing.

- [x] **Step 4: Run targeted tests**

Run: `npm test -- tests/docx-core.test.ts -t "table grid and row height"`

Expected: PASS.

## Task 2: Cell Vertical Merge and Alignment

**Files:**
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`
- Modify: `src/docx-reader.ts`
- Test: `tests/docx-core.test.ts`

- [x] **Step 1: Write failing tests**

Add writer and round-trip tests for `verticalMerge: "restart" | "continue"` and `verticalAlignment: "center"` in cell properties.

- [x] **Step 2: Run targeted tests**

Run: `npm test -- tests/docx-core.test.ts -t "cell vertical merge and alignment"`

Expected: FAIL because those cell properties are not supported yet.

- [x] **Step 3: Implement minimal support**

Add cell schema properties, emit `w:vMerge` and `w:vAlign`, and parse them back.

- [x] **Step 4: Run targeted tests**

Run: `npm test -- tests/docx-core.test.ts -t "cell vertical merge and alignment"`

Expected: PASS.

## Task 3: Cell Shading and Margins

**Files:**
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`
- Modify: `src/docx-reader.ts`
- Test: `tests/docx-core.test.ts`

- [x] **Step 1: Write failing tests**

Add writer and round-trip tests for:

```ts
shading: { fill: "D9EAF7" },
margins: { top: 120, right: 180, bottom: 120, left: 180 }
```

- [x] **Step 2: Run targeted tests**

Run: `npm test -- tests/docx-core.test.ts -t "cell shading and margins"`

Expected: FAIL because shading and cell margins are not supported yet.

- [x] **Step 3: Implement minimal support**

Add cell schema properties, emit `w:shd` and `w:tcMar`, and parse them back.

- [x] **Step 4: Run targeted tests**

Run: `npm test -- tests/docx-core.test.ts -t "cell shading and margins"`

Expected: PASS.

## Task 4: Full Verification and Push

**Files:**
- Modify: all Phase 10 files

- [x] **Step 1: Run full verification**

Run: `npm test`

Expected: all tests pass.

Run: `npm run build`

Expected: TypeScript build exits with code 0.

- [ ] **Step 2: Commit**

```bash
git add .
git commit -m "feat: add phase 10 table fidelity"
```

- [ ] **Step 3: Push**

```bash
git push -u origin phase-10-table-fidelity
```

Expected: branch `phase-10-table-fidelity` is pushed to `https://github.com/codingayice/word2json.git`.

## Self-Review

- Spec coverage: This plan covers explicit table grid, row height, vertical merge, vertical alignment, shading, and margins.
- Placeholder scan: No unresolved placeholders remain.
- Type consistency: New schema names are `TableNode.grid`, `TableRowNode.height`, `TableCellNode.verticalMerge`, `TableCellNode.verticalAlignment`, `TableCellNode.shading`, and `TableCellNode.margins`.
