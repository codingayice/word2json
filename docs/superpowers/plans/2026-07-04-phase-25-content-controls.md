# Phase 25 Content Controls Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Preserve basic Word structured document tags (SDT/content controls) for template and form documents.

**Architecture:** Add optional SDT metadata to paragraph nodes and text runs. The writer wraps paragraphs or runs in `w:sdt` with `w:sdtPr` metadata and `w:sdtContent`; the reader unwraps block-level and run-level SDTs and restores the metadata onto JSON nodes.

**Tech Stack:** TypeScript, Vitest, JSZip, fast-xml-parser.

---

## File Structure

- `src/schema.ts`: add `ContentControl` and optional `contentControl` fields.
- `src/docx-writer.ts`: emit block-level and run-level `w:sdt`.
- `src/docx-reader.ts`: parse block-level and run-level `w:sdt`.
- `tests/docx-core.test.ts`: add writer and round-trip tests for content controls.

## Task 1: Block-Level Paragraph Content Controls

**Files:**
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`
- Modify: `src/docx-reader.ts`
- Test: `tests/docx-core.test.ts`

- [x] **Step 1: Write failing tests**

Add writer and round-trip tests for:

```ts
{
  type: "paragraph",
  contentControl: { alias: "Customer Name", tag: "customer.name", lock: "sdtContentLocked" },
  runs: [{ text: "Acme Inc." }],
}
```

- [x] **Step 2: Run targeted tests**

Run: `npm test -- tests/docx-core.test.ts -t "block content controls"`

Expected: FAIL because block-level `w:sdt` is not supported yet.

- [x] **Step 3: Implement minimal support**

Add schema metadata, emit paragraph wrappers as `w:sdt/w:sdtPr/w:sdtContent`, extend block extraction to treat `w:sdt` as a block, and parse the contained paragraph metadata back.

- [x] **Step 4: Run targeted tests**

Run: `npm test -- tests/docx-core.test.ts -t "block content controls"`

Expected: PASS.

## Task 2: Run-Level Content Controls

**Files:**
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`
- Modify: `src/docx-reader.ts`
- Test: `tests/docx-core.test.ts`

- [x] **Step 1: Write failing tests**

Add writer and round-trip tests for:

```ts
{
  type: "paragraph",
  runs: [
    { text: "INV-001", contentControl: { alias: "Invoice Number", tag: "invoice.number" } },
  ],
}
```

- [x] **Step 2: Run targeted tests**

Run: `npm test -- tests/docx-core.test.ts -t "run content controls"`

Expected: FAIL because run-level `w:sdt` is not supported yet.

- [x] **Step 3: Implement minimal support**

Wrap run XML in `w:sdt`, parse paragraph-level `sdt` run containers, and restore the run metadata.

- [x] **Step 4: Run targeted tests**

Run: `npm test -- tests/docx-core.test.ts -t "run content controls"`

Expected: PASS.

## Task 3: Full Verification and Push

**Files:**
- Modify: all Phase 25 files

- [x] **Step 1: Run full verification**

Run: `npm test`

Expected: all tests pass.

Run: `npm run build`

Expected: TypeScript build exits with code 0.

- [x] **Step 2: Commit**

```bash
git add .
git commit -m "feat: add phase 25 content controls"
```

- [x] **Step 3: Push**

```bash
git push -u origin phase-25-content-controls
```

Expected: branch `phase-25-content-controls` is pushed to `https://github.com/codingayice/word2json.git`.

## Self-Review

- Spec coverage: This plan covers paragraph and run content controls with alias, tag, and lock metadata.
- Placeholder scan: No unresolved placeholders remain.
- Type consistency: both paragraph and run fields use the same `ContentControl` type.
