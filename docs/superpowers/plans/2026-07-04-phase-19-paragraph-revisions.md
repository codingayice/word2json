# Phase 19 Paragraph Revisions Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Preserve paragraph-level tracked changes and paragraph property change metadata in reviewed Word documents.

**Architecture:** Extend `ParagraphNode` with revision metadata and paragraph property revision metadata. The writer wraps paragraph XML with `w:ins`/`w:del` for paragraph-level revisions and emits `w:pPrChange` inside paragraph properties. The reader unwraps paragraph revision containers and parses property change metadata back into JSON.

**Tech Stack:** TypeScript, Vitest, JSZip, fast-xml-parser.

---

## File Structure

- `src/schema.ts`: add `revision` and `propertyRevision` to `ParagraphNode`.
- `src/docx-writer.ts`: emit paragraph-level `w:ins`, `w:del`, and `w:pPrChange`.
- `src/docx-reader.ts`: parse paragraph-level revision containers and `w:pPrChange`.
- `tests/docx-core.test.ts`: add XML writer and round-trip tests.

## Task 1: Paragraph Insert Revisions

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
  revision: { type: "insert", id: 10, author: "Ada", date: "2026-07-04T02:00:00.000Z" },
  runs: [{ text: "Inserted paragraph" }],
}
```

- [x] **Step 2: Run targeted tests**

Run: `npm test -- tests/docx-core.test.ts -t "paragraph insert revisions"`

Expected: FAIL because paragraph-level insert revisions are not supported yet.

- [x] **Step 3: Implement minimal support**

Add paragraph `revision` metadata, wrap paragraph XML in `w:ins`, and parse `w:ins/w:p` back into a paragraph with revision metadata.

- [x] **Step 4: Run targeted tests**

Run: `npm test -- tests/docx-core.test.ts -t "paragraph insert revisions"`

Expected: PASS.

## Task 2: Paragraph Delete Revisions

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
  revision: { type: "delete", id: 11, author: "Lin", date: "2026-07-04T03:00:00.000Z" },
  runs: [{ text: "Deleted paragraph" }],
}
```

- [x] **Step 2: Run targeted tests**

Run: `npm test -- tests/docx-core.test.ts -t "paragraph delete revisions"`

Expected: FAIL because paragraph-level delete revisions are not supported yet.

- [x] **Step 3: Implement minimal support**

Wrap paragraph XML in `w:del` and parse `w:del/w:p` back into a paragraph with delete revision metadata.

- [x] **Step 4: Run targeted tests**

Run: `npm test -- tests/docx-core.test.ts -t "paragraph delete revisions"`

Expected: PASS.

## Task 3: Paragraph Property Change Metadata

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
  spacing: { before: 240 },
  propertyRevision: { id: 12, author: "Mira", date: "2026-07-04T04:00:00.000Z" },
  runs: [{ text: "Changed spacing" }],
}
```

- [x] **Step 2: Run targeted tests**

Run: `npm test -- tests/docx-core.test.ts -t "paragraph property change"`

Expected: FAIL because `w:pPrChange` is not supported yet.

- [x] **Step 3: Implement minimal support**

Add `propertyRevision` metadata to `ParagraphNode`, emit `w:pPrChange` inside `w:pPr`, and parse it back.

- [x] **Step 4: Run targeted tests**

Run: `npm test -- tests/docx-core.test.ts -t "paragraph property change"`

Expected: PASS.

## Task 4: Full Verification and Push

**Files:**
- Modify: all Phase 19 files

- [x] **Step 1: Run full verification**

Run: `npm test`

Expected: all tests pass.

Run: `npm run build`

Expected: TypeScript build exits with code 0.

- [x] **Step 2: Commit**

```bash
git add .
git commit -m "feat: add phase 19 paragraph revisions"
```

- [x] **Step 3: Push**

```bash
git push -u origin phase-19-paragraph-revisions
```

Expected: branch `phase-19-paragraph-revisions` is pushed to `https://github.com/codingayice/word2json.git`.

## Self-Review

- Spec coverage: This plan covers paragraph insert/delete revisions and paragraph property change metadata.
- Placeholder scan: No unresolved placeholders remain.
- Type consistency: paragraph revision metadata reuses `RunRevision`; paragraph property change metadata uses id/author/date.
