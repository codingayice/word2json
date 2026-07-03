# Phase 23 Multi-Run Comments Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Preserve a single Word comment range across multiple adjacent runs in the same paragraph.

**Architecture:** Use the stable `Comment.id` from Phase 22 to group adjacent runs that share the same comment id. The writer emits one `w:commentRangeStart` before the first run and one `w:commentRangeEnd` plus reference after the last run. The reader walks paragraph child order to attach the same comment metadata to every run inside a comment range.

**Tech Stack:** TypeScript, Vitest, JSZip, fast-xml-parser.

---

## File Structure

- `src/docx-writer.ts`: group adjacent paragraph runs with the same `comment.id` into one comment range.
- `src/docx-reader.ts`: parse comment range start/end ordering and attach comments to all runs inside a range.
- `tests/docx-core.test.ts`: add writer and round-trip tests for multi-run comment ranges.

## Task 1: Multi-Run Comment Writer

**Files:**
- Modify: `src/docx-writer.ts`
- Test: `tests/docx-core.test.ts`

- [x] **Step 1: Write failing tests**

Add a writer test for:

```ts
{
  type: "paragraph",
  runs: [
    { text: "First ", comment: { id: 50, author: "Ada", text: "One range." } },
    { text: "second", bold: true, comment: { id: 50, author: "Ada", text: "One range." } },
  ],
}
```

Assert that document XML has one `w:commentRangeStart w:id="50"`, one `w:commentRangeEnd w:id="50"`, one `w:commentReference w:id="50"`, and both run texts inside the range.

- [x] **Step 2: Run targeted tests**

Run: `npm test -- tests/docx-core.test.ts -t "multi-run comments"`

Expected: FAIL because the writer currently emits one comment range per run.

- [x] **Step 3: Implement minimal writer support**

Render paragraph runs through a grouping function that suppresses per-run comment wrapping for grouped runs, emits one range around adjacent same-id runs, and keeps existing single-run behavior.

- [x] **Step 4: Run targeted tests**

Run: `npm test -- tests/docx-core.test.ts -t "multi-run comments"`

Expected: writer assertions pass.

## Task 2: Multi-Run Comment Reader

**Files:**
- Modify: `src/docx-reader.ts`
- Test: `tests/docx-core.test.ts`

- [x] **Step 1: Write failing tests**

Add a round-trip test for the same two-run paragraph. The parsed JSON must attach the same `comment` object to both runs.

- [x] **Step 2: Run targeted tests**

Run: `npm test -- tests/docx-core.test.ts -t "multi-run comments"`

Expected: FAIL because the reader currently matches comments by text/index and does not understand range boundaries.

- [x] **Step 3: Implement minimal reader support**

Parse paragraph child XML order for `commentRangeStart`, `commentRangeEnd`, and `r` nodes, build a map from run text occurrence to active comment id, then attach comments to runs in order.

- [x] **Step 4: Run targeted tests**

Run: `npm test -- tests/docx-core.test.ts -t "multi-run comments"`

Expected: PASS.

## Task 3: Full Verification and Push

**Files:**
- Modify: all Phase 23 files

- [x] **Step 1: Run full verification**

Run: `npm test`

Expected: all tests pass.

Run: `npm run build`

Expected: TypeScript build exits with code 0.

- [x] **Step 2: Commit**

```bash
git add .
git commit -m "feat: add phase 23 multi-run comments"
```

- [x] **Step 3: Push**

```bash
git push -u origin phase-23-multi-run-comments
```

Expected: branch `phase-23-multi-run-comments` is pushed to `https://github.com/codingayice/word2json.git`.

## Self-Review

- Spec coverage: This plan covers adjacent multi-run comment ranges in one paragraph.
- Placeholder scan: No unresolved placeholders remain.
- Type consistency: comment grouping uses existing optional `Comment.id`.
