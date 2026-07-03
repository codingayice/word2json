# Phase 22 Comment IDs Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Preserve stable Word comment ids in JSON so later phases can represent multi-run and cross-block comment ranges.

**Architecture:** Extend `Comment` with an optional `id`. The writer respects an explicit id when emitting comment ranges and `comments.xml`, while still assigning ids for comments without one. The reader includes the Word comment id when attaching comment metadata to runs.

**Tech Stack:** TypeScript, Vitest, JSZip, fast-xml-parser.

---

## File Structure

- `src/schema.ts`: add optional `id` to `Comment`.
- `src/docx-writer.ts`: preserve explicit comment ids and avoid duplicate comment entries for the same id.
- `src/docx-reader.ts`: parse comment ids into run comment metadata.
- `tests/docx-core.test.ts`: add writer and round-trip tests for comment id preservation.

## Task 1: Comment ID Writer Fidelity

**Files:**
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`
- Test: `tests/docx-core.test.ts`

- [x] **Step 1: Write failing tests**

Add a writer test for:

```ts
{
  text: "Reviewed",
  comment: {
    id: 42,
    author: "Ada",
    initials: "AL",
    date: "2026-07-04T13:00:00.000Z",
    text: "Stable comment id.",
  },
}
```

- [x] **Step 2: Run targeted tests**

Run: `npm test -- tests/docx-core.test.ts -t "comment ids"`

Expected: FAIL because explicit comment ids are not supported yet.

- [x] **Step 3: Implement minimal writer support**

Emit `w:commentRangeStart`, `w:commentRangeEnd`, `w:commentReference`, and `comments.xml` using the explicit id. Keep auto-assigned ids for comments without an explicit id.

- [x] **Step 4: Run targeted tests**

Run: `npm test -- tests/docx-core.test.ts -t "comment ids"`

Expected: PASS for writer id assertions.

## Task 2: Comment ID Round-Trip

**Files:**
- Modify: `src/docx-reader.ts`
- Test: `tests/docx-core.test.ts`

- [x] **Step 1: Write failing tests**

Add a round-trip test for the same comment with `id: 42`.

- [x] **Step 2: Run targeted tests**

Run: `npm test -- tests/docx-core.test.ts -t "comment ids"`

Expected: FAIL because parsed comments do not include ids yet.

- [x] **Step 3: Implement reader support**

Parse `w:comment/@w:id` into the `Comment` object and keep existing comment range attachment behavior.

- [x] **Step 4: Run targeted tests**

Run: `npm test -- tests/docx-core.test.ts -t "comment ids"`

Expected: PASS.

## Task 3: Full Verification and Push

**Files:**
- Modify: all Phase 22 files

- [x] **Step 1: Run full verification**

Run: `npm test`

Expected: all tests pass.

Run: `npm run build`

Expected: TypeScript build exits with code 0.

- [x] **Step 2: Commit**

```bash
git add .
git commit -m "feat: add phase 22 comment ids"
```

- [x] **Step 3: Push**

```bash
git push -u origin phase-22-comment-ids
```

Expected: branch `phase-22-comment-ids` is pushed to `https://github.com/codingayice/word2json.git`.

## Self-Review

- Spec coverage: This plan covers stable comment id writing and parsing.
- Placeholder scan: No unresolved placeholders remain.
- Type consistency: comment ids are optional numbers on the existing `Comment` type.
