# Phase 24 Cross-Paragraph Comments Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Preserve comment range boundaries that span multiple paragraphs.

**Architecture:** Add optional paragraph-level comment boundary metadata so JSON can represent `w:commentRangeStart` before a paragraph and `w:commentRangeEnd`/`w:commentReference` after a later paragraph. The writer emits the boundary markers around paragraph XML and creates the comment entry from the start boundary. The reader detects boundary markers around block XML and reconstructs the paragraph metadata.

**Tech Stack:** TypeScript, Vitest, JSZip, fast-xml-parser.

---

## File Structure

- `src/schema.ts`: add paragraph-level `commentRangeStart` and `commentRangeEnd` metadata.
- `src/docx-writer.ts`: emit paragraph-level comment boundary markers and comments.xml entries.
- `src/docx-reader.ts`: extract and parse paragraph-level comment boundary markers.
- `tests/docx-core.test.ts`: add writer and round-trip tests for cross-paragraph comment boundaries.

## Task 1: Cross-Paragraph Comment Writer

**Files:**
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`
- Test: `tests/docx-core.test.ts`

- [x] **Step 1: Write failing tests**

Add a writer test for two paragraphs:

```ts
[
  {
    type: "paragraph",
    commentRangeStart: { id: 60, author: "Ada", text: "Across paragraphs." },
    runs: [{ text: "First paragraph" }],
  },
  {
    type: "paragraph",
    commentRangeEnd: { id: 60 },
    runs: [{ text: "Second paragraph" }],
  },
]
```

Assert that `w:commentRangeStart w:id="60"` appears before the first paragraph, `w:commentRangeEnd w:id="60"` and `w:commentReference w:id="60"` appear after the second paragraph, and `comments.xml` has a single comment id 60.

- [x] **Step 2: Run targeted tests**

Run: `npm test -- tests/docx-core.test.ts -t "cross-paragraph comments"`

Expected: FAIL because paragraph-level comment boundaries are not supported yet.

- [x] **Step 3: Implement minimal writer support**

Add metadata types, emit boundary markers around paragraph XML, and register the start comment in comments.xml.

- [x] **Step 4: Run targeted tests**

Run: `npm test -- tests/docx-core.test.ts -t "cross-paragraph comments"`

Expected: writer assertions pass.

## Task 2: Cross-Paragraph Comment Reader

**Files:**
- Modify: `src/docx-reader.ts`
- Test: `tests/docx-core.test.ts`

- [x] **Step 1: Write failing tests**

Add a round-trip test for the same two paragraphs.

- [x] **Step 2: Run targeted tests**

Run: `npm test -- tests/docx-core.test.ts -t "cross-paragraph comments"`

Expected: FAIL because boundary markers are not parsed into paragraphs yet.

- [x] **Step 3: Implement minimal reader support**

Extend block extraction to include adjacent comment range markers with the paragraph block and parse them into `commentRangeStart`/`commentRangeEnd`.

- [x] **Step 4: Run targeted tests**

Run: `npm test -- tests/docx-core.test.ts -t "cross-paragraph comments"`

Expected: PASS.

## Task 3: Full Verification and Push

**Files:**
- Modify: all Phase 24 files

- [x] **Step 1: Run full verification**

Run: `npm test`

Expected: all tests pass.

Run: `npm run build`

Expected: TypeScript build exits with code 0.

- [x] **Step 2: Commit**

```bash
git add .
git commit -m "feat: add phase 24 cross-paragraph comments"
```

- [x] **Step 3: Push**

```bash
git push -u origin phase-24-cross-paragraph-comments
```

Expected: branch `phase-24-cross-paragraph-comments` is pushed to `https://github.com/codingayice/word2json.git`.

## Self-Review

- Spec coverage: This plan covers paragraph-level start/end boundaries for comment ranges spanning multiple paragraphs.
- Placeholder scan: No unresolved placeholders remain.
- Type consistency: start boundary carries full comment metadata; end boundary carries the stable id.
