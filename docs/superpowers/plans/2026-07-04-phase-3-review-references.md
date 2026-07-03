# Phase 3 Review References Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add review and reference features that appear in real Word documents: hyperlinks, comments, bookmarks, and inline break controls.

**Architecture:** Extend the existing JSON schema and explicit OOXML writer/reader. The writer will produce the needed package parts and relationships; the reader will parse the same supported subset back into stable JSON for high-fidelity reconstruction of supported documents.

**Tech Stack:** TypeScript, Vitest, JSZip, fast-xml-parser, tsx.

---

## File Structure

- `src/schema.ts`: add inline metadata for hyperlinks, comments, bookmarks, and break runs.
- `src/docx-writer.ts`: emit hyperlink relationships, `comments.xml`, bookmark markers, and inline break elements.
- `src/docx-reader.ts`: parse the emitted OOXML subset and relationship/comment parts back into JSON.
- `tests/docx-core.test.ts`: add red/green tests for XML package output and round-trip behavior.

## Task 1: Hyperlinks

**Files:**
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`
- Modify: `src/docx-reader.ts`
- Test: `tests/docx-core.test.ts`

- [ ] **Step 1: Write failing tests**

Add tests proving a run with `link: { url: "https://example.com" }` writes a `w:hyperlink` node and an external relationship, then round-trips back to the same JSON.

- [ ] **Step 2: Run the targeted tests**

Run: `npm test -- tests/docx-core.test.ts -t "hyperlinks"`

Expected: FAIL because links are not supported yet.

- [ ] **Step 3: Implement minimal support**

Add `TextRun.link`, collect hyperlink relationships while writing document XML, emit `word/_rels/document.xml.rels`, and parse relationship ids back to URLs.

- [ ] **Step 4: Run the targeted tests**

Run: `npm test -- tests/docx-core.test.ts -t "hyperlinks"`

Expected: PASS.

## Task 2: Comments

**Files:**
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`
- Modify: `src/docx-reader.ts`
- Test: `tests/docx-core.test.ts`

- [ ] **Step 1: Write failing tests**

Add tests proving a commented run writes `word/comments.xml`, `commentRangeStart`, `commentRangeEnd`, and `commentReference`, then round-trips the author, date, text, and initials.

- [ ] **Step 2: Run the targeted tests**

Run: `npm test -- tests/docx-core.test.ts -t "comments"`

Expected: FAIL because comments are not supported yet.

- [ ] **Step 3: Implement minimal support**

Add `TextRun.comment`, emit comments package parts only when needed, wrap the commented run with comment range markers, and parse comments back by id.

- [ ] **Step 4: Run the targeted tests**

Run: `npm test -- tests/docx-core.test.ts -t "comments"`

Expected: PASS.

## Task 3: Bookmarks and Inline Breaks

**Files:**
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`
- Modify: `src/docx-reader.ts`
- Test: `tests/docx-core.test.ts`

- [ ] **Step 1: Write failing tests**

Add tests proving bookmark markers and inline line/page breaks are emitted and round-trip.

- [ ] **Step 2: Run the targeted tests**

Run: `npm test -- tests/docx-core.test.ts -t "bookmarks and breaks"`

Expected: FAIL because bookmarks and inline break controls are not supported yet.

- [ ] **Step 3: Implement minimal support**

Add `TextRun.bookmark` and `TextRun.break`, emit `w:bookmarkStart`, `w:bookmarkEnd`, `w:br`, and parse them back for the supported writer shape.

- [ ] **Step 4: Run the targeted tests**

Run: `npm test -- tests/docx-core.test.ts -t "bookmarks and breaks"`

Expected: PASS.

## Task 4: Full Verification and Push

**Files:**
- Modify: all Phase 3 files

- [ ] **Step 1: Run full verification**

Run: `npm test`

Expected: all tests pass.

Run: `npm run build`

Expected: TypeScript build exits with code 0.

- [ ] **Step 2: Commit**

```bash
git add .
git commit -m "feat: add phase 3 review references"
```

- [ ] **Step 3: Push**

```bash
git push -u origin phase-3-review-references
```

Expected: branch `phase-3-review-references` is pushed to `https://github.com/codingayice/word2json.git`.

## Self-Review

- Spec coverage: This plan covers the Phase 3 scope: hyperlinks, comments, bookmarks, and inline break controls.
- Placeholder scan: No unresolved placeholders remain.
- Type consistency: Public schema names are `TextRun.link`, `TextRun.comment`, `TextRun.bookmark`, and `TextRun.break`.
