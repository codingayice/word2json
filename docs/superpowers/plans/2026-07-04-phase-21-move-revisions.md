# Phase 21 Move Revisions Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Preserve Word tracked move revisions for moved text and moved paragraphs.

**Architecture:** Extend revision metadata from insert/delete to include `moveFrom` and `moveTo`. The writer emits run-level `w:moveFrom`/`w:moveTo` containers and paragraph-level `w:moveFrom`/`w:moveTo` wrappers. The reader treats those containers as revision-bearing content and round-trips them through JSON.

**Tech Stack:** TypeScript, Vitest, JSZip, fast-xml-parser.

---

## File Structure

- `src/schema.ts`: extend `RunRevision["type"]` to include `moveFrom` and `moveTo`.
- `src/docx-writer.ts`: emit run-level and paragraph-level move revision containers.
- `src/docx-reader.ts`: parse move revision containers at run and block level.
- `tests/docx-core.test.ts`: add writer and round-trip tests for run and paragraph move revisions.

## Task 1: Run Move Revisions

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
    { text: "Moved away", revision: { type: "moveFrom", id: 30, author: "Ada", date: "2026-07-04T09:00:00.000Z" } },
    { text: "Moved here", revision: { type: "moveTo", id: 31, author: "Lin", date: "2026-07-04T10:00:00.000Z" } },
  ],
}
```

- [x] **Step 2: Run targeted tests**

Run: `npm test -- tests/docx-core.test.ts -t "run move revisions"`

Expected: FAIL because run move revision containers are not supported yet.

- [x] **Step 3: Implement minimal support**

Extend revision type union, emit `w:moveFrom` with `w:delText`, emit `w:moveTo` with `w:t`, and parse both containers back into `TextRun.revision`.

- [x] **Step 4: Run targeted tests**

Run: `npm test -- tests/docx-core.test.ts -t "run move revisions"`

Expected: PASS.

## Task 2: Paragraph Move Revisions

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
  revision: { type: "moveFrom", id: 32, author: "Mira", date: "2026-07-04T11:00:00.000Z" },
  runs: [{ text: "Moved paragraph from" }],
}
```

```ts
{
  type: "paragraph",
  revision: { type: "moveTo", id: 33, author: "Noor", date: "2026-07-04T12:00:00.000Z" },
  runs: [{ text: "Moved paragraph to" }],
}
```

- [x] **Step 2: Run targeted tests**

Run: `npm test -- tests/docx-core.test.ts -t "paragraph move revisions"`

Expected: FAIL because paragraph-level move revision wrappers are not supported yet.

- [x] **Step 3: Implement minimal support**

Wrap paragraph XML in `w:moveFrom`/`w:moveTo`, extract those wrappers as top-level blocks, and parse the wrapper metadata onto `ParagraphNode.revision`.

- [x] **Step 4: Run targeted tests**

Run: `npm test -- tests/docx-core.test.ts -t "paragraph move revisions"`

Expected: PASS.

## Task 3: Full Verification and Push

**Files:**
- Modify: all Phase 21 files

- [x] **Step 1: Run full verification**

Run: `npm test`

Expected: all tests pass.

Run: `npm run build`

Expected: TypeScript build exits with code 0.

- [x] **Step 2: Commit**

```bash
git add .
git commit -m "feat: add phase 21 move revisions"
```

- [x] **Step 3: Push**

```bash
git push -u origin phase-21-move-revisions
```

Expected: branch `phase-21-move-revisions` is pushed to `https://github.com/codingayice/word2json.git`.

## Self-Review

- Spec coverage: This plan covers run and paragraph move revision metadata and content.
- Placeholder scan: No unresolved placeholders remain.
- Type consistency: move revisions reuse the existing `revision` property on runs and paragraphs.
