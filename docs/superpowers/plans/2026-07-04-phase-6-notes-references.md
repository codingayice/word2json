# Phase 6 Notes References Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add long-document reference features: footnotes, endnotes, REF fields, and PAGEREF fields.

**Architecture:** Continue the explicit OOXML package writer/reader. Footnotes and endnotes are modeled as run-level references that create package parts; reference fields reuse the existing structured field-run approach and extend it with targets.

**Tech Stack:** TypeScript, Vitest, JSZip, fast-xml-parser, tsx.

---

## File Structure

- `src/schema.ts`: add note reference data and richer field run data.
- `src/docx-writer.ts`: emit footnote/endnote parts, relationships, content types, note reference runs, and REF/PAGEREF field codes.
- `src/docx-reader.ts`: parse footnote/endnote parts and field codes back into JSON.
- `tests/docx-core.test.ts`: add XML and round-trip tests.

## Task 1: Footnotes

**Files:**
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`
- Modify: `src/docx-reader.ts`
- Test: `tests/docx-core.test.ts`

- [ ] **Step 1: Write failing tests**

Add tests proving a run with `footnote` writes `word/footnotes.xml`, a `footnoteReference`, relationships, content type overrides, and round-trips note content.

- [ ] **Step 2: Run targeted tests**

Run: `npm test -- tests/docx-core.test.ts -t "footnotes"`

Expected: FAIL because footnotes are not supported yet.

- [ ] **Step 3: Implement minimal support**

Add run-level `footnote`, create note ids starting at `1`, write the notes part with required separator entries, and parse the supported writer shape back.

- [ ] **Step 4: Run targeted tests**

Run: `npm test -- tests/docx-core.test.ts -t "footnotes"`

Expected: PASS.

## Task 2: Endnotes

**Files:**
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`
- Modify: `src/docx-reader.ts`
- Test: `tests/docx-core.test.ts`

- [ ] **Step 1: Write failing tests**

Add tests proving a run with `endnote` writes `word/endnotes.xml`, an `endnoteReference`, relationships, content type overrides, and round-trips note content.

- [ ] **Step 2: Run targeted tests**

Run: `npm test -- tests/docx-core.test.ts -t "endnotes"`

Expected: FAIL because endnotes are not supported yet.

- [ ] **Step 3: Implement minimal support**

Add run-level `endnote`, create note ids starting at `1`, write the endnotes part with required separator entries, and parse the supported writer shape back.

- [ ] **Step 4: Run targeted tests**

Run: `npm test -- tests/docx-core.test.ts -t "endnotes"`

Expected: PASS.

## Task 3: REF and PAGEREF Fields

**Files:**
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`
- Modify: `src/docx-reader.ts`
- Test: `tests/docx-core.test.ts`

- [ ] **Step 1: Write failing tests**

Add tests proving field runs can emit and round-trip `REF Clause1` and `PAGEREF Clause1`.

- [ ] **Step 2: Run targeted tests**

Run: `npm test -- tests/docx-core.test.ts -t "reference fields"`

Expected: FAIL because field targets are not supported yet.

- [ ] **Step 3: Implement minimal support**

Extend `TextRun.field` from string literals to structured field objects while preserving `page` and `numPages`, emit field instructions, and parse known instructions back.

- [ ] **Step 4: Run targeted tests**

Run: `npm test -- tests/docx-core.test.ts -t "reference fields"`

Expected: PASS.

## Task 4: Full Verification and Push

**Files:**
- Modify: all Phase 6 files

- [ ] **Step 1: Run full verification**

Run: `npm test`

Expected: all tests pass.

Run: `npm run build`

Expected: TypeScript build exits with code 0.

- [ ] **Step 2: Commit**

```bash
git add .
git commit -m "feat: add phase 6 notes references"
```

- [ ] **Step 3: Push**

```bash
git push -u origin phase-6-notes-references
```

Expected: branch `phase-6-notes-references` is pushed to `https://github.com/codingayice/word2json.git`.

## Self-Review

- Spec coverage: This plan covers footnotes, endnotes, REF, and PAGEREF.
- Placeholder scan: No unresolved placeholders remain.
- Type consistency: New schema names are `TextRun.footnote`, `TextRun.endnote`, and structured `TextRun.field`.
