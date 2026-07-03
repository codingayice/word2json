# Phase 36 Font Table Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Preserve DOCX `word/fontTable.xml` declarations for higher-fidelity font handling.

**Architecture:** Add optional top-level `fonts` to `DocumentJson`. The writer emits `word/fontTable.xml`, adds a document relationship and content type override; the reader parses font table entries back into JSON.

**Tech Stack:** TypeScript, Vitest, JSZip, fast-xml-parser.

---

## File Structure

- `src/schema.ts`: add `DocumentFont` and optional `DocumentJson.fonts`.
- `src/docx-writer.ts`: emit font table part, relationship, and content type override.
- `src/docx-reader.ts`: parse font table part.
- `tests/docx-core.test.ts`: add writer and round-trip tests for font table.

## Task 1: Writer Support

**Files:**
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`
- Test: `tests/docx-core.test.ts`

- [x] **Step 1: Write failing writer test**

Add a test named `writes font table` with:

```ts
fonts: [
  { name: "Aptos", family: "swiss", pitch: "variable", charset: "00", panose1: "020F0502020204030204" },
  { name: "SimSun", family: "roman", pitch: "fixed", charset: "86" },
]
```

- [x] **Step 2: Run targeted writer test**

Run: `npm test -- tests/docx-core.test.ts -t "font table"`

Expected: FAIL because `word/fontTable.xml` is not supported yet.

- [x] **Step 3: Implement minimal writer support**

Emit:

```xml
<w:fonts>
  <w:font w:name="Aptos">
    <w:panose1 w:val="020F0502020204030204"/>
    <w:charset w:val="00"/>
    <w:family w:val="swiss"/>
    <w:pitch w:val="variable"/>
  </w:font>
</w:fonts>
```

and add relationship/content type entries.

- [x] **Step 4: Run targeted writer test**

Run: `npm test -- tests/docx-core.test.ts -t "font table"`

Expected: writer test PASS.

## Task 2: Reader Round-Trip Support

**Files:**
- Modify: `src/docx-reader.ts`
- Test: `tests/docx-core.test.ts`

- [x] **Step 1: Write failing round-trip test**

Add a test named `round-trips font table` using the same `fonts` array.

- [x] **Step 2: Run targeted round-trip test**

Run: `npm test -- tests/docx-core.test.ts -t "font table"`

Expected: FAIL because the reader drops font table entries.

- [x] **Step 3: Implement minimal reader support**

Parse `word/fontTable.xml` into `DocumentJson.fonts`, preserving name, family, pitch, charset, and panose1.

- [x] **Step 4: Run targeted round-trip test**

Run: `npm test -- tests/docx-core.test.ts -t "font table"`

Expected: writer and round-trip tests PASS.

## Task 3: Full Verification and Push

**Files:**
- Modify: all Phase 36 files

- [x] **Step 1: Run full verification**

Run: `npm test`

Expected: all tests pass.

Run: `npm run build`

Expected: TypeScript build exits with code 0.

Run: `git diff --check`

Expected: no whitespace errors.

- [x] **Step 2: Commit**

```bash
git add docs/superpowers/plans/2026-07-04-phase-36-font-table.md src/schema.ts src/docx-writer.ts src/docx-reader.ts tests/docx-core.test.ts
git commit -m "feat: add phase 36 font table"
```

- [ ] **Step 3: Push**

```bash
git push -u origin phase-36-font-table
```

Expected: branch `phase-36-font-table` is pushed to `https://github.com/codingayice/word2json.git`.

## Self-Review

- Spec coverage: This plan covers font table part writing, relationship/content type registration, parsing, and round-trip.
- Placeholder scan: No unresolved placeholders remain.
- Type consistency: JSON uses `panose1`, and XML uses `w:panose1`.
