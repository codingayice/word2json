# Phase 12 Run Advanced Formatting Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Preserve common advanced run-level Word formatting including highlight, strike, caps, superscript/subscript, character spacing, and text scale.

**Architecture:** Extend `TextRun` and reusable `StyleRunProperties` with the same advanced formatting subset. The writer emits these fields in `w:rPr`; the reader parses the supported subset from direct runs and style definitions.

**Tech Stack:** TypeScript, Vitest, JSZip, fast-xml-parser.

---

## File Structure

- `src/schema.ts`: add advanced run formatting fields to `TextRun` and `StyleRunProperties`.
- `src/docx-writer.ts`: emit advanced `w:rPr` children for direct runs and run styles.
- `src/docx-reader.ts`: parse advanced `w:rPr` children for direct runs and run styles.
- `tests/docx-core.test.ts`: add XML and round-trip tests.

## Task 1: Highlight and Strike

**Files:**
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`
- Modify: `src/docx-reader.ts`
- Test: `tests/docx-core.test.ts`

- [x] **Step 1: Write failing tests**

Add writer and round-trip tests for:

```ts
{ text: "Marked", highlight: "yellow", strike: true, doubleStrike: true }
```

- [x] **Step 2: Run targeted tests**

Run: `npm test -- tests/docx-core.test.ts -t "run highlight and strike"`

Expected: FAIL because highlight and strike fields are not supported yet.

- [x] **Step 3: Implement minimal support**

Add schema fields, emit `w:highlight`, `w:strike`, `w:dstrike`, and parse them back.

- [x] **Step 4: Run targeted tests**

Run: `npm test -- tests/docx-core.test.ts -t "run highlight and strike"`

Expected: PASS.

## Task 2: Caps and Vertical Align

**Files:**
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`
- Modify: `src/docx-reader.ts`
- Test: `tests/docx-core.test.ts`

- [x] **Step 1: Write failing tests**

Add writer and round-trip tests for:

```ts
{ text: "Formula", smallCaps: true, allCaps: true, verticalAlign: "superscript" }
{ text: "2", verticalAlign: "subscript" }
```

- [x] **Step 2: Run targeted tests**

Run: `npm test -- tests/docx-core.test.ts -t "run caps and vertical align"`

Expected: FAIL because caps and vertical alignment are not supported yet.

- [x] **Step 3: Implement minimal support**

Add schema fields, emit `w:smallCaps`, `w:caps`, `w:vertAlign`, and parse them back.

- [x] **Step 4: Run targeted tests**

Run: `npm test -- tests/docx-core.test.ts -t "run caps and vertical align"`

Expected: PASS.

## Task 3: Character Spacing and Scale

**Files:**
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`
- Modify: `src/docx-reader.ts`
- Test: `tests/docx-core.test.ts`

- [x] **Step 1: Write failing tests**

Add writer and round-trip tests for:

```ts
{ text: "Tracked", characterSpacing: 20, scale: 90 }
```

- [x] **Step 2: Run targeted tests**

Run: `npm test -- tests/docx-core.test.ts -t "run character spacing and scale"`

Expected: FAIL because spacing and scale are not supported yet.

- [x] **Step 3: Implement minimal support**

Add schema fields, emit `w:spacing` and `w:w`, and parse them back.

- [x] **Step 4: Run targeted tests**

Run: `npm test -- tests/docx-core.test.ts -t "run character spacing and scale"`

Expected: PASS.

## Task 4: Style Run Advanced Formatting

**Files:**
- Modify: `src/docx-writer.ts`
- Modify: `src/docx-reader.ts`
- Test: `tests/docx-core.test.ts`

- [x] **Step 1: Write failing tests**

Add writer and round-trip tests for a character style:

```ts
styles: {
  character: [{
    id: "WarningText",
    name: "Warning Text",
    run: { highlight: "yellow", strike: true, verticalAlign: "superscript", characterSpacing: 20, scale: 90 },
  }],
}
```

- [x] **Step 2: Run targeted tests**

Run: `npm test -- tests/docx-core.test.ts -t "style run advanced formatting"`

Expected: FAIL until style `w:rPr` shares the advanced run formatting support.

- [x] **Step 3: Implement minimal support**

Ensure `styleRunPropertiesXml` and `parseStyleRunProperties` share the same advanced fields as direct runs.

- [x] **Step 4: Run targeted tests**

Run: `npm test -- tests/docx-core.test.ts -t "style run advanced formatting"`

Expected: PASS.

## Task 5: Full Verification and Push

**Files:**
- Modify: all Phase 12 files

- [x] **Step 1: Run full verification**

Run: `npm test`

Expected: all tests pass.

Run: `npm run build`

Expected: TypeScript build exits with code 0.

- [x] **Step 2: Commit**

```bash
git add .
git commit -m "feat: add phase 12 run advanced formatting"
```

- [x] **Step 3: Push**

```bash
git push -u origin phase-12-run-advanced-formatting
```

Expected: branch `phase-12-run-advanced-formatting` is pushed to `https://github.com/codingayice/word2json.git`.

## Self-Review

- Spec coverage: This plan covers advanced direct run properties and reusable character style run properties.
- Placeholder scan: No unresolved placeholders remain.
- Type consistency: New schema names are `highlight`, `strike`, `doubleStrike`, `smallCaps`, `allCaps`, `verticalAlign`, `characterSpacing`, and `scale` on both `TextRun` and `StyleRunProperties`.
