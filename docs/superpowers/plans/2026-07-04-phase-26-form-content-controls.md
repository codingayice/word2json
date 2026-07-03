# Phase 26 Form Content Controls Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Preserve common Word form content control types: checkbox, dropdown list, and date picker.

**Architecture:** Extend the existing `ContentControl` metadata with optional form-specific variants. The writer emits the corresponding `w:checkBox`, `w:dropDownList`, or `w:date` metadata inside `w:sdtPr`; the reader parses those nodes back into JSON while preserving existing alias/tag/lock handling.

**Tech Stack:** TypeScript, Vitest, JSZip, fast-xml-parser.

---

## File Structure

- `src/schema.ts`: extend `ContentControl` with checkbox, dropdown, and date properties.
- `src/docx-writer.ts`: emit form SDT properties.
- `src/docx-reader.ts`: parse form SDT properties.
- `tests/docx-core.test.ts`: add writer and round-trip tests for each form control type.

## Task 1: Checkbox Content Controls

**Files:**
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`
- Modify: `src/docx-reader.ts`
- Test: `tests/docx-core.test.ts`

- [x] **Step 1: Write failing tests**

Add writer and round-trip tests for:

```ts
{
  text: "☒",
  contentControl: {
    alias: "Accepted",
    tag: "accepted",
    checkbox: { checked: true, checkedSymbol: "2612", uncheckedSymbol: "2610" },
  },
}
```

- [x] **Step 2: Run targeted tests**

Run: `npm test -- tests/docx-core.test.ts -t "checkbox content controls"`

Expected: FAIL because checkbox SDT metadata is not supported yet.

- [x] **Step 3: Implement minimal support**

Emit and parse `w:checkBox`, `w:checked`, `w:checkedState`, and `w:uncheckedState`.

- [x] **Step 4: Run targeted tests**

Run: `npm test -- tests/docx-core.test.ts -t "checkbox content controls"`

Expected: PASS.

## Task 2: Dropdown Content Controls

**Files:**
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`
- Modify: `src/docx-reader.ts`
- Test: `tests/docx-core.test.ts`

- [x] **Step 1: Write failing tests**

Add writer and round-trip tests for:

```ts
{
  text: "Gold",
  contentControl: {
    alias: "Plan",
    tag: "plan",
    dropdown: {
      items: [
        { displayText: "Silver", value: "silver" },
        { displayText: "Gold", value: "gold" },
      ],
    },
  },
}
```

- [x] **Step 2: Run targeted tests**

Run: `npm test -- tests/docx-core.test.ts -t "dropdown content controls"`

Expected: FAIL because dropdown SDT metadata is not supported yet.

- [x] **Step 3: Implement minimal support**

Emit and parse `w:dropDownList` with `w:listItem` children.

- [x] **Step 4: Run targeted tests**

Run: `npm test -- tests/docx-core.test.ts -t "dropdown content controls"`

Expected: PASS.

## Task 3: Date Content Controls

**Files:**
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`
- Modify: `src/docx-reader.ts`
- Test: `tests/docx-core.test.ts`

- [x] **Step 1: Write failing tests**

Add writer and round-trip tests for:

```ts
{
  text: "2026-07-04",
  contentControl: {
    alias: "Due Date",
    tag: "dueDate",
    date: { fullDate: "2026-07-04T00:00:00Z", format: "yyyy-MM-dd" },
  },
}
```

- [x] **Step 2: Run targeted tests**

Run: `npm test -- tests/docx-core.test.ts -t "date content controls"`

Expected: FAIL because date SDT metadata is not supported yet.

- [x] **Step 3: Implement minimal support**

Emit and parse `w:date`, `w:fullDate`, and `w:dateFormat`.

- [x] **Step 4: Run targeted tests**

Run: `npm test -- tests/docx-core.test.ts -t "date content controls"`

Expected: PASS.

## Task 4: Full Verification and Push

**Files:**
- Modify: all Phase 26 files

- [x] **Step 1: Run full verification**

Run: `npm test`

Expected: all tests pass.

Run: `npm run build`

Expected: TypeScript build exits with code 0.

- [ ] **Step 2: Commit**

```bash
git add .
git commit -m "feat: add phase 26 form content controls"
```

- [ ] **Step 3: Push**

```bash
git push -u origin phase-26-form-content-controls
```

Expected: branch `phase-26-form-content-controls` is pushed to `https://github.com/codingayice/word2json.git`.

## Self-Review

- Spec coverage: This plan covers checkbox, dropdown list, and date picker SDT metadata.
- Placeholder scan: No unresolved placeholders remain.
- Type consistency: all new fields extend the existing `ContentControl` type.
