# Phase 28 Advanced Content Controls Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Preserve additional Word structured document tag metadata for combo boxes and repeating sections.

**Architecture:** Extend the existing shared `ContentControl` shape because block and run SDTs already flow through the same writer and reader helpers. The writer adds `w:comboBox`, `w:repeatingSection`, and `w:repeatingSectionItem` under `w:sdtPr`; the reader parses those nodes through `parseContentControl()` so both block and run content controls round-trip.

**Tech Stack:** TypeScript, Vitest, JSZip, fast-xml-parser.

---

## File Structure

- `src/schema.ts`: add `comboBox`, `repeatingSection`, and `repeatingSectionItem` metadata types.
- `src/docx-writer.ts`: emit combo box and repeating section SDT properties.
- `src/docx-reader.ts`: parse combo box and repeating section SDT properties.
- `tests/docx-core.test.ts`: add writer and round-trip tests for the new metadata.

## Task 1: Combo Box Content Controls

**Files:**
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`
- Modify: `src/docx-reader.ts`
- Test: `tests/docx-core.test.ts`

- [x] **Step 1: Write failing tests**

Add writer and round-trip tests for:

```ts
{
  text: "Custom",
  contentControl: {
    alias: "Choice",
    tag: "choice",
    comboBox: {
      items: [
        { displayText: "Standard", value: "standard" },
        { displayText: "Custom", value: "custom" },
      ],
    },
  },
}
```

- [x] **Step 2: Run targeted tests**

Run: `npm test -- tests/docx-core.test.ts -t "combo box content controls"`

Expected: FAIL because combo box SDT metadata is not supported yet.

- [x] **Step 3: Implement minimal support**

Add `comboBox?: DropdownContentControl` to `ContentControl`, then emit and parse `w:comboBox` with `w:listItem` children.

- [x] **Step 4: Run targeted tests**

Run: `npm test -- tests/docx-core.test.ts -t "combo box content controls"`

Expected: PASS.

## Task 2: Repeating Section Content Controls

**Files:**
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`
- Modify: `src/docx-reader.ts`
- Test: `tests/docx-core.test.ts`

- [x] **Step 1: Write failing tests**

Add writer and round-trip tests for a block content control:

```ts
{
  type: "paragraph",
  contentControl: {
    alias: "Line Items",
    tag: "lineItems",
    repeatingSection: { sectionTitle: "Item", doNotAllowInsertDeleteSection: true },
    repeatingSectionItem: { id: "{11111111-2222-3333-4444-555555555555}" },
  },
  runs: [{ text: "Widget" }],
}
```

- [x] **Step 2: Run targeted tests**

Run: `npm test -- tests/docx-core.test.ts -t "repeating section content controls"`

Expected: FAIL because repeating section SDT metadata is not supported yet.

- [x] **Step 3: Implement minimal support**

Add:

```ts
export type RepeatingSectionContentControl = {
  sectionTitle?: string;
  doNotAllowInsertDeleteSection?: boolean;
};

export type RepeatingSectionItemContentControl = {
  id?: string;
};
```

Emit and parse:

```xml
<w:repeatingSection>
  <w:sectionTitle w:val="Item"/>
  <w:doNotAllowInsertDeleteSection/>
</w:repeatingSection>
<w:repeatingSectionItem>
  <w:id w:val="{11111111-2222-3333-4444-555555555555}"/>
</w:repeatingSectionItem>
```

- [x] **Step 4: Run targeted tests**

Run: `npm test -- tests/docx-core.test.ts -t "repeating section content controls"`

Expected: PASS.

## Task 3: Full Verification and Push

**Files:**
- Modify: all Phase 28 files

- [x] **Step 1: Run full verification**

Run: `npm test`

Expected: all tests pass.

Run: `npm run build`

Expected: TypeScript build exits with code 0.

Run: `git diff --check`

Expected: no whitespace errors.

- [x] **Step 2: Commit**

```bash
git add docs/superpowers/plans/2026-07-04-phase-28-advanced-content-controls.md src/schema.ts src/docx-writer.ts src/docx-reader.ts tests/docx-core.test.ts
git commit -m "feat: add phase 28 advanced content controls"
```

- [ ] **Step 3: Push**

```bash
git push -u origin phase-28-advanced-content-controls
```

Expected: branch `phase-28-advanced-content-controls` is pushed to `https://github.com/codingayice/word2json.git`.

## Self-Review

- Spec coverage: This plan covers combo box, repeating section, and repeating section item SDT metadata.
- Placeholder scan: No unresolved placeholders remain.
- Type consistency: property names match across schema, writer, reader, and tests.
