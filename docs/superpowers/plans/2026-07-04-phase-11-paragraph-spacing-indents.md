# Phase 11 Paragraph Spacing Indents Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Preserve common paragraph layout properties including spacing before/after/line and left/right/first-line/hanging indentation.

**Architecture:** Extend `ParagraphNode` and paragraph style definitions with reusable paragraph layout properties. The writer emits these properties into `w:pPr/w:spacing` and `w:pPr/w:ind`; the reader parses the same supported subset back into JSON.

**Tech Stack:** TypeScript, Vitest, JSZip, fast-xml-parser.

---

## File Structure

- `src/schema.ts`: add paragraph spacing and indentation models.
- `src/docx-writer.ts`: emit paragraph and style-level `w:spacing` and `w:ind`.
- `src/docx-reader.ts`: parse paragraph and style-level `w:spacing` and `w:ind`.
- `tests/docx-core.test.ts`: add XML and round-trip tests.

## Task 1: Paragraph Spacing

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
  spacing: { before: 240, after: 120, line: 360, lineRule: "auto" },
  runs: [{ text: "Spaced paragraph" }],
}
```

- [x] **Step 2: Run targeted tests**

Run: `npm test -- tests/docx-core.test.ts -t "paragraph spacing"`

Expected: FAIL because paragraph spacing is not supported yet.

- [x] **Step 3: Implement minimal support**

Add `ParagraphNode.spacing`, emit `w:spacing`, and parse it back.

- [x] **Step 4: Run targeted tests**

Run: `npm test -- tests/docx-core.test.ts -t "paragraph spacing"`

Expected: PASS.

## Task 2: Paragraph Indentation

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
  indent: { left: 720, right: 360, firstLine: 240, hanging: 120 },
  runs: [{ text: "Indented paragraph" }],
}
```

- [x] **Step 2: Run targeted tests**

Run: `npm test -- tests/docx-core.test.ts -t "paragraph indentation"`

Expected: FAIL because paragraph indentation is not supported yet.

- [x] **Step 3: Implement minimal support**

Add `ParagraphNode.indent`, emit `w:ind`, and parse it back.

- [x] **Step 4: Run targeted tests**

Run: `npm test -- tests/docx-core.test.ts -t "paragraph indentation"`

Expected: PASS.

## Task 3: Paragraph Style Layout Properties

**Files:**
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`
- Modify: `src/docx-reader.ts`
- Test: `tests/docx-core.test.ts`

- [x] **Step 1: Write failing tests**

Add writer and round-trip tests for paragraph style properties:

```ts
styles: {
  paragraph: [{
    id: "BodyText",
    name: "Body Text",
    paragraph: {
      spacing: { before: 120, after: 120 },
      indent: { left: 360, hanging: 180 },
    },
  }],
}
```

- [x] **Step 2: Run targeted tests**

Run: `npm test -- tests/docx-core.test.ts -t "paragraph style layout properties"`

Expected: FAIL because style-level spacing and indent are not supported yet.

- [x] **Step 3: Implement minimal support**

Extend `StyleParagraphProperties` with `spacing` and `indent`, emit them inside style-level `w:pPr`, and parse them back.

- [x] **Step 4: Run targeted tests**

Run: `npm test -- tests/docx-core.test.ts -t "paragraph style layout properties"`

Expected: PASS.

## Task 4: Full Verification and Push

**Files:**
- Modify: all Phase 11 files

- [x] **Step 1: Run full verification**

Run: `npm test`

Expected: all tests pass.

Run: `npm run build`

Expected: TypeScript build exits with code 0.

- [ ] **Step 2: Commit**

```bash
git add .
git commit -m "feat: add phase 11 paragraph spacing indents"
```

- [ ] **Step 3: Push**

```bash
git push -u origin phase-11-paragraph-spacing-indents
```

Expected: branch `phase-11-paragraph-spacing-indents` is pushed to `https://github.com/codingayice/word2json.git`.

## Self-Review

- Spec coverage: This plan covers direct paragraph spacing, direct indentation, and paragraph style-level layout properties.
- Placeholder scan: No unresolved placeholders remain.
- Type consistency: New schema names are `ParagraphSpacing`, `ParagraphIndent`, `ParagraphNode.spacing`, `ParagraphNode.indent`, `StyleParagraphProperties.spacing`, and `StyleParagraphProperties.indent`.
