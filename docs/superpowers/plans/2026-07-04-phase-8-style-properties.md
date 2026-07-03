# Phase 8 Style Properties Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Preserve reusable formatting inside style definitions so DOCX templates can round-trip paragraph, character, and table style properties instead of only style IDs.

**Architecture:** Extend style definition schema with small property subsets that reuse existing document formatting concepts. The writer serializes those properties inside `word/styles.xml`, and the reader parses the same supported subset back into `DocumentJson.styles`.

**Tech Stack:** TypeScript, Vitest, JSZip, fast-xml-parser.

---

## File Structure

- `src/schema.ts`: add run, paragraph, and table style property types to style definitions.
- `src/docx-writer.ts`: emit `w:pPr`, `w:rPr`, and `w:tblPr` inside style definitions.
- `src/docx-reader.ts`: parse supported style properties from `word/styles.xml`.
- `tests/docx-core.test.ts`: add XML and round-trip tests.

## Task 1: Paragraph Style Properties

**Files:**
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`
- Modify: `src/docx-reader.ts`
- Test: `tests/docx-core.test.ts`

- [x] **Step 1: Write failing tests**

Add tests proving a paragraph style can contain paragraph alignment and run formatting:

```ts
styles: {
  paragraph: [{
    id: "ContractTitle",
    name: "Contract Title",
    basedOn: "Normal",
    next: "Normal",
    paragraph: { alignment: "center" },
    run: { bold: true, fontFamily: "Aptos Display", fontSize: 18, color: "1F4E79" },
  }],
}
```

- [x] **Step 2: Run targeted tests**

Run: `npm test -- tests/docx-core.test.ts -t "paragraph style properties"`

Expected: FAIL because style properties are not serialized or parsed yet.

- [x] **Step 3: Implement minimal support**

Add `StyleRunProperties`, `StyleParagraphProperties`, `ParagraphStyleDefinition.run`, and `ParagraphStyleDefinition.paragraph`. Emit style-level `w:pPr/w:jc` and `w:rPr` with existing run formatting subset. Parse them back from custom paragraph styles.

- [x] **Step 4: Run targeted tests**

Run: `npm test -- tests/docx-core.test.ts -t "paragraph style properties"`

Expected: PASS.

## Task 2: Character Style Properties

**Files:**
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`
- Modify: `src/docx-reader.ts`
- Test: `tests/docx-core.test.ts`

- [x] **Step 1: Write failing tests**

Add tests proving character styles can carry run formatting:

```ts
styles: {
  character: [{
    id: "DefinedTerm",
    name: "Defined Term",
    basedOn: "DefaultParagraphFont",
    run: { italic: true, underline: true, color: "C00000" },
  }],
}
```

- [x] **Step 2: Run targeted tests**

Run: `npm test -- tests/docx-core.test.ts -t "character style properties"`

Expected: FAIL because character style `run` properties are not supported yet.

- [x] **Step 3: Implement minimal support**

Add `StyleDefinition.run`, emit style-level `w:rPr`, and parse it back for character styles.

- [x] **Step 4: Run targeted tests**

Run: `npm test -- tests/docx-core.test.ts -t "character style properties"`

Expected: PASS.

## Task 3: Table Style Properties

**Files:**
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`
- Modify: `src/docx-reader.ts`
- Test: `tests/docx-core.test.ts`

- [x] **Step 1: Write failing tests**

Add tests proving table styles can carry table borders:

```ts
styles: {
  table: [{
    id: "ContractTable",
    name: "Contract Table",
    basedOn: "TableNormal",
    table: { borders: "single" },
  }],
}
```

- [x] **Step 2: Run targeted tests**

Run: `npm test -- tests/docx-core.test.ts -t "table style properties"`

Expected: FAIL because table style `table` properties are not supported yet.

- [x] **Step 3: Implement minimal support**

Add `TableStyleDefinition.table`, emit style-level `w:tblPr/w:tblBorders`, and parse it back for table styles.

- [x] **Step 4: Run targeted tests**

Run: `npm test -- tests/docx-core.test.ts -t "table style properties"`

Expected: PASS.

## Task 4: Full Verification and Push

**Files:**
- Modify: all Phase 8 files

- [x] **Step 1: Run full verification**

Run: `npm test`

Expected: all tests pass.

Run: `npm run build`

Expected: TypeScript build exits with code 0.

- [ ] **Step 2: Commit**

```bash
git add .
git commit -m "feat: add phase 8 style properties"
```

- [ ] **Step 3: Push**

```bash
git push -u origin phase-8-style-properties
```

Expected: branch `phase-8-style-properties` is pushed to `https://github.com/codingayice/word2json.git`.

## Self-Review

- Spec coverage: This plan deepens Phase 7 style IDs into reusable formatting properties for paragraph, character, and table styles.
- Placeholder scan: No unresolved placeholders remain.
- Type consistency: New schema names are `StyleRunProperties`, `StyleParagraphProperties`, `ParagraphStyleDefinition.paragraph`, `StyleDefinition.run`, and `TableStyleDefinition.table`.
