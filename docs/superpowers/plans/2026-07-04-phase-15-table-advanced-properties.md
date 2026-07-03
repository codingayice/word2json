# Phase 15 Table Advanced Properties Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Preserve common advanced table and cell properties that strongly affect Word table layout and visual fidelity.

**Architecture:** Extend the existing table schema with table alignment, table cell spacing, per-cell borders, and cell text direction. The writer emits these values into `w:tblPr` and `w:tcPr`; the reader parses the same supported subset from those property nodes.

**Tech Stack:** TypeScript, Vitest, JSZip, fast-xml-parser.

---

## File Structure

- `src/schema.ts`: add table `alignment`, table `cellSpacing`, cell `borders`, and cell `textDirection`.
- `src/docx-writer.ts`: emit `w:jc`, `w:tblCellSpacing`, `w:tcBorders`, and `w:textDirection`.
- `src/docx-reader.ts`: parse the same properties from `w:tblPr` and `w:tcPr`.
- `tests/docx-core.test.ts`: add XML writer and round-trip tests.

## Task 1: Table Alignment and Cell Spacing

**Files:**
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`
- Modify: `src/docx-reader.ts`
- Test: `tests/docx-core.test.ts`

- [x] **Step 1: Write failing writer test**

Add this test under `describe("DOCX writer", ...)`:

```ts
it("writes table alignment and cell spacing", async () => {
  const document = createDocumentJson([
    {
      type: "table",
      alignment: "center",
      cellSpacing: 120,
      rows: [
        { cells: [{ blocks: [{ type: "paragraph", runs: [{ text: "Centered" }] }] }] },
      ],
    },
  ]);

  const buffer = await buildDocx(document);
  const zip = await JSZip.loadAsync(buffer);
  const xml = await zip.file("word/document.xml")!.async("string");

  expect(xml).toContain('<w:jc w:val="center"/>');
  expect(xml).toContain('<w:tblCellSpacing w:w="120" w:type="dxa"/>');
});
```

- [x] **Step 2: Write failing round-trip test**

Add this test under `describe("DOCX reader", ...)`:

```ts
it("round-trips table alignment and cell spacing", async () => {
  const source = createDocumentJson([
    {
      type: "table",
      alignment: "center",
      cellSpacing: 120,
      rows: [
        { cells: [{ blocks: [{ type: "paragraph", runs: [{ text: "Centered" }] }] }] },
      ],
    },
  ]);

  const docx = await buildDocx(source);
  const parsed = await parseDocx(docx);

  expect(parsed).toEqual(source);
});
```

- [x] **Step 3: Run targeted tests**

Run: `npm test -- tests/docx-core.test.ts -t "table alignment and cell spacing"`

Expected: FAIL because these table properties are not supported yet.

- [x] **Step 4: Implement minimal support**

Add `alignment?: ParagraphAlignment` and `cellSpacing?: number` to `TableNode`. Emit `<w:jc w:val="..."/>` and `<w:tblCellSpacing w:w="..." w:type="dxa"/>`; parse them back from `w:tblPr`.

- [x] **Step 5: Run targeted tests**

Run: `npm test -- tests/docx-core.test.ts -t "table alignment and cell spacing"`

Expected: PASS.

## Task 2: Cell Borders

**Files:**
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`
- Modify: `src/docx-reader.ts`
- Test: `tests/docx-core.test.ts`

- [x] **Step 1: Write failing writer test**

Add this test under `describe("DOCX writer", ...)`:

```ts
it("writes cell borders", async () => {
  const document = createDocumentJson([
    {
      type: "table",
      rows: [
        {
          cells: [
            {
              borders: {
                top: { style: "single", size: 8, color: "4472C4", space: 0 },
                bottom: { style: "single", size: 8, color: "4472C4", space: 0 },
              },
              blocks: [{ type: "paragraph", runs: [{ text: "Bordered cell" }] }],
            },
          ],
        },
      ],
    },
  ]);

  const buffer = await buildDocx(document);
  const zip = await JSZip.loadAsync(buffer);
  const xml = await zip.file("word/document.xml")!.async("string");

  expect(xml).toContain("<w:tcBorders>");
  expect(xml).toContain('<w:top w:val="single" w:sz="8" w:space="0" w:color="4472C4"/>');
  expect(xml).toContain('<w:bottom w:val="single" w:sz="8" w:space="0" w:color="4472C4"/>');
});
```

- [x] **Step 2: Write failing round-trip test**

Add this test under `describe("DOCX reader", ...)`:

```ts
it("round-trips cell borders", async () => {
  const source = createDocumentJson([
    {
      type: "table",
      rows: [
        {
          cells: [
            {
              borders: {
                top: { style: "single", size: 8, color: "4472C4", space: 0 },
                bottom: { style: "single", size: 8, color: "4472C4", space: 0 },
              },
              blocks: [{ type: "paragraph", runs: [{ text: "Bordered cell" }] }],
            },
          ],
        },
      ],
    },
  ]);

  const docx = await buildDocx(source);
  const parsed = await parseDocx(docx);

  expect(parsed).toEqual(source);
});
```

- [x] **Step 3: Run targeted tests**

Run: `npm test -- tests/docx-core.test.ts -t "cell borders"`

Expected: FAIL because `TableCellNode.borders` is not supported yet.

- [x] **Step 4: Implement minimal support**

Reuse the existing `ParagraphBorders`/`BorderDefinition` shape for `TableCellNode.borders`. Emit `<w:tcBorders>` in `w:tcPr` and parse supported sides back.

- [x] **Step 5: Run targeted tests**

Run: `npm test -- tests/docx-core.test.ts -t "cell borders"`

Expected: PASS.

## Task 3: Cell Text Direction

**Files:**
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`
- Modify: `src/docx-reader.ts`
- Test: `tests/docx-core.test.ts`

- [x] **Step 1: Write failing writer test**

Add this test under `describe("DOCX writer", ...)`:

```ts
it("writes cell text direction", async () => {
  const document = createDocumentJson([
    {
      type: "table",
      rows: [
        {
          cells: [
            {
              textDirection: "btLr",
              blocks: [{ type: "paragraph", runs: [{ text: "Vertical" }] }],
            },
          ],
        },
      ],
    },
  ]);

  const buffer = await buildDocx(document);
  const zip = await JSZip.loadAsync(buffer);
  const xml = await zip.file("word/document.xml")!.async("string");

  expect(xml).toContain('<w:textDirection w:val="btLr"/>');
});
```

- [x] **Step 2: Write failing round-trip test**

Add this test under `describe("DOCX reader", ...)`:

```ts
it("round-trips cell text direction", async () => {
  const source = createDocumentJson([
    {
      type: "table",
      rows: [
        {
          cells: [
            {
              textDirection: "btLr",
              blocks: [{ type: "paragraph", runs: [{ text: "Vertical" }] }],
            },
          ],
        },
      ],
    },
  ]);

  const docx = await buildDocx(source);
  const parsed = await parseDocx(docx);

  expect(parsed).toEqual(source);
});
```

- [x] **Step 3: Run targeted tests**

Run: `npm test -- tests/docx-core.test.ts -t "cell text direction"`

Expected: FAIL because `TableCellNode.textDirection` is not supported yet.

- [x] **Step 4: Implement minimal support**

Add `textDirection?: "lrTb" | "tbRl" | "btLr"` to `TableCellNode`. Emit `<w:textDirection w:val="..."/>` and parse it back from `w:tcPr`.

- [x] **Step 5: Run targeted tests**

Run: `npm test -- tests/docx-core.test.ts -t "cell text direction"`

Expected: PASS.

## Task 4: Full Verification and Push

**Files:**
- Modify: all Phase 15 files

- [x] **Step 1: Run full verification**

Run: `npm test`

Expected: all tests pass.

Run: `npm run build`

Expected: TypeScript build exits with code 0.

- [x] **Step 2: Commit**

```bash
git add .
git commit -m "feat: add phase 15 table advanced properties"
```

- [x] **Step 3: Push**

```bash
git push -u origin phase-15-table-advanced-properties
```

Expected: branch `phase-15-table-advanced-properties` is pushed to `https://github.com/codingayice/word2json.git`.

## Self-Review

- Spec coverage: This plan covers table alignment, cell spacing, per-cell borders, and cell text direction.
- Placeholder scan: No unresolved placeholders remain.
- Type consistency: Table-level properties live on `TableNode`; cell-level properties live on `TableCellNode`.
