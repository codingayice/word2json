# Phase 17 Complex Fields References Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Preserve common complex field and internal reference structures used in long Word documents.

**Architecture:** Extend `TextRun.field` to carry optional display result text and add TOC field support. Extend hyperlinks so they can target internal bookmarks without external relationships. The writer emits complex field begin/instruction/separate/result/end sequences; the reader groups those sequences back into one JSON run.

**Tech Stack:** TypeScript, Vitest, JSZip, fast-xml-parser.

---

## File Structure

- `src/schema.ts`: add field result objects and internal hyperlink target support.
- `src/docx-writer.ts`: emit field result runs, TOC instructions, and internal hyperlink anchors.
- `src/docx-reader.ts`: parse complex field sequences with results and internal hyperlink anchors.
- `tests/docx-core.test.ts`: add XML writer tests and round-trip tests.

## Task 1: Field Result Text

**Files:**
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`
- Modify: `src/docx-reader.ts`
- Test: `tests/docx-core.test.ts`

- [x] **Step 1: Write failing writer test**

Add this test under `describe("DOCX writer", ...)`:

```ts
it("writes field result text", async () => {
  const document = createDocumentJson([
    {
      type: "paragraph",
      runs: [
        { text: "", field: { type: "page", result: "3" } },
      ],
    },
  ]);

  const buffer = await buildDocx(document);
  const zip = await JSZip.loadAsync(buffer);
  const xml = await zip.file("word/document.xml")!.async("string");

  expect(xml).toContain('<w:instrText xml:space="preserve">PAGE</w:instrText>');
  expect(xml).toContain("<w:t>3</w:t>");
});
```

- [x] **Step 2: Write failing round-trip test**

Add this test under `describe("DOCX reader", ...)`:

```ts
it("round-trips field result text", async () => {
  const source = createDocumentJson([
    {
      type: "paragraph",
      runs: [
        { text: "", field: { type: "page", result: "3" } },
      ],
    },
  ]);

  const docx = await buildDocx(source);
  const parsed = await parseDocx(docx);

  expect(parsed).toEqual(source);
});
```

- [x] **Step 3: Run targeted tests**

Run: `npm test -- tests/docx-core.test.ts -t "field result text"`

Expected: FAIL because field result text is not preserved yet.

- [x] **Step 4: Implement minimal support**

Add `FieldWithResult` objects for `page`, `numPages`, `ref`, and `pageRef`. Emit result text between `w:fldCharType="separate"` and `w:fldCharType="end"`, and parse grouped field sequences back into one run.

- [x] **Step 5: Run targeted tests**

Run: `npm test -- tests/docx-core.test.ts -t "field result text"`

Expected: PASS.

## Task 2: Table of Contents Field

**Files:**
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`
- Modify: `src/docx-reader.ts`
- Test: `tests/docx-core.test.ts`

- [x] **Step 1: Write failing writer test**

Add this test under `describe("DOCX writer", ...)`:

```ts
it("writes table of contents field", async () => {
  const document = createDocumentJson([
    {
      type: "paragraph",
      runs: [
        { text: "", field: { type: "toc", switches: 'o "1-3" h z u', result: "Table of Contents" } },
      ],
    },
  ]);

  const buffer = await buildDocx(document);
  const zip = await JSZip.loadAsync(buffer);
  const xml = await zip.file("word/document.xml")!.async("string");

  expect(xml).toContain('<w:instrText xml:space="preserve">TOC \\o "1-3" \\h \\z \\u</w:instrText>');
  expect(xml).toContain("<w:t>Table of Contents</w:t>");
});
```

- [x] **Step 2: Write failing round-trip test**

Add this test under `describe("DOCX reader", ...)`:

```ts
it("round-trips table of contents field", async () => {
  const source = createDocumentJson([
    {
      type: "paragraph",
      runs: [
        { text: "", field: { type: "toc", switches: 'o "1-3" h z u', result: "Table of Contents" } },
      ],
    },
  ]);

  const docx = await buildDocx(source);
  const parsed = await parseDocx(docx);

  expect(parsed).toEqual(source);
});
```

- [x] **Step 3: Run targeted tests**

Run: `npm test -- tests/docx-core.test.ts -t "table of contents field"`

Expected: FAIL because TOC fields are not supported yet.

- [x] **Step 4: Implement minimal support**

Add `TocField` to the schema. Convert `switches: 'o "1-3" h z u'` to `TOC \o "1-3" \h \z \u` in XML, and parse TOC instructions back into the compact switch string.

- [x] **Step 5: Run targeted tests**

Run: `npm test -- tests/docx-core.test.ts -t "table of contents field"`

Expected: PASS.

## Task 3: Internal Hyperlink Anchors

**Files:**
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`
- Modify: `src/docx-reader.ts`
- Test: `tests/docx-core.test.ts`

- [x] **Step 1: Write failing writer test**

Add this test under `describe("DOCX writer", ...)`:

```ts
it("writes internal hyperlinks", async () => {
  const document = createDocumentJson([
    {
      type: "paragraph",
      runs: [
        { text: "Jump", link: { anchor: "Clause1" } },
      ],
    },
  ]);

  const buffer = await buildDocx(document);
  const zip = await JSZip.loadAsync(buffer);
  const xml = await zip.file("word/document.xml")!.async("string");
  const rels = await zip.file("word/_rels/document.xml.rels")!.async("string");

  expect(xml).toContain('<w:hyperlink w:anchor="Clause1">');
  expect(xml).not.toContain('Target="Clause1"');
  expect(rels).not.toContain("rIdHyperlink1");
});
```

- [x] **Step 2: Write failing round-trip test**

Add this test under `describe("DOCX reader", ...)`:

```ts
it("round-trips internal hyperlinks", async () => {
  const source = createDocumentJson([
    {
      type: "paragraph",
      runs: [
        { text: "Jump", link: { anchor: "Clause1" } },
      ],
    },
  ]);

  const docx = await buildDocx(source);
  const parsed = await parseDocx(docx);

  expect(parsed).toEqual(source);
});
```

- [x] **Step 3: Run targeted tests**

Run: `npm test -- tests/docx-core.test.ts -t "internal hyperlinks"`

Expected: FAIL because hyperlinks only support external relationship URLs.

- [x] **Step 4: Implement minimal support**

Change `Hyperlink` to support `{ url: string }` or `{ anchor: string }`. Emit `w:anchor` without relationship for internal links, and parse hyperlink anchor attributes back.

- [x] **Step 5: Run targeted tests**

Run: `npm test -- tests/docx-core.test.ts -t "internal hyperlinks"`

Expected: PASS.

## Task 4: Full Verification and Push

**Files:**
- Modify: all Phase 17 files

- [x] **Step 1: Run full verification**

Run: `npm test`

Expected: all tests pass.

Run: `npm run build`

Expected: TypeScript build exits with code 0.

- [x] **Step 2: Commit**

```bash
git add .
git commit -m "feat: add phase 17 complex fields references"
```

- [x] **Step 3: Push**

```bash
git push -u origin phase-17-complex-fields-references
```

Expected: branch `phase-17-complex-fields-references` is pushed to `https://github.com/codingayice/word2json.git`.

## Self-Review

- Spec coverage: This plan covers field results, TOC fields, and internal hyperlink anchors.
- Placeholder scan: No unresolved placeholders remain.
- Type consistency: Field objects carry `type` and optional `result`; hyperlinks use either `url` or `anchor`.
