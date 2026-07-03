# Phase 13 Borders Shading Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Preserve common visual borders and shading for paragraphs, direct runs, and paragraph styles.

**Architecture:** Add small reusable border and shading value objects to the JSON schema. The writer emits them into `w:pPr`, `w:rPr`, and paragraph style `w:pPr`; the reader parses the same subset from direct paragraphs, runs, and paragraph styles.

**Tech Stack:** TypeScript, Vitest, JSZip, fast-xml-parser.

---

## File Structure

- `src/schema.ts`: add `BorderDefinition`, `ShadingDefinition`, paragraph `borders`/`shading`, run `border`, and style paragraph `borders`/`shading`.
- `src/docx-writer.ts`: emit paragraph borders, paragraph shading, run border, and paragraph style borders/shading.
- `src/docx-reader.ts`: parse paragraph borders, paragraph shading, run border, and paragraph style borders/shading.
- `tests/docx-core.test.ts`: add writer XML tests and JSON -> DOCX -> JSON round-trip tests.

## Task 1: Paragraph Shading

**Files:**
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`
- Modify: `src/docx-reader.ts`
- Test: `tests/docx-core.test.ts`

- [x] **Step 1: Write failing writer test**

Add this test under `describe("DOCX writer", ...)`:

```ts
it("writes paragraph shading", async () => {
  const document = createDocumentJson([
    {
      type: "paragraph",
      shading: { fill: "FFF2CC" },
      runs: [{ text: "Highlighted paragraph" }],
    },
  ]);

  const buffer = await buildDocx(document);
  const zip = await JSZip.loadAsync(buffer);
  const xml = await zip.file("word/document.xml")!.async("string");

  expect(xml).toContain('<w:shd w:fill="FFF2CC"/>');
});
```

- [x] **Step 2: Write failing round-trip test**

Add this test under `describe("DOCX round-trip", ...)`:

```ts
it("round-trips paragraph shading", async () => {
  const source = createDocumentJson([
    {
      type: "paragraph",
      shading: { fill: "FFF2CC" },
      runs: [{ text: "Highlighted paragraph" }],
    },
  ]);

  const docx = await buildDocx(source);
  const parsed = await parseDocx(docx);

  expect(parsed).toEqual(source);
});
```

- [x] **Step 3: Run targeted tests**

Run: `npm test -- tests/docx-core.test.ts -t "paragraph shading"`

Expected: FAIL because paragraph `shading` is not supported yet.

- [x] **Step 4: Implement minimal support**

Add `ShadingDefinition` and `shading?: ShadingDefinition` to `ParagraphNode` and `StyleParagraphProperties`. Emit `<w:shd w:fill="..."/>` inside paragraph `w:pPr`, and parse `w:shd` using the `fill` attribute.

- [x] **Step 5: Run targeted tests**

Run: `npm test -- tests/docx-core.test.ts -t "paragraph shading"`

Expected: PASS.

## Task 2: Paragraph Borders

**Files:**
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`
- Modify: `src/docx-reader.ts`
- Test: `tests/docx-core.test.ts`

- [x] **Step 1: Write failing writer test**

Add this test under `describe("DOCX writer", ...)`:

```ts
it("writes paragraph borders", async () => {
  const document = createDocumentJson([
    {
      type: "paragraph",
      borders: {
        top: { style: "single", size: 8, color: "4472C4", space: 2 },
        bottom: { style: "single", size: 8, color: "4472C4", space: 2 },
      },
      runs: [{ text: "Bordered paragraph" }],
    },
  ]);

  const buffer = await buildDocx(document);
  const zip = await JSZip.loadAsync(buffer);
  const xml = await zip.file("word/document.xml")!.async("string");

  expect(xml).toContain("<w:pBdr>");
  expect(xml).toContain('<w:top w:val="single" w:sz="8" w:space="2" w:color="4472C4"/>');
  expect(xml).toContain('<w:bottom w:val="single" w:sz="8" w:space="2" w:color="4472C4"/>');
});
```

- [x] **Step 2: Write failing round-trip test**

Add this test under `describe("DOCX round-trip", ...)`:

```ts
it("round-trips paragraph borders", async () => {
  const source = createDocumentJson([
    {
      type: "paragraph",
      borders: {
        top: { style: "single", size: 8, color: "4472C4", space: 2 },
        bottom: { style: "single", size: 8, color: "4472C4", space: 2 },
      },
      runs: [{ text: "Bordered paragraph" }],
    },
  ]);

  const docx = await buildDocx(source);
  const parsed = await parseDocx(docx);

  expect(parsed).toEqual(source);
});
```

- [x] **Step 3: Run targeted tests**

Run: `npm test -- tests/docx-core.test.ts -t "paragraph borders"`

Expected: FAIL because paragraph `borders` is not supported yet.

- [x] **Step 4: Implement minimal support**

Add `BorderDefinition` and `ParagraphBorders` with `top`, `left`, `bottom`, and `right` sides. Emit `<w:pBdr>` with supported sides and parse them back from direct paragraph properties.

- [x] **Step 5: Run targeted tests**

Run: `npm test -- tests/docx-core.test.ts -t "paragraph borders"`

Expected: PASS.

## Task 3: Run Border

**Files:**
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`
- Modify: `src/docx-reader.ts`
- Test: `tests/docx-core.test.ts`

- [x] **Step 1: Write failing writer test**

Add this test under `describe("DOCX writer", ...)`:

```ts
it("writes run border", async () => {
  const document = createDocumentJson([
    {
      type: "paragraph",
      runs: [
        { text: "Boxed", border: { style: "single", size: 6, color: "C00000", space: 1 } },
      ],
    },
  ]);

  const buffer = await buildDocx(document);
  const zip = await JSZip.loadAsync(buffer);
  const xml = await zip.file("word/document.xml")!.async("string");

  expect(xml).toContain('<w:bdr w:val="single" w:sz="6" w:space="1" w:color="C00000"/>');
});
```

- [x] **Step 2: Write failing round-trip test**

Add this test under `describe("DOCX round-trip", ...)`:

```ts
it("round-trips run border", async () => {
  const source = createDocumentJson([
    {
      type: "paragraph",
      runs: [
        { text: "Boxed", border: { style: "single", size: 6, color: "C00000", space: 1 } },
      ],
    },
  ]);

  const docx = await buildDocx(source);
  const parsed = await parseDocx(docx);

  expect(parsed).toEqual(source);
});
```

- [x] **Step 3: Run targeted tests**

Run: `npm test -- tests/docx-core.test.ts -t "run border"`

Expected: FAIL because run `border` is not supported yet.

- [x] **Step 4: Implement minimal support**

Add `border?: BorderDefinition` to `StyleRunProperties`. Emit `<w:bdr>` in direct run and style run properties, and parse it back from `w:rPr`.

- [x] **Step 5: Run targeted tests**

Run: `npm test -- tests/docx-core.test.ts -t "run border"`

Expected: PASS.

## Task 4: Paragraph Style Borders and Shading

**Files:**
- Modify: `src/docx-writer.ts`
- Modify: `src/docx-reader.ts`
- Test: `tests/docx-core.test.ts`

- [x] **Step 1: Write failing writer test**

Add this test under `describe("DOCX writer", ...)`:

```ts
it("writes paragraph style borders and shading", async () => {
  const document = {
    version: "1.0" as const,
    styles: {
      paragraph: [
        {
          id: "Callout",
          name: "Callout",
          paragraph: {
            shading: { fill: "E2F0D9" },
            borders: { left: { style: "single", size: 12, color: "70AD47", space: 4 } },
          },
        },
      ],
    },
    sections: [
      {
        blocks: [{ type: "paragraph" as const, styleId: "Callout", runs: [{ text: "Styled callout" }] }],
      },
    ],
  };

  const buffer = await buildDocx(document);
  const zip = await JSZip.loadAsync(buffer);
  const styles = await zip.file("word/styles.xml")!.async("string");

  expect(styles).toContain('<w:shd w:fill="E2F0D9"/>');
  expect(styles).toContain('<w:left w:val="single" w:sz="12" w:space="4" w:color="70AD47"/>');
});
```

- [x] **Step 2: Write failing round-trip test**

Add this test under `describe("DOCX round-trip", ...)`:

```ts
it("round-trips paragraph style borders and shading", async () => {
  const source = {
    version: "1.0" as const,
    styles: {
      paragraph: [
        {
          id: "Callout",
          name: "Callout",
          paragraph: {
            shading: { fill: "E2F0D9" },
            borders: { left: { style: "single", size: 12, color: "70AD47", space: 4 } },
          },
        },
      ],
    },
    sections: [
      {
        blocks: [{ type: "paragraph" as const, styleId: "Callout", runs: [{ text: "Styled callout" }] }],
      },
    ],
  };

  const docx = await buildDocx(source);
  const parsed = await parseDocx(docx);

  expect(parsed).toEqual(source);
});
```

- [x] **Step 3: Run targeted tests**

Run: `npm test -- tests/docx-core.test.ts -t "paragraph style borders and shading"`

Expected: FAIL until style `w:pPr` shares the same paragraph visual property support.

- [x] **Step 4: Implement minimal support**

Reuse the paragraph border and shading helpers in `paragraphStylePropertiesXml` and `parseStyleParagraphProperties`.

- [x] **Step 5: Run targeted tests**

Run: `npm test -- tests/docx-core.test.ts -t "paragraph style borders and shading"`

Expected: PASS.

## Task 5: Full Verification and Push

**Files:**
- Modify: all Phase 13 files

- [x] **Step 1: Run full verification**

Run: `npm test`

Expected: all tests pass.

Run: `npm run build`

Expected: TypeScript build exits with code 0.

- [x] **Step 2: Commit**

```bash
git add .
git commit -m "feat: add phase 13 borders and shading"
```

- [x] **Step 3: Push**

```bash
git push -u origin phase-13-borders-shading
```

Expected: branch `phase-13-borders-shading` is pushed to `https://github.com/codingayice/word2json.git`.

## Self-Review

- Spec coverage: This plan covers direct paragraph shading, direct paragraph borders, direct run borders, and reusable paragraph style borders/shading.
- Placeholder scan: No unresolved placeholders remain.
- Type consistency: Border fields use `style`, `size`, `color`, and `space`; shading uses `fill`.
