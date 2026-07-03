# Phase 14 Document Settings Defaults Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Preserve common document-level settings and default style properties that influence Word rendering even when individual blocks do not specify them.

**Architecture:** Add `settings` on `DocumentJson` for `word/settings.xml`, and add `styles.defaults` for `w:docDefaults` in `word/styles.xml`. The writer emits these package parts when values exist; the reader parses the same supported subset back into JSON.

**Tech Stack:** TypeScript, Vitest, JSZip, fast-xml-parser.

---

## File Structure

- `src/schema.ts`: add `DocumentSettings` and `DocumentStyleDefaults`.
- `src/docx-writer.ts`: write `word/settings.xml`, content type and relationship entries, plus `w:docDefaults`.
- `src/docx-reader.ts`: parse `word/settings.xml` and `w:docDefaults`.
- `tests/docx-core.test.ts`: add writer XML tests and JSON -> DOCX -> JSON round-trip tests.

## Task 1: Document Settings Part

**Files:**
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`
- Modify: `src/docx-reader.ts`
- Test: `tests/docx-core.test.ts`

- [x] **Step 1: Write failing writer test**

Add this test under `describe("DOCX writer", ...)`:

```ts
it("writes document settings", async () => {
  const document = {
    version: "1.0" as const,
    settings: { defaultTabStop: 720, evenAndOddHeaders: true, updateFields: true },
    sections: [{ blocks: [{ type: "paragraph" as const, runs: [{ text: "Settings" }] }] }],
  };

  const buffer = await buildDocx(document);
  const zip = await JSZip.loadAsync(buffer);
  const settings = await zip.file("word/settings.xml")!.async("string");
  const contentTypes = await zip.file("[Content_Types].xml")!.async("string");
  const rels = await zip.file("word/_rels/document.xml.rels")!.async("string");

  expect(settings).toContain('<w:defaultTabStop w:val="720"/>');
  expect(settings).toContain("<w:evenAndOddHeaders/>");
  expect(settings).toContain("<w:updateFields/>");
  expect(contentTypes).toContain("/word/settings.xml");
  expect(rels).toContain('Target="settings.xml"');
});
```

- [x] **Step 2: Write failing round-trip test**

Add this test under `describe("DOCX reader", ...)`:

```ts
it("round-trips document settings", async () => {
  const source = {
    version: "1.0" as const,
    settings: { defaultTabStop: 720, evenAndOddHeaders: true, updateFields: true },
    sections: [{ blocks: [{ type: "paragraph" as const, runs: [{ text: "Settings" }] }] }],
  };

  const docx = await buildDocx(source);
  const parsed = await parseDocx(docx);

  expect(parsed).toEqual(source);
});
```

- [x] **Step 3: Run targeted tests**

Run: `npm test -- tests/docx-core.test.ts -t "document settings"`

Expected: FAIL because `settings` is not supported yet.

- [x] **Step 4: Implement minimal support**

Add `settings?: DocumentSettings` with `defaultTabStop?: number`, `evenAndOddHeaders?: boolean`, and `updateFields?: boolean`. Emit `word/settings.xml`, add content type and relationship entries, and parse the same fields from `word/settings.xml`.

- [x] **Step 5: Run targeted tests**

Run: `npm test -- tests/docx-core.test.ts -t "document settings"`

Expected: PASS.

## Task 2: Run Doc Defaults

**Files:**
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`
- Modify: `src/docx-reader.ts`
- Test: `tests/docx-core.test.ts`

- [x] **Step 1: Write failing writer test**

Add this test under `describe("DOCX writer", ...)`:

```ts
it("writes run document defaults", async () => {
  const document = {
    version: "1.0" as const,
    styles: {
      defaults: {
        run: { fontFamily: "Aptos", fontSize: 11, color: "1F1F1F" },
      },
    },
    sections: [{ blocks: [{ type: "paragraph" as const, runs: [{ text: "Defaults" }] }] }],
  };

  const buffer = await buildDocx(document);
  const zip = await JSZip.loadAsync(buffer);
  const styles = await zip.file("word/styles.xml")!.async("string");

  expect(styles).toContain("<w:docDefaults>");
  expect(styles).toContain('<w:rPrDefault><w:rPr><w:rFonts w:ascii="Aptos" w:hAnsi="Aptos"/><w:sz w:val="22"/><w:color w:val="1F1F1F"/></w:rPr></w:rPrDefault>');
});
```

- [x] **Step 2: Write failing round-trip test**

Add this test under `describe("DOCX reader", ...)`:

```ts
it("round-trips run document defaults", async () => {
  const source = {
    version: "1.0" as const,
    styles: {
      defaults: {
        run: { fontFamily: "Aptos", fontSize: 11, color: "1F1F1F" },
      },
    },
    sections: [{ blocks: [{ type: "paragraph" as const, runs: [{ text: "Defaults" }] }] }],
  };

  const docx = await buildDocx(source);
  const parsed = await parseDocx(docx);

  expect(parsed).toEqual(source);
});
```

- [x] **Step 3: Run targeted tests**

Run: `npm test -- tests/docx-core.test.ts -t "run document defaults"`

Expected: FAIL because `styles.defaults.run` is not supported yet.

- [x] **Step 4: Implement minimal support**

Add `defaults?: { run?: StyleRunProperties; paragraph?: StyleParagraphProperties }` to `DocumentStyles`. Emit run defaults as `w:docDefaults/w:rPrDefault/w:rPr` using the existing run property helper, and parse them with `parseStyleRunProperties`.

- [x] **Step 5: Run targeted tests**

Run: `npm test -- tests/docx-core.test.ts -t "run document defaults"`

Expected: PASS.

## Task 3: Paragraph Doc Defaults

**Files:**
- Modify: `src/docx-writer.ts`
- Modify: `src/docx-reader.ts`
- Test: `tests/docx-core.test.ts`

- [x] **Step 1: Write failing writer test**

Add this test under `describe("DOCX writer", ...)`:

```ts
it("writes paragraph document defaults", async () => {
  const document = {
    version: "1.0" as const,
    styles: {
      defaults: {
        paragraph: {
          spacing: { after: 160, line: 276, lineRule: "auto" as const },
          indent: { firstLine: 420 },
        },
      },
    },
    sections: [{ blocks: [{ type: "paragraph" as const, runs: [{ text: "Paragraph defaults" }] }] }],
  };

  const buffer = await buildDocx(document);
  const zip = await JSZip.loadAsync(buffer);
  const styles = await zip.file("word/styles.xml")!.async("string");

  expect(styles).toContain('<w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="276" w:lineRule="auto"/><w:ind w:firstLine="420"/></w:pPr></w:pPrDefault>');
});
```

- [x] **Step 2: Write failing round-trip test**

Add this test under `describe("DOCX reader", ...)`:

```ts
it("round-trips paragraph document defaults", async () => {
  const source = {
    version: "1.0" as const,
    styles: {
      defaults: {
        paragraph: {
          spacing: { after: 160, line: 276, lineRule: "auto" as const },
          indent: { firstLine: 420 },
        },
      },
    },
    sections: [{ blocks: [{ type: "paragraph" as const, runs: [{ text: "Paragraph defaults" }] }] }],
  };

  const docx = await buildDocx(source);
  const parsed = await parseDocx(docx);

  expect(parsed).toEqual(source);
});
```

- [x] **Step 3: Run targeted tests**

Run: `npm test -- tests/docx-core.test.ts -t "paragraph document defaults"`

Expected: FAIL because `styles.defaults.paragraph` is not supported yet.

- [x] **Step 4: Implement minimal support**

Reuse paragraph property XML and parsing helpers for `w:docDefaults/w:pPrDefault/w:pPr`.

- [x] **Step 5: Run targeted tests**

Run: `npm test -- tests/docx-core.test.ts -t "paragraph document defaults"`

Expected: PASS.

## Task 4: Full Verification and Push

**Files:**
- Modify: all Phase 14 files

- [x] **Step 1: Run full verification**

Run: `npm test`

Expected: all tests pass.

Run: `npm run build`

Expected: TypeScript build exits with code 0.

- [x] **Step 2: Commit**

```bash
git add .
git commit -m "feat: add phase 14 document settings defaults"
```

- [x] **Step 3: Push**

```bash
git push -u origin phase-14-document-settings-defaults
```

Expected: branch `phase-14-document-settings-defaults` is pushed to `https://github.com/codingayice/word2json.git`.

## Self-Review

- Spec coverage: This plan covers document settings and both run and paragraph doc defaults.
- Placeholder scan: No unresolved placeholders remain.
- Type consistency: `settings` lives on `DocumentJson`; `defaults` lives on `DocumentStyles`.
