# Phase 138 Text Run Complex Script Font Family Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add high-fidelity JSON round-trip support for Word text run complex script font family serialized as `<w:rFonts w:cs="...">`.

**Architecture:** This phase adds `complexScriptFontFamily?: string` to `StyleRunProperties`, keeping it separate from the existing Latin `fontFamily` mapping. The writer uses one `rFonts` helper so `fontFamily` can emit `w:ascii/w:hAnsi`, `complexScriptFontFamily` can emit `w:cs`, and both can coexist on the same `<w:rFonts/>`; the reader maps `w:cs` back to the new JSON field.

**Tech Stack:** TypeScript, JSZip-based DOCX package tests, existing DOCX reader/writer XML helpers, Vitest.

---

### Task 1: Add RED Tests For Text Run Complex Script Font Family

**Files:**
- Modify: `tests/docx-core.test.ts`

- [x] **Step 1: Write writer tests before implementation**

Add two tests near the other text run property writer tests:

```ts
  it("writes text run complex script font family", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "مرحبا", complexScriptFontFamily: "Arial" }],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<w:rPr><w:rFonts w:cs="Arial"/></w:rPr><w:t>مرحبا</w:t>');
  });

  it("writes text run latin and complex script font families", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "Mixed", fontFamily: "Aptos", complexScriptFontFamily: "Arial" }],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<w:rPr><w:rFonts w:ascii="Aptos" w:hAnsi="Aptos" w:cs="Arial"/></w:rPr><w:t>Mixed</w:t>');
  });
```

- [x] **Step 2: Write reader round-trip tests before implementation**

Add two tests near the other text run property round-trip tests:

```ts
  it("round-trips text run complex script font family", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "مرحبا", complexScriptFontFamily: "Arial" }],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips text run latin and complex script font families", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "Mixed", fontFamily: "Aptos", complexScriptFontFamily: "Arial" }],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });
```

- [x] **Step 3: Run writer RED test**

Run: `npm test -- tests/docx-core.test.ts -t "writes text run complex script font family"`

Expected: FAIL because `<w:rFonts w:cs="...">` is not emitted yet.

- [x] **Step 4: Run reader RED test**

Run: `npm test -- tests/docx-core.test.ts -t "round-trips text run complex script font family"`

Expected: FAIL because `complexScriptFontFamily` is not preserved through parsing yet.

### Task 2: Implement Text Run Complex Script Font Family

**Files:**
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`
- Modify: `src/docx-reader.ts`

- [x] **Step 1: Add schema property**

Add the string field to `StyleRunProperties`:

```ts
  complexScriptFontFamily?: string;
```

- [x] **Step 2: Emit writer XML**

Replace the inline `fontFamily` `rFonts` string in `runPropertiesXml` with:

```ts
    runFontsXml(run),
```

Add the helper:

```ts
function runFontsXml(run: TextRun): string {
  const attributes = [
    run.fontFamily ? ` w:ascii="${escapeAttribute(run.fontFamily)}" w:hAnsi="${escapeAttribute(run.fontFamily)}"` : "",
    run.complexScriptFontFamily ? ` w:cs="${escapeAttribute(run.complexScriptFontFamily)}"` : "",
  ].join("");

  return attributes ? `<w:rFonts${attributes}/>` : "";
}
```

- [x] **Step 3: Parse reader XML**

Add parsing to `parseRunFont`:

```ts
    ...(typeof fonts.cs === "string" ? { complexScriptFontFamily: fonts.cs } : {}),
```

- [x] **Step 4: Run targeted GREEN tests**

Run: `npm test -- tests/docx-core.test.ts -t "text run complex script font family"`

Expected: PASS with all four targeted tests passing.

### Task 3: Verify, Commit, Push

**Files:**
- Modify: `docs/superpowers/plans/2026-07-04-phase-138-text-run-complex-script-font-family.md`

- [x] **Step 1: Run full verification**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "text run complex script font family"
npm test
npm run build
git diff --check
```

Expected: Targeted tests, full tests, and build pass. `git diff --check` may print existing LF/CRLF warnings but must exit 0.

- [ ] **Step 2: Commit implementation**

```bash
git add src/schema.ts src/docx-writer.ts src/docx-reader.ts tests/docx-core.test.ts docs/superpowers/plans/2026-07-04-phase-138-text-run-complex-script-font-family.md
git commit -m "feat: add phase 138 text run complex script font family"
```

- [ ] **Step 3: Push branch**

```bash
git push -u origin phase-138-text-run-complex-script-font-family
```

- [ ] **Step 4: Record push metadata**

Append the pushed commit hash and remote branch to this plan, then commit the plan update:

```bash
git add docs/superpowers/plans/2026-07-04-phase-138-text-run-complex-script-font-family.md
git commit -m "docs: mark phase 138 pushed"
git push
```

## Push Record

- Pending.
