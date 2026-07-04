# Phase 141 Text Run Theme Fonts Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add high-fidelity JSON round-trip support for Word text run theme font metadata serialized on `<w:rFonts>`.

**Architecture:** This phase adds explicit theme font fields to `StyleRunProperties` beside the existing concrete font family fields. The writer extends the existing `runFontsXml` helper to emit `w:asciiTheme/w:hAnsiTheme`, `w:eastAsiaTheme`, and `w:cstheme`, while the reader maps those attributes back into the same JSON fields.

**Tech Stack:** TypeScript, JSZip-based DOCX package tests, existing DOCX reader/writer XML helpers, Vitest.

---

### Task 1: Add RED Tests For Text Run Theme Fonts

**Files:**
- Modify: `tests/docx-core.test.ts`

- [x] **Step 1: Write writer tests before implementation**

Add two tests near the other text run font writer tests:

```ts
  it("writes text run theme fonts", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "Theme", fontTheme: "minorHAnsi", eastAsiaFontTheme: "minorEastAsia", complexScriptFontTheme: "minorBidi" }],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<w:rPr><w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi" w:eastAsiaTheme="minorEastAsia" w:cstheme="minorBidi"/></w:rPr><w:t>Theme</w:t>');
  });

  it("writes text run theme fonts with direct fonts and hint", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [{
          text: "Mixed",
          fontFamily: "Aptos",
          eastAsiaFontFamily: "SimSun",
          complexScriptFontFamily: "Arial",
          fontTheme: "majorHAnsi",
          eastAsiaFontTheme: "majorEastAsia",
          complexScriptFontTheme: "majorBidi",
          fontHint: "eastAsia",
        }],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<w:rPr><w:rFonts w:ascii="Aptos" w:hAnsi="Aptos" w:eastAsia="SimSun" w:cs="Arial" w:asciiTheme="majorHAnsi" w:hAnsiTheme="majorHAnsi" w:eastAsiaTheme="majorEastAsia" w:cstheme="majorBidi" w:hint="eastAsia"/></w:rPr><w:t>Mixed</w:t>');
  });
```

- [x] **Step 2: Write reader round-trip tests before implementation**

Add two tests near the other text run font round-trip tests:

```ts
  it("round-trips text run theme fonts", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "Theme", fontTheme: "minorHAnsi", eastAsiaFontTheme: "minorEastAsia", complexScriptFontTheme: "minorBidi" }],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips text run theme fonts with direct fonts and hint", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [{
          text: "Mixed",
          fontFamily: "Aptos",
          eastAsiaFontFamily: "SimSun",
          complexScriptFontFamily: "Arial",
          fontTheme: "majorHAnsi",
          eastAsiaFontTheme: "majorEastAsia",
          complexScriptFontTheme: "majorBidi",
          fontHint: "eastAsia",
        }],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });
```

- [x] **Step 3: Run writer RED test**

Run: `npm test -- tests/docx-core.test.ts -t "writes text run theme fonts"`

Expected: FAIL because `<w:rFonts>` does not emit theme font attributes yet.

- [x] **Step 4: Run reader RED test**

Run: `npm test -- tests/docx-core.test.ts -t "round-trips text run theme fonts"`

Expected: FAIL because theme font attributes are not preserved through parsing yet.

### Task 2: Implement Text Run Theme Fonts

**Files:**
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`
- Modify: `src/docx-reader.ts`

- [x] **Step 1: Add schema fields**

Add the fields to `StyleRunProperties`:

```ts
  fontTheme?: string;
  eastAsiaFontTheme?: string;
  complexScriptFontTheme?: string;
```

- [x] **Step 2: Emit writer XML**

Extend `runFontsXml` with:

```ts
    run.fontTheme ? ` w:asciiTheme="${escapeAttribute(run.fontTheme)}" w:hAnsiTheme="${escapeAttribute(run.fontTheme)}"` : "",
    run.eastAsiaFontTheme ? ` w:eastAsiaTheme="${escapeAttribute(run.eastAsiaFontTheme)}"` : "",
    run.complexScriptFontTheme ? ` w:cstheme="${escapeAttribute(run.complexScriptFontTheme)}"` : "",
```

- [x] **Step 3: Parse reader XML**

Add parsing to `parseRunFont`:

```ts
    ...(typeof fonts.asciiTheme === "string" ? { fontTheme: fonts.asciiTheme } : {}),
    ...(typeof fonts.eastAsiaTheme === "string" ? { eastAsiaFontTheme: fonts.eastAsiaTheme } : {}),
    ...(typeof fonts.cstheme === "string" ? { complexScriptFontTheme: fonts.cstheme } : {}),
```

- [x] **Step 4: Run targeted GREEN tests**

Run: `npm test -- tests/docx-core.test.ts -t "text run theme fonts"`

Expected: PASS with all four targeted tests passing.

### Task 3: Verify, Commit, Push

**Files:**
- Modify: `docs/superpowers/plans/2026-07-04-phase-141-text-run-theme-fonts.md`

- [x] **Step 1: Run full verification**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "text run theme fonts"
npm test
npm run build
git diff --check
```

Expected: Targeted tests, full tests, and build pass. `git diff --check` may print existing LF/CRLF warnings but must exit 0.

- [ ] **Step 2: Commit implementation**

```bash
git add src/schema.ts src/docx-writer.ts src/docx-reader.ts tests/docx-core.test.ts docs/superpowers/plans/2026-07-04-phase-141-text-run-theme-fonts.md
git commit -m "feat: add phase 141 text run theme fonts"
```

- [ ] **Step 3: Push branch**

```bash
git push -u origin phase-141-text-run-theme-fonts
```

- [ ] **Step 4: Record push metadata**

Append the pushed commit hash and remote branch to this plan, then commit the plan update:

```bash
git add docs/superpowers/plans/2026-07-04-phase-141-text-run-theme-fonts.md
git commit -m "docs: mark phase 141 pushed"
git push
```

## Push Record

- Pending.
