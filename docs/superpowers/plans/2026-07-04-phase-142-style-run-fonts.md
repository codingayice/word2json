# Phase 142 Style Run Fonts Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Preserve high-fidelity run font metadata inside Word style definitions, not only inline text runs.

**Architecture:** `StyleRunProperties` already exposes direct font, theme font, and font hint fields. This phase updates the style writer and style reader so `<w:style><w:rPr><w:rFonts .../></w:rPr></w:style>` round-trips `w:eastAsia`, `w:cs`, `w:asciiTheme/w:hAnsiTheme`, `w:eastAsiaTheme`, `w:cstheme`, and `w:hint`.

**Tech Stack:** TypeScript, JSZip-based DOCX package tests, existing DOCX reader/writer XML helpers, Vitest.

---

### Task 1: Add RED Tests For Style Run Fonts

**Files:**
- Modify: `tests/docx-core.test.ts`

- [x] **Step 1: Write writer tests before implementation**

Add a writer test near the existing custom style writer tests:

```ts
  it("writes style run fonts with direct theme and hint attributes", async () => {
    const document = createDocumentJson([{ type: "paragraph", runs: [{ text: "Term", styleId: "DefinedTerm" }] }], {
      styles: {
        character: [{
          id: "DefinedTerm",
          name: "Defined Term",
          run: {
            fontFamily: "Aptos",
            eastAsiaFontFamily: "SimSun",
            complexScriptFontFamily: "Arial",
            fontTheme: "majorHAnsi",
            eastAsiaFontTheme: "majorEastAsia",
            complexScriptFontTheme: "majorBidi",
            fontHint: "eastAsia",
          },
        }],
      },
    });

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const styles = await zip.file("word/styles.xml")!.async("string");

    expect(styles).toContain('<w:style w:type="character" w:styleId="DefinedTerm"><w:name w:val="Defined Term"/><w:rPr><w:rFonts w:ascii="Aptos" w:hAnsi="Aptos" w:eastAsia="SimSun" w:cs="Arial" w:asciiTheme="majorHAnsi" w:hAnsiTheme="majorHAnsi" w:eastAsiaTheme="majorEastAsia" w:cstheme="majorBidi" w:hint="eastAsia"/></w:rPr></w:style>');
  });
```

- [x] **Step 2: Write reader round-trip test before implementation**

Add a reader test near the existing custom style round-trip tests:

```ts
  it("round-trips style run fonts with direct theme and hint attributes", async () => {
    const source = createDocumentJson([{ type: "paragraph", runs: [{ text: "Term", styleId: "DefinedTerm" }] }], {
      styles: {
        character: [{
          id: "DefinedTerm",
          name: "Defined Term",
          run: {
            fontFamily: "Aptos",
            eastAsiaFontFamily: "SimSun",
            complexScriptFontFamily: "Arial",
            fontTheme: "majorHAnsi",
            eastAsiaFontTheme: "majorEastAsia",
            complexScriptFontTheme: "majorBidi",
            fontHint: "eastAsia",
          },
        }],
      },
    });

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });
```

- [x] **Step 3: Run writer RED test**

Run: `npm test -- tests/docx-core.test.ts -t "writes style run fonts"`

Expected: FAIL because style `<w:rFonts>` only emits `w:ascii/w:hAnsi`.

- [x] **Step 4: Run reader RED test**

Run: `npm test -- tests/docx-core.test.ts -t "round-trips style run fonts"`

Expected: FAIL because style parsing drops the East Asia, complex script, theme, and hint font fields.

### Task 2: Implement Style Run Fonts

**Files:**
- Modify: `src/docx-writer.ts`
- Modify: `src/docx-reader.ts`

- [x] **Step 1: Reuse the existing run font writer helper for styles**

Replace the inline `fontFamily` `<w:rFonts>` string inside `styleRunPropertiesXml` with:

```ts
    styleRunFontsXml(run),
```

Add a helper near `styleRunPropertiesXml`:

```ts
function styleRunFontsXml(run: StyleRunProperties): string {
  const attributes = [
    run.fontFamily ? ` w:ascii="${escapeAttribute(run.fontFamily)}" w:hAnsi="${escapeAttribute(run.fontFamily)}"` : "",
    run.eastAsiaFontFamily ? ` w:eastAsia="${escapeAttribute(run.eastAsiaFontFamily)}"` : "",
    run.complexScriptFontFamily ? ` w:cs="${escapeAttribute(run.complexScriptFontFamily)}"` : "",
    run.fontTheme ? ` w:asciiTheme="${escapeAttribute(run.fontTheme)}" w:hAnsiTheme="${escapeAttribute(run.fontTheme)}"` : "",
    run.eastAsiaFontTheme ? ` w:eastAsiaTheme="${escapeAttribute(run.eastAsiaFontTheme)}"` : "",
    run.complexScriptFontTheme ? ` w:cstheme="${escapeAttribute(run.complexScriptFontTheme)}"` : "",
    run.fontHint ? ` w:hint="${run.fontHint}"` : "",
  ].join("");

  return attributes ? `<w:rFonts${attributes}/>` : "";
}
```

- [x] **Step 2: Parse style run font attributes**

Extend `parseStyleRunProperties` with:

```ts
    ...(typeof fonts.eastAsia === "string" ? { eastAsiaFontFamily: fonts.eastAsia } : {}),
    ...(typeof fonts.cs === "string" ? { complexScriptFontFamily: fonts.cs } : {}),
    ...(typeof fonts.asciiTheme === "string" ? { fontTheme: fonts.asciiTheme } : {}),
    ...(typeof fonts.eastAsiaTheme === "string" ? { eastAsiaFontTheme: fonts.eastAsiaTheme } : {}),
    ...(typeof fonts.cstheme === "string" ? { complexScriptFontTheme: fonts.cstheme } : {}),
    ...(typeof fonts.hint === "string" ? { fontHint: fonts.hint as NonNullable<StyleRunProperties["fontHint"]> } : {}),
```

- [x] **Step 3: Run targeted GREEN tests**

Run: `npm test -- tests/docx-core.test.ts -t "style run fonts"`

Expected: PASS with both targeted tests passing.

### Task 3: Verify, Commit, Push

**Files:**
- Modify: `docs/superpowers/plans/2026-07-04-phase-142-style-run-fonts.md`

- [x] **Step 1: Run full verification**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "style run fonts"
npm test
npm run build
git diff --check
```

Expected: Targeted tests, full tests, and build pass. `git diff --check` may print existing LF/CRLF warnings but must exit 0.

- [ ] **Step 2: Commit implementation**

```bash
git add src/docx-writer.ts src/docx-reader.ts tests/docx-core.test.ts docs/superpowers/plans/2026-07-04-phase-142-style-run-fonts.md
git commit -m "feat: add phase 142 style run fonts"
```

- [ ] **Step 3: Push branch**

```bash
git push -u origin phase-142-style-run-fonts
```

- [ ] **Step 4: Record push metadata**

Append the pushed commit hash and remote branch to this plan, then commit the plan update:

```bash
git add docs/superpowers/plans/2026-07-04-phase-142-style-run-fonts.md
git commit -m "docs: mark phase 142 pushed"
git push
```

## Push Record

- Pending.
