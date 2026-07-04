# Phase 139 Text Run East Asia Font Family Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add high-fidelity JSON round-trip support for Word text run East Asia font family serialized as `<w:rFonts w:eastAsia="...">`.

**Architecture:** This phase adds `eastAsiaFontFamily?: string` to `StyleRunProperties`, keeping CJK font selection separate from the existing Latin `fontFamily` and complex script `complexScriptFontFamily`. The writer extends the existing `runFontsXml` helper to emit `w:eastAsia` on the same `<w:rFonts/>`, and the reader maps `w:eastAsia` back to the new JSON field.

**Tech Stack:** TypeScript, JSZip-based DOCX package tests, existing DOCX reader/writer XML helpers, Vitest.

---

### Task 1: Add RED Tests For Text Run East Asia Font Family

**Files:**
- Modify: `tests/docx-core.test.ts`

- [x] **Step 1: Write writer tests before implementation**

Add two tests near the other text run property writer tests:

```ts
  it("writes text run east asia font family", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "你好", eastAsiaFontFamily: "SimSun" }],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<w:rPr><w:rFonts w:eastAsia="SimSun"/></w:rPr><w:t>你好</w:t>');
  });

  it("writes text run latin east asia and complex script font families", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "Mixed", fontFamily: "Aptos", eastAsiaFontFamily: "SimSun", complexScriptFontFamily: "Arial" }],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<w:rPr><w:rFonts w:ascii="Aptos" w:hAnsi="Aptos" w:eastAsia="SimSun" w:cs="Arial"/></w:rPr><w:t>Mixed</w:t>');
  });
```

- [x] **Step 2: Write reader round-trip tests before implementation**

Add two tests near the other text run property round-trip tests:

```ts
  it("round-trips text run east asia font family", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "你好", eastAsiaFontFamily: "SimSun" }],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips text run latin east asia and complex script font families", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "Mixed", fontFamily: "Aptos", eastAsiaFontFamily: "SimSun", complexScriptFontFamily: "Arial" }],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });
```

- [x] **Step 3: Run writer RED test**

Run: `npm test -- tests/docx-core.test.ts -t "writes text run east asia font family"`

Expected: FAIL because `<w:rFonts w:eastAsia="...">` is not emitted yet.

- [x] **Step 4: Run reader RED test**

Run: `npm test -- tests/docx-core.test.ts -t "round-trips text run east asia font family"`

Expected: FAIL because `eastAsiaFontFamily` is not preserved through parsing yet.

### Task 2: Implement Text Run East Asia Font Family

**Files:**
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`
- Modify: `src/docx-reader.ts`

- [x] **Step 1: Add schema property**

Add the string field to `StyleRunProperties`:

```ts
  eastAsiaFontFamily?: string;
```

- [x] **Step 2: Emit writer XML**

Extend `runFontsXml` with:

```ts
    run.eastAsiaFontFamily ? ` w:eastAsia="${escapeAttribute(run.eastAsiaFontFamily)}"` : "",
```

- [x] **Step 3: Parse reader XML**

Add parsing to `parseRunFont`:

```ts
    ...(typeof fonts.eastAsia === "string" ? { eastAsiaFontFamily: fonts.eastAsia } : {}),
```

- [x] **Step 4: Run targeted GREEN tests**

Run: `npm test -- tests/docx-core.test.ts -t "text run east asia font family"`

Expected: PASS with all four targeted tests passing.

### Task 3: Verify, Commit, Push

**Files:**
- Modify: `docs/superpowers/plans/2026-07-04-phase-139-text-run-east-asia-font-family.md`

- [x] **Step 1: Run full verification**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "text run east asia font family"
npm test
npm run build
git diff --check
```

Expected: Targeted tests, full tests, and build pass. `git diff --check` may print existing LF/CRLF warnings but must exit 0.

- [x] **Step 2: Commit implementation**

```bash
git add src/schema.ts src/docx-writer.ts src/docx-reader.ts tests/docx-core.test.ts docs/superpowers/plans/2026-07-04-phase-139-text-run-east-asia-font-family.md
git commit -m "feat: add phase 139 text run east asia font family"
```

- [x] **Step 3: Push branch**

```bash
git push -u origin phase-139-text-run-east-asia-font-family
```

- [x] **Step 4: Record push metadata**

Append the pushed commit hash and remote branch to this plan, then commit the plan update:

```bash
git add docs/superpowers/plans/2026-07-04-phase-139-text-run-east-asia-font-family.md
git commit -m "docs: mark phase 139 pushed"
git push
```

## Push Record

- Branch: `phase-139-text-run-east-asia-font-family`
- Remote: `origin/phase-139-text-run-east-asia-font-family`
- Implementation commit: `4cb9a87 feat: add phase 139 text run east asia font family`
- Pushed to: `https://github.com/codingayice/word2json.git`
