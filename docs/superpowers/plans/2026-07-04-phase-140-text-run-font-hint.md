# Phase 140 Text Run Font Hint Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add high-fidelity JSON round-trip support for Word text run font hint metadata serialized as `<w:rFonts w:hint="...">`.

**Architecture:** This phase adds `fontHint?: RunFontHint` to `StyleRunProperties`, modeling the WordprocessingML hint values `default`, `eastAsia`, and `cs`. The writer extends `runFontsXml` so `fontHint` can coexist with Latin, East Asia, and complex script font family attributes, and the reader maps `w:hint` back to the new JSON field.

**Tech Stack:** TypeScript, JSZip-based DOCX package tests, existing DOCX reader/writer XML helpers, Vitest.

---

### Task 1: Add RED Tests For Text Run Font Hint

**Files:**
- Modify: `tests/docx-core.test.ts`

- [x] **Step 1: Write writer tests before implementation**

Add two tests near the other text run property writer tests:

```ts
  it("writes text run font hint east asia", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "你好", fontHint: "eastAsia" }],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr><w:t>你好</w:t>');
  });

  it("writes text run font hint with all font families", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "Mixed", fontFamily: "Aptos", eastAsiaFontFamily: "SimSun", complexScriptFontFamily: "Arial", fontHint: "cs" }],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<w:rPr><w:rFonts w:ascii="Aptos" w:hAnsi="Aptos" w:eastAsia="SimSun" w:cs="Arial" w:hint="cs"/></w:rPr><w:t>Mixed</w:t>');
  });
```

- [x] **Step 2: Write reader round-trip tests before implementation**

Add two tests near the other text run property round-trip tests:

```ts
  it("round-trips text run font hint east asia", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "你好", fontHint: "eastAsia" }],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips text run font hint with all font families", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "Mixed", fontFamily: "Aptos", eastAsiaFontFamily: "SimSun", complexScriptFontFamily: "Arial", fontHint: "cs" }],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });
```

- [x] **Step 3: Run writer RED test**

Run: `npm test -- tests/docx-core.test.ts -t "writes text run font hint"`

Expected: FAIL because `<w:rFonts w:hint="...">` is not emitted yet.

- [x] **Step 4: Run reader RED test**

Run: `npm test -- tests/docx-core.test.ts -t "round-trips text run font hint"`

Expected: FAIL because `fontHint` is not preserved through parsing yet.

### Task 2: Implement Text Run Font Hint

**Files:**
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`
- Modify: `src/docx-reader.ts`

- [x] **Step 1: Add schema type and property**

Add the field to `StyleRunProperties`:

```ts
  fontHint?: RunFontHint;
```

Add the union type near other run-related types:

```ts
export type RunFontHint = "default" | "eastAsia" | "cs";
```

- [x] **Step 2: Emit writer XML**

Extend `runFontsXml` with:

```ts
    run.fontHint ? ` w:hint="${run.fontHint}"` : "",
```

- [x] **Step 3: Parse reader XML**

Add parsing to `parseRunFont`:

```ts
    ...(typeof fonts.hint === "string" ? { fontHint: fonts.hint as NonNullable<TextRun["fontHint"]> } : {}),
```

- [x] **Step 4: Run targeted GREEN tests**

Run: `npm test -- tests/docx-core.test.ts -t "text run font hint"`

Expected: PASS with all four targeted tests passing.

### Task 3: Verify, Commit, Push

**Files:**
- Modify: `docs/superpowers/plans/2026-07-04-phase-140-text-run-font-hint.md`

- [x] **Step 1: Run full verification**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "text run font hint"
npm test
npm run build
git diff --check
```

Expected: Targeted tests, full tests, and build pass. `git diff --check` may print existing LF/CRLF warnings but must exit 0.

- [x] **Step 2: Commit implementation**

```bash
git add src/schema.ts src/docx-writer.ts src/docx-reader.ts tests/docx-core.test.ts docs/superpowers/plans/2026-07-04-phase-140-text-run-font-hint.md
git commit -m "feat: add phase 140 text run font hint"
```

- [x] **Step 3: Push branch**

```bash
git push -u origin phase-140-text-run-font-hint
```

- [x] **Step 4: Record push metadata**

Append the pushed commit hash and remote branch to this plan, then commit the plan update:

```bash
git add docs/superpowers/plans/2026-07-04-phase-140-text-run-font-hint.md
git commit -m "docs: mark phase 140 pushed"
git push
```

## Push Record

- Branch: `phase-140-text-run-font-hint`
- Remote: `origin`
- Repository: `https://github.com/codingayice/word2json.git`
- Implementation commit: `4548a8f8a6acc88cd480dfef26b0b64afac2e9d5`
- Pull request URL: `https://github.com/codingayice/word2json/pull/new/phase-140-text-run-font-hint`
