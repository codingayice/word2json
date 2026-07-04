# Phase 143 Style Run Advanced Typography Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Preserve advanced run typography metadata inside Word style definitions, matching existing inline text run behavior.

**Architecture:** `StyleRunProperties` already exposes `complexScriptFontSize`, `language`, `characterPosition`, and `kerning`. This phase updates style run XML writing and parsing so style `<w:rPr>` can round-trip `<w:szCs>`, `<w:lang>`, `<w:position>`, and `<w:kern>` just like inline text runs.

**Tech Stack:** TypeScript, JSZip-based DOCX package tests, existing DOCX reader/writer XML helpers, Vitest.

---

### Task 1: Add RED Tests For Style Run Advanced Typography

**Files:**
- Modify: `tests/docx-core.test.ts`

- [x] **Step 1: Write writer test before implementation**

Add a writer test near the existing style run tests:

```ts
  it("writes style run advanced typography properties", async () => {
    const document = {
      version: "1.0" as const,
      styles: {
        character: [{
          id: "EmphasisText",
          name: "Emphasis Text",
          run: {
            complexScriptFontSize: 14,
            language: { value: "en-US", eastAsia: "zh-CN", bidi: "ar-SA" },
            characterPosition: 4,
            kerning: 28,
          },
        }],
      },
      sections: [{ blocks: [{ type: "paragraph" as const, runs: [{ text: "Term", styleId: "EmphasisText" }] }] }],
    };

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const styles = await zip.file("word/styles.xml")!.async("string");

    expect(styles).toContain('<w:rPr><w:szCs w:val="28"/><w:lang w:val="en-US" w:eastAsia="zh-CN" w:bidi="ar-SA"/><w:position w:val="4"/><w:kern w:val="28"/></w:rPr>');
  });
```

- [x] **Step 2: Write reader round-trip test before implementation**

Add a reader test near the existing style run round-trip tests:

```ts
  it("round-trips style run advanced typography properties", async () => {
    const source = {
      version: "1.0" as const,
      styles: {
        character: [{
          id: "EmphasisText",
          name: "Emphasis Text",
          run: {
            complexScriptFontSize: 14,
            language: { value: "en-US", eastAsia: "zh-CN", bidi: "ar-SA" },
            characterPosition: 4,
            kerning: 28,
          },
        }],
      },
      sections: [{ blocks: [{ type: "paragraph" as const, runs: [{ text: "Term", styleId: "EmphasisText" }] }] }],
    };

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });
```

- [x] **Step 3: Run writer RED test**

Run: `npm test -- tests/docx-core.test.ts -t "writes style run advanced typography"`

Expected: FAIL because style `<w:rPr>` does not emit `<w:szCs>`, `<w:lang>`, `<w:position>`, or `<w:kern>` yet.

- [x] **Step 4: Run reader RED test**

Run: `npm test -- tests/docx-core.test.ts -t "round-trips style run advanced typography"`

Expected: FAIL because style parsing drops `complexScriptFontSize`, `language`, `characterPosition`, and `kerning`.

### Task 2: Implement Style Run Advanced Typography

**Files:**
- Modify: `src/docx-writer.ts`
- Modify: `src/docx-reader.ts`

- [x] **Step 1: Emit style run advanced typography XML**

Extend `styleRunPropertiesXml` after `fontSize` with:

```ts
    run.complexScriptFontSize ? `<w:szCs w:val="${run.complexScriptFontSize * 2}"/>` : "",
```

Extend it near the existing run-positioned properties with:

```ts
    run.language ? runLanguageXml(run.language) : "",
    run.characterPosition !== undefined ? `<w:position w:val="${run.characterPosition}"/>` : "",
    run.kerning !== undefined ? `<w:kern w:val="${run.kerning}"/>` : "",
```

- [x] **Step 2: Parse style run advanced typography XML**

Add local objects to `parseStyleRunProperties`:

```ts
  const complexScriptSize = asObject(properties.szCs);
  const language = parseRunLanguage(properties.lang);
  const characterPosition = asObject(properties.position);
  const kerning = asObject(properties.kern);
```

Extend the parsed object with:

```ts
    ...(typeof complexScriptSize.val === "number" ? { complexScriptFontSize: complexScriptSize.val / 2 } : {}),
    ...(typeof complexScriptSize.val === "string" ? { complexScriptFontSize: Number.parseInt(complexScriptSize.val, 10) / 2 } : {}),
    ...(language ? { language } : {}),
    ...(characterPosition.val !== undefined ? { characterPosition: parseNumber(characterPosition.val) } : {}),
    ...(kerning.val !== undefined ? { kerning: parseNumber(kerning.val) } : {}),
```

- [x] **Step 3: Run targeted GREEN tests**

Run: `npm test -- tests/docx-core.test.ts -t "style run advanced typography"`

Expected: PASS with both targeted tests passing.

### Task 3: Verify, Commit, Push

**Files:**
- Modify: `docs/superpowers/plans/2026-07-04-phase-143-style-run-advanced-typography.md`

- [x] **Step 1: Run full verification**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "style run advanced typography"
npm test
npm run build
git diff --check
```

Expected: Targeted tests, full tests, and build pass. `git diff --check` may print existing LF/CRLF warnings but must exit 0.

- [x] **Step 2: Commit implementation**

```bash
git add src/docx-writer.ts src/docx-reader.ts tests/docx-core.test.ts docs/superpowers/plans/2026-07-04-phase-143-style-run-advanced-typography.md
git commit -m "feat: add phase 143 style run advanced typography"
```

- [x] **Step 3: Push branch**

```bash
git push -u origin phase-143-style-run-advanced-typography
```

- [x] **Step 4: Record push metadata**

Append the pushed commit hash and remote branch to this plan, then commit the plan update:

```bash
git add docs/superpowers/plans/2026-07-04-phase-143-style-run-advanced-typography.md
git commit -m "docs: mark phase 143 pushed"
git push
```

## Push Record

- Branch: `phase-143-style-run-advanced-typography`
- Remote: `origin`
- Repository: `https://github.com/codingayice/word2json.git`
- Implementation commit: `4580013f39e6b2cf861c5005f0ade38eefb5af1f`
- Pull request URL: `https://github.com/codingayice/word2json/pull/new/phase-143-style-run-advanced-typography`
