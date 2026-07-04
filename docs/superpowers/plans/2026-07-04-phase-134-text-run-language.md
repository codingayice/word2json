# Phase 134 Text Run Language Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add high-fidelity JSON round-trip support for Word text run language metadata serialized as `<w:lang>`.

**Architecture:** This phase adds a focused `language` object to `StyleRunProperties` for run-level language metadata only. The writer emits `<w:lang>` with `w:val`, `w:eastAsia`, and `w:bidi` attributes when present, while the reader maps those attributes back into the same JSON shape through `parseRunFont`.

**Tech Stack:** TypeScript, JSZip-based DOCX package tests, existing DOCX reader/writer XML helpers, Vitest.

---

### Task 1: Add RED Tests For Text Run Language

**Files:**
- Modify: `tests/docx-core.test.ts`

- [x] **Step 1: Write writer tests before implementation**

Add two tests near the other text run property writer tests:

```ts
  it("writes text run language value", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "Hello", language: { value: "en-US" } }],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<w:rPr><w:lang w:val="en-US"/></w:rPr><w:t>Hello</w:t>');
  });

  it("writes text run language east asia and bidi", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "Mixed", language: { value: "en-US", eastAsia: "zh-CN", bidi: "ar-SA" } }],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<w:rPr><w:lang w:val="en-US" w:eastAsia="zh-CN" w:bidi="ar-SA"/></w:rPr><w:t>Mixed</w:t>');
  });
```

- [x] **Step 2: Write reader round-trip tests before implementation**

Add two tests near the other text run property round-trip tests:

```ts
  it("round-trips text run language value", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "Hello", language: { value: "en-US" } }],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips text run language east asia and bidi", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "Mixed", language: { value: "en-US", eastAsia: "zh-CN", bidi: "ar-SA" } }],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });
```

- [x] **Step 3: Run writer RED test**

Run: `npm test -- tests/docx-core.test.ts -t "writes text run language"`

Expected: FAIL because `<w:lang>` is not emitted yet.

- [x] **Step 4: Run reader RED test**

Run: `npm test -- tests/docx-core.test.ts -t "round-trips text run language"`

Expected: FAIL because `language` is not preserved through parsing yet.

### Task 2: Implement Text Run Language

**Files:**
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`
- Modify: `src/docx-reader.ts`

- [x] **Step 1: Add schema types**

Add the type and field:

```ts
  language?: RunLanguage;
```

```ts
export type RunLanguage = {
  value?: string;
  eastAsia?: string;
  bidi?: string;
};
```

- [x] **Step 2: Emit writer XML**

Add a helper call in `runPropertiesXml`:

```ts
    run.language ? runLanguageXml(run.language) : "",
```

Add the helper:

```ts
function runLanguageXml(language: RunLanguage): string {
  const attributes = [
    language.value ? ` w:val="${escapeAttribute(language.value)}"` : "",
    language.eastAsia ? ` w:eastAsia="${escapeAttribute(language.eastAsia)}"` : "",
    language.bidi ? ` w:bidi="${escapeAttribute(language.bidi)}"` : "",
  ].join("");

  return attributes ? `<w:lang${attributes}/>` : "";
}
```

- [x] **Step 3: Parse reader XML**

Add parsing to `parseRunFont`:

```ts
  const language = parseRunLanguage(properties.lang);
```

```ts
    ...(language ? { language } : {}),
```

Add the helper:

```ts
function parseRunLanguage(value: unknown): RunLanguage | undefined {
  const language = asObject(value);
  const parsed = {
    ...(typeof language.val === "string" ? { value: language.val } : {}),
    ...(typeof language.eastAsia === "string" ? { eastAsia: language.eastAsia } : {}),
    ...(typeof language.bidi === "string" ? { bidi: language.bidi } : {}),
  };

  return Object.keys(parsed).length > 0 ? parsed : undefined;
}
```

- [x] **Step 4: Run targeted GREEN tests**

Run: `npm test -- tests/docx-core.test.ts -t "text run language"`

Expected: PASS with all four targeted tests passing.

### Task 3: Verify, Commit, Push

**Files:**
- Modify: `docs/superpowers/plans/2026-07-04-phase-134-text-run-language.md`

- [x] **Step 1: Run full verification**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "text run language"
npm test
npm run build
git diff --check
```

Expected: Targeted tests, full tests, and build pass. `git diff --check` may print existing LF/CRLF warnings but must exit 0.

- [x] **Step 2: Commit implementation**

```bash
git add src/schema.ts src/docx-writer.ts src/docx-reader.ts tests/docx-core.test.ts docs/superpowers/plans/2026-07-04-phase-134-text-run-language.md
git commit -m "feat: add phase 134 text run language"
```

- [x] **Step 3: Push branch**

```bash
git push -u origin phase-134-text-run-language
```

- [x] **Step 4: Record push metadata**

Append the pushed commit hash and remote branch to this plan, then commit the plan update:

```bash
git add docs/superpowers/plans/2026-07-04-phase-134-text-run-language.md
git commit -m "docs: mark phase 134 pushed"
git push
```

## Push Record

- Branch: `phase-134-text-run-language`
- Remote: `origin/phase-134-text-run-language`
- Implementation commit: `21ab8e1 feat: add phase 134 text run language`
- Pushed to: `https://github.com/codingayice/word2json.git`
