# Phase 145 Style Run Fit Text Emphasis Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Preserve style run fit-text and emphasis metadata in Word style definitions.

**Architecture:** Inline text runs already serialize and parse `fitText` and `emphasis` through `<w:fitText>` and `<w:em>`. This phase extends the style run writer and reader to use the same JSON shape inside style `<w:rPr>`.

**Tech Stack:** TypeScript, JSZip-based DOCX package tests, existing DOCX reader/writer XML helpers, Vitest.

---

### Task 1: Add RED Tests For Style Run Fit Text And Emphasis

**Files:**
- Modify: `tests/docx-core.test.ts`

- [x] **Step 1: Write writer test before implementation**

Add a writer test near the existing style run tests:

```ts
  it("writes style run fit text and emphasis properties", async () => {
    const document = {
      version: "1.0" as const,
      styles: {
        character: [{
          id: "CompressedText",
          name: "Compressed Text",
          run: {
            fitText: { width: 1440, id: 11 },
            emphasis: "underDot" as const,
          },
        }],
      },
      sections: [{ blocks: [{ type: "paragraph" as const, runs: [{ text: "Fit", styleId: "CompressedText" }] }] }],
    };

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const styles = await zip.file("word/styles.xml")!.async("string");

    expect(styles).toContain('<w:rPr><w:fitText w:val="1440" w:id="11"/><w:em w:val="underDot"/></w:rPr>');
  });
```

- [x] **Step 2: Write reader round-trip test before implementation**

Add a reader test near the existing style run round-trip tests:

```ts
  it("round-trips style run fit text and emphasis properties", async () => {
    const source = {
      version: "1.0" as const,
      styles: {
        character: [{
          id: "CompressedText",
          name: "Compressed Text",
          run: {
            fitText: { width: 1440, id: 11 },
            emphasis: "underDot" as const,
          },
        }],
      },
      sections: [{ blocks: [{ type: "paragraph" as const, runs: [{ text: "Fit", styleId: "CompressedText" }] }] }],
    };

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });
```

- [x] **Step 3: Run writer RED test**

Run: `npm test -- tests/docx-core.test.ts -t "writes style run fit text and emphasis"`

Expected: FAIL because style `<w:rPr>` does not emit `<w:fitText>` or `<w:em>` yet.

- [x] **Step 4: Run reader RED test**

Run: `npm test -- tests/docx-core.test.ts -t "round-trips style run fit text and emphasis"`

Expected: FAIL because style parsing drops `fitText` and `emphasis`.

### Task 2: Implement Style Run Fit Text And Emphasis

**Files:**
- Modify: `src/docx-writer.ts`
- Modify: `src/docx-reader.ts`

- [x] **Step 1: Emit style run fitText and emphasis XML**

Extend `styleRunPropertiesXml` after `scale` with:

```ts
    run.fitText ? `<w:fitText w:val="${run.fitText.width}"${run.fitText.id !== undefined ? ` w:id="${run.fitText.id}"` : ""}/>` : "",
    run.emphasis ? `<w:em w:val="${run.emphasis}"/>` : "",
```

- [x] **Step 2: Parse style run fitText and emphasis XML**

Add local objects to `parseStyleRunProperties`:

```ts
  const fitText = asObject(properties.fitText);
  const emphasis = asObject(properties.em);
```

Extend the parsed object with:

```ts
    ...(fitText.val !== undefined
      ? { fitText: { width: parseNumber(fitText.val), ...(fitText.id !== undefined ? { id: parseNumber(fitText.id) } : {}) } }
      : {}),
    ...(typeof emphasis.val === "string" ? { emphasis: emphasis.val as NonNullable<StyleRunProperties["emphasis"]> } : {}),
```

- [x] **Step 3: Run targeted GREEN tests**

Run: `npm test -- tests/docx-core.test.ts -t "style run fit text and emphasis"`

Expected: PASS with both targeted tests passing.

### Task 3: Verify, Commit, Push

**Files:**
- Modify: `docs/superpowers/plans/2026-07-04-phase-145-style-run-fit-text-emphasis.md`

- [x] **Step 1: Run full verification**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "style run fit text and emphasis"
npm test
npm run build
git diff --check
```

Expected: Targeted tests, full tests, and build pass. `git diff --check` may print existing LF/CRLF warnings but must exit 0.

- [x] **Step 2: Commit implementation**

```bash
git add src/docx-writer.ts src/docx-reader.ts tests/docx-core.test.ts docs/superpowers/plans/2026-07-04-phase-145-style-run-fit-text-emphasis.md
git commit -m "feat: add phase 145 style run fit text emphasis"
```

- [x] **Step 3: Push branch**

```bash
git push -u origin phase-145-style-run-fit-text-emphasis
```

- [x] **Step 4: Record push metadata**

Append the pushed commit hash and remote branch to this plan, then commit the plan update:

```bash
git add docs/superpowers/plans/2026-07-04-phase-145-style-run-fit-text-emphasis.md
git commit -m "docs: mark phase 145 pushed"
git push
```

## Push Record

- Branch: `phase-145-style-run-fit-text-emphasis`
- Remote: `origin`
- Repository: `https://github.com/codingayice/word2json.git`
- Implementation commit: `433cf0b6c2941a822640c3cc6573968eec60c5b8`
- Pull request URL: `https://github.com/codingayice/word2json/pull/new/phase-145-style-run-fit-text-emphasis`
